#!/usr/bin/env python3
"""
Checagem mensal automática de Notas Fiscais duplicadas na ZLFIB — Fitted Units
(sub-projeto "Fitted Recuperação").

Pedido da Juliana em 2026-08-13: rodar a análise do mês anterior assim que
ela logar no SAP no primeiro dia útil de cada mês, e mandar um e-mail pra
ela mesma (juliana.silveira@pirelli.com) só se encontrar duplicidade real
(com o anexo das notas duplicadas). Se não houver duplicidade, não notifica.

Como não existe um jeito confiável de "escutar" o login do SAP de fora, este
script é pensado pra rodar em polling: o Agendador de Tarefas do Windows
chama `watcher()` de 1 em 1 hora (ver scripts/sap/watcher_mensal_zlfib.bat).
A cada chamada:
  1. Se hoje não for o primeiro dia útil do mês, não faz nada (saída rápida).
  2. Se a checagem deste mês já rodou (estado salvo em
     data/processed/zlfib_mensal_estado.json — fora do Git, é estado local),
     não faz nada.
  3. Se o SAP não estiver aberto/logado ainda, não faz nada — tenta de novo
     na próxima hora. Se chegar um horário limite (18h) sem conseguir, manda
     um e-mail de aviso (sem anexo) uma única vez, pra lembrar de rodar
     manualmente.
  4. Assim que achar o SAP logado, roda a análise do mês anterior (todas as
     4 filiais, mesma regra de 2026-08-13: Direção=Entrada, exclui Tipo NF
     R8/transferência de material) e marca o mês como concluído.

Limitação conhecida: "primeiro dia útil" aqui considera só segunda-sexta,
sem calendário de feriados nacionais/municipais. Se cair num feriado, roda
mesmo assim nesse dia (não pula pro próximo dia útil real).
"""
import calendar
import datetime as dt
import json
import sys
from pathlib import Path

import win32com.client as win32

sys.path.insert(0, str(Path(__file__).resolve().parent))
from analisar_zlfib_duplicidade import FILIAIS, analisar  # noqa: E402

RAIZ_PROJETO = Path(__file__).resolve().parent.parent.parent
ARQUIVO_ESTADO = RAIZ_PROJETO / "data" / "processed" / "zlfib_mensal_estado.json"
EMAIL_DESTINATARIO = "juliana.silveira@pirelli.com"
HORA_LIMITE_AVISO = 18  # se passar dessa hora no 1o dia util sem achar o SAP logado, avisa por e-mail


def log(msg):
    carimbo = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{carimbo}] {msg}")


def carregar_estado() -> dict:
    if not ARQUIVO_ESTADO.exists():
        return {}
    try:
        return json.loads(ARQUIVO_ESTADO.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return {}


def salvar_estado(estado: dict):
    ARQUIVO_ESTADO.parent.mkdir(parents=True, exist_ok=True)
    ARQUIVO_ESTADO.write_text(json.dumps(estado, ensure_ascii=False, indent=2), encoding="utf-8")


def primeiro_dia_util_do_mes(ano: int, mes: int) -> dt.date:
    dia = 1
    while dt.date(ano, mes, dia).weekday() >= 5:  # 5=sabado, 6=domingo
        dia += 1
    return dt.date(ano, mes, dia)


def mes_anterior_range(hoje: dt.date):
    primeiro_dia_mes_atual = hoje.replace(day=1)
    ultimo_dia_mes_anterior = primeiro_dia_mes_atual - dt.timedelta(days=1)
    ano, mes = ultimo_dia_mes_anterior.year, ultimo_dia_mes_anterior.month
    ultimo_dia = calendar.monthrange(ano, mes)[1]
    data_de = f"01.{mes:02d}.{ano}"
    data_ate = f"{ultimo_dia:02d}.{mes:02d}.{ano}"
    nomes_mes = ["", "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho",
                 "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
    return data_de, data_ate, nomes_mes[mes], ano


def sap_esta_logado():
    """Tenta conectar numa sessão do SAP GUI já aberta e checa se está
    logada (Info.User preenchido). Devolve a sessão se sim, None se não."""
    try:
        import win32com.client as win32com_client
        sap_gui_auto = win32com_client.GetObject("SAPGUI")
        application = sap_gui_auto.GetScriptingEngine
        connection = application.Children(0)
        session = connection.Children(0)
        if session.Info.User:
            return session
    except Exception:
        pass
    return None


def enviar_email(assunto: str, corpo: str, anexo: Path = None):
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)  # olMailItem
    mail.To = EMAIL_DESTINATARIO
    mail.Subject = assunto
    mail.Body = corpo
    if anexo is not None:
        mail.Attachments.Add(str(anexo))
    mail.Send()


def rodar_verificacao_mensal(hoje: dt.date, log=log):
    data_de, data_ate, nome_mes, ano = mes_anterior_range(hoje)
    log(f"Rodando checagem de duplicidade ZLFIB do mês anterior: {nome_mes}/{ano} ({data_de} a {data_ate}).")

    resultado = analisar(filiais=FILIAIS, data_de=data_de, data_ate=data_ate, log=log)

    if resultado["tem_duplicidade"]:
        log(f"Duplicidade encontrada: {resultado['grupos_dup_chave'] + resultado['grupos_dup_sem_chave']} "
            f"grupo(s), {resultado['notas_envolvidas']} nota(s), R$ {resultado['valor_total']:.2f}. Enviando e-mail...")
        corpo = (
            f"Oi Juliana,\n\n"
            f"A checagem mensal automática da ZLFIB (Fitted Units) encontrou notas fiscais "
            f"possivelmente duplicadas no período de {nome_mes}/{ano}.\n\n"
            f"Grupos por Chave de Acesso: {resultado['grupos_dup_chave']}\n"
            f"Grupos por Parceiro+NF+Série+Valor: {resultado['grupos_dup_sem_chave']}\n"
            f"Notas envolvidas: {resultado['notas_envolvidas']}\n"
            f"Valor total envolvido: R$ {resultado['valor_total']:.2f}\n\n"
            f"Detalhe completo em anexo. Isso é uma triagem automática — vale revisar manualmente "
            f"antes de confirmar como erro real.\n\n"
            f"(e-mail automático — checagem mensal ZLFIB)"
        )
        enviar_email(
            f"Duplicidade de NF (ZLFIB) — Fitted Units — {nome_mes}/{ano}",
            corpo,
            anexo=resultado["arquivo"],
        )
        log("E-mail enviado.")
    else:
        log("Nenhuma duplicidade encontrada este mês — sem notificação, conforme combinado.")

    return resultado


def watcher():
    hoje = dt.date.today()
    agora = dt.datetime.now()
    estado = carregar_estado()
    chave_mes = hoje.strftime("%Y-%m")

    if hoje != primeiro_dia_util_do_mes(hoje.year, hoje.month):
        return  # nao e' o dia certo, nada a fazer

    if estado.get("ultimo_mes_verificado") == chave_mes:
        return  # ja rodou com sucesso este mes

    session = sap_esta_logado()
    if session is None:
        log(f"SAP ainda não está logado (checagem de {agora.strftime('%H:%M')}).")
        if agora.hour >= HORA_LIMITE_AVISO and estado.get("aviso_indisponibilidade_mes") != chave_mes:
            log(f"Passou das {HORA_LIMITE_AVISO}h sem o SAP logado — enviando aviso por e-mail.")
            enviar_email(
                "Checagem mensal ZLFIB não rodou — SAP não estava logado",
                "Oi Juliana,\n\n"
                "Hoje é o primeiro dia útil do mês e a checagem automática de duplicidade da ZLFIB "
                "(Fitted Units) não conseguiu rodar porque não encontrou o SAP aberto/logado até agora.\n\n"
                "Quando puder, abra o SAP e a checagem roda sozinha na próxima hora — ou peça pro Claude "
                "rodar manualmente.\n\n"
                "(e-mail automático — checagem mensal ZLFIB)",
            )
            estado["aviso_indisponibilidade_mes"] = chave_mes
            salvar_estado(estado)
        return

    log("SAP logado — rodando a checagem mensal.")
    try:
        rodar_verificacao_mensal(hoje, log=log)
        estado["ultimo_mes_verificado"] = chave_mes
        salvar_estado(estado)
    except Exception as e:
        log(f"ERRO ao rodar a checagem mensal: {e}")
        raise


if __name__ == "__main__":
    watcher()
