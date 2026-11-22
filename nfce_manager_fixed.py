# -*- coding: utf-8 -*-
"""
Coletor de Notas Fiscais (NFC-e)
================================
Lista os XMLs de NFC-e de uma pasta, por mês/ano, e permite:
  - exportar um ZIP na Área de Trabalho,
  - gerar um relatório em PDF,
  - enviar os XMLs por e-mail (ZIPs divididos em partes, se necessário).

NOVIDADE (versão 3.5): as notas AUTORIZADAS e as CANCELADAS são separadas.
  - A busca também reconhece as canceladas (evento de cancelamento, cStat 135/155).
  - O ZIP exportado traz as pastas "Autorizadas" e "Canceladas".
  - O e-mail envia ZIPs separados: um conjunto só de autorizadas e outro só de canceladas.

Este arquivo foi reconstruído a partir da estrutura do programa original.
"""
import tkinter as tk
from tkinter import messagebox, filedialog, ttk
from datetime import datetime
from pathlib import Path
import os
import locale
import zipfile
import smtplib
import ssl
import json
import tempfile
import threading
import time
import random
import sys
import shutil
import unicodedata
import traceback
import socket
import math
import webbrowser
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from email.utils import getaddresses
import xml.etree.ElementTree as ET

try:
    import winreg
except ImportError:  # fora do Windows (apenas para testes)
    winreg = None

try:
    from PIL import Image, ImageDraw, ImageFont
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

try:
    from urllib import request as urllib_request
    import urllib.request  # noqa: F401
except ImportError:  # pragma: no cover
    urllib_request = None

try:
    from packaging import version as _versao
    PACKAGING_AVAILABLE = True
except ImportError:
    PACKAGING_AVAILABLE = False

# ----------------------------------------------------------------------------
# Constantes
# ----------------------------------------------------------------------------
LIMITE_EMAIL_BYTES = 18 * 1024 * 1024      # tamanho máximo (sem compactar) de cada ZIP enviado
JITTER_MAX_SEG = 4                          # variação aleatória na pausa entre e-mails
PAUSA_ENTRE_PARTES_SEG = 8                  # pausa mínima entre e-mails (evita limite do Gmail)

SMTP_REMETENTE_FIXO = "fiscais.arquivos@gmail.com"
# IMPORTANTE: a senha de aplicativo NÃO fica mais escrita no código.
# Informe-a em  Configurações > Conta de envio  (ela é salva no computador do usuário).
# Se quiser embutir de novo, coloque entre aspas aqui — mas isso expõe a senha a quem tiver o instalador.
SMTP_SENHA_FIXA = ""
SMTP_SERVIDOR_FIXO = "smtp.gmail.com"
SMTP_PORTA_FIXA = 587
SMTP_USAR_SSL_FIXO = True

APP_VERSION = "3.5"
GITHUB_REPO = "luuiz224/coletor-notas-fiscais"
APP_MUTEX_NAME = "ColetorNotasFiscaisMutex"
_mutex_handle = None

PASTA_PADRAO_XML = r"C:\CUSTODIA\NFCE\XML"

MESES = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
         "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]

# Tipos de nota
AUTORIZADA = "autorizada"
CANCELADA = "cancelada"
OUTRO = "outro"
ROTULO_TIPO = {AUTORIZADA: "Autorizadas", CANCELADA: "Canceladas", OUTRO: "Outros"}
ROTULO_TIPO_UNICO = {AUTORIZADA: "Autorizada", CANCELADA: "Cancelada", OUTRO: "Outro"}

# True  -> o XML da nota que foi cancelada vai para "Canceladas", junto do evento.
# False -> o XML da nota fica em "Autorizadas" e só o evento de cancelamento vai para "Canceladas".
NOTA_CANCELADA_VAI_PARA_CANCELADAS = True

CSTAT_AUTORIZADA = {"100", "150"}
CSTAT_CANCELADA_PROTOCOLO = {"101", "151"}
CSTAT_EVENTO_OK = {"135", "136", "155"}
TP_EVENTO_CANCELAMENTO = "110111"

try:
    import customtkinter as ctk
    CTK_AVAILABLE = True
except ImportError:
    ctk = None
    CTK_AVAILABLE = False


# ----------------------------------------------------------------------------
# Funções de apoio
# ----------------------------------------------------------------------------
def resource_path(rel):
    """Caminho de um arquivo de recurso (imagem etc.), funcionando tanto no
    modo desenvolvimento quanto empacotado (PyInstaller)."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


def contexto_ssl():
    """Contexto SSL robusto para app EMPACOTADO (usa certifi se disponível).
    Sem isso, o HTTPS pode falhar por falta de certificados raiz."""
    try:
        import certifi
        return ssl.create_default_context(cafile=certifi.where())
    except Exception:
        return ssl.create_default_context()


def criar_mutex():
    """Cria um mutex nomeado para o instalador conseguir detectar o app aberto."""
    global _mutex_handle
    try:
        import ctypes
        _mutex_handle = ctypes.windll.kernel32.CreateMutexW(None, False, APP_MUTEX_NAME)
    except Exception:
        _mutex_handle = None


def definir_app_id():
    """Define uma identidade própria no Windows (AppUserModelID). Sem isso, a
    BARRA DE TAREFAS herda o ícone do Python em vez do ícone do app."""
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("Automacao2000.ColetorNotasFiscais")
    except Exception:
        pass


def parsear_destinatarios(raw):
    """Interpreta o campo de e-mail destinatário e devolve uma lista de endereços
    válidos (com '@' e domínio). Aceita separação por vírgula ou ponto e vírgula."""
    if not raw:
        return []
    texto = str(raw).replace(";", ",").replace("\n", ",")
    enderecos = []
    for _, addr in getaddresses([texto]):
        addr = addr.strip()
        if not addr or "@" not in addr:
            continue
        _, _, dominio = addr.partition("@")
        if "." not in dominio or dominio.startswith(".") or dominio.endswith("."):
            continue
        if addr not in enderecos:
            enderecos.append(addr)
    return enderecos


def classificar_erro(exc):
    """Traduz qualquer exceção para (mensagem_amigavel, codigo_curto).

    A mensagem é curta e sem jargão técnico; o código serve para o suporte
    identificar o problema rapidamente."""
    if isinstance(exc, smtplib.SMTPAuthenticationError):
        return ("Não foi possível entrar na conta de e-mail. Confira o e-mail e a "
                "senha de aplicativo em Configurações.", "EMAIL-AUTH")
    if isinstance(exc, smtplib.SMTPRecipientsRefused):
        return ("O servidor recusou o e-mail destinatário. Confira o endereço em Configurações.",
                "EMAIL-DEST")
    if isinstance(exc, smtplib.SMTPSenderRefused):
        return ("O servidor recusou o e-mail remetente. Confira a conta em Configurações.",
                "EMAIL-REM")
    if isinstance(exc, smtplib.SMTPResponseException):
        codigo = getattr(exc, "smtp_code", 0) or 0
        texto = getattr(exc, "smtp_error", b"")
        if isinstance(texto, bytes):
            texto = texto.decode("utf-8", "ignore")
        texto = str(texto)
        if "5.5.2" in texto:
            return ("O endereço do destinatário parece inválido. Confira em Configurações.",
                    "EMAIL-DEST")
        if codigo in (421, 450, 451, 452, 454) or "5.4.5" in texto or "4.7.0" in texto:
            return ("O Gmail está limitando os envios no momento (muitos e-mails em pouco tempo). "
                    "Aguarde alguns minutos e tente de novo.", f"EMAIL-LIMITE-{codigo}")
        return (f"O servidor de e-mail recusou o envio (código {codigo}). "
                "Tente novamente; se continuar, avise o suporte.", f"EMAIL-{codigo}")
    if isinstance(exc, smtplib.SMTPServerDisconnected):
        return ("A conexão com o servidor de e-mail foi interrompida. Tente novamente.", "EMAIL-DESCONECTOU")
    if isinstance(exc, socket.gaierror):
        return ("Sem acesso à internet (ou o servidor de e-mail não foi encontrado). "
                "Verifique a conexão.", "REDE-DNS")
    if isinstance(exc, smtplib.SMTPConnectError):
        return ("Não foi possível conectar ao servidor de e-mail. Verifique a internet.", "EMAIL-CONEXAO")
    if isinstance(exc, (TimeoutError, socket.timeout)):
        return ("A conexão demorou demais. Verifique a internet e tente de novo.", "REDE-TIMEOUT")
    if isinstance(exc, ConnectionError):
        return ("A conexão com a internet foi interrompida. Tente novamente.", "REDE-CONEXAO")
    if isinstance(exc, smtplib.SMTPException):
        return ("Houve um problema no envio do e-mail. Tente novamente.", "EMAIL-GERAL")
    if isinstance(exc, PermissionError):
        return ("O programa não tem permissão para acessar um arquivo ou pasta. "
                "Feche outros programas que o estejam usando.", "ARQ-PERMISSAO")
    if isinstance(exc, FileNotFoundError):
        return ("Um arquivo ou pasta não foi encontrado. Confira os caminhos em Configurações.",
                "ARQ-NAOENCONTRADO")
    if isinstance(exc, zipfile.BadZipFile):
        return ("Não foi possível criar/ler o arquivo ZIP.", "ZIP")
    if isinstance(exc, OSError):
        errno = getattr(exc, "errno", None)
        if errno == 28:
            return ("Não há espaço livre no disco.", "DISCO-CHEIO")
        return ("O Windows devolveu um erro ao acessar um arquivo.", f"SO-{errno}")
    if isinstance(exc, ImportError):
        return ("Um componente do programa não está instalado. Reinstale o programa.", "COMPONENTE")
    if isinstance(exc, ET.ParseError):
        return ("Um dos arquivos XML está corrompido ou incompleto.", "XML")
    if isinstance(exc, (ValueError, TypeError)):
        return ("Um dado informado está em formato inesperado.", "DADO")
    return ("Ocorreu um erro inesperado.", f"ERRO-{type(exc).__name__}")


# ----------------------------------------------------------------------------
# Classificação das notas: AUTORIZADA x CANCELADA
# ----------------------------------------------------------------------------
def _local(tag):
    return tag.rsplit("}", 1)[-1] if isinstance(tag, str) else ""


def _texto(no, caminho):
    el = no.find(caminho)
    return el.text.strip() if el is not None and el.text else None


def classificar_xml(caminho):
    """Devolve (tipo, chave). Nunca levanta exceção: arquivo ilegível vira (OUTRO, None).

    - nota com protocolo 100/150               -> AUTORIZADA
    - evento de cancelamento (110111) aceito   -> CANCELADA
    - cancelamento no formato antigo           -> CANCELADA
    - denegada, rejeitada, sem protocolo, outros eventos, XML inválido -> OUTRO
    """
    try:
        raiz = ET.parse(caminho).getroot()
    except Exception:
        return OUTRO, None

    nome = _local(raiz.tag)

    if nome in ("nfeProc", "NFe"):
        inf = raiz.find(".//{*}infNFe")
        chave = None
        if inf is not None and inf.get("Id"):
            chave = inf.get("Id").replace("NFe", "", 1)
        cstat = _texto(raiz, ".//{*}protNFe/{*}infProt/{*}cStat")
        if cstat in CSTAT_AUTORIZADA:
            return AUTORIZADA, chave
        if cstat in CSTAT_CANCELADA_PROTOCOLO:
            return CANCELADA, chave
        return OUTRO, chave

    if nome in ("procEventoNFe", "evento", "envEvento"):
        ev = raiz.find(".//{*}evento/{*}infEvento")
        if ev is None and nome == "evento":
            ev = raiz.find("{*}infEvento")
        if ev is None:
            return OUTRO, None
        tp = _texto(ev, "{*}tpEvento")
        chave = _texto(ev, "{*}chNFe")
        if tp != TP_EVENTO_CANCELAMENTO:
            return OUTRO, chave
        cstat = _texto(raiz, ".//{*}retEvento/{*}infEvento/{*}cStat")
        if cstat is not None and cstat not in CSTAT_EVENTO_OK:
            return OUTRO, chave
        return CANCELADA, chave

    if nome in ("procCancNFe", "cancNFe", "retCancNFe"):
        return CANCELADA, _texto(raiz, ".//{*}chNFe")

    return OUTRO, None


def aplicar_pareamento(arquivos):
    """Se uma nota autorizada tem evento de cancelamento no mesmo período, o XML
    dela também é marcado como CANCELADA (quando NOTA_CANCELADA_VAI_PARA_CANCELADAS)."""
    if NOTA_CANCELADA_VAI_PARA_CANCELADAS:
        canceladas = {a["chave"] for a in arquivos if a.get("tipo") == CANCELADA and a.get("chave")}
        for a in arquivos:
            if a.get("tipo") == AUTORIZADA and a.get("chave") in canceladas:
                a["tipo"] = CANCELADA
    return arquivos


def resumo_por_tipo(arquivos):
    r = {AUTORIZADA: 0, CANCELADA: 0, OUTRO: 0}
    for a in arquivos:
        r[a.get("tipo", OUTRO)] = r.get(a.get("tipo", OUTRO), 0) + 1
    return r


# ----------------------------------------------------------------------------
# Temas / paleta
# ----------------------------------------------------------------------------
class Themes:
    LIGHT = {"bg": "#E9E9E9", "bg_secondary": "#C2C2C2", "bg_card": "#FAFAFA", "border": "#E5E5E5",
             "text": "#000000", "text_secondary": "#333333", "accent": "#7E34F7",
             "accent_hover": "#6A2CD4", "success": "#2D9F5F", "fg_button": "#FFFFFF"}
    DARK = {"bg": "#0D0D0D", "bg_secondary": "#1A1A1A", "bg_card": "#161616", "border": "#2D2D2D",
            "text": "#FFFFFF", "text_secondary": "#BBBBBB", "accent": "#965CF9",
            "accent_hover": "#B1B1B1", "success": "#2D9F5F", "fg_button": "#FFFFFF"}


class P:
    """Paleta no formato (claro, escuro). Os widgets do CustomTkinter aceitam essa
    tupla diretamente e trocam de cor sozinhos quando o tema muda."""
    BG = ("#F5F6F8", "#0F1115")
    SURFACE = ("#FFFFFF", "#171A21")
    SURFACE_2 = ("#EDEFF3", "#222630")
    INPUT = ("#FFFFFF", "#1C2029")
    BORDER = ("#E2E5EB", "#2B303B")
    TEXT = ("#171A20", "#ECEEF1")
    TEXT_SUB = ("#6A7280", "#98A0AC")
    ACCENT = ("#A4161A", "#E5484D")
    ACCENT_HV = ("#7F1114", "#EF5A5A")
    ACCENT_SOFT = ("#FBEAEA", "#3A1E20")
    SUCCESS = ("#15A24A", "#22C55E")
    SUCCESS_HV = ("#118A3E", "#16A34A")
    ON_ACCENT = "#FFFFFF"


def cor(tup, escuro):
    """Resolve uma tupla (claro, escuro) para uma cor única (uso no ttk/Treeview,
    que não entendem tuplas)."""
    if isinstance(tup, tuple):
        return tup[1] if escuro else tup[0]
    return tup


# ----------------------------------------------------------------------------
# Ícones desenhados com PIL (sem depender de arquivos de imagem)
# ----------------------------------------------------------------------------
_SS = 4  # supersampling para suavizar as bordas


def _nova(px):
    s = px * _SS
    img = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    return s, img, ImageDraw.Draw(img)


def _fim(img, px):
    return img.resize((px, px), Image.LANCZOS)


def desenhar_nota(px, cor_icone):
    """Ícone de nota fiscal / recibo: corpo com topo arredondado, linhas de
    'texto' e base serrilhada."""
    s, img, d = _nova(px)
    x0, x1 = int(s * 0.20), int(s * 0.80)
    y0, y1 = int(s * 0.08), int(s * 0.88)
    d.rounded_rectangle([x0, y0, x1, y1 - int(s * 0.08)], radius=int(s * 0.08), fill=cor_icone)
    dente = (x1 - x0) / 6.0
    pts = [(x0, y1 - int(s * 0.08))]
    for i in range(6):
        pts.append((x0 + dente * (i + 0.5), y1))
        pts.append((x0 + dente * (i + 1), y1 - int(s * 0.08)))
    d.polygon(pts, fill=cor_icone)
    for k in range(4):
        yy = int(s * (0.24 + 0.13 * k))
        d.rounded_rectangle([x0 + int(s * 0.10), yy, x1 - int(s * 0.10) - (int(s * 0.12) if k == 3 else 0),
                             yy + int(s * 0.05)], radius=int(s * 0.02), fill=(0, 0, 0, 0))
    return _fim(img, px)


def desenhar_sol(px, cor_icone):
    s, img, d = _nova(px)
    c = s // 2
    r = int(s * 0.18)
    d.ellipse([c - r, c - r, c + r, c + r], fill=cor_icone)
    for i in range(8):
        a = math.radians(i * 45)
        r1, r2 = int(s * 0.28), int(s * 0.42)
        d.line([c + r1 * math.cos(a), c + r1 * math.sin(a), c + r2 * math.cos(a), c + r2 * math.sin(a)],
               fill=cor_icone, width=max(2, int(s * 0.06)))
    return _fim(img, px)


def desenhar_lua(px, cor_icone):
    s, img, d = _nova(px)
    d.ellipse([int(s * 0.18), int(s * 0.14), int(s * 0.82), int(s * 0.86)], fill=cor_icone)
    d.ellipse([int(s * 0.38), int(s * 0.04), int(s * 1.00), int(s * 0.70)], fill=(0, 0, 0, 0))
    return _fim(img, px)


def _fonte_win(nome, size):
    """Carrega uma fonte do Windows; se faltar, cai na fonte padrão do PIL."""
    try:
        return ImageFont.truetype("C:/Windows/Fonts/" + nome, size)
    except Exception:
        try:
            return ImageFont.load_default()
        except Exception:
            return None


def desenhar_xml(px):
    """Ícone do sistema: documento XML branco com </>, linhas, selo 'XML' e
    chaves { } vermelhas."""
    s, img, d = _nova(px)
    vermelho = "#A4161A"
    d.rounded_rectangle([int(s * 0.16), int(s * 0.06), int(s * 0.84), int(s * 0.94)],
                        radius=int(s * 0.08), fill="#FFFFFF", outline="#D5D8DE", width=max(1, int(s * 0.015)))
    d.polygon([(int(s * 0.62), int(s * 0.06)), (int(s * 0.84), int(s * 0.28)), (int(s * 0.62), int(s * 0.28))],
              fill="#E5E7EB")
    f1 = _fonte_win("consolab.ttf", int(s * 0.17))
    d.text((s // 2, int(s * 0.38)), "</>", fill=vermelho, font=f1, anchor="mm")

    def linha(yf, wf, c):
        d.rounded_rectangle([int(s * 0.26), int(s * yf), int(s * (0.26 + wf)), int(s * (yf + 0.035))],
                            radius=int(s * 0.015), fill=c)

    linha(0.52, 0.48, "#C9CED6")
    linha(0.59, 0.36, "#C9CED6")
    d.rounded_rectangle([int(s * 0.26), int(s * 0.72), int(s * 0.74), int(s * 0.86)],
                        radius=int(s * 0.03), fill=vermelho)
    f2 = _fonte_win("arialbd.ttf", int(s * 0.13))
    d.text((s // 2, int(s * 0.79)), "XML", fill="#FFFFFF", font=f2, anchor="mm")
    return _fim(img, px)


def desenhar_engrenagem(px, cor_icone):
    """Ícone de engrenagem (Configurações): disco central com dentes ao redor
    e um furo no meio."""
    s, img, d = _nova(px)
    c = s / 2
    r_ext, r_int = s * 0.44, s * 0.33
    pts, n = [], 8
    for i in range(n):
        a0 = 2 * math.pi * i / n
        w = math.pi / n * 0.55
        for ang, rr in ((a0 - w, r_int), (a0 - w * 0.6, r_ext), (a0 + w * 0.6, r_ext), (a0 + w, r_int)):
            pts.append((c + rr * math.cos(ang), c + rr * math.sin(ang)))
    d.polygon(pts, fill=cor_icone)
    d.ellipse([c - s * 0.30, c - s * 0.30, c + s * 0.30, c + s * 0.30], fill=cor_icone)
    r = s * 0.13
    d.ellipse([c - r, c - r, c + r, c + r], fill=(0, 0, 0, 0))
    return _fim(img, px)


def desenhar_info(px, cor_icone):
    """Ícone de informação (Sobre): círculo com um 'i' dentro."""
    s, img, d = _nova(px)
    d.ellipse([int(s * 0.08), int(s * 0.08), int(s * 0.92), int(s * 0.92)], fill=cor_icone)
    d.ellipse([int(s * 0.45), int(s * 0.22), int(s * 0.55), int(s * 0.32)], fill=(0, 0, 0, 0))
    d.rounded_rectangle([int(s * 0.45), int(s * 0.40), int(s * 0.55), int(s * 0.76)],
                        radius=int(s * 0.03), fill=(0, 0, 0, 0))
    return _fim(img, px)


# ----------------------------------------------------------------------------
# Janelas auxiliares
# ----------------------------------------------------------------------------
def _centralizar(win, parent, largura, altura):
    """Centraliza `win` sobre `parent` (ou na tela, se não houver pai visível)."""
    try:
        parent.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width() - largura) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - altura) // 2
        win.geometry(f"{largura}x{altura}+{max(x, 0)}+{max(y, 0)}")
    except Exception:
        win.geometry(f"{largura}x{altura}")


class LoadingDialog:
    def __init__(self, parent, theme, title):
        self.dialog = ctk.CTkToplevel(parent)
        self.dialog.title(title)
        self.dialog.geometry("350x120")
        self.dialog.resizable(False, False)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.protocol("WM_DELETE_WINDOW", lambda: None)
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() - 350) // 2
        y = (self.dialog.winfo_screenheight() - 120) // 2
        self.dialog.geometry(f"350x120+{x}+{y}")
        frame = ctk.CTkFrame(self.dialog, fg_color="transparent")
        frame.pack(expand=True, fill=tk.BOTH, padx=20, pady=20)
        self.label = ctk.CTkLabel(frame, text=title, font=ctk.CTkFont(family="Inter", size=14, weight="bold"))
        self.label.pack(pady=(0, 10))
        self.progress = ctk.CTkProgressBar(frame, width=310, mode="indeterminate")
        self.progress.pack(pady=(0, 8))
        self.progress.start()
        self.status_label = ctk.CTkLabel(frame, text="Iniciando...", font=ctk.CTkFont(family="Inter", size=11),
                                         text_color=theme["text_secondary"])
        self.status_label.pack()

    def update_status(self, text):
        self.status_label.configure(text=text)

    def close(self):
        self.progress.stop()
        self.dialog.grab_release()
        self.dialog.destroy()


class DialogoProgresso:
    """Janela MODAL de progresso mostrada durante o envio. Trava a interface (o
    usuário não consegue mexer no app até terminar)."""

    def __init__(self, parent, titulo, icon_path=None, indeterminado=False):
        self.indeterminado = indeterminado
        self.win = ctk.CTkToplevel(parent)
        self.win.title(titulo)
        self.win.resizable(False, False)
        self.win.configure(fg_color=P.SURFACE)
        self.win.transient(parent)
        self.win.protocol("WM_DELETE_WINDOW", lambda: None)
        _centralizar(self.win, parent, 430, 150)

        def _set_icon():
            try:
                if icon_path:
                    self.win.iconbitmap(icon_path)
            except Exception:
                pass
        self.win.after(200, _set_icon)

        frame = ctk.CTkFrame(self.win, fg_color="transparent")
        frame.pack(fill=tk.BOTH, expand=True, padx=22, pady=18)
        ctk.CTkLabel(frame, text=titulo, text_color=P.TEXT,
                     font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold")).pack(anchor="w")
        self.lbl = ctk.CTkLabel(frame, text="Preparando...", text_color=P.TEXT_SUB,
                                font=ctk.CTkFont(family="Segoe UI", size=12))
        self.lbl.pack(anchor="w", pady=(4, 0))
        self.bar = ctk.CTkProgressBar(frame, height=10, corner_radius=5,
                                      progress_color=P.ACCENT, fg_color=P.SURFACE_2)
        self.bar.pack(fill=tk.X, pady=(14, 0))
        if indeterminado:
            self.bar.configure(mode="indeterminate")
            self.bar.start()
        else:
            self.bar.set(0)
        self.win.lift()
        self.win.after(100, self._grab)

    def _grab(self):
        try:
            self.win.grab_set()
        except Exception:
            pass

    def atualizar(self, pct, texto):
        try:
            if not self.indeterminado:
                self.bar.set(max(0.0, min(1.0, pct)))
            self.lbl.configure(text=texto)
        except Exception:
            pass

    def fechar(self):
        try:
            if self.indeterminado:
                self.bar.stop()
            self.win.grab_release()
            self.win.destroy()
        except Exception:
            pass


class DialogoBusca:
    """Janela NÃO-modal mostrada durante a busca de XMLs, com uma barra de
    progresso e três contadores (Encontrados, Processados, Erros)."""

    def __init__(self, parent, pasta, icon_path=None):
        self.cancelado = False
        self.win = ctk.CTkToplevel(parent)
        self.win.title("Buscando XMLs")
        self.win.resizable(False, False)
        self.win.configure(fg_color=P.SURFACE)
        self.win.transient(parent)
        self.win.protocol("WM_DELETE_WINDOW", self._pedir_cancelar)
        _centralizar(self.win, parent, 440, 300)

        def _set_icon():
            try:
                if icon_path:
                    self.win.iconbitmap(icon_path)
            except Exception:
                pass
        self.win.after(200, _set_icon)

        frame = ctk.CTkFrame(self.win, fg_color="transparent")
        frame.pack(fill=tk.BOTH, expand=True, padx=22, pady=18)
        ctk.CTkLabel(frame, text="Buscando arquivos na pasta selecionada...", text_color=P.TEXT,
                     font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold")).pack(anchor="w")
        self.lbl_pasta = ctk.CTkLabel(frame, text=pasta, text_color=P.TEXT_SUB, wraplength=390, justify="left",
                                      font=ctk.CTkFont(family="Segoe UI", size=11))
        self.lbl_pasta.pack(anchor="w", pady=(2, 0))
        self.bar = ctk.CTkProgressBar(frame, height=10, corner_radius=5,
                                      progress_color=P.ACCENT, fg_color=P.SURFACE_2)
        self.bar.pack(fill=tk.X, pady=(14, 0))
        self.bar.set(0)

        caixa = ctk.CTkFrame(frame, fg_color=P.SURFACE_2, corner_radius=10)
        caixa.pack(fill=tk.X, pady=(16, 0))

        def contador(titulo, cor_num=None):
            col = ctk.CTkFrame(caixa, fg_color="transparent")
            col.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=6, pady=10)
            num = ctk.CTkLabel(col, text="0", text_color=cor_num or P.TEXT,
                               font=ctk.CTkFont(family="Segoe UI", size=20, weight="bold"))
            num.pack()
            ctk.CTkLabel(col, text=titulo, text_color=P.TEXT_SUB,
                         font=ctk.CTkFont(family="Segoe UI", size=11)).pack()
            return num

        self.num_encontrados = contador("Encontrados")
        self.num_processados = contador("Processados")
        self.num_erros = contador("Erros", P.ACCENT)

        self.btn_cancelar = ctk.CTkButton(frame, text="Cancelar", height=34, width=110, corner_radius=8,
                                          command=self._pedir_cancelar, fg_color=P.SURFACE_2,
                                          hover_color=P.BORDER, text_color=P.TEXT,
                                          font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"))
        self.btn_cancelar.pack(anchor="e", pady=(16, 0))

    def _pedir_cancelar(self):
        self.cancelado = True
        try:
            self.btn_cancelar.configure(text="Cancelando...", state="disabled")
        except Exception:
            pass

    def atualizar(self, encontrados, processados, erros, pct):
        try:
            self.num_encontrados.configure(text=str(encontrados))
            self.num_processados.configure(text=str(processados))
            self.num_erros.configure(text=str(erros))
            self.bar.set(max(0.0, min(1.0, pct)))
        except Exception:
            pass

    def fechar(self):
        try:
            self.win.destroy()
        except Exception:
            pass


# ----------------------------------------------------------------------------
# Janela principal
# ----------------------------------------------------------------------------
class ListadorNFCe:
    def __init__(self, root):
        self.root = root
        self.root.title("Coletor de Notas Fiscais")

        frozen = getattr(sys, "frozen", False)
        self.app_dir = os.path.dirname(sys.executable) if frozen else os.path.dirname(os.path.abspath(__file__))

        # ícone da janela
        self.icon_path = resource_path("icon.ico")
        if not os.path.exists(self.icon_path):
            self.icon_path = os.path.join(self.app_dir, "icon.ico")
        try:
            if os.path.exists(self.icon_path):
                self.root.iconbitmap(self.icon_path)
        except Exception as e:
            print("Icon warning: ", e)

        # janela centralizada
        w, h = 1040, 660
        x = max((self.root.winfo_screenwidth() - w) // 2, 0)
        y = max((self.root.winfo_screenheight() - h) // 2, 0)
        self.root.geometry(f"{w}x{h}+{x}+{y}")
        self.root.minsize(940, 580)

        # pastas de configuração
        self.config_dir = os.path.join(os.environ.get("APPDATA", self.app_dir), "NFCe Manager")
        os.makedirs(self.config_dir, exist_ok=True)
        self.config_file = os.path.join(self.config_dir, "config_email.json")
        self.config_files_legados = [os.path.join(self.app_dir, "config_email.json")]

        self.arquivos_encontrados = []
        self._destinatarios_envio = []
        self.var_salvar_senha = None
        self.carregar_config()

        self.tema_escuro = bool(self.config.get("tema_escuro", False))
        self.theme = Themes.DARK if self.tema_escuro else Themes.LIGHT
        ctk.set_appearance_mode("dark" if self.tema_escuro else "light")
        ctk.set_default_color_theme("blue")

        self.criar_interface()
        self.root.protocol("WM_DELETE_WINDOW", self.ao_fechar)
        criar_mutex()
        # verifica atualização em segundo plano, logo após abrir
        self.root.after(3000, lambda: self.verificar_atualizacao(silencioso=True))

    # ------------------------------------------------------------------ config
    def carregar_config(self):
        self.config = {
            "email_remetente": SMTP_REMETENTE_FIXO,
            "senha_email": SMTP_SENHA_FIXA,
            "servidor_smtp": SMTP_SERVIDOR_FIXO,
            "porta_smtp": SMTP_PORTA_FIXA,
            "email_destinatario": "",
            "usar_ssl": SMTP_USAR_SSL_FIXO,
            "tema_escuro": False,
            "pasta_xml": PASTA_PADRAO_XML,
            # novos na versão 3.5
            "pasta_xml_cancelados": "",     # opcional: outra pasta onde ficam as canceladas
            "buscar_subpastas": True,       # procura também dentro das subpastas
        }
        try:
            for caminho in [self.config_file] + list(self.config_files_legados):
                if os.path.exists(caminho):
                    with open(caminho, "r", encoding="utf-8") as f:
                        salvo = json.load(f)
                    for k, v in salvo.items():
                        if k in self.config and v is not None and not (k == "senha_email" and v == ""):
                            self.config[k] = v
                    break
        except Exception as e:
            print("Erro ao carregar configurações: ", e)
        if not self.config.get("pasta_xml"):
            self.config["pasta_xml"] = PASTA_PADRAO_XML

    def salvar_config(self):
        try:
            cfg = dict(self.config)
            var = getattr(self, "var_salvar_senha", None)
            if var is not None and not var.get():
                cfg["senha_email"] = ""
            with open(self.config_file, "w", encoding="utf-8") as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.mostrar_erro("CONFIG-SALVAR", e, "Erro ao salvar configurações")

    # ------------------------------------------------------------------ erros / logs
    def mostrar_erro(self, contexto, exc, titulo="Erro"):
        msg, cod = classificar_erro(exc)
        codigo_completo = f"{contexto}/{cod}"
        self._registrar_log_erro(codigo_completo, exc)

        def _show():
            try:
                messagebox.showerror(
                    titulo,
                    msg + "\n\nCódigo do erro: " + codigo_completo +
                    "\nSe precisar de ajuda, mande este código (ou um print desta tela) para o suporte.")
            except Exception:
                pass
        try:
            self.root.after(0, _show)
        except Exception:
            _show()

    def _registrar_log_erro(self, codigo_completo, exc):
        """Grava o traceback completo no log de erros (silencioso se falhar)."""
        try:
            pasta = getattr(self, "config_dir", None) or getattr(self, "app_dir", ".")
            with open(os.path.join(pasta, "erros.log"), "a", encoding="utf-8") as f:
                f.write(datetime.now().strftime("%d/%m/%Y %H:%M:%S") + "  [" + codigo_completo + "]\n")
                f.write(type(exc).__name__ + ": " + str(exc) + "\n")
                f.write("".join(traceback.format_exception(type(exc), exc, exc.__traceback__)))
                f.write("=" * 70 + "\n")
        except Exception:
            pass

    def _prog(self, pct, texto):
        """Atualiza a barra de progresso do envio (seguro chamar de qualquer thread)."""
        dlg = getattr(self, "_dlg", None)
        if dlg is not None:
            self.root.after(0, lambda: dlg.atualizar(pct, texto))

    def _fechar_dlg(self):
        dlg = getattr(self, "_dlg", None)
        if dlg is not None:
            dlg.fechar()
        self._dlg = None

    def _registrar_perf(self, linha):
        """Grava uma linha de desempenho em %APPDATA%\\NFCe Manager\\perf.log.
        Serve para provar onde o tempo é gasto em cada envio."""
        try:
            pasta = getattr(self, "config_dir", None) or getattr(self, "app_dir", ".")
            with open(os.path.join(pasta, "perf.log"), "a", encoding="utf-8") as f:
                f.write(datetime.now().strftime("%d/%m/%Y %H:%M:%S") + "  " + linha + "\n")
        except Exception:
            pass

    # ------------------------------------------------------------------ atualização
    def _consultar_versao_remota(self):
        """Consulta a 'latest release' no GitHub. Devolve (versao, url_exe, notas)
        ou None se não houver release / não for possível comparar."""
        if not PACKAGING_AVAILABLE or "SEU_USUARIO" in GITHUB_REPO:
            return None
        req = urllib_request.Request(
            "https://api.github.com/repos/" + GITHUB_REPO + "/releases/latest",
            headers={"User-Agent": "ColetorNotasFiscais", "Accept": "application/vnd.github+json"})
        with urllib_request.urlopen(req, timeout=10, context=contexto_ssl()) as resp:
            dados = json.loads(resp.read().decode("utf-8"))
        versao = str(dados.get("tag_name", "")).lstrip("vV").strip()
        notas = dados.get("body") or ""
        url_exe = None
        for asset in dados.get("assets", []):
            if str(asset.get("name", "")).lower().endswith(".exe"):
                url_exe = asset.get("browser_download_url")
                break
        if not versao or not url_exe:
            return None
        return versao, url_exe, notas

    def verificar_atualizacao(self, silencioso=False):
        """Checa se há versão nova. silencioso=True (na abertura) não incomoda o
        usuário quando já está atualizado ou quando não há internet."""
        def worker():
            try:
                info = self._consultar_versao_remota()
            except Exception as e:
                self._registrar_log_erro("ATUALIZACAO/CHECK", e)
                if not silencioso:
                    detalhe = "\n\nDetalhe técnico: " + type(e).__name__ + ": " + str(e)
                    self.root.after(0, lambda: messagebox.showinfo(
                        "Atualização",
                        "Não foi possível verificar atualizações agora.\n"
                        "Verifique a internet e tente mais tarde." + detalhe))
                return
            try:
                nova = bool(info) and _versao.parse(info[0]) > _versao.parse(APP_VERSION)
            except Exception:
                nova = False
            if not nova:
                if not silencioso:
                    self.root.after(0, lambda: messagebox.showinfo(
                        "Atualização", "Você já está na versão mais recente (" + APP_VERSION + ")."))
                return
            versao, url, notas = info

            def perguntar():
                if messagebox.askyesno(
                        "Atualização disponível",
                        "Há uma nova versão disponível: " + versao + "\n(você está na " + APP_VERSION + ").\n\n"
                        + ("Novidades:\n" + notas.strip() + "\n\n" if notas.strip() else "")
                        + "Deseja baixar e instalar agora?"):
                    self._baixar_e_instalar(url)
            self.root.after(0, perguntar)
        threading.Thread(target=worker, daemon=True).start()

    def _baixar_e_instalar(self, url):
        """Baixa o instalador (com barra de progresso), executa e fecha o app
        para que a instalação possa substituir os arquivos."""
        if not str(url).startswith("https://"):
            self.mostrar_erro("ATUALIZACAO", ValueError("endereço de download inválido"),
                              "Erro ao baixar atualização")
            return
        self._dlg = DialogoProgresso(self.root, "Baixando atualização", getattr(self, "icon_path", None))
        self._prog(0.0, "Iniciando download...")

        def worker():
            try:
                destino = os.path.join(tempfile.gettempdir(), "Coletor_de_Notas_Fiscais_Setup.exe")
                req = urllib_request.Request(url, headers={"User-Agent": "ColetorNotasFiscais"})
                with urllib_request.urlopen(req, timeout=30, context=contexto_ssl()) as resp:
                    total = int(resp.headers.get("Content-Length") or 0)
                    baixado = 0
                    with open(destino, "wb") as f:
                        while True:
                            bloco = resp.read(64 * 1024)
                            if not bloco:
                                break
                            f.write(bloco)
                            baixado += len(bloco)
                            if total:
                                self._prog(baixado / total, "Baixando atualização... " + f"{baixado * 100 // total}%")
                            else:
                                self._prog(0.5, "Baixando... " + f"{baixado // 1024} KB")
                self._prog(1.0, "Download concluído. Abrindo instalador...")

                def abrir_e_fechar():
                    self._fechar_dlg()
                    try:
                        os.startfile(destino)
                    except Exception as e:
                        self.mostrar_erro("ATUALIZACAO", e, "Erro ao abrir o instalador")
                        return
                    self.root.after(500, self.root.destroy)
                self.root.after(600, abrir_e_fechar)
            except Exception as e:
                self.root.after(0, self._fechar_dlg)
                self.mostrar_erro("ATUALIZACAO", e, "Erro ao baixar atualização")
        threading.Thread(target=worker, daemon=True).start()

    # ------------------------------------------------------------------ suporte
    def abrir_suporte(self):
        """Janela de Ajuda/Suporte com o contato e botão para abrir o WhatsApp."""
        win = ctk.CTkToplevel(self.root)
        win.title("Ajuda / Suporte")
        win.resizable(False, False)
        win.configure(fg_color=P.SURFACE)
        win.transient(self.root)
        _centralizar(win, self.root, 400, 390)

        def _icone():
            try:
                win.iconbitmap(self.icon_path)
            except Exception:
                pass
        win.after(200, _icone)

        frame = ctk.CTkFrame(win, fg_color="transparent")
        frame.pack(fill=tk.BOTH, expand=True, padx=24, pady=20)
        if getattr(self, "ic_automacao", None) is not None:
            ctk.CTkLabel(frame, text="", image=self.ic_automacao, fg_color="#FFFFFF", corner_radius=12,
                         width=self._automacao_size[0] + 20, height=self._automacao_size[1] + 10).pack(pady=(0, 14))
        ctk.CTkLabel(frame, text="Suporte Automação 2000", text_color=P.TEXT,
                     font=self._fonte(17, "bold")).pack()
        ctk.CTkLabel(frame, text="Precisa de ajuda? Fale com a gente:", text_color=P.TEXT_SUB,
                     font=self._fonte(12)).pack(pady=(10, 0))
        ctk.CTkLabel(frame, text="(11) 95413-6081", text_color=P.ACCENT,
                     font=self._fonte(20, "bold")).pack(pady=(6, 0))
        ctk.CTkLabel(frame, text="WhatsApp", text_color=P.TEXT_SUB, font=self._fonte(12)).pack()
        ctk.CTkButton(frame, text="Abrir WhatsApp", height=40, corner_radius=8, fg_color=P.SUCCESS,
                      hover_color=P.SUCCESS_HV, text_color=P.ON_ACCENT, font=self._fonte(13, "bold"),
                      command=lambda: webbrowser.open("https://wa.me/5511954136081")).pack(fill=tk.X, pady=(18, 0))
        ctk.CTkButton(frame, text="Fechar", height=36, corner_radius=8, fg_color=P.SURFACE_2,
                      hover_color=P.BORDER, text_color=P.TEXT, font=self._fonte(12, "bold"),
                      command=win.destroy).pack(fill=tk.X, pady=(8, 0))
        ctk.CTkLabel(frame, text="Coletor de Notas Fiscais — versão " + APP_VERSION, text_color=P.TEXT_SUB,
                     font=self._fonte(11)).pack(side=tk.BOTTOM, pady=(10, 0))
        win.lift()
        win.after(150, lambda: self._grab_seguro(win))

    @staticmethod
    def _grab_seguro(win):
        try:
            win.grab_set()
        except Exception:
            pass

    def ao_fechar(self):
        if hasattr(self, "entry_pasta"):
            try:
                self.config["pasta_xml"] = self.entry_pasta.get().strip()
            except Exception:
                pass
        self.salvar_config()
        try:
            self.root.quit()
        finally:
            self.root.destroy()

    # ------------------------------------------------------------------ tema / estilo
    def alternar_tema(self):
        self.tema_escuro = not self.tema_escuro
        self.theme = Themes.DARK if self.tema_escuro else Themes.LIGHT
        self.config["tema_escuro"] = self.tema_escuro
        ctk.set_appearance_mode("dark" if self.tema_escuro else "light")
        self.btn_theme.configure(text="  " + ("Tema claro" if self.tema_escuro else "Tema escuro"))
        self.aplicar_estilo_treeview()
        self.salvar_config()

    def _fonte(self, size, weight="normal"):
        return ctk.CTkFont(family="Segoe UI", size=size, weight=weight)

    def aplicar_estilo_treeview(self):
        esc = self.tema_escuro
        estilo = ttk.Style()
        estilo.theme_use("clam")
        estilo.configure("Custom.Treeview", background=cor(P.SURFACE, esc), foreground=cor(P.TEXT, esc),
                         fieldbackground=cor(P.SURFACE, esc), borderwidth=0, rowheight=34,
                         font=("Segoe UI", 10))
        estilo.configure("Custom.Treeview.Heading", background=cor(P.SURFACE_2, esc),
                         foreground=cor(P.TEXT_SUB, esc), borderwidth=0, relief="flat", padding=(10, 8),
                         font=("Segoe UI Semibold", 10))
        estilo.map("Custom.Treeview.Heading", background=[("active", cor(P.SURFACE_2, esc))])
        estilo.map("Custom.Treeview", background=[("selected", cor(P.ACCENT, esc))],
                   foreground=[("selected", "#FFFFFF")])
        if hasattr(self, "tree"):
            self.tree.tag_configure("par", background=cor(P.SURFACE, esc))
            self.tree.tag_configure("impar", background=cor(P.BG, esc))
            # notas canceladas aparecem em vermelho
            self.tree.tag_configure("cancelada", foreground=cor(P.ACCENT, esc))

    def _criar_icones(self):
        """Cria os CTkImage dos ícones. O ícone de tema usa light_image/dark_image,
        então a lua (modo claro) vira sol (modo escuro) automaticamente."""
        self.ic_logo = self.ic_tema = self.ic_automacao = self.ic_automacao_header = None
        self.ic_nav_arquivos = self.ic_nav_config = self.ic_nav_sobre = None
        self._automacao_size = (220, 110)
        self._automacao_header_size = (80, 40)
        if not PIL_AVAILABLE:
            return
        try:
            self.ic_logo = ctk.CTkImage(light_image=desenhar_xml(76), dark_image=desenhar_xml(76), size=(38, 38))
            self.ic_tema = ctk.CTkImage(light_image=desenhar_lua(68, cor(P.TEXT, False)),
                                        dark_image=desenhar_sol(68, cor(P.TEXT, True)), size=(17, 17))

            def par(fn):
                return dict(light_image=fn(72, cor(P.TEXT_SUB, False)),
                            dark_image=fn(72, cor(P.TEXT_SUB, True)), size=(18, 18))
            self.ic_nav_arquivos = ctk.CTkImage(**par(desenhar_nota))
            self.ic_nav_config = ctk.CTkImage(**par(desenhar_engrenagem))
            self.ic_nav_sobre = ctk.CTkImage(**par(desenhar_info))
        except Exception as e:
            print("Erro nos ícones: ", e)
        try:
            caminho = resource_path("logo_automacao2000.png")
            if os.path.exists(caminho):
                img = Image.open(caminho).convert("RGBA")
                w, h = img.size
                bw, bh = self._automacao_size
                esc = min(bw / w, bh / h)
                self._automacao_size = (max(int(w * esc), 1), max(int(h * esc), 1))
                hh = 40
                self._automacao_header_size = (max(int(w * hh / h), 1), hh)
                self.ic_automacao = ctk.CTkImage(light_image=img, dark_image=img, size=self._automacao_size)
                self.ic_automacao_header = ctk.CTkImage(light_image=img, dark_image=img,
                                                        size=self._automacao_header_size)
        except Exception as e:
            print("Erro no logo: ", e)

    # ------------------------------------------------------------------ interface
    def criar_interface(self):
        self._criar_icones()
        self.root.configure(fg_color=P.BG)

        sidebar = ctk.CTkFrame(self.root, width=232, corner_radius=0, fg_color=P.SURFACE)
        sidebar.pack(side=tk.LEFT, fill=tk.Y)
        sidebar.pack_propagate(False)
        ctk.CTkFrame(self.root, width=1, fg_color=P.BORDER).pack(side=tk.LEFT, fill=tk.Y)

        # topo: logo + nome
        topo = ctk.CTkFrame(sidebar, fg_color="transparent")
        topo.pack(side=tk.TOP, fill=tk.X, padx=18, pady=(18, 12))
        if self.ic_automacao_header is not None:
            ctk.CTkLabel(topo, text="", image=self.ic_automacao_header, fg_color="#FFFFFF",
                         corner_radius=8).pack(anchor="w")
        else:
            ctk.CTkLabel(topo, text="Automação 2000", text_color=P.ACCENT,
                         font=self._fonte(15, "bold")).pack(anchor="w")
        ctk.CTkLabel(topo, text="Coletor de Notas Fiscais", text_color=P.TEXT,
                     font=self._fonte(14, "bold")).pack(anchor="w", pady=(8, 0))
        ctk.CTkLabel(topo, text="Sistema de envio automático de XMLs", text_color=P.TEXT_SUB,
                     font=self._fonte(11), wraplength=190, justify="left").pack(anchor="w", pady=(2, 0))
        ctk.CTkFrame(sidebar, height=1, fg_color=P.BORDER).pack(side=tk.TOP, fill=tk.X, padx=18, pady=(4, 12))

        # navegação
        self.nav_buttons = {}

        def nav_btn(chave, texto, icone):
            b = ctk.CTkButton(sidebar, text="  " + texto, image=icone, compound="left", anchor="w",
                              command=lambda c=chave: self._mostrar_pagina(c), height=40, corner_radius=8,
                              fg_color="transparent", hover_color=P.SURFACE_2, text_color=P.TEXT_SUB,
                              font=self._fonte(13, "bold"))
            b.pack(side=tk.TOP, fill=tk.X, padx=12, pady=(0, 4))
            self.nav_buttons[chave] = b
        nav_btn("arquivos", "Arquivos", self.ic_nav_arquivos)
        nav_btn("config", "Configurações", self.ic_nav_config)
        nav_btn("sobre", "Sobre", self.ic_nav_sobre)

        # rodapé da barra lateral
        rodape = ctk.CTkFrame(sidebar, fg_color="transparent")
        rodape.pack(side=tk.BOTTOM, fill=tk.X, padx=12, pady=(0, 16))
        self.btn_theme = ctk.CTkButton(
            rodape, text="  " + ("Tema claro" if self.tema_escuro else "Tema escuro"), image=self.ic_tema,
            compound="left", command=self.alternar_tema, height=36, corner_radius=8, fg_color=P.SURFACE_2,
            hover_color=P.BORDER, text_color=P.TEXT, font=self._fonte(12, "bold"))
        self.btn_theme.pack(fill=tk.X, pady=(0, 8))
        ctk.CTkButton(rodape, text="Ajuda", command=self.abrir_suporte, height=36, corner_radius=8,
                      fg_color=P.SURFACE_2, hover_color=P.BORDER, text_color=P.TEXT,
                      font=self._fonte(12, "bold")).pack(fill=tk.X)
        info = ctk.CTkFrame(sidebar, fg_color="transparent")
        info.pack(side=tk.BOTTOM, fill=tk.X, padx=18, pady=(0, 12))
        ctk.CTkLabel(info, text="Versão " + APP_VERSION, text_color=P.TEXT_SUB,
                     font=self._fonte(11)).pack(anchor="w")
        self.label_info = ctk.CTkLabel(info, text="Pronto", text_color=P.SUCCESS, font=self._fonte(11, "bold"))
        self.label_info.pack(anchor="w", pady=(2, 0))

        # área de conteúdo com as páginas
        conteudo = ctk.CTkFrame(self.root, fg_color="transparent")
        conteudo.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=22, pady=18)
        self.paginas = {}
        for chave in ("arquivos", "config", "sobre"):
            self.paginas[chave] = ctk.CTkFrame(conteudo, fg_color="transparent")
        self.criar_pagina_arquivos(self.paginas["arquivos"])
        self.criar_pagina_config(self.paginas["config"])
        self.criar_pagina_sobre(self.paginas["sobre"])
        self.pagina_atual = None
        self._mostrar_pagina("arquivos")
        self.aplicar_estilo_treeview()

    def _mostrar_pagina(self, chave):
        """Troca a página visível e realça o item de menu correspondente."""
        if getattr(self, "pagina_atual", None) == chave:
            return
        for k, frame in self.paginas.items():
            frame.pack_forget()
        self.paginas[chave].pack(fill=tk.BOTH, expand=True)
        for k, b in self.nav_buttons.items():
            if k == chave:
                b.configure(fg_color=P.ACCENT_SOFT, text_color=P.ACCENT)
            else:
                b.configure(fg_color="transparent", text_color=P.TEXT_SUB)
        self.pagina_atual = chave


    # ------------------------------------------------------------------ página: Arquivos
    def criar_pagina_arquivos(self, parent):
        # --- cartão de filtros ------------------------------------------------
        card = ctk.CTkFrame(parent, fg_color=P.SURFACE, corner_radius=12, border_width=1, border_color=P.BORDER)
        card.pack(fill=tk.X)
        interno = ctk.CTkFrame(card, fg_color="transparent")
        interno.pack(fill=tk.X, padx=18, pady=14)

        def coluna(rotulo):
            col = ctk.CTkFrame(interno, fg_color="transparent")
            ctk.CTkLabel(col, text=rotulo, text_color=P.TEXT_SUB, font=self._fonte(10, "bold")).pack(anchor="w")
            return col

        col_pasta = coluna("PASTA DOS XMLS")
        col_pasta.pack(side=tk.LEFT, fill=tk.X, expand=True)
        linha = ctk.CTkFrame(col_pasta, fg_color="transparent")
        linha.pack(fill=tk.X, pady=(4, 0))
        self.entry_pasta = ctk.CTkEntry(linha, height=36, corner_radius=8, fg_color=P.INPUT, border_color=P.BORDER,
                                        text_color=P.TEXT, font=self._fonte(12),
                                        placeholder_text="Ex.: C:\\CUSTODIA\\NFCE\\XML")
        self.entry_pasta.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        self.entry_pasta.insert(0, self.config.get("pasta_xml", ""))
        ctk.CTkButton(linha, text="Procurar", width=84, height=36, corner_radius=8, command=self.selecionar_pasta,
                      fg_color=P.SURFACE, hover_color=P.SURFACE_2, text_color=P.TEXT, border_width=1,
                      border_color=P.BORDER, font=self._fonte(12, "bold")).pack(side=tk.LEFT)

        agora = datetime.now()
        col_mes = coluna("MÊS")
        col_mes.pack(side=tk.LEFT, padx=(14, 0))
        self.combo_mes = ctk.CTkComboBox(
            col_mes, values=[f"{i + 1:02d} - {n}" for i, n in enumerate(MESES)], width=150, height=36,
            corner_radius=8, fg_color=P.INPUT, border_color=P.BORDER, button_color=P.SURFACE_2,
            button_hover_color=P.BORDER, text_color=P.TEXT, font=self._fonte(12), dropdown_fg_color=P.SURFACE,
            dropdown_text_color=P.TEXT, dropdown_hover_color=P.SURFACE_2, state="readonly")
        self.combo_mes.pack(pady=(4, 0))
        self.combo_mes.set(f"{agora.month:02d} - {MESES[agora.month - 1]}")

        col_ano = coluna("ANO")
        col_ano.pack(side=tk.LEFT, padx=(10, 0))
        self.combo_ano = ctk.CTkComboBox(
            col_ano, values=[str(a) for a in range(agora.year - 6, agora.year + 1)], width=90, height=36,
            corner_radius=8, fg_color=P.INPUT, border_color=P.BORDER, button_color=P.SURFACE_2,
            button_hover_color=P.BORDER, text_color=P.TEXT, font=self._fonte(12), dropdown_fg_color=P.SURFACE,
            dropdown_text_color=P.TEXT, dropdown_hover_color=P.SURFACE_2, state="readonly")
        self.combo_ano.pack(pady=(4, 0))
        self.combo_ano.set(str(agora.year))

        self.btn_buscar = ctk.CTkButton(interno, text="Buscar XMLs", width=120, height=36, corner_radius=8,
                                        command=self.buscar_arquivos_thread, fg_color=P.ACCENT,
                                        hover_color=P.ACCENT_HV, text_color=P.ON_ACCENT,
                                        font=self._fonte(12, "bold"))
        self.btn_buscar.pack(side=tk.LEFT, padx=(14, 0), pady=(18, 0))

        # --- cartões de resumo --------------------------------------------------
        faixa = ctk.CTkFrame(parent, fg_color="transparent")
        faixa.pack(fill=tk.X, pady=(12, 0))

        def cartao_status(valor_ini, rotulo, cor_valor=None):
            c = ctk.CTkFrame(faixa, fg_color=P.SURFACE, corner_radius=12, border_width=1, border_color=P.BORDER)
            c.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
            lbl = ctk.CTkLabel(c, text=valor_ini, text_color=cor_valor or P.TEXT, font=self._fonte(20, "bold"))
            lbl.pack(anchor="w", padx=16, pady=(10, 0))
            ctk.CTkLabel(c, text=rotulo, text_color=P.TEXT_SUB, font=self._fonte(11)).pack(
                anchor="w", padx=16, pady=(0, 10))
            return lbl

        self.stat_qtd = cartao_status("0", "XMLs encontrados")
        self.stat_aut = cartao_status("0", "Autorizadas", P.SUCCESS)
        self.stat_canc = cartao_status("0", "Canceladas", P.ACCENT)
        self.stat_busca = cartao_status("—", "Última busca")

        # --- lista --------------------------------------------------------------
        lista = ctk.CTkFrame(parent, fg_color=P.SURFACE, corner_radius=12, border_width=1, border_color=P.BORDER)
        lista.pack(fill=tk.BOTH, expand=True, pady=(2, 0))
        cab = ctk.CTkFrame(lista, fg_color="transparent")
        cab.pack(fill=tk.X, padx=16, pady=(12, 6))
        ctk.CTkLabel(cab, text="Arquivos encontrados", text_color=P.TEXT,
                     font=self._fonte(14, "bold")).pack(side=tk.LEFT)
        self.lbl_cab_count = ctk.CTkLabel(cab, text="(0)", text_color=P.TEXT_SUB, font=self._fonte(13))
        self.lbl_cab_count.pack(side=tk.LEFT, padx=(6, 0))
        self.btn_relatorio = ctk.CTkButton(cab, text="Relatório PDF", width=116, height=34, corner_radius=8,
                                           command=self.gerar_relatorio_thread, fg_color=P.SURFACE_2,
                                           hover_color=P.BORDER, text_color=P.TEXT, font=self._fonte(12, "bold"))
        self.btn_relatorio.pack(side=tk.RIGHT)
        self.btn_exportar = ctk.CTkButton(cab, text="Exportar ZIP", width=110, height=34, corner_radius=8,
                                          command=self.exportar_xmls_zip, fg_color=P.SURFACE_2,
                                          hover_color=P.BORDER, text_color=P.TEXT, font=self._fonte(12, "bold"))
        self.btn_exportar.pack(side=tk.RIGHT, padx=(0, 8))
        self.btn_enviar = ctk.CTkButton(cab, text="Enviar E-mail", width=118, height=34, corner_radius=8,
                                        command=self.enviar_email_thread, fg_color=P.ACCENT,
                                        hover_color=P.ACCENT_HV, text_color=P.ON_ACCENT,
                                        font=self._fonte(12, "bold"))
        self.btn_enviar.pack(side=tk.RIGHT, padx=(0, 8))

        corpo = ctk.CTkFrame(lista, fg_color="transparent")
        corpo.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 12))
        self.tree = ttk.Treeview(corpo, columns=("nome", "tipo", "data", "tamanho"), show="headings",
                                 style="Custom.Treeview", selectmode="extended")
        self.tree.heading("nome", text="  Nome do arquivo", anchor="w")
        self.tree.heading("tipo", text="Tipo", anchor="w")
        self.tree.heading("data", text="Data", anchor="w")
        self.tree.heading("tamanho", text="Tamanho", anchor="e")
        self.tree.column("nome", width=360, anchor="w")
        self.tree.column("tipo", width=110, anchor="w", stretch=False)
        self.tree.column("data", width=140, anchor="w", stretch=False)
        self.tree.column("tamanho", width=100, anchor="e", stretch=False)
        barra = ctk.CTkScrollbar(corpo, command=self.tree.yview)
        self.tree.configure(yscrollcommand=barra.set)
        barra.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    # ------------------------------------------------------------------ página: Configurações
    def criar_pagina_config(self, parent):
        # barra de botões (rodapé)
        rodape = ctk.CTkFrame(parent, fg_color=P.SURFACE, corner_radius=12, border_width=1, border_color=P.BORDER)
        rodape.pack(side=tk.BOTTOM, fill=tk.X, pady=(12, 0))
        ctk.CTkButton(rodape, text="Salvar Configurações", height=38, corner_radius=8,
                      command=self.salvar_configuracoes, fg_color=P.ACCENT, hover_color=P.ACCENT_HV,
                      text_color=P.ON_ACCENT, font=self._fonte(12, "bold")).pack(side=tk.RIGHT, padx=14, pady=12)
        ctk.CTkButton(rodape, text="Testar Conexão", height=38, corner_radius=8, command=self.testar_conexao,
                      fg_color=P.SURFACE, hover_color=P.SURFACE_2, text_color=P.TEXT, border_width=1,
                      border_color=P.BORDER, font=self._fonte(12, "bold")).pack(side=tk.RIGHT, pady=12)

        corpo = ctk.CTkFrame(parent, fg_color="transparent")
        corpo.pack(fill=tk.BOTH, expand=True)

        lateral = ctk.CTkFrame(corpo, width=190, fg_color=P.SURFACE, corner_radius=12, border_width=1,
                               border_color=P.BORDER)
        lateral.pack(side=tk.LEFT, fill=tk.Y)
        lateral.pack_propagate(False)
        self.cfg_nav = {}
        self.cfg_secoes = {}
        self.cfg_conteudo = ctk.CTkFrame(corpo, fg_color=P.SURFACE, corner_radius=12, border_width=1,
                                         border_color=P.BORDER)
        self.cfg_conteudo.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(12, 0))

        def cfg_btn(chave, texto):
            b = ctk.CTkButton(lateral, text="  " + texto, anchor="w", height=38, corner_radius=8,
                              fg_color="transparent", hover_color=P.SURFACE_2, text_color=P.TEXT_SUB,
                              font=self._fonte(12, "bold"), command=lambda c=chave: self._mostrar_config_secao(c))
            b.pack(side=tk.TOP, fill=tk.X, padx=10, pady=(8 if not self.cfg_nav else 2, 0))
            self.cfg_nav[chave] = b

        def nova_secao(titulo, subtitulo):
            f = ctk.CTkFrame(self.cfg_conteudo, fg_color="transparent")
            ctk.CTkLabel(f, text=titulo, text_color=P.TEXT, font=self._fonte(16, "bold")).pack(
                anchor="w", padx=22, pady=(18, 0))
            ctk.CTkLabel(f, text=subtitulo, text_color=P.TEXT_SUB, font=self._fonte(12)).pack(
                anchor="w", padx=22, pady=(2, 10))
            return f

        def campo(dono, rotulo, var_key, hide=False, ajuda=None):
            ctk.CTkLabel(dono, text=rotulo, text_color=P.TEXT_SUB, font=self._fonte(10, "bold")).pack(
                anchor="w", padx=22, pady=(8, 0))
            e = ctk.CTkEntry(dono, height=36, corner_radius=8, fg_color=P.INPUT, border_color=P.BORDER,
                             text_color=P.TEXT, font=self._fonte(12), show="*" if hide else "")
            e.pack(fill=tk.X, padx=22, pady=(4, 0))
            e.insert(0, str(self.config.get(var_key, "") or ""))
            if ajuda:
                ctk.CTkLabel(dono, text=ajuda, text_color=P.TEXT_SUB, font=self._fonte(11)).pack(
                    anchor="w", padx=22, pady=(4, 0))
            return e

        # --- Conta de envio
        s = nova_secao("Conta de envio", "Conta que envia os e-mails (já configurada de fábrica).")
        self.entry_email_remetente = campo(s, "E-MAIL REMETENTE", "email_remetente")
        self.entry_senha = campo(s, "SENHA DE APLICATIVO", "senha_email", hide=True,
                                 ajuda="Senha de 16 letras gerada na conta Google (não é a senha normal do e-mail).")
        self.var_salvar_senha = tk.BooleanVar(value=True)

        def alternar_senha():
            if self.entry_senha.cget("show") == "*":
                self.entry_senha.configure(show="")
                self.btn_ver_senha.configure(text="Ocultar")
            else:
                self.entry_senha.configure(show="*")
                self.btn_ver_senha.configure(text="Mostrar")
        linha = ctk.CTkFrame(s, fg_color="transparent")
        linha.pack(fill=tk.X, padx=22, pady=(10, 0))
        self.btn_ver_senha = ctk.CTkButton(linha, text="Mostrar", width=90, height=32, corner_radius=8,
                                           command=alternar_senha, fg_color=P.SURFACE_2, hover_color=P.BORDER,
                                           text_color=P.TEXT, font=self._fonte(12, "bold"))
        self.btn_ver_senha.pack(side=tk.LEFT)
        ctk.CTkCheckBox(linha, text="Salvar senha localmente", variable=self.var_salvar_senha,
                        text_color=P.TEXT, fg_color=P.ACCENT, hover_color=P.ACCENT_HV,
                        font=self._fonte(12)).pack(side=tk.LEFT, padx=(14, 0))
        self.cfg_secoes["conta"] = s

        # --- Servidor SMTP
        s = nova_secao("Servidor SMTP", "Endereço e porta do servidor de e-mail.")
        self.entry_smtp = campo(s, "SERVIDOR SMTP", "servidor_smtp")
        self.entry_porta = campo(s, "PORTA", "porta_smtp")
        self.var_ssl = tk.BooleanVar(value=bool(self.config.get("usar_ssl", True)))
        ctk.CTkCheckBox(s, text="Usar SSL/TLS", variable=self.var_ssl, text_color=P.TEXT, fg_color=P.ACCENT,
                        hover_color=P.ACCENT_HV, font=self._fonte(12)).pack(anchor="w", padx=22, pady=(12, 0))
        self.cfg_secoes["smtp"] = s

        # --- Destinatários
        s = nova_secao("Destinatários", "Quem recebe os XMLs enviados.")
        self.entry_destinatario = campo(
            s, "E-MAILS DESTINATÁRIOS", "email_destinatario",
            ajuda="Para vários e-mails, separe por vírgula (ex.: um@x.com, outro@y.com).")
        self.cfg_secoes["dest"] = s

        # --- Pastas (novo na 3.1)
        s = nova_secao("Pastas das notas", "Onde o programa procura as notas autorizadas e canceladas.")
        self.entry_pasta_cancelados = campo(
            s, "PASTA DAS NOTAS CANCELADAS (OPCIONAL)", "pasta_xml_cancelados",
            ajuda="Deixe em branco se as canceladas ficam na mesma pasta dos XMLs (ou em uma subpasta dela).")
        self.var_subpastas = tk.BooleanVar(value=bool(self.config.get("buscar_subpastas", True)))
        ctk.CTkCheckBox(s, text="Procurar também dentro das subpastas", variable=self.var_subpastas,
                        text_color=P.TEXT, fg_color=P.ACCENT, hover_color=P.ACCENT_HV,
                        font=self._fonte(12)).pack(anchor="w", padx=22, pady=(14, 0))
        self.cfg_secoes["pastas"] = s

        # --- Atualização
        s = nova_secao("Atualização", "Verifique se existe uma versão mais nova do programa.")
        ctk.CTkLabel(s, text="Versão instalada: " + APP_VERSION, text_color=P.TEXT,
                     font=self._fonte(12)).pack(anchor="w", padx=22, pady=(4, 0))
        ctk.CTkButton(s, text="Verificar atualização", height=36, width=170, corner_radius=8,
                      command=lambda: self.verificar_atualizacao(silencioso=False), fg_color=P.ACCENT,
                      hover_color=P.ACCENT_HV, text_color=P.ON_ACCENT,
                      font=self._fonte(12, "bold")).pack(anchor="w", padx=22, pady=(14, 0))
        self.cfg_secoes["atualizacao"] = s

        cfg_btn("conta", "Conta de envio")
        cfg_btn("smtp", "Servidor SMTP")
        cfg_btn("dest", "Destinatários")
        cfg_btn("pastas", "Pastas das notas")
        cfg_btn("atualizacao", "Atualização")
        self.cfg_secao_atual = None
        self._mostrar_config_secao("conta")

    def _mostrar_config_secao(self, chave):
        """Troca a seção visível dentro de Configurações e realça o sub-menu."""
        for k, f in self.cfg_secoes.items():
            f.pack_forget()
        self.cfg_secoes[chave].pack(fill=tk.BOTH, expand=True, anchor="n")
        for k, b in self.cfg_nav.items():
            if k == chave:
                b.configure(fg_color=P.ACCENT_SOFT, text_color=P.ACCENT)
            else:
                b.configure(fg_color="transparent", text_color=P.TEXT_SUB)
        self.cfg_secao_atual = chave

    # ------------------------------------------------------------------ página: Sobre
    def criar_pagina_sobre(self, parent):
        card = ctk.CTkFrame(parent, fg_color=P.SURFACE, corner_radius=14, border_width=1, border_color=P.BORDER)
        card.pack(fill=tk.BOTH, expand=True)
        meio = ctk.CTkFrame(card, fg_color="transparent")
        meio.pack(expand=True, padx=30, pady=30)
        if getattr(self, "ic_automacao", None) is not None:
            ctk.CTkLabel(meio, text="", image=self.ic_automacao, fg_color="#FFFFFF", corner_radius=12).pack(
                pady=(0, 14))
        ctk.CTkLabel(meio, text="Coletor de Notas Fiscais", text_color=P.TEXT,
                     font=self._fonte(22, "bold")).pack()
        ctk.CTkLabel(meio, text="Sistema de envio automático de XMLs de NFCe", text_color=P.TEXT_SUB,
                     font=self._fonte(13)).pack(pady=(4, 0))
        ctk.CTkLabel(meio, text="Versão " + APP_VERSION, text_color=P.TEXT_SUB,
                     font=self._fonte(12)).pack(pady=(10, 0))
        botoes = ctk.CTkFrame(meio, fg_color="transparent")
        botoes.pack(pady=(20, 0))
        ctk.CTkButton(botoes, text="Verificar atualização", height=38, width=170, corner_radius=8,
                      command=lambda: self.verificar_atualizacao(silencioso=False), fg_color=P.ACCENT,
                      hover_color=P.ACCENT_HV, text_color=P.ON_ACCENT,
                      font=self._fonte(12, "bold")).pack(side=tk.LEFT)
        ctk.CTkButton(botoes, text="Ajuda / Suporte", height=38, width=150, corner_radius=8,
                      command=self.abrir_suporte, fg_color=P.SURFACE_2, hover_color=P.BORDER,
                      text_color=P.TEXT, font=self._fonte(12, "bold")).pack(side=tk.LEFT, padx=(10, 0))


    # ------------------------------------------------------------------ utilidades
    def obter_desktop(self):
        """Lê o caminho real do Desktop direto do registro do Windows.
        Funciona tanto com OneDrive quanto sem."""
        caminho = None
        try:
            if winreg is not None:
                chave = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                                       r"Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders")
                caminho, _ = winreg.QueryValueEx(chave, "Desktop")
                winreg.CloseKey(chave)
        except Exception:
            caminho = None
        if caminho and os.path.exists(caminho):
            return str(caminho)
        for cand in (Path.home() / "Desktop", Path.home() / "OneDrive" / "Desktop",
                     Path.home() / "OneDrive" / "Área de Trabalho"):
            if cand.exists():
                return str(cand)
        raise FileNotFoundError("Não foi possível localizar a pasta Desktop: " + str(Path.home()))

    def selecionar_pasta(self):
        pasta = filedialog.askdirectory(initialdir=self.entry_pasta.get().strip() or None)
        if pasta:
            self.entry_pasta.delete(0, tk.END)
            self.entry_pasta.insert(0, os.path.normpath(pasta))
            self.config["pasta_xml"] = self.entry_pasta.get().strip()

    def _mes_ano(self):
        """Devolve (numero_do_mes, nome_do_mes, ano) conforme os seletores da tela."""
        partes = self.combo_mes.get().split(" - ")
        return int(partes[0]), partes[1].strip(), int(self.combo_ano.get())

    def _nome_arquivo_seguro(self, s):
        """Remove acentos para uso em NOMES DE ARQUIVO (ex.: 'Março' -> 'Marco').
        Caracteres como 'ç' e 'ã' podem dar problema em e-mails e no Windows."""
        return "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))

    def _fmt_moeda(self, v):
        """Formata valor no padrão brasileiro: 1234.5 -> R$ 1.234,50"""
        return "R$ " + f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    def _fmt_emissao(self, txt):
        """Converte a data de emissão do XML para dd/mm/aaaa HH:MM."""
        if not txt:
            return ""
        try:
            return datetime.fromisoformat(txt).strftime("%d/%m/%Y %H:%M")
        except Exception:
            pass
        for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
            try:
                return datetime.strptime(txt[:10], fmt).strftime("%d/%m/%Y")
            except Exception:
                continue
        return txt

    def _ler_campos_config(self):
        """Copia o que está digitado nas telas de Configurações para self.config."""
        c = self.config
        c["email_remetente"] = self.entry_email_remetente.get().strip()
        c["senha_email"] = self.entry_senha.get().replace(" ", "").strip()
        c["servidor_smtp"] = self.entry_smtp.get().strip()
        try:
            c["porta_smtp"] = int(self.entry_porta.get().strip())
        except ValueError:
            pass
        c["email_destinatario"] = self.entry_destinatario.get().strip()
        c["usar_ssl"] = bool(self.var_ssl.get())
        c["pasta_xml_cancelados"] = self.entry_pasta_cancelados.get().strip()
        c["buscar_subpastas"] = bool(self.var_subpastas.get())

    # ------------------------------------------------------------------ busca
    def buscar_arquivos_thread(self):
        try:
            pasta = self.entry_pasta.get().strip()
            mes, _, ano = self._mes_ano()
        except Exception as e:
            self.mostrar_erro("BUSCA", e, "Erro ao buscar arquivos")
            return
        self.config["pasta_xml"] = pasta
        if hasattr(self, "entry_smtp"):
            self._ler_campos_config()
        extra = (self.config.get("pasta_xml_cancelados") or "").strip()
        recursivo = bool(self.config.get("buscar_subpastas", True))
        self._busca_params = (pasta, extra, mes, ano, recursivo)

        self.btn_buscar.configure(state="disabled", text="Buscando...")
        self.label_info.configure(text="Buscando...")
        self._dlg_busca = DialogoBusca(self.root, pasta, getattr(self, "icon_path", None))
        threading.Thread(target=self.buscar_arquivos, daemon=True).start()

    def buscar_arquivos(self):
        """Roda em segundo plano. Lista os .xml do período (pasta principal, pasta das
        canceladas e subpastas) e classifica cada um em AUTORIZADA / CANCELADA / OUTRO."""
        dlg = self._dlg_busca
        pasta, extra, mes, ano, recursivo = self._busca_params

        def _upd(**kw):
            self.root.after(0, lambda: dlg.atualizar(**kw))

        def _restaurar():
            self.btn_buscar.configure(state="normal", text="Buscar XMLs")
            self.label_info.configure(text="Pronto")

        try:
            if not pasta or not os.path.exists(pasta):
                def _erro():
                    dlg.fechar()
                    messagebox.showerror("Erro", "Pasta não encontrada")
                    _restaurar()
                self.root.after(0, _erro)
                return

            raizes = [pasta]
            if extra and os.path.isdir(extra) and os.path.normcase(os.path.abspath(extra)) != \
                    os.path.normcase(os.path.abspath(pasta)):
                raizes.append(extra)

            # 1) lista os candidatos (.xml)
            candidatos, vistos = [], set()
            for raiz in raizes:
                if recursivo:
                    for diretorio, _subs, nomes in os.walk(raiz):
                        for nome in nomes:
                            if nome.lower().endswith(".xml"):
                                candidatos.append(os.path.join(diretorio, nome))
                else:
                    for nome in os.listdir(raiz):
                        if nome.lower().endswith(".xml"):
                            candidatos.append(os.path.join(raiz, nome))

            # 2) filtra por mês/ano e classifica
            achados, erros, total = [], 0, max(len(candidatos), 1)
            for i, caminho in enumerate(candidatos, 1):
                if dlg.cancelado:
                    break
                chave_norm = os.path.normcase(os.path.abspath(caminho))
                if chave_norm in vistos:
                    continue
                vistos.add(chave_norm)
                try:
                    st = os.stat(caminho)
                    data = datetime.fromtimestamp(st.st_mtime)
                    if data.month == mes and data.year == ano:
                        tipo, chave = classificar_xml(caminho)
                        if tipo == OUTRO and chave is None:
                            erros += 1  # não é NFC-e nem evento reconhecido (ou está corrompido)
                        achados.append({"nome": os.path.basename(caminho), "data": data, "tamanho": st.st_size,
                                        "caminho": caminho, "tipo": tipo, "chave": chave})
                except Exception:
                    erros += 1
                if i % 20 == 0 or i == len(candidatos):
                    _upd(encontrados=len(achados), processados=i, erros=erros, pct=i / total)

            aplicar_pareamento(achados)
            achados.sort(key=lambda a: a["data"])
            cancelou = dlg.cancelado

            def atualizar_ui():
                dlg.fechar()
                self._dlg_busca = None
                self.arquivos_encontrados = achados
                self.tree.delete(*self.tree.get_children())
                for i, a in enumerate(achados):
                    tags = ["par" if i % 2 == 0 else "impar"]
                    if a["tipo"] == CANCELADA:
                        tags.append("cancelada")
                    self.tree.insert("", tk.END, tags=tuple(tags), values=(
                        "  " + a["nome"], ROTULO_TIPO_UNICO.get(a["tipo"], "Outro"),
                        a["data"].strftime("%d/%m/%Y %H:%M"), f"{a['tamanho'] / 1024:.1f} KB  "))
                r = resumo_por_tipo(achados)
                self.stat_qtd.configure(text=str(len(achados)))
                self.stat_aut.configure(text=str(r[AUTORIZADA]))
                self.stat_canc.configure(text=str(r[CANCELADA]))
                self.lbl_cab_count.configure(text="(" + str(len(achados)) + ")")
                self.stat_busca.configure(text=datetime.now().strftime("%d/%m/%Y %H:%M"))
                _restaurar()
                if cancelou:
                    messagebox.showinfo("Busca cancelada",
                                        "Busca interrompida. " + str(len(achados)) + " arquivo(s) listados até aqui.")
                elif not achados:
                    messagebox.showinfo("Resultado", "Nenhum arquivo encontrado neste período.")
            self.root.after(0, atualizar_ui)
        except Exception as e:
            def _falha():
                dlg.fechar()
                _restaurar()
            self.root.after(0, _falha)
            self.mostrar_erro("BUSCA", e, "Erro ao buscar arquivos")

    def _notas_para_envio(self):
        """Só entram em ZIP/e-mail as notas AUTORIZADAS e as CANCELADAS."""
        return [a for a in self.arquivos_encontrados if a.get("tipo") in (AUTORIZADA, CANCELADA)]

    # ------------------------------------------------------------------ exportar ZIP
    def exportar_xmls_zip(self):
        """Exporta os XMLs para um ZIP na Área de Trabalho (com suporte a OneDrive).
        O ZIP tem duas pastas: Autorizadas/ e Canceladas/."""
        if not self.arquivos_encontrados:
            messagebox.showwarning("Aviso", "Busque arquivos primeiro")
            return
        notas = self._notas_para_envio()
        if not notas:
            messagebox.showwarning("Aviso", "Não há notas autorizadas ou canceladas na lista.")
            return
        try:
            desktop = self.obter_desktop()
            _, mes_nome, ano = self._mes_ano()
            base = "NFCe_" + self._nome_arquivo_seguro(mes_nome) + " - " + str(ano)
            destino, n = os.path.join(desktop, base + ".zip"), 1
            while os.path.exists(destino):
                n += 1
                destino = os.path.join(desktop, f"{base} ({n}).zip")
            usados = {AUTORIZADA: set(), CANCELADA: set()}
            qtd = {AUTORIZADA: 0, CANCELADA: 0}
            pasta_zip = {AUTORIZADA: "Autorizadas", CANCELADA: "Canceladas"}
            with zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED) as z:
                for a in notas:
                    nome = os.path.basename(a["caminho"])
                    raiz, ext = os.path.splitext(nome)
                    k = 1
                    while nome in usados[a["tipo"]]:
                        k += 1
                        nome = f"{raiz} ({k}){ext}"
                    usados[a["tipo"]].add(nome)
                    z.write(a["caminho"], pasta_zip[a["tipo"]] + "/" + nome)
                    qtd[a["tipo"]] += 1
            ignorados = len(self.arquivos_encontrados) - len(notas)
            msg = ("Arquivo salvo na Área de Trabalho:\n" + os.path.basename(destino) +
                   f"\n\nAutorizadas: {qtd[AUTORIZADA]}\nCanceladas: {qtd[CANCELADA]}")
            if ignorados:
                msg += f"\n\n({ignorados} arquivo(s) fora do ZIP: denegadas, rejeitadas ou que não são NFC-e.)"
            messagebox.showinfo("Sucesso", msg)
        except Exception as e:
            self.mostrar_erro("EXPORTAR-ZIP", e, "Erro ao exportar ZIP")

    # ------------------------------------------------------------------ relatório PDF
    def _extrair_dados_nfce(self, caminho):
        """Lê um XML de NFCe e devolve um dict com numero, emissao, chave, valor e
        emitente. Usa curingas de namespace para funcionar com qualquer versão do
        layout. Para eventos de cancelamento, devolve apenas número/chave/data."""
        try:
            raiz = ET.parse(caminho).getroot()
        except Exception:
            return None

        def _txt(caminho_busca):
            el = raiz.find(caminho_busca)
            return el.text.strip() if el is not None and el.text else ""

        if _local(raiz.tag) in ("procEventoNFe", "evento", "envEvento", "procCancNFe", "cancNFe", "retCancNFe"):
            chave = _txt(".//{*}chNFe")
            numero = str(int(chave[25:34])) if len(chave) == 44 and chave[25:34].isdigit() else ""
            return {"numero": numero, "emissao": _txt(".//{*}dhEvento"), "chave": chave, "valor": 0.0,
                    "emit_nome": "", "emit_cnpj": ""}

        inf = raiz.find(".//{*}infNFe")
        if inf is None:
            return None
        try:
            valor = float(_txt(".//{*}total/{*}ICMSTot/{*}vNF").replace(",", "."))
        except (ValueError, TypeError):
            valor = 0.0
        return {
            "numero": _txt(".//{*}ide/{*}nNF"),
            "emissao": _txt(".//{*}ide/{*}dhEmi") or _txt(".//{*}ide/{*}dEmi"),
            "chave": (inf.get("Id") or "").replace("NFe", "", 1),
            "valor": valor,
            "emit_nome": _txt(".//{*}emit/{*}xNome"),
            "emit_cnpj": _txt(".//{*}emit/{*}CNPJ"),
        }

    def gerar_relatorio_thread(self):
        if not self.arquivos_encontrados:
            messagebox.showwarning("Aviso", "Busque arquivos primeiro")
            return
        self.btn_relatorio.configure(state="disabled")
        threading.Thread(target=self.gerar_relatorio_pdf, daemon=True).start()

    def gerar_relatorio_pdf(self):
        """Botão: gera o PDF na Área de Trabalho e mostra o resumo."""
        try:
            _, mes_nome, ano = self._mes_ano()
            nome = "Relatorio_NFCe_" + self._nome_arquivo_seguro(mes_nome) + " - " + str(ano) + ".pdf"
            destino = os.path.join(self.obter_desktop(), nome)
            qtd, soma, canc = self._construir_pdf_relatorio(
                destino, lambda t: self.root.after(0, lambda: self.label_info.configure(text=t)))

            def ok():
                messagebox.showinfo(
                    "Relatório gerado",
                    "PDF salvo na Área de Trabalho:\n" + nome + "\n\nCupons: " + str(qtd) +
                    "\nValor total: " + self._fmt_moeda(soma) +
                    ("\nCanceladas: " + str(canc) if canc else ""))
            self.root.after(0, ok)
        except ImportError:
            self.root.after(0, lambda: messagebox.showerror(
                "Biblioteca ausente",
                "A biblioteca 'reportlab' não está instalada.\n\nInstale com:  pip install reportlab"))
        except Exception as e:
            self.mostrar_erro("RELATORIO", e, "Erro ao gerar relatório")
        finally:
            self.root.after(0, lambda: self.btn_relatorio.configure(state="normal"))
            self.root.after(0, lambda: self.label_info.configure(text="Pronto"))

    def _construir_pdf_relatorio(self, path_pdf, atualizar_label):
        """Lê todos os XMLs encontrados e gera o PDF em path_pdf.
        Retorna (qtd_cupons_autorizados, soma_valores_autorizados, qtd_canceladas).
        Os totais de valor consideram SOMENTE as notas autorizadas; as canceladas
        aparecem numa tabela separada, sem valor."""
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.lib.units import mm
        from reportlab.lib import colors
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, LongTable, TableStyle
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

        arquivos = self.arquivos_encontrados
        total = max(len(arquivos), 1)
        atualizar_label("Lendo XMLs para o relatório...")

        autorizadas, canceladas, ignorados = [], [], 0
        emitente = cnpj = ""
        for i, a in enumerate(arquivos, 1):
            if a.get("tipo") not in (AUTORIZADA, CANCELADA):
                ignorados += 1
                continue
            d = self._extrair_dados_nfce(a["caminho"])
            if d is None:
                ignorados += 1
                continue
            if a["tipo"] == AUTORIZADA:
                autorizadas.append(d)
                if not emitente and d["emit_nome"]:
                    emitente, cnpj = d["emit_nome"], d["emit_cnpj"]
            else:
                # a nota cancelada pode ser o XML da própria nota ou só o evento
                if not d["numero"] and a.get("chave") and len(a["chave"]) == 44:
                    d["numero"] = str(int(a["chave"][25:34]))
                canceladas.append(d)
                if not emitente and d.get("emit_nome"):
                    emitente, cnpj = d["emit_nome"], d["emit_cnpj"]
            if i % 25 == 0:
                atualizar_label("Lendo XMLs... " + f"{i * 100 // total}%")

        # remove duplicatas de canceladas (nota + evento da mesma chave)
        vistas, unicas = set(), []
        for d in canceladas:
            k = d["chave"] or id(d)
            if k not in vistas:
                vistas.add(k)
                unicas.append(d)
        canceladas = unicas

        atualizar_label("Gerando PDF...")
        autorizadas.sort(key=lambda d: (int(d["numero"]) if str(d["numero"]).isdigit() else 0))
        canceladas.sort(key=lambda d: (int(d["numero"]) if str(d["numero"]).isdigit() else 0))
        soma = sum(d["valor"] for d in autorizadas)

        _, mes_nome, ano = self._mes_ano()
        estilos = getSampleStyleSheet()
        tit = ParagraphStyle("tit", parent=estilos["Title"], fontSize=18, spaceAfter=6)
        normal = estilos["Normal"]
        doc = SimpleDocTemplate(path_pdf, pagesize=landscape(A4), leftMargin=14 * mm, rightMargin=14 * mm,
                                topMargin=12 * mm, bottomMargin=12 * mm)
        el = [Paragraph("Relatório de NFCe", tit),
              Paragraph("<b>Emitente:</b> " + (emitente or "-") + " &nbsp;&nbsp; <b>CNPJ:</b> " + (cnpj or "-"),
                        normal),
              Paragraph("<b>Período:</b> " + mes_nome + " - " + str(ano) + " &nbsp;&nbsp; <b>Gerado em:</b> " +
                        datetime.now().strftime("%d/%m/%Y %H:%M"), normal),
              Spacer(1, 6),
              Paragraph("<b>Total de cupons:</b> " + str(len(autorizadas)) +
                        " &nbsp;&nbsp;&nbsp; <b>Valor total:</b> " + self._fmt_moeda(soma) +
                        " &nbsp;&nbsp;&nbsp; <b>Canceladas:</b> " + str(len(canceladas)), normal)]
        if ignorados:
            el.append(Paragraph("<font size=8 color=grey>(" + str(ignorados) +
                                " arquivo(s) ignorado(s) por não serem NFCe autorizada ou cancelada — "
                                "ex.: denegadas, rejeitadas ou outros eventos)</font>", normal))
        el.append(Spacer(1, 10))

        # tabela das autorizadas
        linhas = [["Número", "Emissão", "Chave de acesso", "Valor"]]
        for d in autorizadas:
            linhas.append([d["numero"], self._fmt_emissao(d["emissao"]), d["chave"], self._fmt_moeda(d["valor"])])
        linhas.append(["", "", "TOTAL", self._fmt_moeda(soma)])
        t = LongTable(linhas, repeatRows=1, colWidths=[28 * mm, 40 * mm, 150 * mm, 36 * mm])
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#7C3AED")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("ALIGN", (3, 0), (3, -1), "RIGHT"),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#E5E7EB")),
            ("ROWBACKGROUNDS", (0, 1), (-1, -2), [colors.white, colors.HexColor("#F3F0FF")]),
        ]))
        el.append(t)

        # tabela das canceladas (sem valor)
        if canceladas:
            el.append(Spacer(1, 14))
            el.append(Paragraph("<b>Notas canceladas (" + str(len(canceladas)) + ")</b>", normal))
            el.append(Spacer(1, 4))
            linhas = [["Número", "Cancelamento", "Chave de acesso", "Situação"]]
            for d in canceladas:
                linhas.append([d["numero"], self._fmt_emissao(d["emissao"]), d["chave"], "Cancelada"])
            t2 = LongTable(linhas, repeatRows=1, colWidths=[28 * mm, 40 * mm, 150 * mm, 36 * mm])
            t2.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#A4161A")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#E5E7EB")),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#FBEAEA")]),
            ]))
            el.append(t2)

        doc.build(el)
        return len(autorizadas), soma, len(canceladas)


    # ------------------------------------------------------------------ ZIPs e e-mail
    def criar_zips_em_partes(self, arquivos, temp_dir, mes_nome, ano):
        """Cria os ZIPs do envio, SEPARADOS por tipo: primeiro as AUTORIZADAS, depois as
        CANCELADAS. Cada grupo é dividido em várias partes se o ZIP ultrapassar
        LIMITE_EMAIL_BYTES.

        Retorna uma lista de dicts, na ordem de envio:
            {'caminho', 'nome', 'tipo', 'parte', 'total_partes', 'qtd', 'qtd_tipo'}
        Se um tipo não tem arquivos, nenhum ZIP é criado para ele."""
        saida = []
        for tipo in (AUTORIZADA, CANCELADA):
            grupo = [a for a in arquivos if a.get("tipo") == tipo]
            if not grupo:
                continue
            partes = []          # [(caminho, qtd)]
            zf, caminho, qtd, usados = None, None, 0, set()

            def abrir():
                n = len(partes) + 1
                nome = f"NFCe_{ROTULO_TIPO[tipo]}_{mes_nome}_{ano}_parte{n}.zip"
                cam = os.path.join(temp_dir, nome)
                return zipfile.ZipFile(cam, "w", zipfile.ZIP_DEFLATED), cam

            for a in grupo:
                if zf is None:
                    zf, caminho = abrir()
                    qtd, usados = 0, set()
                nome = os.path.basename(a["caminho"])
                raiz, ext = os.path.splitext(nome)
                k = 1
                while nome in usados:
                    k += 1
                    nome = f"{raiz} ({k}){ext}"
                usados.add(nome)
                zf.write(a["caminho"], nome)
                qtd += 1
                if zf.fp is not None and zf.fp.tell() >= LIMITE_EMAIL_BYTES:
                    zf.close()
                    partes.append((caminho, qtd))
                    zf = None
            if zf is not None:
                zf.close()
                partes.append((caminho, qtd))

            for i, (cam, q) in enumerate(partes, 1):
                saida.append({"caminho": cam, "nome": os.path.basename(cam), "tipo": tipo, "parte": i,
                              "total_partes": len(partes), "qtd": q, "qtd_tipo": len(grupo)})
        return saida

    def enviar_zip_por_email(self, zip_path, zip_name, mes_nome, ano, parte, total_partes,
                             total_arquivos, pdf_path=None, tipo=None):
        """Envia um único ZIP por e-mail.
        Ajusta o assunto e o corpo automaticamente quando há múltiplas partes e
        informa se o ZIP é de notas AUTORIZADAS ou CANCELADAS."""
        cfg = self.config
        rotulo = ROTULO_TIPO.get(tipo, "")
        destinatarios = getattr(self, "_destinatarios_envio", None) or \
            parsear_destinatarios(cfg.get("email_destinatario", ""))
        msg = MIMEMultipart()
        msg["From"] = cfg["email_remetente"]
        msg["To"] = ", ".join(destinatarios)
        assunto = "XMLs NFCe" + (" " + rotulo if rotulo else "") + " - " + mes_nome + "/" + str(ano)
        if total_partes > 1:
            assunto += " (Parte " + str(parte) + " de " + str(total_partes) + ")"
        msg["Subject"] = assunto

        de_que = "das notas " + rotulo.upper() if rotulo else "dos XMLs"
        if total_partes > 1:
            corpo = ("Prezados,\n\nSegue a parte " + str(parte) + " de " + str(total_partes) +
                     " dos XMLs " + de_que + " referentes a " + mes_nome + "/" + str(ano) +
                     ".\nTotal de arquivos no lote completo: " + str(total_arquivos) + "\n\nAtenciosamente.")
        else:
            corpo = ("Prezados,\n\nSeguem os XMLs " + de_que + " referentes a " + mes_nome + "/" + str(ano) +
                     ".\nTotal de arquivos: " + str(total_arquivos) + "\n\nAtenciosamente.")
        if pdf_path and os.path.exists(pdf_path):
            corpo = corpo.replace("\n\nAtenciosamente.",
                                  "\n\nEm anexo também o relatório (PDF) com o total de cupons e valores."
                                  "\n\nAtenciosamente.")
        msg.attach(MIMEText(corpo, "plain", "utf-8"))

        with open(zip_path, "rb") as f:
            anexo = MIMEBase("application", "zip")
            anexo.set_payload(f.read())
        encoders.encode_base64(anexo)
        anexo.add_header("Content-Disposition", 'attachment; filename="' + zip_name + '"')
        msg.attach(anexo)

        if pdf_path and os.path.exists(pdf_path):
            with open(pdf_path, "rb") as f:
                pdf = MIMEBase("application", "pdf")
                pdf.set_payload(f.read())
            encoders.encode_base64(pdf)
            pdf.add_header("Content-Disposition", 'attachment; filename="' + os.path.basename(pdf_path) + '"')
            msg.attach(pdf)

        self._enviar_msg_smtp(msg, destinatarios=destinatarios)

    def _abrir_smtp(self, cfg=None):
        """Abre a conexão com o servidor SMTP e faz o login. Devolve o servidor."""
        cfg = cfg or self.config
        host, porta = cfg["servidor_smtp"], int(cfg["porta_smtp"])
        if porta == 465:
            srv = smtplib.SMTP_SSL(host, porta, timeout=60, context=contexto_ssl())
        else:
            srv = smtplib.SMTP(host, porta, timeout=60)
            srv.ehlo()
            if cfg.get("usar_ssl"):
                srv.starttls(context=contexto_ssl())
                srv.ehlo()
        srv.login(cfg["email_remetente"], cfg["senha_email"])
        return srv

    def _enviar_msg_smtp(self, msg, max_tentativas=3, destinatarios=None):
        """Envia a mensagem via SMTP com reconexão e retry.

        O servidor (Gmail etc.) frequentemente derruba conexões ociosas ou limita
        envios seguidos; por isso cada tentativa abre uma conexão nova."""
        ultimo = None
        for tentativa in range(1, max_tentativas + 1):
            srv = None
            try:
                srv = self._abrir_smtp()
                srv.send_message(msg, to_addrs=destinatarios or None)
                srv.quit()
                return
            except smtplib.SMTPResponseException as e:
                ultimo = e
                transitorio = (getattr(e, "smtp_code", 0) or 0) in (421, 450, 451, 452, 454)
                if not transitorio or tentativa == max_tentativas:
                    raise
            except (smtplib.SMTPServerDisconnected, smtplib.SMTPConnectError, ConnectionError,
                    TimeoutError, OSError) as e:
                ultimo = e
                if tentativa == max_tentativas:
                    raise
            finally:
                try:
                    if srv is not None:
                        srv.close()
                except Exception:
                    pass
            self._prog(getattr(self, "_pct_atual", 0.5),
                       "Conexão instável, tentando novamente (" + str(tentativa) + "/" + str(max_tentativas) + ")...")
            time.sleep(3 * tentativa)
        if ultimo:
            raise ultimo

    def enviar_email_thread(self):
        if not self.arquivos_encontrados:
            messagebox.showwarning("Aviso", "Busque arquivos primeiro")
            return
        self._ler_campos_config()
        destinatarios = parsear_destinatarios(self.config.get("email_destinatario", ""))
        if not destinatarios:
            messagebox.showwarning(
                "E-mail destinatário inválido",
                "O e-mail destinatário está em branco ou com formato inválido.\n\n"
                "Vá na aba 'Configurações' e informe um e-mail válido.")
            return
        if not self.config.get("senha_email"):
            messagebox.showwarning(
                "Senha do e-mail",
                "A senha de aplicativo do e-mail remetente não foi informada.\n\n"
                "Vá em Configurações > Conta de envio e preencha o campo 'Senha de aplicativo'.")
            return
        r = resumo_por_tipo(self.arquivos_encontrados)
        if r[AUTORIZADA] + r[CANCELADA] == 0:
            messagebox.showwarning("Aviso", "Não há notas autorizadas ou canceladas na lista para enviar.")
            return
        self._destinatarios_envio = destinatarios
        try:
            _, mes_nome, ano = self._mes_ano()
        except Exception as e:
            self.mostrar_erro("ENVIO", e, "Erro no envio")
            return
        if not messagebox.askyesno(
                "Confirmar envio",
                "Confira antes de enviar:\n\nPeríodo:  " + mes_nome + " - " + str(ano) +
                "\nArquivos:  " + str(r[AUTORIZADA] + r[CANCELADA]) +
                "  (" + str(r[AUTORIZADA]) + " autorizadas, " + str(r[CANCELADA]) + " canceladas)" +
                "\nPara:  " + ", ".join(destinatarios) + "\n\nDeseja enviar agora?"):
            return
        self.btn_enviar.configure(state="disabled")
        self._dlg = DialogoProgresso(self.root, "Enviando e-mail", getattr(self, "icon_path", None))
        self._dlg.atualizar(0.0, "Preparando envio...")
        threading.Thread(target=self.enviar_email_smtp, daemon=True).start()

    def enviar_email_smtp(self):
        """Orquestra a criação dos ZIPs em partes e o envio de cada um.
        A barra de progresso já foi aberta por enviar_email_thread.
        Envia primeiro os ZIPs das AUTORIZADAS e depois os das CANCELADAS."""
        t0 = time.monotonic()
        temp_dir = None
        enviados = 0
        try:
            _, mes_nome, ano = self._mes_ano()
            mes_seguro = self._nome_arquivo_seguro(mes_nome)
            temp_dir = tempfile.mkdtemp(prefix="nfce_")
            notas = self._notas_para_envio()

            self._prog(0.0, "Compactando arquivos...")
            t = time.monotonic()
            zips = self.criar_zips_em_partes(notas, temp_dir, mes_seguro, ano)
            t_zip = time.monotonic() - t
            if not zips:
                raise ValueError("Nenhuma nota autorizada ou cancelada para enviar.")

            self._prog(0.08, "Gerando relatório PDF...")
            t = time.monotonic()
            pdf_path = os.path.join(temp_dir, "Relatorio_NFCe_" + mes_seguro + " - " + str(ano) + ".pdf")
            try:
                self._construir_pdf_relatorio(pdf_path, atualizar_label=lambda txt: None)
            except ImportError:
                pdf_path = None  # sem reportlab: envia só os ZIPs
            t_pdf = time.monotonic() - t

            total = len(zips)
            t_espera = t_envio = 0.0
            for i, z in enumerate(zips, 1):
                if i > 1:
                    espera = PAUSA_ENTRE_PARTES_SEG + random.uniform(0, JITTER_MAX_SEG)
                    t = time.monotonic()
                    while time.monotonic() - t < espera:
                        restante = int(espera - (time.monotonic() - t)) + 1
                        self._pct_atual = 0.36 + 0.62 * (i - 1) / total
                        self._prog(self._pct_atual, "Aguardando " + str(restante) + "s para evitar congestionamento...")
                        time.sleep(1.0)
                    t_espera += time.monotonic() - t
                self._pct_atual = 0.36 + 0.62 * (i - 1) / total
                self._prog(self._pct_atual,
                           "Enviando e-mail... (parte " + str(i) + " de " + str(total) + ") — " +
                           ROTULO_TIPO[z["tipo"]])
                t = time.monotonic()
                self.enviar_zip_por_email(z["caminho"], z["nome"], mes_nome, ano, z["parte"], z["total_partes"],
                                          z["qtd_tipo"], pdf_path if i == 1 else None, tipo=z["tipo"])
                t_envio += time.monotonic() - t
                enviados += 1

            self._prog(1.0, "Concluído!")
            r = resumo_por_tipo(notas)
            self._registrar_perf(
                "ENVIO OK | arquivos=" + str(len(notas)) + " partes=" + str(total) +
                " | compactar=" + f"{t_zip:.1f}" + "s pdf=" + f"{t_pdf:.1f}" + "s espera=" + f"{t_espera:.1f}" +
                "s envio_smtp=" + f"{t_envio:.1f}" + "s | TOTAL=" + f"{time.monotonic() - t0:.1f}" + "s")
            texto = (str(enviados) + " email(s) enviado(s) com sucesso!\n\nAutorizadas: " + str(r[AUTORIZADA]) +
                     "\nCanceladas: " + str(r[CANCELADA]))
            self.root.after(300, lambda: messagebox.showinfo("Sucesso", texto))
        except Exception as e:
            self.mostrar_erro("ENVIO", e, "Erro no envio")
        finally:
            if temp_dir and os.path.exists(temp_dir):
                shutil.rmtree(temp_dir, ignore_errors=True)
            self.root.after(0, self._fechar_dlg)
            self.root.after(0, lambda: self.label_info.configure(text="Pronto"))
            self.root.after(0, lambda: self.btn_enviar.configure(state="normal"))

    # ------------------------------------------------------------------ configurações
    def salvar_configuracoes(self):
        self._ler_campos_config()
        self.salvar_config()
        messagebox.showinfo("Config", "Salvo com sucesso!")

    def testar_conexao(self):
        self._ler_campos_config()
        cfg = dict(self.config)
        self._dlg = DialogoProgresso(self.root, "Testando conexão", getattr(self, "icon_path", None),
                                     indeterminado=True)
        self._dlg.atualizar(0, "Conectando ao servidor de e-mail...")

        def worker():
            try:
                srv = self._abrir_smtp(cfg)
                srv.quit()
                self.root.after(0, self._fechar_dlg)
                self.root.after(0, lambda: messagebox.showinfo("Sucesso", "Conexão com o servidor de e-mail OK!"))
            except Exception as e:
                self.root.after(0, self._fechar_dlg)
                self.mostrar_erro("CONEXAO", e, "Falha na conexão")
        threading.Thread(target=worker, daemon=True).start()


# ----------------------------------------------------------------------------
# Tratamento global de erros e ponto de entrada
# ----------------------------------------------------------------------------
app = None


def _mostrar_erro_global(app, contexto, exc):
    """Roteia um erro não tratado para a mensagem amigável do app; se até isso
    falhar, mostra uma janela simples."""
    try:
        if app is not None:
            app.mostrar_erro(contexto, exc, "Ops, algo deu errado")
            return
        raise RuntimeError("app indisponível")
    except Exception:
        try:
            msg, cod = classificar_erro(exc)
            messagebox.showerror("Ops, algo deu errado",
                                 msg + "\n\nCódigo do erro: " + contexto + "/" + cod +
                                 "\nMande este código (ou um print) para o suporte.")
        except Exception:
            pass


def _tk_callback_exc(exc, val, tb):
    _mostrar_erro_global(app, "APP", val)


def _thread_exc(args):
    _mostrar_erro_global(app, "THREAD", args.exc_value)


def main():
    global app
    try:
        locale.setlocale(locale.LC_TIME, "pt_BR.UTF-8")
    except Exception:
        try:
            locale.setlocale(locale.LC_TIME, "Portuguese_Brazil.1252")
        except Exception:
            pass
    if not CTK_AVAILABLE:
        raiz = tk.Tk()
        raiz.withdraw()
        messagebox.showerror(
            "Componente faltando",
            "Um componente do programa (CustomTkinter) não está instalado.\n\n"
            "A instalação pode estar incompleta — reinstale o programa.")
        return
    definir_app_id()
    root = ctk.CTk()
    app = ListadorNFCe(root)
    root.report_callback_exception = _tk_callback_exc
    threading.excepthook = _thread_exc
    root.mainloop()


if __name__ == "__main__":
    main()
