"""Price — análise local de rentabilidade e sugestão de preços.

Python 3.10+ | Instalação: python -m pip install pandas numpy openpyxl
Execução: python price.py | Demonstração: python price.py --demo
Tkinter e uma sessão gráfica são necessários para a interface.

Versão para portfólio: não inclui bases reais, credenciais, serviços remotos,
telemetria ou caminhos de uma organização. A demonstração é inteiramente fictícia.
Exportações públicas usam SOMENTE a base fictícia, nunca a base do usuário.
Exportações operacionais e configurações podem conter informações da sua base:
não as publique no Git. Use dados fictícios também em capturas de tela.
Sugestão de .gitignore: .venv/, __pycache__/, *.csv, *.xlsx, *.json, .env*, *.log.

Margem = (preço líquido - custo local) / preço líquido.
Sugestão = maior custo válido do item / (1 - margem) / (1 - imposto).
O imposto inicial é apenas ilustrativo e deve ser ajustado para cada operação.
Não há atualização de ERP, envio de dados ou salvamento automático.
"""
from __future__ import annotations

import argparse
import csv
import io
import json
import math
import os
import queue
import re
import tempfile
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any, Callable

import numpy as np
import pandas as pd

APP_TITLE = "Price | Inteligência de Precificação"
APP_SUBTITLE = "Análise local de rentabilidade • Categorias dinâmicas • Demonstração fictícia"
DEFAULT_TAX_RATE = 0.0925  # Premissa demonstrativa, editável na interface.
RESULT_PREVIEW_ROWS = 2000
MAX_CSV_BYTES = 50 * 1024 * 1024
MAX_CONFIG_BYTES = 2 * 1024 * 1024
MAX_HISTORY = 2000
REQUIRED_COLUMNS = ("COD_ITEM", "DES_CATEGORIA", "CUSTO_MEDIO", "PRECO_OFICINA")
RESULT_COLUMNS = (*REQUIRED_COLUMNS, "DES_ITEM", "NOME_FANTASIA",
                  "RENTABILIDADE_ATUAL_CALCULADA", "MARGEM_MINIMA",
                  "REALIZAR_AJUSTE", "MAIOR_CUSTO_MEDIO_ITEM", "STATUS",
                  "NOVO_PRECO_SUGERIDO", "IMPOSTO_UTILIZADO")
PERCENT_COLUMNS = frozenset({"RENTABILIDADE_ATUAL_CALCULADA", "MARGEM_MINIMA"})
MONEY_COLUMNS = frozenset({"CUSTO_MEDIO", "PRECO_OFICINA", "MAIOR_CUSTO_MEDIO_ITEM",
                           "NOVO_PRECO_SUGERIDO"})
STATUS_TAGS = {"DENTRO DA MARGEM": "inside", "FORA DA MARGEM": "outside",
               "MANTER - SEM AJUSTE": "keep", "SEM PARÂMETRO": "noparam",
               "DADO INVÁLIDO": "invalid"}
TRUE_VALUES = {"1", "TRUE", "SIM", "S", "YES", "Y", "ATIVO", "AJUSTAR"}
FALSE_VALUES = {"0", "FALSE", "NÃO", "NAO", "N", "NO", "INATIVO", "MANTER"}


class ValidationError(ValueError):
    """Mensagem segura, sem valores de células ou caminhos locais."""


def normalize_category(value: Any) -> str:
    return "" if pd.isna(value) else str(value).strip().upper()


def normalized_columns(columns) -> list[str]:
    result = [str(c).strip().upper() for c in columns]
    if not all(result) or len(result) != len(set(result)):
        raise ValidationError("Cabeçalhos vazios ou duplicados. Revise o CSV.")
    return result


def read_csv_flexible(path: str | Path) -> pd.DataFrame:
    """Leitura limitada; preserva identificadores textuais e zeros à esquerda."""
    with open(path, "rb") as stream:
        raw = stream.read(MAX_CSV_BYTES + 1)
    if len(raw) > MAX_CSV_BYTES:
        raise ValidationError("O CSV excede o limite de 50 MiB.")
    text = None
    for encoding in ("utf-8-sig", "cp1252", "latin1"):
        try:
            text = raw.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    if not text or "\x00" in text:
        raise ValidationError("CSV vazio ou codificação não suportada.")
    try:
        dialect = csv.Sniffer().sniff(text[:65536], delimiters=";,\t|")
        reader = csv.reader(io.StringIO(text), dialect)
        header = next(reader)
        columns = normalized_columns(header)
        frame = pd.read_csv(io.StringIO(text), sep=dialect.delimiter,
                            quotechar=dialect.quotechar, doublequote=dialect.doublequote,
                            dtype=str, keep_default_na=False, on_bad_lines="error")
        if not isinstance(frame.index, pd.RangeIndex):
            raise ValidationError("CSV contém linhas com mais campos que o cabeçalho.")
        if len(frame.columns) != len(columns):
            raise ValidationError("Estrutura de cabeçalhos inconsistente.")
        frame.columns = columns
    except (csv.Error, StopIteration, pd.errors.ParserError, pd.errors.EmptyDataError):
        raise ValidationError("CSV malformado. Verifique separadores e aspas.") from None
    return frame


def parse_number(series: pd.Series) -> pd.Series:
    """Aceita decimais BR/US; não transforma texto arbitrário em número.

    Com um único tipo de separador, ele representa casas decimais.
    Para milhar, use ambos os separadores (1.234,56 ou 1,234.56).
    """
    if pd.api.types.is_numeric_dtype(series) and not pd.api.types.is_bool_dtype(series):
        return pd.to_numeric(series, errors="coerce").astype(float).replace([np.inf, -np.inf], np.nan)
    values = series.astype("string").str.strip()
    values = values.str.replace(r"^R\$\s*", "", regex=True).str.strip()
    simple = values.str.fullmatch(r"[+-]?(?:\d+(?:[.,]\d+)?|[.,]\d+)", na=False)
    br = values.str.fullmatch(r"[+-]?\d{1,3}(?:\.\d{3})+,\d+", na=False)
    us = values.str.fullmatch(r"[+-]?\d{1,3}(?:,\d{3})+\.\d+", na=False)
    normalized = values.where(simple | br | us)
    normalized.loc[br] = values.loc[br].str.replace(".", "", regex=False)
    normalized.loc[us] = values.loc[us].str.replace(",", "", regex=False)
    normalized.loc[~us] = normalized.loc[~us].str.replace(",", ".", regex=False)
    return pd.to_numeric(normalized, errors="coerce").astype(float).replace([np.inf, -np.inf], np.nan)


def fraction(value: Any, *, percent: bool = False) -> float:
    """Percentuais da UI são sempre percentuais; parâmetros JSON são frações."""
    if isinstance(value, (bool, np.bool_)):
        raise ValidationError("Informe um percentual numérico entre 0 e menos de 100%.")
    number = parse_number(pd.Series([value])).iloc[0]
    if percent:
        number /= 100
    if not math.isfinite(number) or not 0 <= number < 1:
        raise ValidationError("Informe um percentual numérico entre 0 e menos de 100%.")
    return float(number)


def parse_bool(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    text = str(value).strip().upper()
    if text in TRUE_VALUES:
        return True
    if text in FALSE_VALUES:
        return False
    raise ValidationError("Opção de ajuste inválida. Utilize SIM ou NAO.")


def normalize_params(params: dict) -> dict[str, dict[str, Any]]:
    if not isinstance(params, dict) or len(params) > 10000:
        raise ValidationError("Estrutura de parâmetros inválida.")
    result = {}
    for category, entry in params.items():
        if not isinstance(category, str) or not category.strip() or len(category) > 200:
            raise ValidationError("Nome de categoria ausente ou muito longo.")
        key = normalize_category(category)
        if key in result:
            raise ValidationError("Categorias duplicadas após normalização.")
        if not isinstance(entry, dict):
            entry = {"margem": entry, "ajustar": True}
        margin = entry.get("margem")
        result[key] = {"margem": None if margin is None else fraction(margin),
                       "ajustar": parse_bool(entry.get("ajustar", False))}
    return result


def load_parameters(path: str | Path) -> dict:
    frame = read_csv_flexible(path)
    def column(*options):
        return next((name for name in options if name in frame), None)
    cat = column("DES_CATEGORIA", "CATEGORIA")
    margin = column("MARGEM_MINIMA", "MARGEM", "RENTABILIDADE_MINIMA", "RENT_MIN")
    adjust = column("AJUSTAR", "REALIZAR_AJUSTE", "ATIVO")
    if not cat or not margin or frame.empty:
        raise ValidationError("Parâmetros exigem categoria e margem mínima, com pelo menos uma linha.")
    params = {}
    for row in frame.to_dict("records"):
        key = normalize_category(row[cat])
        if key in params:
            raise ValidationError("Categorias duplicadas no arquivo de parâmetros.")
        raw = str(row[margin]).strip()
        explicit_percent = raw.endswith("%")
        value = parse_number(pd.Series([raw.rstrip("%").strip()])).iloc[0]
        # CSV: 30 ou 30% => 30%; 0,30 => 30%; 1 => 100% (inválido).
        parsed = fraction(value, percent=explicit_percent or value > 1)
        params[key] = {"margem": parsed,
                       "ajustar": True if adjust is None else parse_bool(row[adjust])}
    return normalize_params(params)


def prepare_base(df: pd.DataFrame) -> pd.DataFrame:
    frame = df.copy()
    frame.columns = normalized_columns(frame.columns)
    if any(col not in frame for col in REQUIRED_COLUMNS):
        raise ValidationError("Base exige COD_ITEM, DES_CATEGORIA, CUSTO_MEDIO e PRECO_OFICINA.")
    if frame.empty:
        raise ValidationError("A base não possui registros.")
    # Allowlist: campos extras (CPF, e-mail, contatos etc.) não seguem para resultados.
    frame = frame[[c for c in RESULT_COLUMNS[:6] if c in frame]].copy()
    frame["COD_ITEM"] = frame["COD_ITEM"].fillna("").astype(str).str.strip()
    frame["DES_CATEGORIA"] = frame["DES_CATEGORIA"].map(normalize_category)
    for col in ("CUSTO_MEDIO", "PRECO_OFICINA"):
        frame[col] = parse_number(frame[col])
    return frame


def calculate(df: pd.DataFrame, params: dict, tax_rate: float = DEFAULT_TAX_RATE) -> pd.DataFrame:
    """Função pura. Não altera a base de entrada nem os parâmetros."""
    tax = fraction(tax_rate)
    params = normalize_params(params)
    x = prepare_base(df)
    cost, price = x["CUSTO_MEDIO"], x["PRECO_OFICINA"]
    valid_cost = cost.notna() & cost.ge(0)
    valid = x["COD_ITEM"].ne("") & x["DES_CATEGORIA"].ne("") & valid_cost & price.gt(0)
    net = price * (1 - tax)
    with np.errstate(divide="ignore", invalid="ignore", over="ignore"):
        profitability = (net - cost) / net
    valid &= net.gt(0) & np.isfinite(profitability)
    x["RENTABILIDADE_ATUAL_CALCULADA"] = profitability.where(valid)
    x["MARGEM_MINIMA"] = x["DES_CATEGORIA"].map({k: v["margem"] for k, v in params.items()})
    x["REALIZAR_AJUSTE"] = x["DES_CATEGORIA"].map({k: v["ajustar"] for k, v in params.items()}).eq(True)
    x["MAIOR_CUSTO_MEDIO_ITEM"] = cost.where(valid_cost & x["COD_ITEM"].ne("")).groupby(x["COD_ITEM"], sort=False).transform("max")
    configured = x["MARGEM_MINIMA"].notna()
    compare = valid & configured & x["REALIZAR_AJUSTE"]
    outside = compare & profitability.lt(x["MARGEM_MINIMA"])
    x["STATUS"] = np.select(
        [~valid, valid & configured & ~x["REALIZAR_AJUSTE"], outside, compare],
        ["DADO INVÁLIDO", "MANTER - SEM AJUSTE", "FORA DA MARGEM", "DENTRO DA MARGEM"],
        default="SEM PARÂMETRO")
    x["NOVO_PRECO_SUGERIDO"] = price
    with np.errstate(over="ignore", invalid="ignore", divide="ignore"):
        suggested = x.loc[outside, "MAIOR_CUSTO_MEDIO_ITEM"] / (1 - x.loc[outside, "MARGEM_MINIMA"]) / (1 - tax)
        # Arredonda para cima em centavos para não ficar abaixo da margem-alvo.
        suggested = np.ceil(suggested * 100) / 100
    finite = np.isfinite(suggested)
    x.loc[suggested.index[~finite], "STATUS"] = "DADO INVÁLIDO"
    x.loc[suggested.index[finite], "NOVO_PRECO_SUGERIDO"] = np.maximum(suggested[finite], price.loc[suggested.index[finite]])
    x["IMPOSTO_UTILIZADO"] = tax
    return x


def demo_data() -> tuple[pd.DataFrame, dict]:
    """Dados fixos e fictícios, independentes de qualquer entrada do usuário."""
    rows = [
        ("0001", "FILTROS", 50, 70, "Peça fictícia A", "Loja exemplo A"),
        ("0001", "FILTROS", 55, 75, "Peça fictícia A", "Loja exemplo B"),
        ("0002", "FILTROS", 20, 80, "Peça fictícia B", "Loja exemplo A"),
        ("0003", "ACESSÓRIOS", 35, 40, "Peça fictícia C", "Loja exemplo A"),
        ("0004", "OUTROS", 10, 30, "Peça fictícia D", "Loja exemplo B"),
        ("0005", "FILTROS", -5, 0, "Registro inválido fictício", "Loja exemplo B"),
    ]
    return pd.DataFrame(rows, columns=RESULT_COLUMNS[:6]), {
        "FILTROS": {"margem": .30, "ajustar": True},
        "ACESSÓRIOS": {"margem": .25, "ajustar": False}}


def safe_cell(value: Any) -> Any:
    """Neutraliza fórmulas em células textuais; números continuam numéricos."""
    if not isinstance(value, str):
        return value
    value = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", value)
    if value.lstrip().startswith(("=", "+", "-", "@")) or value.startswith(("\t", "\r", "\n")):
        value = "'" + value
    return value[:32767]


def safe_frame(frame: pd.DataFrame) -> pd.DataFrame:
    result = frame.copy()
    for col in result.columns:
        if pd.api.types.is_object_dtype(result[col]) or pd.api.types.is_string_dtype(result[col]):
            result[col] = result[col].map(safe_cell)
    return result


def atomic_write(path: str | Path, writer: Callable[[Path], None]) -> None:
    """Só substitui o destino depois de concluir a escrita com sucesso."""
    target = Path(path)
    fd, temporary = tempfile.mkstemp(prefix=".price-", suffix=target.suffix, dir=target.parent)
    os.close(fd)
    try:
        writer(Path(temporary))
        os.replace(temporary, target)
    finally:
        Path(temporary).unlink(missing_ok=True)


def history_records(history: Any) -> list[dict]:
    if not isinstance(history, list) or len(history) > MAX_HISTORY:
        raise ValidationError("Histórico de configuração inválido.")
    keys = ("data", "categoria", "margem", "acao")
    result = []
    for entry in history:
        if not isinstance(entry, dict) or any(not isinstance(entry.get(k), str) or len(entry[k]) > 200 for k in keys):
            raise ValidationError("Registro de histórico inválido.")
        result.append({k: entry[k] for k in keys})
    return result


def save_configuration(path: str | Path, params: dict, tax_rate: float, history: list) -> None:
    payload = {"schema_version": 2, "imposto": fraction(tax_rate),
               "parametros": normalize_params(params), "historico": history_records(history)}
    # Não persiste caminhos, base original ou resultados. Nomes de categorias
    # ainda podem ser internos: este JSON é uma configuração operacional local.
    data = json.dumps(payload, ensure_ascii=False, indent=2, allow_nan=False)
    if len(data.encode("utf-8")) > MAX_CONFIG_BYTES:
        raise ValidationError("Configuração excede o limite de 2 MiB.")
    atomic_write(path, lambda temp: temp.write_text(data, encoding="utf-8"))


def load_configuration(path: str | Path) -> tuple[dict, float, list]:
    with open(path, "rb") as stream:
        data = stream.read(MAX_CONFIG_BYTES + 1)
    if len(data) > MAX_CONFIG_BYTES:
        raise ValidationError("Configuração excede o limite de 2 MiB.")
    try:
        payload = json.loads(data.decode("utf-8-sig"))
    except (ValueError, UnicodeError):
        raise ValidationError("Arquivo JSON inválido.") from None
    if not isinstance(payload, dict) or payload.get("schema_version", 1) not in (1, 2):
        raise ValidationError("Versão de configuração não suportada.")
    return (normalize_params(payload.get("parametros", {})),
            fraction(payload.get("imposto", DEFAULT_TAX_RATE)),
            history_records(payload.get("historico", [])))


def export_data(path: str | Path, result: pd.DataFrame, params: dict,
                history: list, *, public_demo: bool = True) -> None:
    """Modo público gera um resultado fictício novo; não anonimiza dados reais."""
    if public_demo:
        base, params = demo_data()
        result = calculate(base, params, DEFAULT_TAX_RATE)
        history = []
    else:
        params = normalize_params(params)
    result = safe_frame(result[[c for c in RESULT_COLUMNS if c in result]])
    parameters = safe_frame(pd.DataFrame([
        {"DES_CATEGORIA": cat, "MARGEM_MINIMA": entry["margem"], "REALIZAR_AJUSTE": entry["ajustar"]}
        for cat, entry in params.items()]))
    history_frame = safe_frame(pd.DataFrame(history_records(history), columns=["data", "categoria", "margem", "acao"]))
    suffix = Path(path).suffix.lower()
    if suffix not in (".xlsx", ".csv"):
        raise ValidationError("Escolha uma extensão .xlsx ou .csv.")
    def write(temp):
        if suffix == ".csv":
            result.to_csv(temp, index=False, sep=";", encoding="utf-8-sig", quoting=csv.QUOTE_ALL)
        else:
            with pd.ExcelWriter(temp, engine="openpyxl") as writer:
                for name, frame in (("Resultado", result), ("Parametros", parameters), ("Historico", history_frame)):
                    frame.to_excel(writer, sheet_name=name, index=False)
                    ws = writer.sheets[name]
                    ws.freeze_panes = "A2"
                    ws.auto_filter.ref = ws.dimensions
                    for cells in ws.iter_cols(min_row=1, max_row=1):
                        ws.column_dimensions[cells[0].column_letter].width = 25
    atomic_write(path, write)


def safe_error(exc: Exception) -> str:
    if isinstance(exc, ValidationError):
        return str(exc)
    if isinstance(exc, PermissionError):
        return "Acesso negado. Verifique permissões e se o arquivo está aberto em outro programa."
    if isinstance(exc, FileNotFoundError):
        return "Arquivo não encontrado ou pasta de destino indisponível."
    if isinstance(exc, ImportError):
        return "Dependência indisponível. Instale pandas, numpy e openpyxl."
    return "Não foi possível concluir a operação. Verifique formato, espaço em disco e permissões."


class PricingApp(tk.Tk):
    BG = "#101114"
    PANEL = "#17191F"
    PANEL_DARK = "#13161C"
    CARD = "#1B1F27"
    CARD_ALT = "#232833"
    LOG_BG = "#0E1015"
    ACCENT = "#E10600"
    TEXT = "#FFFFFF"
    TEXT_SOFT = "#D5D7DC"
    TEXT_MUTED = "#8F96A3"
    BORDER = "#2A2D34"
    SUCCESS = "#22C55E"
    WARNING = "#F59E0B"
    ERROR = "#FF4D4F"

    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1400x900")
        self.minsize(1180, 800)
        self.configure(bg=self.BG)
        self.df = self.result = None
        self.params = {}
        self.history = []
        self._busy = False
        self._closed = False
        self._queue = queue.Queue()
        self._render_id = 0
        self._poll_id = None
        self.cat_var = tk.StringVar()
        self.margin_var = tk.StringVar()
        self.adjust_var = tk.BooleanVar(value=False)
        self.tax_var = tk.StringVar(value=f"{DEFAULT_TAX_RATE * 100:g}")
        self.public_export_var = tk.BooleanVar(value=True)
        self.tax_var.trace_add("write", self._tax_changed)
        self.badge_var = tk.StringVar(value="AGUARDANDO BASE")
        self.status_var = tk.StringVar(value="Pronto para iniciar")
        self.detail_var = tk.StringVar(value="Use a demonstração fictícia para apresentar o projeto.")
        self.file_var = tk.StringVar(value="Nenhuma base carregada")
        self.summary_var = tk.StringVar(value="Nenhuma análise executada.")
        self.kpi_total_var = tk.StringVar(value="0")
        self.kpi_in_var = tk.StringVar(value="0")
        self.kpi_out_var = tk.StringVar(value="0")
        self.kpi_invalid_var = tk.StringVar(value="0")
        self._configure_styles()
        self._build()
        self.protocol("WM_DELETE_WINDOW", self.close)
    def _configure_styles(self):
        style = ttk.Style(self)
        style.theme_use("clam")

        style.configure(
            "Dark.TFrame",
            background=self.BG,
        )

        style.configure(
            "Card.TFrame",
            background=self.CARD,
        )

        style.configure(
            "Panel.TFrame",
            background=self.PANEL,
        )

        style.configure(
            "Dark.TLabel",
            background=self.BG,
            foreground=self.TEXT_SOFT,
            font=("Segoe UI", 10),
        )

        style.configure(
            "Card.TLabel",
            background=self.CARD,
            foreground=self.TEXT_SOFT,
            font=("Segoe UI", 10),
        )

        style.configure(
            "CardTitle.TLabel",
            background=self.CARD,
            foreground=self.TEXT,
            font=("Segoe UI", 11, "bold"),
        )

        style.configure(
            "Accent.TButton",
            background=self.ACCENT,
            foreground=self.TEXT,
            borderwidth=0,
            focusthickness=0,
            padding=(12, 8),
            font=("Segoe UI", 9, "bold"),
        )
        style.map(
            "Accent.TButton",
            background=[("active", "#B80500"), ("pressed", "#930400")],
            foreground=[("disabled", "#AAAAAA"), ("active", self.TEXT)],
        )

        style.configure(
            "Secondary.TButton",
            background=self.BORDER,
            foreground=self.TEXT,
            borderwidth=0,
            focusthickness=0,
            padding=(11, 8),
            font=("Segoe UI", 9, "bold"),
        )
        style.map(
            "Secondary.TButton",
            background=[("active", "#3A3E47"), ("pressed", "#20232A")],
            foreground=[("active", self.TEXT)],
        )

        style.configure(
            "Danger.TButton",
            background="#3A2022",
            foreground="#FFD7D7",
            borderwidth=0,
            focusthickness=0,
            padding=(11, 8),
            font=("Segoe UI", 9, "bold"),
        )
        style.map(
            "Danger.TButton",
            background=[("active", "#5A2629"), ("pressed", "#2B1718")],
        )

        style.configure(
            "Dark.TEntry",
            fieldbackground=self.LOG_BG,
            foreground=self.TEXT,
            bordercolor=self.BORDER,
            insertcolor=self.TEXT,
            padding=6,
        )

        style.configure(
            "Dark.TCombobox",
            fieldbackground=self.LOG_BG,
            background=self.BORDER,
            foreground=self.TEXT,
            arrowcolor=self.TEXT_SOFT,
            bordercolor=self.BORDER,
            padding=5,
        )
        style.map(
            "Dark.TCombobox",
            fieldbackground=[("readonly", self.LOG_BG)],
            foreground=[("readonly", self.TEXT)],
            selectbackground=[("readonly", self.LOG_BG)],
            selectforeground=[("readonly", self.TEXT)],
        )

        style.configure(
            "Dark.TCheckbutton",
            background=self.CARD,
            foreground=self.TEXT_SOFT,
            font=("Segoe UI", 9),
        )
        style.map(
            "Dark.TCheckbutton",
            background=[("active", self.CARD)],
            foreground=[("active", self.TEXT)],
        )

        style.configure(
            "Dark.Treeview",
            background=self.LOG_BG,
            fieldbackground=self.LOG_BG,
            foreground="#D8DCE5",
            rowheight=26,
            bordercolor=self.BORDER,
            borderwidth=0,
            font=("Segoe UI", 9),
        )
        style.map(
            "Dark.Treeview",
            background=[("selected", "#3A1B1A")],
            foreground=[("selected", self.TEXT)],
        )

        style.configure(
            "Dark.Treeview.Heading",
            background=self.BORDER,
            foreground=self.TEXT,
            relief="flat",
            font=("Segoe UI", 9, "bold"),
        )
        style.map(
            "Dark.Treeview.Heading",
            background=[("active", "#353942")],
        )

        style.configure(
            "Vertical.TScrollbar",
            troughcolor=self.PANEL_DARK,
            background=self.BORDER,
            bordercolor=self.PANEL_DARK,
            arrowcolor=self.TEXT_SOFT,
        )
        style.configure(
            "Horizontal.TScrollbar",
            troughcolor=self.PANEL_DARK,
            background=self.BORDER,
            bordercolor=self.PANEL_DARK,
            arrowcolor=self.TEXT_SOFT,
        )

    def _build(self):
        # Cabeçalho
        header = tk.Frame(self, bg=self.BG)
        header.pack(fill="x", padx=18, pady=(16, 8))

        self.badge_label = tk.Label(
            header,
            textvariable=self.badge_var,
            bg=self.WARNING,
            fg=self.TEXT,
            font=("Segoe UI", 9, "bold"),
            padx=10,
            pady=4,
        )
        self.badge_label.pack(anchor="w", pady=(0, 8))

        tk.Label(
            header,
            text=APP_TITLE,
            bg=self.BG,
            fg=self.TEXT,
            font=("Segoe UI", 25, "bold"),
        ).pack(anchor="w")

        tk.Label(
            header,
            text=APP_SUBTITLE,
            bg=self.BG,
            fg=self.TEXT_SOFT,
            font=("Segoe UI", 12),
        ).pack(anchor="w", pady=(4, 0))

        # Barra de ações
        actions = tk.Frame(self, bg=self.BG)
        actions.pack(fill="x", padx=18, pady=(4, 8))

        ttk.Button(
            actions, text="1. Carregar Base CSV",
            command=self.load_base, style="Accent.TButton"
        ).pack(side="left")

        ttk.Button(
            actions, text="2. Carregar Parâmetros",
            command=self.load_param_file, style="Secondary.TButton"
        ).pack(side="left", padx=(6, 0))

        ttk.Button(
            actions, text="Salvar Configuração",
            command=self.save_config, style="Secondary.TButton"
        ).pack(side="left", padx=(6, 0))

        ttk.Button(
            actions, text="Carregar Configuração",
            command=self.load_config, style="Secondary.TButton"
        ).pack(side="left", padx=(6, 0))

        ttk.Button(
            actions, text="Executar / Atualizar Análise",
            command=self.run_analysis, style="Accent.TButton"
        ).pack(side="left", padx=(14, 0))

        ttk.Button(
            actions, text="Exportar Resultado",
            command=self.export_result, style="Secondary.TButton"
        ).pack(side="left", padx=(6, 0))

        options = tk.Frame(self, bg=self.BG)
        options.pack(fill="x", padx=18, pady=(0, 8))
        ttk.Button(options, text="Carregar demonstração fictícia", command=self.load_demo,
                   style="Secondary.TButton").pack(side="left")
        ttk.Label(options, text="  Imposto (%)", style="Dark.TLabel").pack(side="left")
        ttk.Entry(options, textvariable=self.tax_var, width=8,
                  style="Dark.TEntry").pack(side="left", padx=6)
        ttk.Checkbutton(options, text="Exportar somente demonstração fictícia",
                        variable=self.public_export_var,
                        style="Dark.TCheckbutton").pack(side="left", padx=10)

        # Corpo
        body = tk.Frame(self, bg=self.BG)
        body.pack(fill="both", expand=True, padx=18, pady=(0, 10))

        left = tk.Frame(body, bg=self.PANEL, width=405)
        left.pack(side="left", fill="y")
        left.pack_propagate(False)

        right = tk.Frame(body, bg=self.PANEL_DARK)
        right.pack(side="right", fill="both", expand=True, padx=(14, 0))

        # Painel esquerdo - parâmetros
        tk.Label(
            left,
            text="Configuração operacional",
            bg=self.PANEL,
            fg=self.TEXT,
            font=("Segoe UI", 14, "bold"),
        ).pack(anchor="w", padx=16, pady=(16, 4))

        tk.Label(
            left,
            text="Defina a rentabilidade mínima e indique quais categorias deverão sofrer ajuste.",
            bg=self.PANEL,
            fg=self.TEXT_MUTED,
            font=("Segoe UI", 9),
            wraplength=360,
            justify="left",
        ).pack(anchor="w", padx=16, pady=(0, 12))

        form_card = tk.Frame(left, bg=self.CARD)
        form_card.pack(fill="x", padx=16, pady=(0, 10))

        tk.Label(
            form_card, text="Categoria",
            bg=self.CARD, fg=self.TEXT_MUTED,
            font=("Segoe UI", 9)
        ).grid(row=0, column=0, sticky="w", padx=12, pady=(12, 4))

        category_combo = ttk.Combobox(
            form_card,
            textvariable=self.cat_var,
            values=[],
            state="readonly",
            style="Dark.TCombobox",
            width=31,
        )
        self.category_combo = category_combo
        category_combo.grid(row=1, column=0, columnspan=2, sticky="ew", padx=12)
        category_combo.bind("<<ComboboxSelected>>", self._load_selected_category)

        tk.Label(
            form_card, text="Margem mínima (%)",
            bg=self.CARD, fg=self.TEXT_MUTED,
            font=("Segoe UI", 9)
        ).grid(row=2, column=0, sticky="w", padx=12, pady=(12, 4))

        ttk.Entry(
            form_card,
            textvariable=self.margin_var,
            style="Dark.TEntry",
            width=14,
        ).grid(row=3, column=0, sticky="ew", padx=(12, 6))

        ttk.Checkbutton(
            form_card,
            text="Realizar ajuste",
            variable=self.adjust_var,
            style="Dark.TCheckbutton",
        ).grid(row=3, column=1, sticky="w", padx=(6, 12))

        ttk.Button(
            form_card,
            text="Adicionar / Atualizar ajuste",
            command=self.apply_parameter,
            style="Accent.TButton",
        ).grid(row=4, column=0, columnspan=2, sticky="ew", padx=12, pady=(12, 6))

        ttk.Button(
            form_card,
            text="Remover ajuste da categoria",
            command=self.remove_parameter,
            style="Danger.TButton",
        ).grid(row=5, column=0, columnspan=2, sticky="ew", padx=12, pady=(0, 12))

        form_card.grid_columnconfigure(0, weight=1)
        form_card.grid_columnconfigure(1, weight=1)

        # Grade de parâmetros
        param_card = tk.Frame(left, bg=self.CARD)
        param_card.pack(fill="both", expand=True, padx=16, pady=(0, 10))

        tk.Label(
            param_card,
            text="Parâmetros cadastrados",
            bg=self.CARD,
            fg=self.TEXT,
            font=("Segoe UI", 11, "bold"),
        ).pack(anchor="w", padx=12, pady=(12, 8))

        param_grid = tk.Frame(param_card, bg=self.CARD)
        param_grid.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        self.param_tree = ttk.Treeview(
            param_grid,
            columns=("categoria", "margem", "acao"),
            show="headings",
            height=9,
            style="Dark.Treeview",
        )
        self.param_tree.heading("categoria", text="Categoria")
        self.param_tree.heading("margem", text="Margem")
        self.param_tree.heading("acao", text="Ação")
        self.param_tree.column("categoria", width=155)
        self.param_tree.column("margem", width=75, anchor="center")
        self.param_tree.column("acao", width=125, anchor="center")
        self.param_tree.pack(side="left", fill="both", expand=True)
        self.param_tree.bind("<<TreeviewSelect>>", self._select_parameter_row)

        sb = ttk.Scrollbar(
            param_grid, orient="vertical",
            command=self.param_tree.yview
        )
        sb.pack(side="right", fill="y")
        self.param_tree.configure(yscrollcommand=sb.set)
        self.param_tree.tag_configure("adjust", foreground="#FFD0CE")
        self.param_tree.tag_configure("keep", foreground="#F5C96B")
        self.param_tree.tag_configure("empty", foreground=self.TEXT_MUTED)

        # Resumo inferior no painel esquerdo
        info_box = tk.Frame(left, bg=self.CARD_ALT)
        info_box.pack(fill="x", padx=16, pady=(0, 16))

        tk.Label(
            info_box,
            text="Base atualmente carregada",
            bg=self.CARD_ALT,
            fg=self.TEXT_MUTED,
            font=("Segoe UI", 9),
        ).pack(anchor="w", padx=12, pady=(10, 2))

        tk.Label(
            info_box,
            textvariable=self.file_var,
            bg=self.CARD_ALT,
            fg=self.TEXT,
            font=("Segoe UI", 10, "bold"),
            wraplength=345,
            justify="left",
        ).pack(anchor="w", padx=12, pady=(0, 10))

        # Painel direito - status
        status_card = tk.Frame(right, bg=self.CARD)
        status_card.pack(fill="x", padx=16, pady=(16, 10))

        tk.Label(
            status_card,
            text="Status da análise",
            bg=self.CARD,
            fg=self.TEXT_MUTED,
            font=("Segoe UI", 9),
        ).pack(anchor="w", padx=14, pady=(12, 2))

        tk.Label(
            status_card,
            textvariable=self.status_var,
            bg=self.CARD,
            fg=self.TEXT,
            font=("Segoe UI", 16, "bold"),
        ).pack(anchor="w", padx=14)

        self.detail_label = tk.Label(
            status_card,
            textvariable=self.detail_var,
            bg=self.CARD,
            fg=self.ACCENT,
            font=("Segoe UI", 10, "bold"),
            wraplength=850,
            justify="left",
        )
        self.detail_label.pack(anchor="w", padx=14, pady=(4, 12))

        # Indicadores rápidos
        metrics = tk.Frame(right, bg=self.PANEL_DARK)
        metrics.pack(fill="x", padx=16, pady=(0, 10))

        kpis = [
            ("TOTAL ANALISADO", self.kpi_total_var, self.TEXT),
            ("DENTRO DA MARGEM", self.kpi_in_var, self.SUCCESS),
            ("FORA DA MARGEM", self.kpi_out_var, self.ERROR),
            ("DADOS INVÁLIDOS", self.kpi_invalid_var, self.WARNING),
        ]

        for idx, (label_text, value_var, value_color) in enumerate(kpis):
            card = tk.Frame(metrics, bg=self.CARD)
            card.grid(
                row=0,
                column=idx,
                sticky="nsew",
                padx=(0 if idx == 0 else 5, 0 if idx == len(kpis) - 1 else 5),
            )

            tk.Label(
                card,
                text=label_text,
                bg=self.CARD,
                fg=self.TEXT_MUTED,
                font=("Segoe UI", 8, "bold"),
            ).pack(anchor="w", padx=12, pady=(10, 2))

            tk.Label(
                card,
                textvariable=value_var,
                bg=self.CARD,
                fg=value_color,
                font=("Segoe UI", 18, "bold"),
            ).pack(anchor="w", padx=12, pady=(0, 10))

            metrics.grid_columnconfigure(idx, weight=1)

        # Histórico
        hist_card = tk.Frame(right, bg=self.CARD)
        hist_card.pack(fill="x", padx=16, pady=(0, 10))

        tk.Label(
            hist_card,
            text="Histórico de ajustes realizados",
            bg=self.CARD,
            fg=self.TEXT,
            font=("Segoe UI", 11, "bold"),
        ).pack(anchor="w", padx=14, pady=(12, 8))

        self.history_tree = ttk.Treeview(
            hist_card,
            columns=("data", "categoria", "margem", "acao"),
            show="headings",
            height=4,
            style="Dark.Treeview",
        )
        self.history_tree.heading("data", text="Data/Hora")
        self.history_tree.heading("categoria", text="Categoria")
        self.history_tree.heading("margem", text="Margem")
        self.history_tree.heading("acao", text="Ação")
        self.history_tree.column("data", width=150)
        self.history_tree.column("categoria", width=240)
        self.history_tree.column("margem", width=95, anchor="center")
        self.history_tree.column("acao", width=180, anchor="center")
        self.history_tree.pack(fill="x", padx=14, pady=(0, 12))

        # Resultado
        result_card = tk.Frame(right, bg=self.CARD)
        result_card.pack(fill="both", expand=True, padx=16, pady=(0, 10))

        top_result = tk.Frame(result_card, bg=self.CARD)
        top_result.pack(fill="x", padx=14, pady=(12, 8))

        tk.Label(
            top_result,
            text="Resultado da análise",
            bg=self.CARD,
            fg=self.TEXT,
            font=("Segoe UI", 11, "bold"),
        ).pack(side="left")

        tk.Label(
            top_result,
            text=f"Prévia de até {RESULT_PREVIEW_ROWS:,} linhas".replace(",", "."),
            bg=self.CARD,
            fg=self.TEXT_MUTED,
            font=("Segoe UI", 9),
        ).pack(side="right")

        table_frame = tk.Frame(result_card, bg=self.CARD)
        table_frame.pack(fill="both", expand=True, padx=14, pady=(0, 6))

        columns = (
            "COD_ITEM",
            "DES_ITEM",
            "NOME_FANTASIA",
            "DES_CATEGORIA",
            "CUSTO_MEDIO",
            "PRECO_OFICINA",
            "RENTABILIDADE_ATUAL_CALCULADA",
            "MARGEM_MINIMA",
            "MAIOR_CUSTO_MEDIO_ITEM",
            "STATUS",
            "NOVO_PRECO_SUGERIDO",
        )

        self.result_tree = ttk.Treeview(
            table_frame,
            columns=columns,
            show="headings",
            style="Dark.Treeview",
        )

        headings = {
            "COD_ITEM": "Item",
            "DES_ITEM": "Descrição",
            "NOME_FANTASIA": "Loja",
            "DES_CATEGORIA": "Categoria",
            "CUSTO_MEDIO": "Custo Médio",
            "PRECO_OFICINA": "Preço Atual",
            "RENTABILIDADE_ATUAL_CALCULADA": "Rent. Atual",
            "MARGEM_MINIMA": "Margem Mín.",
            "MAIOR_CUSTO_MEDIO_ITEM": "Maior Custo Item",
            "STATUS": "Status",
            "NOVO_PRECO_SUGERIDO": "Novo Preço Sugerido",
        }

        widths = {
            "COD_ITEM": 85,
            "DES_ITEM": 220,
            "NOME_FANTASIA": 145,
            "DES_CATEGORIA": 175,
            "CUSTO_MEDIO": 105,
            "PRECO_OFICINA": 105,
            "RENTABILIDADE_ATUAL_CALCULADA": 105,
            "MARGEM_MINIMA": 100,
            "MAIOR_CUSTO_MEDIO_ITEM": 125,
            "STATUS": 170,
            "NOVO_PRECO_SUGERIDO": 130,
        }

        for c in columns:
            self.result_tree.heading(c, text=headings[c])
            self.result_tree.column(c, width=widths[c], anchor="center")

        self.result_tree.pack(side="left", fill="both", expand=True)

        yscroll = ttk.Scrollbar(
            table_frame, orient="vertical", command=self.result_tree.yview
        )
        yscroll.pack(side="right", fill="y")

        xscroll = ttk.Scrollbar(
            result_card, orient="horizontal", command=self.result_tree.xview
        )
        xscroll.pack(fill="x", padx=14, pady=(0, 8))

        self.result_tree.configure(
            yscrollcommand=yscroll.set,
            xscrollcommand=xscroll.set,
        )
        self.result_tree.tag_configure("inside", foreground="#9EE8B0")
        self.result_tree.tag_configure("outside", foreground="#FF9A9A")
        self.result_tree.tag_configure("keep", foreground="#F5C96B")
        self.result_tree.tag_configure("noparam", foreground=self.TEXT_MUTED)
        self.result_tree.tag_configure("invalid", foreground="#FFC46B")

        # Rodapé/resumo
        footer = tk.Frame(self, bg=self.BG)
        footer.pack(fill="x", padx=18, pady=(0, 14))

        tk.Label(
            footer,
            textvariable=self.summary_var,
            bg=self.BG,
            fg=self.TEXT_MUTED,
            font=("Consolas", 9),
        ).pack(side="left")

        ttk.Button(
            footer,
            text="Fechar",
            command=self.close,
            style="Secondary.TButton"
        ).pack(side="right")

    def _set_badge(self, text, color):
        self.badge_var.set(text)
        self.badge_label.config(bg=color)

    def _set_status(self, title, detail="", badge=None, badge_color=None):
        self.status_var.set(title)
        self.detail_var.set(detail)
        if badge:
            self._set_badge(badge, badge_color or self.ACCENT)

    def _controls(self, widget=None):
        for child in (widget or self).winfo_children():
            if isinstance(child, (ttk.Button, ttk.Entry, ttk.Combobox, ttk.Checkbutton)):
                yield child
            yield from self._controls(child)

    def _set_busy(self, busy):
        self._busy = busy
        for widget in self._controls():
            widget.state(["disabled"] if busy else ["!disabled"])
        self.configure(cursor="watch" if busy else "")

    def _submit(self, label, work, completed):
        """Worker não acessa Tk. A fila é consultada exclusivamente pela UI."""
        if self._busy or self._closed:
            return
        self._set_busy(True)
        self._set_status(label, "Processando em segundo plano…", "PROCESSANDO", self.WARNING)
        def worker():
            try:
                outcome = (True, work())
            except Exception as exc:
                outcome = (False, safe_error(exc))
            self._queue.put(outcome)
        threading.Thread(target=worker, daemon=True, name="price-worker").start()
        def poll():
            if self._closed:
                return
            try:
                ok, value = self._queue.get_nowait()
            except queue.Empty:
                self._poll_id = self.after(60, poll)
                return
            self._poll_id = None
            self._set_busy(False)
            if ok:
                try:
                    completed(value)
                except Exception as exc:
                    self._show_error(safe_error(exc))
            else:
                self._show_error(value)
        self._poll_id = self.after(60, poll)

    def _show_error(self, message):
        self._set_status("Operação não concluída", message, "ERRO", self.ERROR)
        messagebox.showerror("Price", message, parent=self)

    def report_callback_exception(self, exc_type, exc_value, traceback):
        # Tk normalmente imprime o traceback completo; não expõe dados locais.
        self._show_error(safe_error(exc_value))

    def close(self):
        if self._busy:
            # Evita interromper gravação atômica em andamento.
            messagebox.showinfo("Processamento em andamento", "Aguarde a operação terminar para fechar.", parent=self)
            return
        self._closed = True
        self._render_id += 1
        if self._poll_id:
            self.after_cancel(self._poll_id)
        self.destroy()

    def _tax_changed(self, *_args):
        if hasattr(self, "result_tree"):
            self._invalidate_result()

    def _invalidate_result(self):
        self.result = None
        self._render_id += 1
        self.result_tree.delete(*self.result_tree.get_children())
        for var in (self.kpi_total_var, self.kpi_in_var, self.kpi_out_var, self.kpi_invalid_var):
            var.set("0")
        self.summary_var.set("Execute a análise para atualizar os resultados.")

    def _categories(self):
        names = set(self.params)
        if self.df is not None:
            names.update(self.df["DES_CATEGORIA"].dropna().unique())
        return sorted(name for name in names if name)

    def _refresh_parameter_grid(self):
        self.param_tree.delete(*self.param_tree.get_children())
        categories = self._categories()
        self.category_combo.configure(values=categories)
        if self.cat_var.get() not in categories:
            self.cat_var.set(categories[0] if categories else "")
        for cat in categories:
            data = self.params.get(cat, {})
            margin = data.get("margem")
            adjust = data.get("ajustar", False)
            action = "NÃO CONFIGURADO" if margin is None else "AJUSTAR" if adjust else "MANTER — SEM AJUSTE"
            tag = "empty" if margin is None else "adjust" if adjust else "keep"
            self.param_tree.insert("", "end", values=(cat, "" if margin is None else f"{margin:.2%}", action), tags=(tag,))
        self._load_selected_category()

    def _load_selected_category(self, _event=None):
        data = self.params.get(self.cat_var.get(), {})
        margin = data.get("margem")
        self.margin_var.set("" if margin is None else f"{margin * 100:g}")
        self.adjust_var.set(bool(data.get("ajustar", False)))

    def _select_parameter_row(self, _event=None):
        selected = self.param_tree.selection()
        if selected:
            self.cat_var.set(self.param_tree.item(selected[0], "values")[0])
            self._load_selected_category()

    def _refresh_history(self):
        self.history_tree.delete(*self.history_tree.get_children())
        for entry in self.history[-200:]:
            self.history_tree.insert("", "end", values=tuple(entry[k] for k in ("data", "categoria", "margem", "acao")))

    def _record(self, category, margin, action):
        self.history.append({"data": datetime.now().isoformat(timespec="seconds"),
                             "categoria": category,
                             "margem": "-" if margin is None else f"{margin:.2%}", "acao": action})
        self.history = self.history[-MAX_HISTORY:]

    def apply_parameter(self):
        if self._busy:
            return
        category = self.cat_var.get()
        try:
            if not category:
                raise ValidationError("Carregue uma base ou um arquivo de parâmetros primeiro.")
            margin = fraction(self.margin_var.get().strip().rstrip("%"), percent=True)
            params = dict(self.params)
            params[category] = {"margem": margin, "ajustar": self.adjust_var.get()}
            self.params = normalize_params(params)
        except ValidationError as exc:
            self._show_error(str(exc))
            return
        self._record(category, margin, "AJUSTAR" if self.adjust_var.get() else "MANTER — SEM AJUSTE")
        self._refresh_parameter_grid()
        self._refresh_history()
        self._invalidate_result()
        if self.df is not None:
            self.run_analysis(silent=True)

    def remove_parameter(self):
        if self._busy:
            return
        category = self.cat_var.get()
        if category not in self.params:
            return
        del self.params[category]
        self._record(category, None, "REMOVIDO / SEM PARÂMETRO")
        self._refresh_parameter_grid()
        self._refresh_history()
        self._invalidate_result()
        if self.df is not None:
            self.run_analysis(silent=True)

    def load_base(self):
        if self._busy:
            return
        path = filedialog.askopenfilename(filetypes=[("CSV", "*.csv")], parent=self)
        if not path:
            return
        def done(frame):
            self.df = frame
            self._invalidate_result()
            self.file_var.set(f"Base local • {len(frame):,} registros")
            self._refresh_parameter_grid()
            self.run_analysis(silent=True)
        self._submit("Carregando base", lambda: prepare_base(read_csv_flexible(path)), done)

    def load_demo(self):
        if self._busy:
            return
        self.df, self.params = demo_data()
        self.history = []
        self.tax_var.set(f"{DEFAULT_TAX_RATE * 100:g}")
        self.public_export_var.set(True)
        self._invalidate_result()
        self.file_var.set("DEMONSTRAÇÃO • Dados inteiramente fictícios")
        self._refresh_parameter_grid()
        self._refresh_history()
        self.run_analysis(silent=True)

    def load_param_file(self):
        if self._busy:
            return
        path = filedialog.askopenfilename(filetypes=[("CSV", "*.csv")], parent=self)
        if not path:
            return
        def done(params):
            merged = normalize_params({**self.params, **params})
            self.params = merged
            for cat, entry in params.items():
                self._record(cat, entry["margem"], "IMPORTADO / AJUSTAR" if entry["ajustar"] else "IMPORTADO / MANTER")
            self._refresh_parameter_grid()
            self._refresh_history()
            self._invalidate_result()
            if self.df is not None:
                self.run_analysis(silent=True)
            else:
                self._set_status("Parâmetros carregados", "Carregue a base para analisar.", "PRONTO", self.SUCCESS)
        self._submit("Importando parâmetros", lambda: load_parameters(path), done)

    def run_analysis(self, silent=False):
        if self._busy:
            return
        if self.df is None:
            if not silent:
                self._show_error("Carregue uma base ou a demonstração primeiro.")
            return
        try:
            tax = fraction(self.tax_var.get(), percent=True)
            params = normalize_params(self.params)
        except ValidationError as exc:
            self._show_error(str(exc))
            return
        base = self.df  # Bloqueio dos controles garante snapshot estável durante o job.
        self._invalidate_result()
        def done(result):
            self.result = result
            counts = result["STATUS"].value_counts()
            self.kpi_total_var.set(f"{len(result):,}".replace(",", "."))
            for var, key in ((self.kpi_in_var, "DENTRO DA MARGEM"), (self.kpi_out_var, "FORA DA MARGEM"), (self.kpi_invalid_var, "DADO INVÁLIDO")):
                var.set(str(counts.get(key, 0)))
            self.summary_var.set(f"{len(result):,} registros • Prévia: {min(len(result), RESULT_PREVIEW_ROWS):,} • Sem parâmetro: {counts.get('SEM PARÂMETRO', 0)}")
            self._set_status("Análise concluída", "A prévia é limitada; a exportação operacional contém todos os registros.", "CONCLUÍDO", self.SUCCESS)
            self._refresh_result()
        self._submit("Analisando rentabilidade", lambda: calculate(base, params, tax), done)

    def _refresh_result(self):
        self._render_id += 1
        token = self._render_id
        self.result_tree.delete(*self.result_tree.get_children())
        columns = tuple(self.result_tree["columns"])
        preview = self.result.head(RESULT_PREVIEW_ROWS).reindex(columns=columns)
        rows = iter(preview.itertuples(index=False, name=None))
        status_index = columns.index("STATUS")
        def batch():
            if self._closed or token != self._render_id:
                return
            for _ in range(100):
                row = next(rows, None)
                if row is None:
                    return
                values = []
                for column, value in zip(columns, row):
                    if pd.isna(value):
                        value = ""
                    elif column in PERCENT_COLUMNS:
                        value = f"{value:.2%}"
                    elif column in MONEY_COLUMNS:
                        value = self._money(value)
                    values.append(value)
                self.result_tree.insert("", "end", values=values, tags=(STATUS_TAGS.get(row[status_index], ""),))
            self.after(1, batch)
        batch()

    @staticmethod
    def _money(value):
        return f"R$ {float(value):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    def save_config(self):
        if self._busy:
            return
        try:
            params = normalize_params(self.params)
            tax = fraction(self.tax_var.get(), percent=True)
            history = history_records(self.history)
        except ValidationError as exc:
            self._show_error(str(exc))
            return
        path = filedialog.asksaveasfilename(title="Salvar configuração local (pode conter categorias internas)",
                                          defaultextension=".json", filetypes=[("JSON", "*.json")], parent=self)
        if path:
            self._submit("Salvando configuração", lambda: save_configuration(path, params, tax, history),
                         lambda _: self._set_status("Configuração salva", "Arquivo local com parâmetros e histórico; sem caminhos ou base de origem.", "SALVO", self.SUCCESS))

    def load_config(self):
        if self._busy:
            return
        path = filedialog.askopenfilename(filetypes=[("JSON", "*.json")], parent=self)
        if not path:
            return
        def done(values):
            self.params, tax, self.history = values
            self.tax_var.set(f"{tax * 100:g}")
            self._refresh_parameter_grid()
            self._refresh_history()
            self._invalidate_result()
            if self.df is not None:
                self.run_analysis(silent=True)
            else:
                self._set_status("Configuração restaurada", "Carregue a base para analisar.", "PRONTO", self.SUCCESS)
        self._submit("Carregando configuração", lambda: load_configuration(path), done)

    def export_result(self):
        if self._busy:
            return
        demo = self.public_export_var.get()
        if not demo and self.result is None:
            self._show_error("Execute uma análise atualizada antes de exportar dados operacionais.")
            return
        if not demo and not messagebox.askyesno("Exportação operacional", "Este arquivo terá códigos, categorias, lojas e valores reais da base carregada.\nDestine-o ao uso interno; ele não é anonimizado.\n\nContinuar?", parent=self):
            return
        path = filedialog.asksaveasfilename(
            title="Exportar demonstração fictícia" if demo else "Exportar resultado operacional",
            initialfile="price_demonstracao.xlsx" if demo else "price_resultado_local.xlsx",
            defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv")], parent=self)
        if not path:
            return
        result, params, history = self.result, self.params, self.history
        self._submit("Exportando resultado",
                     lambda: export_data(path, result, params, history, public_demo=demo),
                     lambda _: self._set_status("Exportação concluída", "Arquivo gerado somente com a demonstração fictícia padrão." if demo else "Arquivo operacional gerado para uso interno.", "EXPORTADO", self.SUCCESS))


def main():
    parser = argparse.ArgumentParser(description="Price — análise local de precificação")
    parser.add_argument("--demo", action="store_true", help="inicia com dados inteiramente fictícios")
    args = parser.parse_args()
    try:
        app = PricingApp()
    except tk.TclError:
        parser.exit(1, "Interface gráfica indisponível. Execute em uma sessão desktop com Tkinter.\n")
    if args.demo:
        app.after(100, app.load_demo)
    app.mainloop()


if __name__ == "__main__":
    main()
