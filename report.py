import logging
import os
import re
from datetime import datetime
from pathlib import Path
from tkinter import Tk, filedialog, messagebox

import pandas as pd

from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# ============================================================
# CONFIGURAZIONE
# ============================================================

APP_NAME = "Smart Sales Report"
VERSION = "4.0"
LOG_FILE = Path(__file__).resolve().parent / "smart_sales_report.log"

logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    encoding="utf-8",
)

logger = logging.getLogger(__name__)

COLORE_PRINCIPALE = "1F4E78"
COLORE_SECONDARIO = "4472C4"
COLORE_CHIARO = "D9EAF7"
COLORE_VERDE = "E2F0D9"
COLORE_BIANCO = "FFFFFF"


# ============================================================
# INTERFACCIA
# ============================================================

def crea_root():
    root = Tk()
    root.withdraw()
    return root


def scegli_file():

    root = crea_root()

    percorso = filedialog.askopenfilename(
        title="Seleziona il file del cliente",
        filetypes=[
            ("File supportati", "*.csv *.xlsx"),
            ("File CSV", "*.csv"),
            ("File Excel", "*.xlsx"),
        ],
    )

    root.destroy()

    if not percorso:
        return None

    return Path(percorso)


# ============================================================
# LETTURA FILE
# ============================================================

def leggi_csv(percorso):

    codifiche = [
        "utf-8-sig",
        "utf-8",
        "cp1252",
        "latin-1",
    ]

    ultimo_errore = None

    for codifica in codifiche:

        try:

            return pd.read_csv(
                percorso,
                encoding=codifica,
                sep=None,
                engine="python",
            )

        except Exception as errore:

            ultimo_errore = errore

    raise ValueError(
        f"Impossibile leggere il file CSV: {ultimo_errore}"
    )


def carica_dati(percorso):

    estensione = percorso.suffix.lower()

    if estensione == ".csv":
        return leggi_csv(percorso)

    if estensione == ".xlsx":
        return pd.read_excel(percorso)

    raise ValueError(
        f"Formato {estensione} non supportato."
    )


# ============================================================
# NORMALIZZAZIONE TESTO
# ============================================================

def normalizza_testo(testo):

    testo = str(testo).strip().lower()

    sostituzioni = {
        "à": "a",
        "è": "e",
        "é": "e",
        "ì": "i",
        "ò": "o",
        "ù": "u",
    }

    for carattere, sostituzione in sostituzioni.items():
        testo = testo.replace(
            carattere,
            sostituzione
        )

    testo = re.sub(
        r"[^a-z0-9]+",
        "_",
        testo
    )

    return testo.strip("_")


def trova_colonna(df, possibili_nomi):

    colonne = {
        normalizza_testo(colonna): colonna
        for colonna in df.columns
    }

    for nome in possibili_nomi:

        nome = normalizza_testo(nome)

        if nome in colonne:
            return colonne[nome]

    return None


# ============================================================
# PULIZIA DATI
# ============================================================

def pulisci_dati(df):

    df = df.copy()

    righe_originali = len(df)

    # Elimina righe completamente vuote
    df.dropna(
        how="all",
        inplace=True
    )

    # Elimina colonne completamente vuote
    df.dropna(
        axis=1,
        how="all",
        inplace=True
    )

    # Sistema i nomi delle colonne
    df.columns = [
        str(colonna).strip()
        for colonna in df.columns
    ]

    # Pulisce le celle di testo
    colonne_testo = df.select_dtypes(
        include=["object", "string"]
    ).columns

    for colonna in colonne_testo:

        df[colonna] = df[colonna].apply(
            lambda valore:
            valore.strip()
            if isinstance(valore, str)
            else valore
        )

    # Conta duplicati
    duplicati = int(
        df.duplicated().sum()
    )

    # Elimina duplicati
    df.drop_duplicates(
        inplace=True
    )

    df.reset_index(
        drop=True,
        inplace=True
    )

    informazioni = {
        "Righe originali": righe_originali,
        "Righe finali": len(df),
        "Duplicati rimossi": duplicati,
        "Celle vuote": int(
            df.isna().sum().sum()
        ),
    }

    return df, informazioni


# ============================================================
# RICONOSCIMENTO COLONNE
# ============================================================

def riconosci_colonne(df):

    return {

        "quantita": trova_colonna(
            df,
            [
                "quantita",
                "quantità",
                "quantity",
                "qty",
                "pezzi",
                "units",
            ],
        ),

        "prezzo": trova_colonna(
            df,
            [
                "prezzo",
                "price",
                "prezzo_unitario",
                "unit_price",
                "unitprice",
            ],
        ),

        "prodotto": trova_colonna(
            df,
            [
                "prodotto",
                "product",
                "articolo",
                "item",
            ],
        ),

        "categoria": trova_colonna(
            df,
            [
                "categoria",
                "category",
                "reparto",
            ],
        ),

        "data": trova_colonna(
            df,
            [
                "data",
                "date",
                "data_vendita",
                "sale_date",
            ],
        ),

        "cliente": trova_colonna(
            df,
            [
                "cliente",
                "customer",
                "client",
            ],
        ),
    }


# ============================================================
# ANALISI VENDITE
# ============================================================

def converti_numero(valore):
    if pd.isna(valore):
        return None

    if isinstance(valore, (int, float)):
        return valore

    testo = str(valore).strip()

    # Rimuove simboli e spazi comuni
    testo = (
        testo.replace("€", "")
        .replace("$", "")
        .replace("£", "")
        .replace(" ", "")
    )

    # Formato europeo: 1.250,50
    if "," in testo and "." in testo:
        if testo.rfind(",") > testo.rfind("."):
            testo = testo.replace(".", "")
            testo = testo.replace(",", ".")
        else:
            testo = testo.replace(",", "")

    # Formato europeo semplice: 49,90
    elif "," in testo:
        testo = testo.replace(",", ".")

    try:
        return float(testo)
    except (ValueError, TypeError):
        return None

def analizza_vendite(df):

    df = df.copy()

    colonne = riconosci_colonne(df)

    quantita = colonne["quantita"]
    prezzo = colonne["prezzo"]
    prodotto = colonne["prodotto"]
    categoria = colonne["categoria"]

    analisi = {
        "attiva": False,
        "colonne": colonne,
        "kpi": {},
        "prodotti": pd.DataFrame(),
        "categorie": pd.DataFrame(),
    }

    # Servono almeno quantità e prezzo
    if not quantita or not prezzo:
        return df, analisi

    # Converte quantità e prezzo in numeri
    df[quantita] = df[quantita].apply(converti_numero)
    df[prezzo] = df[prezzo].apply(converti_numero)

    # Calcolo fatturato
    df["Fatturato"] = (
        df[quantita]
        * df[prezzo]
    )

    analisi["attiva"] = True

    fatturato_totale = float(
        df["Fatturato"].sum()
    )

    quantita_totale = float(
        df[quantita].sum()
    )

    numero_operazioni = len(df)

    if numero_operazioni > 0:

        valore_medio = (
            fatturato_totale
            / numero_operazioni
        )

    else:

        valore_medio = 0

    analisi["kpi"] = {
        "Fatturato totale": fatturato_totale,
        "Quantità totale": quantita_totale,
        "Numero operazioni": numero_operazioni,
        "Valore medio operazione": valore_medio,
    }

    # --------------------------------------------------------
    # ANALISI PRODOTTI
    # --------------------------------------------------------

    if prodotto:

        prodotti = (
            df.groupby(
                prodotto,
                dropna=False
            )
            .agg(
                Quantita_totale=(
                    quantita,
                    "sum"
                ),
                Fatturato=(
                    "Fatturato",
                    "sum"
                ),
            )
            .reset_index()
            .sort_values(
                "Fatturato",
                ascending=False
            )
        )

        analisi["prodotti"] = prodotti

        if not prodotti.empty:

            analisi["kpi"][
                "Prodotto con maggior fatturato"
            ] = str(
                prodotti.iloc[0][prodotto]
            )

    # --------------------------------------------------------
    # ANALISI CATEGORIE
    # --------------------------------------------------------

    if categoria:

        categorie = (
            df.groupby(
                categoria,
                dropna=False
            )
            .agg(
                Quantita_totale=(
                    quantita,
                    "sum"
                ),
                Fatturato=(
                    "Fatturato",
                    "sum"
                ),
            )
            .reset_index()
            .sort_values(
                "Fatturato",
                ascending=False
            )
        )

        analisi["categorie"] = categorie

    return df, analisi


# ============================================================
# STATISTICHE
# ============================================================

def crea_statistiche(df):

    risultati = []

    colonne_numeriche = df.select_dtypes(
        include="number"
    ).columns

    for colonna in colonne_numeriche:

        serie = df[colonna].dropna()

        if serie.empty:
            continue

        risultati.append({
            "Colonna": colonna,

            "Valori": int(
                serie.count()
            ),

            "Valori mancanti": int(
                df[colonna].isna().sum()
            ),

            "Totale": serie.sum(),

            "Media": serie.mean(),

            "Mediana": serie.median(),

            "Minimo": serie.min(),

            "Massimo": serie.max(),
        })

    return pd.DataFrame(
        risultati
    )


# ============================================================
# QUALITÀ DATI
# ============================================================

def crea_analisi_qualita(df):

    risultati = []

    for colonna in df.columns:

        valori_mancanti = int(
            df[colonna].isna().sum()
        )

        if len(df) > 0:

            percentuale = (
                valori_mancanti
                / len(df)
                * 100
            )

        else:

            percentuale = 0

        risultati.append({

            "Colonna": colonna,

            "Tipo": str(
                df[colonna].dtype
            ),

            "Valori mancanti":
                valori_mancanti,

            "% mancanti": round(
                percentuale,
                2
            ),

            "Valori unici": int(
                df[colonna].nunique(
                    dropna=True
                )
            ),
        })

    return pd.DataFrame(
        risultati
    )


# ============================================================
# FORMATTAZIONE FOGLI
# ============================================================

def formatta_foglio(ws):

    riempimento_header = PatternFill(
        "solid",
        fgColor=COLORE_PRINCIPALE
    )

    font_header = Font(
        color=COLORE_BIANCO,
        bold=True
    )

    bordo = Side(
        style="thin",
        color="D9E2F3"
    )

    # Intestazioni
    for cella in ws[1]:

        cella.fill = riempimento_header
        cella.font = font_header

        cella.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

    # Bordi
    for riga in ws.iter_rows():

        for cella in riga:

            cella.border = Border(
                bottom=bordo
            )

            cella.alignment = Alignment(
                vertical="center"
            )

    # Larghezza automatica
    for colonna in ws.columns:

        lettera = get_column_letter(
            colonna[0].column
        )

        lunghezza_massima = 0

        for cella in colonna:

            if cella.value is not None:

                lunghezza_massima = max(
                    lunghezza_massima,
                    len(str(cella.value))
                )

        ws.column_dimensions[
            lettera
        ].width = min(
            lunghezza_massima + 3,
            35
        )

    ws.freeze_panes = "A2"

    if ws.max_row > 1:

        ws.auto_filter.ref = (
            ws.dimensions
        )


# ============================================================
# DASHBOARD
# ============================================================

def crea_dashboard(workbook, analisi):

    if not analisi["attiva"]:
        return

    ws = workbook.create_sheet(
        "Dashboard",
        0
    )

    ws.sheet_view.showGridLines = False

    # --------------------------------------------------------
    # TITOLO
    # --------------------------------------------------------

    ws.merge_cells(
        "A1:L2"
    )

    ws["A1"] = (
        "SMART SALES REPORT"
    )

    ws["A1"].font = Font(
        bold=True,
        size=24,
        color=COLORE_BIANCO
    )

    ws["A1"].fill = PatternFill(
        "solid",
        fgColor=COLORE_PRINCIPALE
    )

    ws["A1"].alignment = Alignment(
        horizontal="center",
        vertical="center"
    )

    ws.row_dimensions[1].height = 30
    ws.row_dimensions[2].height = 15

    # --------------------------------------------------------
    # KPI
    # --------------------------------------------------------

    kpi = analisi["kpi"]

    cards = [

        (
            "A4:C4",
            "A5:C6",
            "FATTURATO TOTALE",
            kpi.get(
                "Fatturato totale",
                0
            ),
            '€ #,##0.00'
        ),

        (
            "D4:F4",
            "D5:F6",
            "QUANTITÀ VENDUTA",
            kpi.get(
                "Quantità totale",
                0
            ),
            '#,##0'
        ),

        (
            "G4:I4",
            "G5:I6",
            "OPERAZIONI",
            kpi.get(
                "Numero operazioni",
                0
            ),
            '#,##0'
        ),

        (
            "J4:L4",
            "J5:L6",
            "VALORE MEDIO",
            kpi.get(
                "Valore medio operazione",
                0
            ),
            '€ #,##0.00'
        ),
    ]

    for (
        intervallo_titolo,
        intervallo_valore,
        nome,
        valore,
        formato
    ) in cards:

        ws.merge_cells(
            intervallo_titolo
        )

        ws.merge_cells(
            intervallo_valore
        )

        cella_titolo = ws[
            intervallo_titolo.split(":")[0]
        ]

        cella_valore = ws[
            intervallo_valore.split(":")[0]
        ]

        cella_titolo.value = nome

        cella_titolo.font = Font(
            bold=True,
            color=COLORE_BIANCO,
            size=11
        )

        cella_titolo.fill = PatternFill(
            "solid",
            fgColor=COLORE_SECONDARIO
        )

        cella_titolo.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

        cella_valore.value = valore

        cella_valore.font = Font(
            bold=True,
            size=18
        )

        cella_valore.fill = PatternFill(
            "solid",
            fgColor=COLORE_CHIARO
        )

        cella_valore.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

        cella_valore.number_format = (
            formato
        )

    # --------------------------------------------------------
    # PRODOTTO MIGLIORE
    # --------------------------------------------------------

    ws.merge_cells(
        "A8:L8"
    )

    prodotto_migliore = kpi.get(
        "Prodotto con maggior fatturato",
        "Non disponibile"
    )

    ws["A8"] = (
        "Prodotto con maggior fatturato: "
        f"{prodotto_migliore}"
    )

    ws["A8"].font = Font(
        bold=True,
        size=13
    )

    ws["A8"].fill = PatternFill(
        "solid",
        fgColor=COLORE_VERDE
    )

    ws["A8"].alignment = Alignment(
        horizontal="center",
        vertical="center"
    )

    # --------------------------------------------------------
    # DATI PER GRAFICO PRODOTTI
    # --------------------------------------------------------

    prodotti = analisi["prodotti"]

    if not prodotti.empty:

        riga_inizio = 40

        nome_colonna_prodotto = (
            prodotti.columns[0]
        )

        ws.cell(
            riga_inizio,
            1,
            nome_colonna_prodotto
        )

        ws.cell(
            riga_inizio,
            2,
            "Fatturato"
        )

        for numero, (_, riga) in enumerate(
            prodotti.iterrows(),
            start=riga_inizio + 1
        ):

            ws.cell(
                numero,
                1,
                riga[
                    nome_colonna_prodotto
                ]
            )

            ws.cell(
                numero,
                2,
                float(
                    riga["Fatturato"]
                )
            )

        grafico_prodotti = BarChart()

        grafico_prodotti.type = "col"

        grafico_prodotti.style = 10

        grafico_prodotti.title = (
            "Fatturato per prodotto"
        )

        grafico_prodotti.y_axis.title = (
            "Fatturato (€)"
        )

        grafico_prodotti.x_axis.title = (
            "Prodotto"
        )

        grafico_prodotti.height = 10
        grafico_prodotti.width = 16

        dati = Reference(
            ws,
            min_col=2,
            min_row=riga_inizio,
            max_row=(
                riga_inizio
                + len(prodotti)
            )
        )

        categorie = Reference(
            ws,
            min_col=1,
            min_row=riga_inizio + 1,
            max_row=(
                riga_inizio
                + len(prodotti)
            )
        )

        grafico_prodotti.add_data(
            dati,
            titles_from_data=True
        )

        grafico_prodotti.set_categories(
            categorie
        )

        grafico_prodotti.legend = None

        ws.add_chart(
            grafico_prodotti,
            "A10"
        )

    # --------------------------------------------------------
    # DATI PER GRAFICO CATEGORIE
    # --------------------------------------------------------

    categorie_df = analisi[
        "categorie"
    ]

    if not categorie_df.empty:

        riga_inizio = 40

        nome_colonna_categoria = (
            categorie_df.columns[0]
        )

        ws.cell(
            riga_inizio,
            4,
            nome_colonna_categoria
        )

        ws.cell(
            riga_inizio,
            5,
            "Fatturato"
        )

        for numero, (_, riga) in enumerate(
            categorie_df.iterrows(),
            start=riga_inizio + 1
        ):

            ws.cell(
                numero,
                4,
                riga[
                    nome_colonna_categoria
                ]
            )

            ws.cell(
                numero,
                5,
                float(
                    riga["Fatturato"]
                )
            )

        grafico_categorie = PieChart()

        grafico_categorie.title = (
            "Fatturato per categoria"
        )

        grafico_categorie.height = 10
        grafico_categorie.width = 14

        dati = Reference(
            ws,
            min_col=5,
            min_row=riga_inizio,
            max_row=(
                riga_inizio
                + len(categorie_df)
            )
        )

        categorie = Reference(
            ws,
            min_col=4,
            min_row=riga_inizio + 1,
            max_row=(
                riga_inizio
                + len(categorie_df)
            )
        )

        grafico_categorie.add_data(
            dati,
            titles_from_data=True
        )

        grafico_categorie.set_categories(
            categorie
        )

        # Percentuali nel grafico a torta
        grafico_categorie.dataLabels = (
            DataLabelList()
        )

        grafico_categorie.dataLabels.showPercent = True

        grafico_categorie.dataLabels.showLeaderLines = True

        ws.add_chart(
            grafico_categorie,
            "G10"
        )

    # Nasconde le righe tecniche
    for riga in range(
        40,
        100
    ):
        ws.row_dimensions[
            riga
        ].hidden = True

    # Larghezza dashboard
    for colonna in range(
        1,
        13
    ):

        lettera = get_column_letter(
            colonna
        )

        ws.column_dimensions[
            lettera
        ].width = 13


# ============================================================
# GRAFICO PRODOTTI
# ============================================================

def aggiungi_grafico_prodotti(
    workbook,
    analisi
):

    prodotti = analisi["prodotti"]

    if prodotti.empty:
        return

    ws = workbook[
        "Vendite per Prodotto"
    ]

    grafico = BarChart()

    grafico.type = "col"

    grafico.style = 10

    grafico.title = (
        "Fatturato per prodotto"
    )

    grafico.y_axis.title = (
        "Fatturato (€)"
    )

    grafico.x_axis.title = (
        "Prodotto"
    )

    grafico.height = 11
    grafico.width = 18

    # Solo Fatturato come serie
    dati = Reference(
        ws,
        min_col=3,
        min_row=1,
        max_row=ws.max_row
    )

    # Prodotti sull'asse X
    categorie = Reference(
        ws,
        min_col=1,
        min_row=2,
        max_row=ws.max_row
    )

    grafico.add_data(
        dati,
        titles_from_data=True
    )

    grafico.set_categories(
        categorie
    )

    grafico.legend = None

    ws.add_chart(
        grafico,
        "E4"
    )

    # Formato valuta
    for cella in ws["C"][1:]:

        cella.number_format = (
            '€ #,##0.00'
        )


# ============================================================
# GRAFICO CATEGORIE
# ============================================================

def aggiungi_grafico_categorie(
    workbook,
    analisi
):

    categorie_df = analisi[
        "categorie"
    ]

    if categorie_df.empty:
        return

    ws = workbook[
        "Vendite per Categoria"
    ]

    grafico = PieChart()

    grafico.title = (
        "Distribuzione fatturato per categoria"
    )

    grafico.height = 11
    grafico.width = 16

    dati = Reference(
        ws,
        min_col=3,
        min_row=1,
        max_row=ws.max_row
    )

    categorie = Reference(
        ws,
        min_col=1,
        min_row=2,
        max_row=ws.max_row
    )

    grafico.add_data(
        dati,
        titles_from_data=True
    )

    grafico.set_categories(
        categorie
    )

    grafico.dataLabels = (
        DataLabelList()
    )

    grafico.dataLabels.showPercent = True

    grafico.dataLabels.showLeaderLines = True

    ws.add_chart(
        grafico,
        "E4"
    )

    for cella in ws["C"][1:]:

        cella.number_format = (
            '€ #,##0.00'
        )


# ============================================================
# CREAZIONE REPORT
# ============================================================

def genera_report(
    df,
    statistiche,
    qualita,
    pulizia,
    analisi,
    percorso_originale
):

    cartella_output = (
        percorso_originale.parent
        / "output"
    )

    cartella_output.mkdir(
        exist_ok=True
    )

    timestamp = datetime.now().strftime(
        "%Y%m%d_%H%M%S"
    )

    nome_file = (
        f"REPORT_"
        f"{percorso_originale.stem}_"
        f"{timestamp}.xlsx"
    )

    percorso_output = (
        cartella_output
        / nome_file
    )

    informazioni = pd.DataFrame({

        "Informazione": [

            "File originale",

            "Data generazione",

            "Righe originali",

            "Righe finali",

            "Duplicati rimossi",

            "Celle vuote",

            "Numero colonne",

            "Analisi vendite",
        ],

        "Valore": [

            percorso_originale.name,

            datetime.now().strftime(
                "%d/%m/%Y %H:%M:%S"
            ),

            pulizia[
                "Righe originali"
            ],

            pulizia[
                "Righe finali"
            ],

            pulizia[
                "Duplicati rimossi"
            ],

            pulizia[
                "Celle vuote"
            ],

            len(df.columns),

            (
                "Disponibile"
                if analisi["attiva"]
                else "Non rilevata"
            ),
        ],
    })

    with pd.ExcelWriter(
        percorso_output,
        engine="openpyxl"
    ) as writer:

        # ----------------------------------------------------
        # DATI
        # ----------------------------------------------------

        df.to_excel(
            writer,
            sheet_name="Dati Puliti",
            index=False
        )

        # ----------------------------------------------------
        # STATISTICHE
        # ----------------------------------------------------

        statistiche.to_excel(
            writer,
            sheet_name="Statistiche",
            index=False
        )

        # ----------------------------------------------------
        # QUALITÀ
        # ----------------------------------------------------

        qualita.to_excel(
            writer,
            sheet_name="Qualita Dati",
            index=False
        )

        # ----------------------------------------------------
        # INFORMAZIONI
        # ----------------------------------------------------

        informazioni.to_excel(
            writer,
            sheet_name="Informazioni",
            index=False
        )

        # ----------------------------------------------------
        # PRODOTTI
        # ----------------------------------------------------

        if (
            analisi["attiva"]
            and not analisi[
                "prodotti"
            ].empty
        ):

            analisi[
                "prodotti"
            ].to_excel(
                writer,
                sheet_name=(
                    "Vendite per Prodotto"
                ),
                index=False
            )

        # ----------------------------------------------------
        # CATEGORIE
        # ----------------------------------------------------

        if (
            analisi["attiva"]
            and not analisi[
                "categorie"
            ].empty
        ):

            analisi[
                "categorie"
            ].to_excel(
                writer,
                sheet_name=(
                    "Vendite per Categoria"
                ),
                index=False
            )

        workbook = writer.book

        # ----------------------------------------------------
        # DASHBOARD
        # ----------------------------------------------------

        crea_dashboard(
            workbook,
            analisi
        )

        # ----------------------------------------------------
        # FORMATTAZIONE
        # ----------------------------------------------------

        for nome_foglio in (
            workbook.sheetnames
        ):

            if nome_foglio != "Dashboard":

                formatta_foglio(
                    workbook[
                        nome_foglio
                    ]
                )

        # ----------------------------------------------------
        # FORMATO FATTURATO DATI
        # ----------------------------------------------------

        ws_dati = workbook[
            "Dati Puliti"
        ]

        intestazioni = {
            cella.value: cella.column
            for cella in ws_dati[1]
        }

        if "Fatturato" in intestazioni:

            colonna_fatturato = (
                intestazioni[
                    "Fatturato"
                ]
            )

            for riga in range(
                2,
                ws_dati.max_row + 1
            ):

                ws_dati.cell(
                    riga,
                    colonna_fatturato
                ).number_format = (
                    '€ #,##0.00'
                )

        # ----------------------------------------------------
        # GRAFICI
        # ----------------------------------------------------

        if (
            "Vendite per Prodotto"
            in workbook.sheetnames
        ):

            aggiungi_grafico_prodotti(
                workbook,
                analisi
            )

        if (
            "Vendite per Categoria"
            in workbook.sheetnames
        ):

            aggiungi_grafico_categorie(
                workbook,
                analisi
            )

        # ----------------------------------------------------
        # QUALITÀ DATI
        # ----------------------------------------------------

        if (
            "Qualita Dati"
            in workbook.sheetnames
        ):

            ws_qualita = workbook[
                "Qualita Dati"
            ]

            if ws_qualita.max_row >= 2:

                ws_qualita.conditional_formatting.add(

                    f"D2:D{ws_qualita.max_row}",

                    ColorScaleRule(

                        start_type="min",
                        start_color="63BE7B",

                        mid_type="percentile",
                        mid_value=50,
                        mid_color="FFEB84",

                        end_type="max",
                        end_color="F8696B",
                    )
                )

    return percorso_output


# ============================================================
# APERTURA REPORT
# ============================================================

def apri_report(percorso):

    try:

        if os.name == "nt":

            os.startfile(
                percorso
            )

    except Exception:

        pass


# ============================================================
# PROGRAMMA PRINCIPALE
# ============================================================

def main():

    logger.info("Avvio Smart Sales Report versione %s", VERSION)
    
    print(
        "=" * 65
    )

    print(
        f"{APP_NAME} - Versione {VERSION}"
    )

    print(
        "=" * 65
    )

    # --------------------------------------------------------
    # SELEZIONE FILE
    # --------------------------------------------------------

    percorso = scegli_file()

    if percorso is None:

        print(
            "Nessun file selezionato."
        )

        return

    try:

        print(
            f"\nFile selezionato: "
            f"{percorso.name}"
        )

        # ----------------------------------------------------
        # CARICAMENTO
        # ----------------------------------------------------

        print(
            "\nCaricamento dati..."
        )

        logger.info("File selezionato: %s", percorso)

        df_originale = carica_dati(
            percorso
        )

        if df_originale.empty:

            raise ValueError(
                "Il file selezionato è vuoto."
            )

        # ----------------------------------------------------
        # PULIZIA
        # ----------------------------------------------------

        print(
            "Pulizia dati..."
        )

        df, pulizia = pulisci_dati(
            df_originale
        )

        # ----------------------------------------------------
        # ANALISI VENDITE
        # ----------------------------------------------------

        print(
            "Analisi dati..."
        )

        df, analisi = analizza_vendite(
            df
        )

        # ----------------------------------------------------
        # STATISTICHE
        # ----------------------------------------------------

        statistiche = crea_statistiche(
            df
        )

        # ----------------------------------------------------
        # QUALITÀ DATI
        # ----------------------------------------------------

        qualita = crea_analisi_qualita(
            df
        )

        # ----------------------------------------------------
        # ANTEPRIMA
        # ----------------------------------------------------

        print(
            "\n=== ANTEPRIMA DATI ==="
        )

        print(
            df.head(10).to_string(
                index=False
            )
        )

        # ----------------------------------------------------
        # PULIZIA
        # ----------------------------------------------------

        print(
            "\n=== PULIZIA DATI ==="
        )

        for nome, valore in (
            pulizia.items()
        ):

            print(
                f"{nome}: {valore}"
            )

        # ----------------------------------------------------
        # RISULTATI
        # ----------------------------------------------------

        if analisi["attiva"]:

            print(
                "\n=== ANALISI VENDITE ==="
            )

            for nome, valore in (
                analisi["kpi"].items()
            ):

                if isinstance(
                    valore,
                    float
                ):

                    print(
                        f"{nome}: "
                        f"{valore:,.2f}"
                    )

                else:

                    print(
                        f"{nome}: "
                        f"{valore}"
                    )

        else:

            print(
                "\nIl file è stato analizzato, "
                "ma non sono state trovate "
                "colonne sufficienti per "
                "l'analisi delle vendite."
            )

        # ----------------------------------------------------
        # REPORT
        # ----------------------------------------------------

        print(
            "\nGenerazione report Excel..."
        )

        output = genera_report(
            df,
            statistiche,
            qualita,
            pulizia,
            analisi,
            percorso
        )

        print(
            "\n" + "=" * 65
        )

        print(
            "REPORT CREATO CON SUCCESSO"
        )

        print(
            "=" * 65
        )

        print(
            f"\nFile salvato in:\n"
            f"{output}"
        )

        # ----------------------------------------------------
        # MESSAGGIO WINDOWS
        # ----------------------------------------------------

        root = crea_root()

        messagebox.showinfo(
            "Report completato",

            "Analisi completata "
            "con successo.\n\n"
            f"Report salvato in:\n"
            f"{output}"
        )

        root.destroy()

        # ----------------------------------------------------
        # APERTURA AUTOMATICA
        # ----------------------------------------------------

        apri_report(
            output
        )

    except PermissionError:

        print(
            "\nERRORE: impossibile "
            "scrivere il file."
        )

        print(
            "Chiudi eventuali report "
            "Excel già aperti e riprova."
        )

    except Exception as errore:

        logger.exception("Errore durante l'elaborazione: %s", errore)
        
        print(
            "\nERRORE DURANTE "
            "L'ELABORAZIONE:"
        )

        print(
            str(errore)
        )

        root = crea_root()

        messagebox.showerror(
            "Errore",
            str(errore)
        )

        root.destroy()


# ============================================================
# AVVIO
# ============================================================

if __name__ == "__main__":
    main()