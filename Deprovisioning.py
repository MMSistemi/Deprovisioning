# -*- coding: utf-8 -*-

import re
from datetime import date, datetime, timedelta

import pandas as pd
import streamlit as st


# ============================================================
# CONFIGURAZIONE PAGINA
# ============================================================

st.set_page_config(
    page_title="Deprovisioning Utente",
    layout="wide"
)

st.title("Deprovisioning Utente")

st.caption(
    "Generazione dei task di deprovisioning sulla base dello "
    "sAMAccountName e della data di fine rapporto."
)


# ============================================================
# POSSIBILI NOMI DELLE COLONNE
# ============================================================

CAND_MG_MEMBER = [
    "Member",
    "Membro",
    "MemberSamAccountName",
    "MembroSamAccountName",
    "SamAccountName",
    "sAMAccountName",
    "MemberUserPrincipalName",
    "UserPrincipalNameMembro",
    "userPrincipalName",
    "UPN"
]

CAND_MG_GROUP = [
    "Group",
    "Gruppo",
    "GroupName",
    "NomeGruppo",
    "DisplayName",
    "NomeVisualizzato",
    "cn",
    "CN",
    "Name",
    "Nome"
]


# ============================================================
# FUNZIONI GENERALI
# ============================================================

_SPACES_UNDERS = re.compile(r"[\s_]+")


def norm_col(value):
    """
    Normalizza il nome di una colonna rimuovendo
    spazi e underscore.
    """
    return _SPACES_UNDERS.sub(
        "",
        str(value).strip().lower()
    )


def find_col(df, candidates):
    """
    Cerca una colonna nel DataFrame utilizzando
    prima il match esatto e poi il match parziale.
    """
    if df is None or df.empty:
        return None

    normalized = {
        norm_col(col): col
        for col in df.columns
    }

    # Match esatto
    for candidate in candidates:
        key = norm_col(candidate)

        if key in normalized:
            return normalized[key]

    # Match parziale
    for candidate in candidates:
        key = norm_col(candidate)

        for col in df.columns:
            normalized_col = norm_col(col)

            if key and key in normalized_col:
                return col

    return None


def read_excel(uploaded_file):
    """
    Legge il file Excel caricato.
    """
    if uploaded_file is None:
        return pd.DataFrame()

    try:
        uploaded_file.seek(0)

        return pd.read_excel(
            uploaded_file,
            engine="openpyxl"
        )

    except Exception as error:
        st.warning(
            f"Errore nella lettura di "
            f"{uploaded_file.name}: {error}"
        )

        return pd.DataFrame()


def clean_list(series):
    """
    Restituisce una lista ordinata senza duplicati.
    """
    if series is None or series.empty:
        return []

    result = []
    seen = set()

    for value in series.dropna():
        value = str(value).strip()
        key = value.lower()

        if value and key not in seen:
            result.append(value)
            seen.add(key)

    return sorted(
        result,
        key=str.lower
    )


def unique_in_order(values):
    """
    Rimuove i duplicati conservando l'ordine originale.
    """
    result = []
    seen = set()

    for value in values:
        value = str(value).strip()
        key = value.lower()

        if value and key not in seen:
            result.append(value)
            seen.add(key)

    return result


# ============================================================
# GESTIONE SAM
# ============================================================

def normalize_sam(sam):
    """
    Normalizza lo sAMAccountName.
    """
    return str(sam).strip().lower()


def clean_sam(sam):
    """
    Rimuove il suffisso .ext, se presente.
    """
    sam = normalize_sam(sam)

    if sam.endswith(".ext"):
        return sam[:-4]

    return sam


def user_candidates(sam):
    """
    Genera le possibili forme con cui l'utente
    può comparire nel file Estr_MembriGruppi.
    """
    sam = normalize_sam(sam)
    clean = clean_sam(sam)

    candidates = {
        sam,
        clean,
        f"{sam}@consip.it",
        f"{clean}@consip.it"
    }

    return {
        value.lower()
        for value in candidates
        if value
    }


def member_mask(series, sam):
    """
    Cerca lo sAMAccountName nella colonna dei membri.
    """
    candidates = user_candidates(sam)

    values = (
        series
        .fillna("")
        .astype(str)
        .str.strip()
        .str.lower()
    )

    # Prima ricerca con corrispondenza esatta
    mask = values.isin(candidates)

    if mask.any():
        return mask

    # Seconda ricerca per celle contenenti più membri
    escaped_candidates = [
        re.escape(value)
        for value in candidates
    ]

    pattern = (
        r"(?:^|[;,\s])(?:"
        + "|".join(escaped_candidates)
        + r")(?:$|[;,\s])"
    )

    return values.str.contains(
        pattern,
        regex=True,
        na=False
    )


# ============================================================
# ESTRAZIONE GRUPPI AD
# ============================================================

def extract_ad_groups(sam, df):
    """
    Estrae i gruppi AD associati all'utente.
    Esclude Domain Users e Utenti del dominio.
    """
    if df is None or df.empty:
        return [], [
            "File Estr_MembriGruppi non caricato. "
            "Il task 04 verrà generato senza il dettaglio "
            "dei gruppi AD."
        ]

    member_col = find_col(
        df,
        CAND_MG_MEMBER
    )

    group_col = find_col(
        df,
        CAND_MG_GROUP
    )

    if not member_col or not group_col:
        return [], [
            "Colonne Member/Group non trovate nel file "
            "Estr_MembriGruppi."
        ]

    mask = member_mask(
        df[member_col],
        sam
    )

    groups = clean_list(
        df.loc[
            mask,
            group_col
        ]
    )

    excluded_groups = {
        "domain users",
        "utenti del dominio"
    }

    groups = [
        group
        for group in groups
        if group.strip().lower() not in excluded_groups
    ]

    warnings = []

    if not groups:
        warnings.append(
            "Nessun gruppo AD aggiuntivo trovato per l'utente."
        )

    return groups, warnings


# ============================================================
# VALIDAZIONE E FORMATTAZIONE DATA
# ============================================================

def format_datetime(value):
    """
    Formatta una data nel formato richiesto:
    mm/gg/aaaa 00:00
    """
    return value.strftime(
        "%m/%d/%Y 00:00"
    )


def calculate_dates(end_date):
    """
    Calcola la data di fine rapporto e il giorno
    di calendario successivo.

    Esempio:
    11/30/2026 -> 12/01/2026 00:00
    """
    next_date = end_date + timedelta(days=1)

    end_date_formatted = format_datetime(
        end_date
    )

    next_date_formatted = format_datetime(
        next_date
    )

    return (
        end_date_formatted,
        next_date_formatted
    )


# ============================================================
# GENERAZIONE TASK 01
# ============================================================

def genera_task_01(
    sam,
    data_fine
):
    return f"""Ciao, si richiede per l'utenza {sam}:

1. Il giorno {data_fine}:
- Ritirare dall'utente gli asset associati facendolo firmare
- Dismissione Utenza On-prem AD

2. Deprovisioning PC e consegna in armadio 599

Chiudere il task quando tutte le attività sono completate."""


# ============================================================
# GENERAZIONE TASK 02
# ============================================================

def genera_task_02(
    sam,
    data_fine_piu_uno
):
    return f"""Ciao,

si richiede per {sam}:

1) Da svolgere subito:
- Impostare su AD, per il campo "ExpireDate", il valore "{data_fine_piu_uno}"

2) Il giorno {data_fine_piu_uno}:

- Dismissione Casella di posta:
  - Disabilitare invio ad utente (Message Delivery Restrictions)
  - Impostare Hide dalla Rubrica
  - Disabilitare accesso Mailbox (Mailbox features – Disable Protocolli/OWA)
  - Estrarre il PST (O365 eDiscovery) da archiviare in:
    \\\\nasconsip2....\\backuppst\\03 - backup email cancellate
    Il file deve essere archiviato in formato 7z con password condivisa.

- Rimozione profilazione:
  - Rimuovere tutti i Ruoli assegnati
  - Rimuovere tutte le applicazioni dall'utenza Azure
  - Disabilitare l'account di Azure
  - Rimuovere l'abilitazione da tutte le SM, se applicabile
  - Rimuovere l'abilitazione da tutte le DL, se applicabile
  - Rimuovere tutti i gruppi Azure
  - Rimuovere tutti i gruppi AD
  - Cancellare la foto da Azure, se applicabile
  - Rimuovere il Wi-Fi

- Disabilitare l'utenza di dominio e spostarla nella OU Dismessi/Utenti

- Rimuovere i riferimenti utente dal PC su AD:
  - unset attributo mobile
  - unset attributo UPN
  - unset attributo mail

Chiudere il task quando tutte le attività sono completate."""


# ============================================================
# GENERAZIONE TASK 03
# ============================================================

def genera_task_03(sam):
    return f"""GSI

Gestione firma digitale - Gestione Asset

Revocare la firma e rimuovere il diritto a richiedere la Firma Digitale alla risorsa {sam}.

Gestire l'asset e il suo stato."""


# ============================================================
# GENERAZIONE TASK 04
# ============================================================

def genera_task_04(
    sam,
    gruppi_ad
):
    lines = [
        f"Attivare il deprovisioning dei Servizi IT per la risorsa {sam} "
        "sulla base delle evidenze trovate nei gruppi AD."
    ]

    if gruppi_ad:
        lines.extend([
            "",
            "Evidenze trovate nei gruppi AD:"
        ])

        for group in gruppi_ad:
            lines.append(
                f"- {group}"
            )

    else:
        lines.extend([
            "",
            "Nessuna evidenza relativa a gruppi AD aggiuntivi "
            "è stata rilevata."
        ])

    return "\n".join(lines)


# ============================================================
# FUNZIONE DI VISUALIZZAZIONE TASK
# ============================================================

def visualizza_task(
    titolo,
    destinatario,
    contenuto,
    key
):
    """
    Visualizza ogni task in un contenitore separato.
    Il testo viene inserito in una text area per consentire
    una selezione e copia agevole.
    """
    with st.container(border=True):

        st.markdown(
            f"## {titolo}"
        )

        if destinatario:
            st.markdown(
                f"**Destinatario/Competenza: {destinatario}**"
            )

        st.text_area(
            label=f"Testo del task {key}",
            value=contenuto,
            height=420,
            key=f"text_area_{key}",
            label_visibility="collapsed"
        )


# ============================================================
# INTERFACCIA
# ============================================================

st.subheader(
    "Modulo Deprovisioning"
)

col1, col2 = st.columns(2)

with col1:
    sam = st.text_input(
        "Nome utente (sAMAccountName)",
        value="",
        placeholder="Esempio: mario.rossi oppure mario.rossi.ext",
        help=(
            "Inserire lo sAMAccountName dell'utente "
            "da deprovisionare."
        )
    ).strip().lower()

with col2:
    data_fine = st.date_input(
        "Data fine rapporto",
        value=date.today(),
        format="DD/MM/YYYY",
        help=(
            "Il sistema calcolerà automaticamente il giorno "
            "successivo e lo formatterà come mm/gg/aaaa 00:00."
        )
    )


# ============================================================
# UPLOAD FILE GRUPPI AD
# ============================================================

st.markdown(
    "### Evidenze gruppi AD"
)

mg_file = st.file_uploader(
    "Carica Estr_MembriGruppi (Excel) - facoltativo",
    type=["xlsx"],
    key="mg",
    help=(
        "Il file viene utilizzato esclusivamente per individuare "
        "i gruppi AD associati all'utente e compilare il task 04."
    )
)


# ============================================================
# GENERAZIONE
# ============================================================

if st.button(
    "Genera Deprovisioning",
    type="primary",
    use_container_width=True
):

    # --------------------------------------------------------
    # Validazione sAMAccountName
    # --------------------------------------------------------

    if not sam:
        st.error(
            "Inserire lo sAMAccountName."
        )
        st.stop()

    if " " in sam:
        st.error(
            "Lo sAMAccountName non può contenere spazi."
        )
        st.stop()

    # --------------------------------------------------------
    # Calcolo date
    # --------------------------------------------------------

    (
        data_fine_formattata,
        data_fine_piu_uno_formattata
    ) = calculate_dates(
        data_fine
    )

    # --------------------------------------------------------
    # Lettura file gruppi AD
    # --------------------------------------------------------

    mg_df = read_excel(
        mg_file
    )

    gruppi_ad, warnings = extract_ad_groups(
        sam,
        mg_df
    )

    gruppi_ad = unique_in_order(
        gruppi_ad
    )

    # --------------------------------------------------------
    # Generazione task
    # --------------------------------------------------------

    task_01 = genera_task_01(
        sam,
        data_fine_formattata
    )

    task_02 = genera_task_02(
        sam,
        data_fine_piu_uno_formattata
    )

    task_03 = genera_task_03(
        sam
    )

    task_04 = genera_task_04(
        sam,
        gruppi_ad
    )

    # --------------------------------------------------------
    # Riepilogo date
    # --------------------------------------------------------

    st.divider()

    st.subheader(
        "Riepilogo"
    )

    riepilogo_col1, riepilogo_col2, riepilogo_col3 = st.columns(3)

    with riepilogo_col1:
        st.metric(
            "sAMAccountName",
            sam
        )

    with riepilogo_col2:
        st.metric(
            "Data fine",
            data_fine_formattata
        )

    with riepilogo_col3:
        st.metric(
            "Data fine + 1 / ExpireDate",
            data_fine_piu_uno_formattata
        )

    st.info(
        "La data fine + 1 è calcolata aggiungendo un giorno "
        "di calendario alla data di fine rapporto."
    )

    # --------------------------------------------------------
    # Visualizzazione task
    # --------------------------------------------------------

    st.divider()

    st.header(
        "Struttura task da creare"
    )

    visualizza_task(
        titolo="01 - Ritiro e Deprovisioning Asset → GHDW",
        destinatario="GHDW",
        contenuto=task_01,
        key="01"
    )

    visualizza_task(
        titolo="02 - Deprovisioning Utente → GIIT",
        destinatario="GIIT",
        contenuto=task_02,
        key="02"
    )

    visualizza_task(
        titolo="03 - Gestione Asset - Firma Digitale",
        destinatario="GSI",
        contenuto=task_03,
        key="03"
    )

    visualizza_task(
        titolo="04 - Deprovisioning Servizi IT",
        destinatario="Da individuare sulla base dei gruppi AD",
        contenuto=task_04,
        key="04"
    )

    # --------------------------------------------------------
    # Evidenze gruppi AD
    # --------------------------------------------------------

    with st.expander(
        "Evidenze gruppi AD rilevate",
        expanded=bool(gruppi_ad)
    ):

        if gruppi_ad:
            for group in gruppi_ad:
                st.write(
                    f"- {group}"
                )

        else:
            st.info(
                "Nessun gruppo AD aggiuntivo rilevato."
            )

    # --------------------------------------------------------
    # Avvisi
    # --------------------------------------------------------

    if warnings:
        with st.expander(
            "⚠️ Avvisi",
            expanded=True
        ):
            for warning in unique_in_order(warnings):
                st.warning(
                    warning
                )

    # --------------------------------------------------------
    # Diagnostica file
    # --------------------------------------------------------

    if mg_file is not None:
        with st.expander(
            "Diagnostica file Estr_MembriGruppi"
        ):
            st.write(
                "Colonne rilevate:",
                (
                    mg_df.columns.tolist()
                    if not mg_df.empty
                    else []
                )
            )

            st.write(
                "Numero gruppi AD rilevati:",
                len(gruppi_ad)
            )

    st.success(
        f"Struttura di deprovisioning generata correttamente "
        f"per '{sam}'."
    )
