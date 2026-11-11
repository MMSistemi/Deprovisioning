# -*- coding: utf-8 -*-

import re
import csv
import io
import zipfile
from datetime import datetime
from typing import List, Set, Tuple, Optional

import streamlit as st
import pandas as pd


# ============================================================
# CONFIGURAZIONE PAGINA
# ============================================================

st.set_page_config(
    page_title="Deprovisioning Utente",
    layout="wide"
)

st.title("Deprovisioning Utente")


# ============================================================
# COSTANTI CSV
# ============================================================

HEADER_UTENTE = [
    "sAMAccountName",
    "Creation",
    "OU",
    "Name",
    "DisplayName",
    "cn",
    "GivenName",
    "Surname",
    "employeeNumber",
    "employeeID",
    "department",
    "Description",
    "passwordNeverExpired",
    "ExpireDate",
    "userprincipalname",
    "mail",
    "mobile",
    "RimozioneGruppo",
    "InserimentoGruppo",
    "disable",
    "moveToOU",
    "telephoneNumber",
    "company"
]


HEADER_COMPUTER = [
    "Computer",
    "OU",
    "add_mail",
    "remove_mail",
    "add_mobile",
    "remove_mobile",
    "add_userprincipalname",
    "remove_userprincipalname",
    "disable",
    "moveToOU"
]


# ============================================================
# CANDIDATI COLONNE EN/IT
# ============================================================

# ------------------------------------------------------------
# Membri e gruppi AD
# ------------------------------------------------------------

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


# ------------------------------------------------------------
# Distribution List
# ------------------------------------------------------------

CAND_DL_GROUP_ADDR = [
    "Distribution Group Primary SMTP address",
    "PrimarySmtpAddress",
    "Primary SMTP address",
    "SMTP",
    "Email",
    "E-mail",
    "GroupName",
    "DisplayName",
    "Group",
    "Gruppo",
    "NomeGruppo",
    "Nome",
    "NomeVisualizzato"
]


CAND_DL_MEMBER = [
    "Member Alias",
    "Alias Membro",
    "MemberAlias",
    "MemberUserPrincipalName",
    "UserPrincipalNameMembro",
    "Member",
    "Membro",
    "userPrincipalName",
    "UPN",
    "MemberEmail",
    "EmailMembro",
    "Email",
    "E-mail"
]


CAND_DL_GROUP = [
    "Distribution Group",
    "Gruppo di distribuzione",
    "Group",
    "Gruppo",
    "GroupName",
    "NomeGruppo",
    "DisplayName",
    "Nome",
    "NomeVisualizzato"
]


# ------------------------------------------------------------
# Shared Mailbox
# ------------------------------------------------------------

CAND_SM_MEMBER = [
    "MemberUserPrincipalName",
    "UserPrincipalNameMembro",
    "Member",
    "Membro",
    "userPrincipalName",
    "UPN",
    "MemberEmail",
    "EmailMembro",
    "Member Mail",
    "Email",
    "E-mail"
]


CAND_SM_MAILBOX_ADDR = [
    "EmailAddress",
    "PrimarySmtpAddress",
    "Primary SMTP address",
    "SMTP",
    "Email",
    "E-mail",
    "DisplayName",
    "Mailbox",
    "SharedMailbox",
    "Cassetta postale",
    "Casella condivisa"
]


CAND_SM_GROUP_NAME = [
    "Group",
    "Gruppo",
    "DisplayName",
    "Nome",
    "NomeVisualizzato",
    "Mailbox",
    "SharedMailbox",
    "Cassetta postale",
    "Casella condivisa",
    "PrimarySmtpAddress",
    "SMTP",
    "Email",
    "E-mail"
]


# ------------------------------------------------------------
# Entra ID
# ------------------------------------------------------------

CAND_ENTRA_MEMBER_UPN = [
    "MemberUserPrincipalName",
    "UserPrincipalNameMembro",
    "userPrincipalName",
    "UPN",
    "Email",
    "E-mail"
]


CAND_ENTRA_GROUP_NAME = [
    "GroupName",
    "NomeGruppo",
    "DisplayName",
    "Nome",
    "NomeVisualizzato",
    "Group",
    "Gruppo"
]


# ------------------------------------------------------------
# Device
# ------------------------------------------------------------

CAND_DEV_ENABLED = [
    "Enabled",
    "Abilitato"
]


CAND_DEV_DESC = [
    "Description",
    "Descrizione"
]


CAND_DEV_NAME = [
    "Name",
    "Nome",
    "Computer",
    "NomeComputer"
]


CAND_DEV_MAIL = [
    "Mail",
    "Email",
    "E-mail",
    "Posta"
]


CAND_DEV_MOBILE = [
    "Mobile",
    "Cellulare",
    "Telefono",
    "Phone"
]


CAND_DEV_UPN = [
    "userPrincipalName",
    "UPN",
    "NomePrincipaleUtente"
]


# ============================================================
# FUNZIONI DI NORMALIZZAZIONE
# ============================================================

_SPACES_UNDERS = re.compile(r"[\s_]+")


def _norm_key(value: str) -> str:
    """
    Normalizza il nome di una colonna:
    - minuscolo;
    - rimozione spazi;
    - rimozione underscore.
    """

    return _SPACES_UNDERS.sub(
        "",
        str(value).strip().lower()
    )


def _find_col(
    df: pd.DataFrame,
    candidates: List[str]
) -> Optional[str]:
    """
    Cerca una colonna tra quelle candidate.

    Prima prova una corrispondenza esatta normalizzata.
    Successivamente prova una corrispondenza parziale.
    """

    if df is None or df.empty:
        return None

    norm_map = {
        _norm_key(column): column
        for column in df.columns
    }

    # Match esatto normalizzato
    for candidate in candidates:

        normalized_candidate = _norm_key(
            candidate
        )

        if normalized_candidate in norm_map:
            return norm_map[
                normalized_candidate
            ]

    # Match parziale
    for candidate in candidates:

        normalized_candidate = _norm_key(
            candidate
        )

        for column in df.columns:

            normalized_column = _norm_key(
                column
            )

            if normalized_candidate in normalized_column:
                return column

    return None


def _find_col_preferred(
    df: pd.DataFrame,
    preferred: List[str],
    fallback: List[str]
) -> Optional[str]:
    """
    Cerca prima nelle colonne preferite.
    Se non trova nulla utilizza le colonne di fallback.
    """

    if df is None or df.empty:
        return None

    for candidate in preferred:

        column = _find_col(
            df,
            [candidate]
        )

        if column:
            return column

    return _find_col(
        df,
        fallback
    )


def _get_any(
    df: pd.DataFrame,
    candidates: List[str]
) -> pd.Series:
    """
    Restituisce la Series corrispondente alla prima
    colonna trovata tra quelle candidate.
    """

    column = _find_col(
        df,
        candidates
    )

    if not column:
        raise KeyError(
            f"Colonne non trovate: {candidates}"
        )

    return df[column]


def _require_any(
    df: pd.DataFrame,
    required_map: dict
) -> Tuple[bool, List[str]]:
    """
    Verifica la presenza delle colonne logiche richieste.
    """

    missing = []

    for logical_name, candidates in required_map.items():

        if _find_col(
            df,
            candidates
        ) is None:

            missing.append(
                logical_name
            )

    return (
        len(missing) == 0,
        missing
    )


def _clean_series_to_list(
    series: pd.Series
) -> List[str]:
    """
    Converte una Series in una lista:
    - senza NaN;
    - senza valori vuoti;
    - senza duplicati;
    - ordinata alfabeticamente.
    """

    if series is None or series.empty:
        return []

    values = [
        str(value).strip()
        for value in series
        .dropna()
        .astype(str)
        .tolist()
    ]

    values = [
        value
        for value in values
        if value
    ]

    unique_values = {}

    for value in values:

        normalized_value = value.lower()

        if normalized_value not in unique_values:

            unique_values[
                normalized_value
            ] = value

    return sorted(
        unique_values.values(),
        key=lambda value: value.lower()
    )


def unique_in_order(
    values: List[str]
) -> List[str]:
    """
    Rimuove i duplicati mantenendo l'ordine originale.
    """

    result = []
    seen = set()

    for value in values:

        clean_value = str(
            value
        ).strip()

        normalized_value = clean_value.lower()

        if (
            clean_value
            and normalized_value not in seen
        ):

            result.append(
                clean_value
            )

            seen.add(
                normalized_value
            )

    return result


# ============================================================
# LETTURA FILE EXCEL
# ============================================================

def _read_excel_or_empty(
    uploaded_file
) -> pd.DataFrame:
    """
    Legge un file Excel.

    In caso di file mancante o errore,
    restituisce un DataFrame vuoto.
    """

    if not uploaded_file:
        return pd.DataFrame()

    try:

        uploaded_file.seek(0)

        return pd.read_excel(
            uploaded_file,
            engine="openpyxl"
        )

    except Exception:

        try:

            uploaded_file.seek(0)

            return pd.read_excel(
                uploaded_file
            )

        except Exception as error:

            st.warning(
                "Non è stato possibile leggere il file "
                f"'{uploaded_file.name}': {error}"
            )

            return pd.DataFrame()


# ============================================================
# GESTIONE STRINGHE UTENTE
# ============================================================

def normalize_sam(
    sam: str
) -> str:
    """
    Normalizza lo sAMAccountName.
    """

    return str(
        sam
    ).strip().lower()


def get_clean_sam(
    sam: str
) -> str:
    """
    Rimuove il suffisso .ext dallo sAMAccountName.
    """

    sam_lower = normalize_sam(
        sam
    )

    if sam_lower.endswith(
        ".ext"
    ):

        return sam_lower[:-4]

    return sam_lower


def infer_nome_cognome_from_sam(
    sam: str
) -> Tuple[str, str]:
    """
    Prova a ricavare nome e cognome da uno sAMAccountName
    nel formato nome.cognome oppure iniziale.cognome.
    """

    clean_sam = get_clean_sam(
        sam
    )

    parts = clean_sam.split(
        ".",
        1
    )

    if len(parts) == 2:

        nome = parts[0].strip().title()
        cognome = parts[1].strip().title()

        return nome, cognome

    return "", clean_sam.title()


def build_display_name(
    sam: str,
    nome: str = "",
    cognome: str = ""
) -> str:
    """
    Costruisce il nome da utilizzare negli oggetti,
    nei titoli e nei nomi file.
    """

    nome = str(
        nome
    ).strip().title()

    cognome = str(
        cognome
    ).strip().title()

    if nome and cognome:

        display_name = (
            f"{cognome} {nome}"
        )

    else:

        inferred_nome, inferred_cognome = (
            infer_nome_cognome_from_sam(
                sam
            )
        )

        if inferred_nome and inferred_cognome:

            display_name = (
                f"{inferred_cognome} "
                f"{inferred_nome}"
            )

        else:

            display_name = (
                get_clean_sam(
                    sam
                ).title()
            )

    if normalize_sam(
        sam
    ).endswith(
        ".ext"
    ):

        display_name += " (esterno)"

    return display_name


def build_file_person_name(
    sam: str,
    nome: str = "",
    cognome: str = ""
) -> str:
    """
    Genera il nome persona da usare nei nomi file.
    """

    nome = str(
        nome
    ).strip().title()

    cognome = str(
        cognome
    ).strip().title()

    if nome and cognome:

        return f"{cognome} {nome}"

    inferred_nome, inferred_cognome = (
        infer_nome_cognome_from_sam(
            sam
        )
    )

    if inferred_nome and inferred_cognome:

        return (
            f"{inferred_cognome} "
            f"{inferred_nome}"
        )

    return get_clean_sam(
        sam
    )


def sanitize_filename_part(
    value: str
) -> str:
    """
    Rimuove dai nomi file i caratteri non compatibili.
    """

    value = re.sub(
        r'[\\/:*?"<>|]',
        "",
        str(value)
    )

    value = re.sub(
        r"\s+",
        " ",
        value
    )

    return value.strip()


# ============================================================
# COSTRUZIONE CSV
# ============================================================

def auto_quote(
    fields,
    quotechar='"',
    predicate=lambda value: " " in value
):
    """
    Racchiude tra doppi apici ogni valore
    che contiene almeno uno spazio.
    """

    output = []

    for field in fields:

        value = str(
            field
        )

        if predicate(
            value
        ):

            escaped_value = value.replace(
                quotechar,
                quotechar * 2
            )

            output.append(
                f"{quotechar}"
                f"{escaped_value}"
                f"{quotechar}"
            )

        else:

            output.append(
                value
            )

    return output


def prepara_csv_standard(
    row,
    header
) -> str:
    """
    Costruisce un CSV mantenendo la regola esistente:
    i campi contenenti spazi vengono racchiusi
    tra doppi apici.
    """

    buffer = io.StringIO()

    writer = csv.writer(
        buffer,
        quoting=csv.QUOTE_NONE,
        escapechar="\\",
        lineterminator="\n"
    )

    quoted_row = auto_quote(
        row,
        quotechar='"',
        predicate=lambda value: " " in str(value)
    )

    writer.writerow(
        header
    )

    writer.writerow(
        quoted_row
    )

    return buffer.getvalue()


def prepara_csv_device(
    row,
    include_eof=True
) -> str:
    """
    Costruisce il CSV del computer.

    La riga EOF viene mantenuta come controllo per evitare
    spostamenti accidentali dei PC nella OU dei dismessi.
    """

    buffer = io.StringIO()

    writer = csv.writer(
        buffer,
        quoting=csv.QUOTE_NONE,
        escapechar="\\",
        lineterminator="\n"
    )

    quoted_row = auto_quote(
        row,
        quotechar='"',
        predicate=lambda value: " " in str(value)
    )

    writer.writerow(
        HEADER_COMPUTER
    )

    writer.writerow(
        quoted_row
    )

    if include_eof:

        eof_row = [
            (
                "EOF-riga lasciata appositamente scritta così "
                "per verificare che nessun PC sia andato "
                "nella OU/dismessi/computer"
            )
        ]

        eof_row.extend(
            [""] * (
                len(HEADER_COMPUTER) - 1
            )
        )

        quoted_eof_row = auto_quote(
            eof_row,
            quotechar='"',
            predicate=lambda value: " " in str(value)
        )

        writer.writerow(
            quoted_eof_row
        )

    return buffer.getvalue()


# ============================================================
# RICERCA UTENTE NEI FILE
# ============================================================

def build_user_candidates(
    sam: str
) -> Set[str]:
    """
    Genera i possibili identificativi dell'utente.
    """

    sam_lower = normalize_sam(
        sam
    )

    clean_sam = get_clean_sam(
        sam
    )

    candidates = {
        sam_lower,
        clean_sam,
        f"{sam_lower}@consip.it",
        f"{clean_sam}@consip.it"
    }

    return {
        candidate.lower()
        for candidate in candidates
        if candidate
    }


def build_member_mask(
    series: pd.Series,
    sam: str
) -> pd.Series:
    """
    Cerca l'utente in una colonna membership.

    Gestisce:
    - sAMAccountName;
    - UPN;
    - alias;
    - celle con più membri separati da ; , spazi o ritorni a capo.
    """

    if series is None or series.empty:

        return pd.Series(
            False,
            index=series.index
            if series is not None
            else None
        )

    candidates = build_user_candidates(
        sam
    )

    normalized_series = (
        series
        .fillna("")
        .astype(str)
        .str.strip()
        .str.lower()
    )

    exact_mask = normalized_series.isin(
        candidates
    )

    if exact_mask.any():
        return exact_mask

    escaped_candidates = [
        re.escape(
            candidate
        )
        for candidate in candidates
    ]

    pattern = (
        r"(?:^|[;,\s])(?:"
        + "|".join(
            escaped_candidates
        )
        + r")(?:$|[;,\s])"
    )

    return normalized_series.str.contains(
        pattern,
        regex=True,
        na=False
    )


# ============================================================
# ESTRAZIONE GRUPPI AD
# ============================================================

def estrai_rimozione_gruppi(
    sam: str,
    mg_df: pd.DataFrame
) -> Tuple[List[str], List[str]]:
    """
    Estrae i gruppi AD da rimuovere.

    Restituisce:
    - lista gruppi AD;
    - lista avvisi.
    """

    warnings = []

    if mg_df is None or mg_df.empty:

        warnings.append(
            "Non è stato caricato il file "
            "Estr_MembriGruppi."
        )

        return [], warnings

    required = {
        "membro": CAND_MG_MEMBER,
        "gruppo": CAND_MG_GROUP
    }

    ok, missing = _require_any(
        mg_df,
        required
    )

    if not ok:

        warnings.append(
            "Nel file Estr_MembriGruppi non sono stati "
            "trovati i campi: "
            + ", ".join(
                missing
            )
            + "."
        )

        return [], warnings

    member_column_name = _find_col(
        mg_df,
        CAND_MG_MEMBER
    )

    group_column_name = _find_col(
        mg_df,
        CAND_MG_GROUP
    )

    member_mask = build_member_mask(
        mg_df[
            member_column_name
        ],
        sam
    )

    groups = _clean_series_to_list(
        mg_df.loc[
            member_mask,
            group_column_name
        ]
    )

    excluded_groups = {
        "domain users",
        "utenti del dominio"
    }

    groups = [
        group
        for group in groups
        if group.strip().lower()
        not in excluded_groups
    ]

    groups = unique_in_order(
        groups
    )

    if not groups:

        warnings.append(
            "Non sono stati trovati gruppi AD "
            "da rimuovere per l'utente indicato."
        )

    return groups, warnings


# ============================================================
# ESTRAZIONE DISTRIBUTION LIST
# ============================================================

def extract_dl_for_user(
    dl_df: pd.DataFrame,
    sam: str
) -> Tuple[List[str], List[str]]:
    """
    Estrae le Distribution List associate all'utente.
    """

    warnings = []

    if dl_df is None or dl_df.empty:

        warnings.append(
            "Non è stato caricato il file DL."
        )

        return [], warnings

    group_column = _find_col_preferred(
        dl_df,
        CAND_DL_GROUP_ADDR,
        CAND_DL_GROUP
    )

    member_column = _find_col(
        dl_df,
        CAND_DL_MEMBER
    )

    if not member_column:

        warnings.append(
            "Nel file DL non è stata trovata "
            "la colonna del membro."
        )

    if not group_column:

        warnings.append(
            "Nel file DL non è stata trovata "
            "la colonna della Distribution List."
        )

    if not member_column or not group_column:
        return [], warnings

    member_mask = build_member_mask(
        dl_df[
            member_column
        ],
        sam
    )

    dl_list = _clean_series_to_list(
        dl_df.loc[
            member_mask,
            group_column
        ]
    )

    if not dl_list:

        warnings.append(
            "Non sono state trovate Distribution List "
            "associate all'utente indicato."
        )

    return dl_list, warnings


# ============================================================
# ESTRAZIONE SHARED MAILBOX
# ============================================================

def extract_sm_for_user(
    sm_df: pd.DataFrame,
    sam: str
) -> Tuple[List[str], List[str]]:
    """
    Estrae le Shared Mailbox associate all'utente.
    """

    warnings = []

    if sm_df is None or sm_df.empty:

        warnings.append(
            "Non è stato caricato il file SM."
        )

        return [], warnings

    member_column = _find_col(
        sm_df,
        CAND_SM_MEMBER
    )

    mailbox_column = _find_col_preferred(
        sm_df,
        CAND_SM_MAILBOX_ADDR,
        CAND_SM_GROUP_NAME
    )

    if not member_column:

        warnings.append(
            "Nel file SM non è stata trovata "
            "la colonna del membro."
        )

    if not mailbox_column:

        warnings.append(
            "Nel file SM non è stata trovata "
            "la colonna della Shared Mailbox."
        )

    if not member_column or not mailbox_column:
        return [], warnings

    member_mask = build_member_mask(
        sm_df[
            member_column
        ],
        sam
    )

    sm_list = _clean_series_to_list(
        sm_df.loc[
            member_mask,
            mailbox_column
        ]
    )

    if not sm_list:

        warnings.append(
            "Non sono state trovate Shared Mailbox "
            "associate all'utente indicato."
        )

    return sm_list, warnings


# ============================================================
# ESTRAZIONE GRUPPI GENERICI
# ============================================================

def extract_group_names_from_df(
    df: pd.DataFrame,
    candidates: List[str]
) -> Set[str]:
    """
    Estrae tutti i nomi dei gruppi presenti in un DataFrame.
    """

    if df is None or df.empty:
        return set()

    group_column = _find_col(
        df,
        candidates
    )

    if not group_column:
        return set()

    groups = _clean_series_to_list(
        df[
            group_column
        ]
    )

    return set(
        groups
    )


# ============================================================
# ESTRAZIONE GRUPPI ENTRA ID
# ============================================================

def extract_entra_groups_for_user(
    entra_df: pd.DataFrame,
    sam: str
) -> Tuple[List[str], List[str]]:
    """
    Estrae i gruppi Entra ID direttamente associati all'utente.
    """

    warnings = []

    if entra_df is None or entra_df.empty:

        warnings.append(
            "Non è stato caricato il file Entra."
        )

        return [], warnings

    required = {
        "UPN membro": CAND_ENTRA_MEMBER_UPN,
        "nome gruppo": CAND_ENTRA_GROUP_NAME
    }

    ok, missing = _require_any(
        entra_df,
        required
    )

    if not ok:

        warnings.append(
            "Nel file Entra non sono stati trovati "
            "i campi: "
            + ", ".join(
                missing
            )
            + "."
        )

        return [], warnings

    member_column = _find_col(
        entra_df,
        CAND_ENTRA_MEMBER_UPN
    )

    group_column = _find_col(
        entra_df,
        CAND_ENTRA_GROUP_NAME
    )

    member_mask = build_member_mask(
        entra_df[
            member_column
        ],
        sam
    )

    entra_groups = _clean_series_to_list(
        entra_df.loc[
            member_mask,
            group_column
        ]
    )

    if not entra_groups:

        warnings.append(
            "Non sono stati trovati gruppi Entra ID "
            "associati all'utente indicato."
        )

    return entra_groups, warnings


def filter_entra_groups(
    entra_groups: List[str],
    dl_df: pd.DataFrame,
    sm_df: pd.DataFrame,
    mg_df: pd.DataFrame
) -> List[str]:
    """
    Dalla lista dei gruppi Entra rimuove:

    - gruppi già identificati come DL;
    - gruppi già identificati come Shared Mailbox;
    - gruppi già presenti nei gruppi AD;
    - O365 Copilot Plus;
    - O365 Teams Premium.
    """

    dl_groups = extract_groupROUP
    )

    sm_groups = extract_group_names_from_df(
        sm_df,
        CAND_SM_MAILBOX_ADDR
        + CAND_SM_GROUP_NAME
    )

    mg_groups = extract_group_names_from_df(
        mg_df,
        CAND_MG_GROUP
    )

    other_groups = (
        dl_groups
        | sm_groups
        | mg_groups
    )

    other_groups_lower = {
        str(group).strip().lower()
        for group in other_groups
        if str(group).strip()
    }

    excluded_specific = {
        "o365 copilot plus",
        "o365 teams premium"
    }

    filtered_groups = []

    for group in entra_groups:

        clean_group = str(
            group
        ).strip()

        normalized_group = clean_group.lower()

        if not clean_group:
            continue

        if normalized_group in other_groups_lower:
            continue

        if normalized_group in excluded_specific:
            continue

        filtered_groups.append(
            clean_group
        )

    return sorted(
        unique_in_order(
            filtered_groups
        ),
        key=lambda value: value.lower()
    )


# ============================================================
# GENERAZIONE CSV UTENTE
# ============================================================

def genera_csv_utente(
    sam: str,
    mg_df: pd.DataFrame
) -> Tuple[
    str,
    List[str],
    List[str],
    List[str]
]:
    """
    Genera il CSV dell'utente per:

    - rimozione gruppi AD;
    - disabilitazione;
    - spostamento nella OU prevista.
    """

    gruppi_ad, warnings = estrai_rimozione_gruppi(
        sam,
        mg_df
    )

    rimozione_gruppi = ";".join(
        gruppi_ad
    )

    row_map = {
        header: ""
        for header in HEADER_UTENTE
    }

    row_map[
        "sAMAccountName"
    ] = normalize_sam(
        sam
    )

    row_map[
        "RimozioneGruppo"
    ] = rimozione_gruppi

    row_map[
        "disable"
    ] = "SI"

    row_map[
        "moveToOU"
    ] = "SI"

    row = [
        row_map.get(
            header,
            ""
        )
        for header in HEADER_UTENTE
    ]

    csv_content = prepara_csv_standard(
        row,
        HEADER_UTENTE
    )

    return (
        csv_content,
        row,
        gruppi_ad,
        warnings
    )


# ============================================================
# GENERAZIONE CSV COMPUTER
# ============================================================

def genera_device_csv(
    sam: str,
    device_df: pd.DataFrame
) -> Tuple[
    Optional[str],
    Optional[List[str]],
    Optional[str],
    List[str]
]:
    """
    Genera il CSV del computer associato all'utente.

    Restituisce:
    - contenuto CSV;
    - riga CSV;
    - nome computer;
    - lista avvisi.
    """

    warnings = []

    if device_df is None or device_df.empty:

        warnings.append(
            "Non è stato caricato il file Estr_Device."
        )

        return None, None, None, warnings

    enabled_column = _find_col(
        device_df,
        CAND_DEV_ENABLED
    )

    description_column = _find_col(
        device_df,
        CAND_DEV_DESC
    )

    name_column = _find_col(
        device_df,
        CAND_DEV_NAME
    )

    mail_column = _find_col(
        device_df,
        CAND_DEV_MAIL
    )

    mobile_column = _find_col(
        device_df,
        CAND_DEV_MOBILE
    )

    upn_column = _find_col(
        device_df,
        CAND_DEV_UPN
    )

    if not enabled_column:

        warnings.append(
            "Nel file Estr_Device non è stata trovata "
            "la colonna Enabled/Abilitato."
        )

    if not description_column:

        warnings.append(
            "Nel file Estr_Device non è stata trovata "
            "la colonna Description/Descrizione."
        )

    if not name_column:

        warnings.append(
            "Nel file Estr_Device non è stata trovata "
            "la colonna Name/Nome/Computer."
        )

    if (
        not enabled_column
        or not description_column
        or not name_column
    ):

        return None, None, None, warnings

    enabled_series = device_df[
        enabled_column
    ]

    if pd.api.types.is_bool_dtype(
        enabled_series
    ):

        enabled_mask = (
            enabled_series == True
        )

    else:

        enabled_mask = (
            enabled_series
            .fillna("")
            .astype(str)
            .str.strip()
            .str.lower()
            .isin(
                [
                    "true",
                    "1",
                    "yes",
                    "si",
                    "sì",
                    "vero"
                ]
            )
        )

    enabled_devices = device_df[
        enabled_mask
    ].copy()

    if enabled_devices.empty:

        warnings.append(
            "Nel file Estr_Device non sono presenti "
            "computer abilitati."
        )

        return None, None, None, warnings

    sam_lower = normalize_sam(
        sam
    )

    clean_sam = get_clean_sam(
        sam
    )

    description_values = (
        enabled_devices[
            description_column
        ]
        .fillna("")
        .astype(str)
        .str.lower()
    )

    exact_pattern = (
        rf"(?:^|\s|[-;,{re.escape('|')},])"
        rf"(?:{re.escape(sam_lower)}|"
        rf"{re.escape(clean_sam)})"
        rf"(?:$|\s|[-;,{re.escape('|')},])"
    )

    device_mask = description_values.str.contains(
        exact_pattern,
        regex=True,
        na=False
    )

    if not device_mask.any():

        device_mask = (
            description_values.str.contains(
                re.escape(
                    sam_lower
                ),
                regex=True,
                na=False
            )
        )

    if (
        not device_mask.any()
        and clean_sam != sam_lower
    ):

        device_mask = (
            description_values.str.contains(
                re.escape(
                    clean_sam
                ),
                regex=True,
                na=False
            )
        )

    if not device_mask.any():

        warnings.append(
            "Non è stato trovato un computer associato "
            "all'utente indicato."
        )

        return None, None, None, warnings

    matched_devices = enabled_devices[
        device_mask
    ]

    if len(
        matched_devices
    ) > 1:

        warnings.append(
            "Sono stati trovati più computer associati "
            "all'utente. Nel CSV è stato utilizzato "
            "il primo risultato."
        )

    row_data = matched_devices.iloc[
        0
    ]

    computer_name = str(
        row_data.get(
            name_column,
            ""
        )
    ).strip()

    if not computer_name:

        warnings.append(
            "Il computer trovato non contiene "
            "un nome valido."
        )

        return None, None, None, warnings

    def has_value(
        column_name: Optional[str]
    ) -> str:

        if not column_name:
            return ""

        value = row_data.get(
            column_name,
            ""
        )

        if pd.isna(
            value
        ):
            return ""

        clean_value = str(
            value
        ).strip()

        if not clean_value:
            return ""

        return "SI"

    remove_mail = has_value(
        mail_column
    )

    remove_mobile = has_value(
        mobile_column
    )

    remove_upn = has_value(
        upn_column
    )

    if not any(
        [
            remove_mail,
            remove_mobile,
            remove_upn
        ]
    ):

        warnings.append(
            "Il computer trovato non presenta riferimenti "
            "Mail, Mobile o userPrincipalName da rimuovere."
        )

        return None, None, computer_name, warnings

    row = [
        computer_name,
        "",
        "",
        remove_mail,
        "",
        remove_mobile,
        "",
        remove_upn,
        "",
        ""
    ]

    csv_content = prepara_csv_device(
        row,
        include_eof=True
    )

    return (
        csv_content,
        row,
        computer_name,
        warnings
    )


# ============================================================
# GENERAZIONE ISTRUZIONI DEPROVISIONING
# ============================================================

def genera_istruzioni_deprovisioning(
    sam: str,
    dl_list: List[str],
    sm_list: List[str],
    entra_groups: List[str],
    gruppi_ad: List[str]
) -> str:
    """
    Genera il file testuale contenente tutte
    le istruzioni di deprovisioning.
    """

    sam_lower = normalize_sam(
        sam
    )

    user_email = (
        f"{sam_lower}@consip.it"
    )

    lines = [
        "Ciao,",
        (
            f"per l'utenza {sam_lower} "
            f"({user_email}) si richiedono "
            "le seguenti attività:"
        ),
        ""
    ]

    step = 1

    fixed_steps = [
        (
            "Disabilitare l'invio verso l'utente "
            "(Message Delivery Restrictions)."
        ),
        (
            "Impostare la casella come nascosta "
            "dalla Rubrica."
        ),
        (
            "Disabilitare l'accesso alla Mailbox "
            "(Mailbox features: disabilitare "
            "protocolli e OWA)."
        ),
        (
            "Estrarre il PST tramite O365 eDiscovery "
            "e archiviarlo nel percorso previsto per "
            "il backup delle email cancellate, "
            "utilizzando un archivio 7z protetto "
            "dalla password condivisa."
        ),
        (
            "Rimuovere i ruoli assegnati all'utenza."
        ),
        (
            "Rimuovere le applicazioni assegnate "
            "all'utenza Azure/Entra ID."
        )
    ]

    for description in fixed_steps:

        lines.append(
            f"{step}. {description}"
        )

        step += 1

    if dl_list:

        lines.append(
            f"{step}. Rimuovere l'utente "
            "dalle seguenti Liste di Distribuzione:"
        )

        for dl in dl_list:

            lines.append(
                f"   - {dl}"
            )

    else:

        lines.append(
            f"{step}. Nessuna Lista di Distribuzione "
            "rilevata per l'utente."
        )

    step += 1

    lines.append(
        f"{step}. Disabilitare l'account "
        "Azure/Entra ID."
    )

    step += 1

    if sm_list:

        lines.append(
            f"{step}. Rimuovere le abilitazioni "
            "dalle seguenti Shared Mailbox:"
        )

        for shared_mailbox in sm_list:

            lines.append(
                f"   - {shared_mailbox}"
            )

    else:

        lines.append(
            f"{step}. Nessuna Shared Mailbox "
            "rilevata per l'utente."
        )

    step += 1

    if entra_groups:

        lines.append(
            f"{step}. Rimuovere l'utente "
            "dai seguenti gruppi Entra ID:"
        )

        for group in entra_groups:

            lines.append(
                f"   - {group}"
            )

    else:

        lines.append(
            f"{step}. Nessun gruppo Entra ID "
            "aggiuntivo rilevato dopo lo scremamento."
        )

    step += 1

    if gruppi_ad:

        lines.append(
            f"{step}. Rimuovere l'utente "
            "dai gruppi AD indicati nel CSV utente."
        )

    else:

        lines.append(
            f"{step}. Nel file Estr_MembriGruppi "
            "non sono stati rilevati gruppi AD "
            "da rimuovere."
        )

    step += 1

    final_steps = [
        "Cancellare la foto dell'utente da Azure, se presente.",
        "Rimuovere l'abilitazione Wi-Fi.",
        "Disabilitare l'utenza Active Directory.",
        "Spostare l'utenza nella OU prevista per gli utenti dismessi."
    ]

    for description in final_steps:

        lines.append(
            f"{step}. {description}"
        )

        step += 1

    return "\n".join(
        lines
    )


# ============================================================
# INTERFACCIA STREAMLIT
# ============================================================

st.subheader(
    "Modulo Deprovisioning Utente"
)

colonna_1, colonna_2 = st.columns(
    2
)


with colonna_1:

    sam = st.text_input(
        "Nome utente (sAMAccountName)",
        "",
        help=(
            "Esempi: mario.rossi oppure "
            "m.rossi.ext"
        )
    ).strip().lower()

    nome = st.text_input(
        "Nome",
        "",
        help=(
            "Campo facoltativo. Se lasciato vuoto "
            "verrà ricavato dallo sAMAccountName."
        )
    ).strip().title()

    cognome = st.text_input(
        "Cognome",
        "",
        help=(
            "Campo facoltativo. Se lasciato vuoto "
            "verrà ricavato dallo sAMAccountName."
        )
    ).strip().title()


with colonna_2:

    data_deprovisioning = st.text_input(
        "Data deprovisioning (gg/mm/aaaa)",
        datetime.now().strftime(
            "%d/%m/%Y"
        )
    ).strip()

    note_aggiuntive = st.text_area(
        "Note aggiuntive",
        "",
        help=(
            "Le note verranno inserite "
            "nelle istruzioni destinate a GIIT."
        )
    ).strip()


st.markdown("---")


# ============================================================
# CARICAMENTO FILE DI EXPORT
# ============================================================

st.markdown(
    "### File necessari per la verifica delle profilazioni"
)

col_file_1, col_file_2 = st.columns(
    2
)


with col_file_1:

    mg_file = st.file_uploader(
        "Carica Estr_MembriGruppi (Excel)",
        type=["xlsx"],
        key="deprovisioning_mg"
    )

    dl_file = st.file_uploader(
        "Carica file DL (Excel)",
        type=["xlsx"],
        key="deprovisioning_dl"
    )

    sm_file = st.file_uploader(
        "Carica file SM (Excel)",
        type=["xlsx"],
        key="deprovisioning_sm"
    )


with col_file_2:

    entra_file = st.file_uploader(
        "Carica file Entra (Excel)",
        type=["xlsx"],
        key="deprovisioning_entra"
    )

    device_file = st.file_uploader(
        "Carica Estr_Device (Excel)",
        type=["xlsx"],
        key="deprovisioning_device"
    )


st.caption(
    "I file possono avere intestazioni in italiano o inglese. "
    "Se un file non viene caricato, il flusso viene comunque "
    "generato e viene mostrato un avviso."
)


# ============================================================
# GENERAZIONE FLUSSO
# ============================================================

if st.button(
    "Genera Deprovisioning",
    type="primary"
):

    # --------------------------------------------------------
    # Validazioni
    # --------------------------------------------------------

    if not sam:

        st.error(
            "Inserire lo sAMAccountName "
            "prima di generare il deprovisioning."
        )

        st.stop()

    if " " in sam:

        st.error(
            "Lo sAMAccountName non può contenere spazi."
        )

        st.stop()

    if not data_deprovisioning:

        st.error(
            "Inserire la data di deprovisioning."
        )

        st.stop()

    # --------------------------------------------------------
    # Lettura file
    # --------------------------------------------------------

    mg_df = _read_excel_or_empty(
        mg_file
    )

    dl_df = _read_excel_or_empty(
        dl_file
    )

    sm_df = _read_excel_or_empty(
        sm_file
    )

    entra_df = _read_excel_or_empty(
        entra_file
    )

    device_df = _read_excel_or_empty(
        device_file
    )

    # --------------------------------------------------------
    # Dati principali
    # --------------------------------------------------------

    sam_lower = normalize_sam(
        sam
    )

    user_email = (
        f"{sam_lower}@consip.it"
    )

    display_name = build_display_name(
        sam_lower,
        nome,
        cognome
    )

    persona_nome_file = sanitize_filename_part(
        build_file_person_name(
            sam_lower,
            nome,
            cognome
        )
    )

    persona_nome_underscore = (
        persona_nome_file.replace(
            " ",
            "_"
        )
    )

    # --------------------------------------------------------
    # Estrazione gruppi AD e generazione CSV utente
    # --------------------------------------------------------

    (
        csv_utente,
        row_utente,
        gruppi_ad,
        warnings_utente
    ) = genera_csv_utente(
        sam_lower,
        mg_df
    )

    # --------------------------------------------------------
    # Estrazione DL
    # --------------------------------------------------------

    dl_list, warnings_dl = extract_dl_for_user(
        dl_df,
        sam_lower
    )

    # --------------------------------------------------------
    # Estrazione Shared Mailbox
    # --------------------------------------------------------

    sm_list, warnings_sm = extract_sm_for_user(
        sm_df,
        sam_lower
    )

    # --------------------------------------------------------
    # Estrazione gruppi Entra ID
    # --------------------------------------------------------

    (
        entra_groups_raw,
        warnings_entra
    ) = extract_entra_groups_for_user(
        entra_df,
        sam_lower
    )

    entra_groups = filter_entra_groups(
        entra_groups_raw,
        dl_df,
        sm_df,
        mg_df
    )

    if (
        entra_groups_raw
        and not entra_groups
    ):

        warnings_entra.append(
            "I gruppi Entra ID trovati sono stati "
            "esclusi perché già presenti nei file "
            "DL, SM o gruppi AD, oppure perché "
            "appartenenti alle esclusioni previste."
        )

    # --------------------------------------------------------
    # Generazione CSV computer
    # --------------------------------------------------------

    (
        csv_computer,
        row_computer,
        computer_name,
        warnings_device
    ) = genera_device_csv(
        sam_lower,
        device_df
    )

    # --------------------------------------------------------
    # Generazione istruzioni
    # --------------------------------------------------------

    istruzioni_deprovisioning = (
        genera_istruzioni_deprovisioning(
            sam_lower,
            dl_list,
            sm_list,
            entra_groups,
            gruppi_ad
        )
    )

    if note_aggiuntive:

        istruzioni_deprovisioning += (
            "\n\nNote aggiuntive:\n"
            f"{note_aggiuntive}"
        )

    # --------------------------------------------------------
    # Nomi file
    # --------------------------------------------------------

    filename_utente = (
        f"1_{persona_nome_file}"
        "_deprovisioning_utenza.csv"
    )

    filename_istruzioni = (
        f"2_{persona_nome_file}"
        "_istruzioni_deprovisioning.txt"
    )

    if csv_computer:

        filename_computer = (
            f"3_{persona_nome_underscore}"
            "_computer_remove.csv"
        )

    else:

        filename_computer = ""

    # --------------------------------------------------------
    # Testo richiesta GIIT
    # --------------------------------------------------------

    subject_giit = (
        "01 - Deprovisioning Utente "
        "(Utenza/Casella/Profilazioni/Computer) - "
        f"{display_name} → GIIT"
    )

    message_giit_lines = [
        "Ciao,",
        (
            f"si richiede il deprovisioning "
            f"dell'utenza {sam_lower} "
            f"relativa a {display_name}, "
            f"da effettuare in data "
            f"{data_deprovisioning}."
        ),
        "",
        (
            "Elaborare i file contenuti nella "
            "cartella compressa .zip "
            "nel seguente ordine:"
        ),
        (
            f"1. {filename_utente} per la rimozione "
            "dei gruppi AD, la disabilitazione "
            "dell'utenza e lo spostamento "
            "nella OU prevista."
        ),
        (
            f"2. {filename_istruzioni} per le attività "
            "relative alla casella, alle Liste "
            "di Distribuzione, alle Shared Mailbox "
            "e alle profilazioni Entra ID."
        )
    ]

    if csv_computer:

        message_giit_lines.append(
            f"3. {filename_computer} per la rimozione "
            f"dei riferimenti dell'utente dal computer "
            f"{computer_name}."
        )

    else:

        message_giit_lines.append(
            "3. Non è stato generato alcun CSV computer "
            "perché non è stato trovato un dispositivo "
            "associato con riferimenti da rimuovere."
        )

    message_giit = "\n".join(
        message_giit_lines
    )

    # --------------------------------------------------------
    # Guida completa
    # --------------------------------------------------------

    guida_flusso = "\n\n".join(
        [
            subject_giit,
            message_giit,
            (
                "ISTRUZIONI OPERATIVE DI DEPROVISIONING"
            ),
            istruzioni_deprovisioning
        ]
    )

    filename_guida = (
        f"Guida_Flusso_Deprovisioning_"
        f"{persona_nome_underscore}.txt"
    )

    # --------------------------------------------------------
    # Creazione ZIP
    # --------------------------------------------------------

    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(
        zip_buffer,
        "w",
        zipfile.ZIP_DEFLATED
    ) as zip_file:

        zip_file.writestr(
            filename_utente,
            csv_utente
        )

        zip_file.writestr(
            filename_istruzioni,
            istruzioni_deprovisioning
        )

        if (
            csv_computer
            and filename_computer
        ):

            zip_file.writestr(
                filename_computer,
                csv_computer
            )

        zip_file.writestr(
            filename_guida,
            guida_flusso
        )

    zip_buffer.seek(
        0
    )

    bundle_name = (
        f"Deprovisioning_"
        f"{persona_nome_underscore}.zip"
    )

    # ========================================================
    # VISUALIZZAZIONE FLUSSO
    # ========================================================

    st.markdown("---")

    st.subheader(
        subject_giit
    )

    st.text(
        message_giit
    )

    st.markdown(
        "**File di competenza GIIT:**"
    )

    # --------------------------------------------------------
    # File 1 - Utente
    # --------------------------------------------------------

    st.markdown(
        f"**1. {filename_utente}**"
    )

    st.info(
        "Il file gestisce la rimozione dei gruppi AD, "
        "la disabilitazione dell'utenza e lo spostamento "
        "nella OU prevista."
    )

    st.dataframe(
        pd.DataFrame(
            [row_utente],
            columns=HEADER_UTENTE
        ),
        use_container_width=True
    )

    st.download_button(
        "Scarica 1 - Deprovisioning Utenza",
        data=csv_utente,
        file_name=filename_utente,
        mime="text/csv",
        key="download_deprovisioning_utente"
    )

    # --------------------------------------------------------
    # Gruppi AD rilevati
    # --------------------------------------------------------

    with st.expander(
        "Gruppi AD rilevati"
    ):

        if gruppi_ad:

            for group in gruppi_ad:

                st.write(
                    f"- {group}"
                )

        else:

            st.info(
                "Nessun gruppo AD rilevato."
            )

    # --------------------------------------------------------
    # File 2 - Istruzioni
    # --------------------------------------------------------

    st.markdown(
        f"**2. {filename_istruzioni}**"
    )

    st.text(
        istruzioni_deprovisioning
    )

    st.download_button(
        "Scarica 2 - Istruzioni Deprovisioning",
        data=istruzioni_deprovisioning,
        file_name=filename_istruzioni,
        mime="text/plain",
        key="download_istruzioni_deprovisioning"
    )

    # --------------------------------------------------------
    # File 3 - Computer
    # --------------------------------------------------------

    if (
        csv_computer
        and row_computer
        and filename_computer
    ):

        st.markdown(
            f"**3. {filename_computer}**"
        )

        st.info(
            "Il file rimuove dal computer i riferimenti "
            "Mail, Mobile e userPrincipalName "
            "associati all'utente."
        )

        st.dataframe(
            pd.DataFrame(
                [row_computer],
                columns=HEADER_COMPUTER
            ),
            use_container_width=True
        )

        st.download_button(
            "Scarica 3 - Computer Remove",
            data=csv_computer,
            file_name=filename_computer,
            mime="text/csv",
            key="download_deprovisioning_computer"
        )

    else:

        st.markdown(
            "**3. CSV Computer**"
        )

        st.warning(
            "Nessun CSV computer generato. "
            "Non è stato trovato un computer associato "
            "con riferimenti da rimuovere."
        )

    # --------------------------------------------------------
    # Bundle completo
    # --------------------------------------------------------

    st.markdown("---")

    st.markdown(
        "**Bundle completo del deprovisioning:**"
    )

    st.download_button(
        "📦 Scarica bundle completo ZIP",
        data=zip_buffer.getvalue(),
        file_name=bundle_name,
        mime="application/zip",
        key="download_deprovisioning_bundle"
    )

    # --------------------------------------------------------
    # Avvisi
    # --------------------------------------------------------

    all_warnings = unique_in_order(
        warnings_utente
        + warnings_dl
        + warnings_sm
        + warnings_entra
        + warnings_device
    )

    if all_warnings:

        st.markdown("---")

        with st.expander(
            "Avvisi e controlli da effettuare",
            expanded=True
        ):

            for warning in all_warnings:

                st.warning(
                    warning
                )

    # --------------------------------------------------------
    # Diagnostica file
    # --------------------------------------------------------

    with st.expander(
        "Diagnostica file caricati"
    ):

        st.write(
            {
                "Estr_MembriGruppi_colonne": (
                    mg_df.columns.tolist()
                    if not mg_df.empty
                    else []
                ),
                "DL_colonne": (
                    dl_df.columns.tolist()
                    if not dl_df.empty
                    else []
                ),
                "SM_colonne": (
                    sm_df.columns.tolist()
                    if not sm_df.empty
                    else []
                ),
                "Entra_colonne": (
                    entra_df.columns.tolist()
                    if not entra_df.empty
                    else []
                ),
                "Device_colonne": (
                    device_df.columns.tolist()
                    if not device_df.empty
                    else []
                )
            }
        )

        st.write(
            {
                "MG_member_col": (
                    _find_col(
                        mg_df,
                        CAND_MG_MEMBER
                    )
                    if not mg_df.empty
                    else None
                ),
                "MG_group_col": (
                    _find_col(
                        mg_df,
                        CAND_MG_GROUP
                    )
                    if not mg_df.empty
                    else None
                ),
                "DL_member_col": (
                    _find_col(
                        dl_df,
                        CAND_DL_MEMBER
                    )
                    if not dl_df.empty
                    else None
                ),
                "DL_group_col": (
                    _find_col_preferred(
                        dl_df,
                        CAND_DL_GROUP_ADDR,
                        CAND_DL_GROUP
                    )
                    if not dl_df.empty
                    else None
                ),
                "SM_member_col": (
                    _find_col(
                        sm_df,
                        CAND_SM_MEMBER
                    )
                    if not sm_df.empty
                    else None
                ),
                "SM_mailbox_col": (
                    _find_col_preferred(
                        sm_df,
                        CAND_SM_MAILBOX_ADDR,
                        CAND_SM_GROUP_NAME
                    )
                    if not sm_df.empty
                    else None
                ),
                "Entra_member_col": (
                    _find_col(
                        entra_df,
                        CAND_ENTRA_MEMBER_UPN
                    )
                    if not entra_df.empty
                    else None
                ),
                "Entra_group_col": (
                    _find_col(
                        entra_df,
                        CAND_ENTRA_GROUP_NAME
                    )
                    if not entra_df.empty
                    else None
                ),
                "Device_name_col": (
                    _find_col(
                        device_df,
                        CAND_DEV_NAME
                    )
                    if not device_df.empty
                    else None
                ),
                "Device_description_col": (
                    _find_col(
                        device_df,
                        CAND_DEV_DESC
                    )
                    if not device_df.empty
                    else None
                )
            }
        )

    st.success(
        "Flusso completo di deprovisioning "
        f"generato correttamente per '{sam_lower}'."
    )
