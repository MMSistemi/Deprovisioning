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
# HEADER CSV
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

CAND_DL_GROUP_ADDR = [
    "Distribution Group Primary SMTP address",
    "PrimarySmtpAddress",
    "SMTP",
    "Email",
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
    "Email"
]

CAND_SM_MAILBOX_ADDR = [
    "EmailAddress",
    "PrimarySmtpAddress",
    "SMTP",
    "Email",
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
    "Email"
]

CAND_ENTRA_MEMBER = [
    "MemberUserPrincipalName",
    "UserPrincipalNameMembro",
    "userPrincipalName",
    "UPN",
    "Email",
    "E-mail"
]

CAND_ENTRA_GROUP = [
    "GroupName",
    "NomeGruppo",
    "DisplayName",
    "Nome",
    "NomeVisualizzato",
    "Group",
    "Gruppo"
]

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
# FUNZIONI GENERALI
# ============================================================

_SPACES_UNDERS = re.compile(r"[\s_]+")


def norm_col(value):
    return _SPACES_UNDERS.sub(
        "",
        str(value).strip().lower()
    )


def find_col(df, candidates):
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
            if key in norm_col(col):
                return col

    return None


def find_col_preferred(df, preferred, fallback):
    for candidate in preferred:
        result = find_col(
            df,
            [candidate]
        )

        if result:
            return result

    return find_col(
        df,
        fallback
    )


def read_excel(uploaded_file):
    if not uploaded_file:
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
    if series is None or series.empty:
        return []

    result = []

    for value in series.dropna():
        value = str(value).strip()

        if (
            value
            and value.lower()
            not in [x.lower() for x in result]
        ):
            result.append(value)

    return sorted(
        result,
        key=str.lower
    )


def unique_in_order(values):
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
    return str(sam).strip().lower()


def clean_sam(sam):
    sam = normalize_sam(sam)

    if sam.endswith(".ext"):
        return sam[:-4]

    return sam


def user_candidates(sam):
    sam = normalize_sam(sam)
    clean = clean_sam(sam)

    return {
        sam,
        clean,
        f"{sam}@consip.it",
        f"{clean}@consip.it"
    }


def member_mask(series, sam):
    candidates = {
        value.lower()
        for value in user_candidates(sam)
    }

    values = (
        series
        .fillna("")
        .astype(str)
        .str.strip()
        .str.lower()
    )

    mask = values.isin(candidates)

    if mask.any():
        return mask

    escaped = [
        re.escape(value)
        for value in candidates
    ]

    pattern = (
        r"(?:^|[;,\s])(?:"
        + "|".join(escaped)
        + r")(?:$|[;,\s])"
    )

    return values.str.contains(
        pattern,
        regex=True,
        na=False
    )


# ============================================================
# NOME VISUALIZZATO
# ============================================================

def build_display_name(sam):
    clean = clean_sam(sam)

    parts = clean.split(
        ".",
        1
    )

    if len(parts) == 2:
        nome = parts[0].capitalize()
        cognome = parts[1].capitalize()

        display = f"{cognome} {nome}"

    else:
        display = clean.capitalize()

    if normalize_sam(sam).endswith(".ext"):
        display += " (esterno)"

    return display


def build_filename_person(sam):
    display = build_display_name(sam)

    display = display.replace(
        " (esterno)",
        ""
    )

    display = re.sub(
        r'[\\/:*?"<>|]',
        "",
        display
    )

    return display.strip()


# ============================================================
# CSV STANDARD
# ============================================================

def auto_quote(fields):
    result = []

    for field in fields:
        value = str(field)

        if " " in value:
            value = value.replace(
                '"',
                '""'
            )

            result.append(
                f'"{value}"'
            )

        else:
            result.append(value)

    return result


def prepara_csv(row, header):
    buffer = io.StringIO()

    writer = csv.writer(
        buffer,
        quoting=csv.QUOTE_NONE,
        escapechar="\\",
        lineterminator="\n"
    )

    writer.writerow(header)
    writer.writerow(auto_quote(row))

    return buffer.getvalue()


# ============================================================
# GRUPPI AD
# ============================================================

def extract_ad_groups(sam, df):
    if df.empty:
        return [], [
            "File Estr_MembriGruppi non caricato."
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
            "Colonne Member/Group non trovate "
            "nel file Estr_MembriGruppi."
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

    excluded = {
        "domain users",
        "utenti del dominio"
    }

    groups = [
        group
        for group in groups
        if group.lower() not in excluded
    ]

    warnings = []

    if not groups:
        warnings.append(
            "Nessun gruppo AD trovato "
            "per l'utente."
        )

    return groups, warnings


# ============================================================
# DISTRIBUTION LIST
# ============================================================

def extract_dl(sam, df):
    if df.empty:
        return [], [
            "File DL non caricato."
        ]

    member_col = find_col(
        df,
        CAND_DL_MEMBER
    )

    group_col = find_col_preferred(
        df,
        CAND_DL_GROUP_ADDR,
        CAND_DL_GROUP
    )

    if not member_col or not group_col:
        return [], [
            "Colonne membro/DL non trovate "
            "nel file DL."
        ]

    mask = member_mask(
        df[member_col],
        sam
    )

    result = clean_list(
        df.loc[
            mask,
            group_col
        ]
    )

    warnings = []

    if not result:
        warnings.append(
            "Nessuna Lista di Distribuzione "
            "trovata per l'utente."
        )

    return result, warnings


# ============================================================
# SHARED MAILBOX
# ============================================================

def extract_sm(sam, df):
    if df.empty:
        return [], [
            "File SM non caricato."
        ]

    member_col = find_col(
        df,
        CAND_SM_MEMBER
    )

    mailbox_col = find_col_preferred(
        df,
        CAND_SM_MAILBOX_ADDR,
        CAND_SM_GROUP_NAME
    )

    if not member_col or not mailbox_col:
        return [], [
            "Colonne Member/EmailAddress "
            "non trovate nel file SM."
        ]

    mask = member_mask(
        df[member_col],
        sam
    )

    result = clean_list(
        df.loc[
            mask,
            mailbox_col
        ]
    )

    warnings = []

    if not result:
        warnings.append(
            "Nessuna Shared Mailbox "
            "trovata per l'utente."
        )

    return result, warnings


# ============================================================
# ENTRA ID
# ============================================================

def extract_entra(sam, df):
    if df.empty:
        return [], [
            "File Entra non caricato."
        ]

    member_col = find_col(
        df,
        CAND_ENTRA_MEMBER
    )

    group_col = find_col(
        df,
        CAND_ENTRA_GROUP
    )

    if not member_col or not group_col:
        return [], [
            "Colonne Member/GroupName "
            "non trovate nel file Entra."
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

    warnings = []

    if not groups:
        warnings.append(
            "Nessun gruppo Entra ID "
            "trovato per l'utente."
        )

    return groups, warnings


def all_group_names(df, candidates):
    if df.empty:
        return set()

    col = find_col(
        df,
        candidates
    )

    if not col:
        return set()

    return {
        str(value).strip().lower()
        for value in df[col].dropna()
        if str(value).strip()
    }


def filter_entra_groups(
    entra_groups,
    dl_df,
    sm_df,
    mg_df
):
    other_groups = set()

    other_groups |= all_group_names(
        dl_df,
        CAND_DL_GROUP_ADDR
        + CAND_DL_GROUP
    )

    other_groups |= all_group_names(
        sm_df,
        CAND_SM_MAILBOX_ADDR
        + CAND_SM_GROUP_NAME
    )

    other_groups |= all_group_names(
        mg_df,
        CAND_MG_GROUP
    )

    excluded = {
        "o365 copilot plus",
        "o365 teams premium"
    }

    result = []

    for group in entra_groups:
        key = group.strip().lower()

        if key in other_groups:
            continue

        if key in excluded:
            continue

        result.append(group)

    return unique_in_order(result)


# ============================================================
# CSV UTENTE
# ============================================================

def genera_csv_utente(
    sam,
    gruppi_ad
):
    row = [
        ""
        for _ in HEADER_UTENTE
    ]

    row[
        HEADER_UTENTE.index(
            "sAMAccountName"
        )
    ] = sam

    row[
        HEADER_UTENTE.index(
            "RimozioneGruppo"
        )
    ] = ";".join(gruppi_ad)

    row[
        HEADER_UTENTE.index(
            "disable"
        )
    ] = "SI"

    row[
        HEADER_UTENTE.index(
            "moveToOU"
        )
    ] = "SI"

    csv_content = prepara_csv(
        row,
        HEADER_UTENTE
    )

    return csv_content, row


# ============================================================
# CSV COMPUTER
# ============================================================

def genera_device_csv(
    sam,
    df
):
    if df.empty:
        return None, None, None, [
            "File Estr_Device non caricato."
        ]

    enabled_col = find_col(
        df,
        CAND_DEV_ENABLED
    )

    desc_col = find_col(
        df,
        CAND_DEV_DESC
    )

    name_col = find_col(
        df,
        CAND_DEV_NAME
    )

    mail_col = find_col(
        df,
        CAND_DEV_MAIL
    )

    mobile_col = find_col(
        df,
        CAND_DEV_MOBILE
    )

    upn_col = find_col(
        df,
        CAND_DEV_UPN
    )

    if not enabled_col:
        return None, None, None, [
            "Colonna Enabled non trovata "
            "nel file Estr_Device."
        ]

    if not desc_col or not name_col:
        return None, None, None, [
            "Colonne Description/Name "
            "non trovate nel file Estr_Device."
        ]

    enabled = (
        df[enabled_col]
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

    devices = df[
        enabled
    ].copy()

    if devices.empty:
        return None, None, None, [
            "Nessun computer abilitato trovato."
        ]

    descriptions = (
        devices[desc_col]
        .fillna("")
        .astype(str)
        .str.lower()
    )

    sam_lower = normalize_sam(sam)
    base_sam = clean_sam(sam)

    mask = descriptions.str.contains(
        re.escape(sam_lower),
        regex=True,
        na=False
    )

    if (
        not mask.any()
        and base_sam != sam_lower
    ):
        mask = descriptions.str.contains(
            re.escape(base_sam),
            regex=True,
            na=False
        )

    if not mask.any():
        return None, None, None, [
            "Nessun computer associato "
            "all'utente."
        ]

    matched = devices[
        mask
    ]

    record = matched.iloc[0]

    computer = str(
        record[name_col]
    ).strip()

    def has_value(column):
        if not column:
            return ""

        value = record.get(
            column,
            ""
        )

        if pd.isna(value):
            return ""

        return (
            "SI"
            if str(value).strip()
            else ""
        )

    remove_mail = has_value(
        mail_col
    )

    remove_mobile = has_value(
        mobile_col
    )

    remove_upn = has_value(
        upn_col
    )

    if not any([
        remove_mail,
        remove_mobile,
        remove_upn
    ]):
        return None, None, computer, [
            "Il computer trovato non contiene "
            "riferimenti utente da rimuovere."
        ]

    row = [
        computer,
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

    csv_content = prepara_csv(
        row,
        HEADER_COMPUTER
    )

    # Manteniamo la riga di sicurezza EOF
    csv_content += (
        '"EOF-riga lasciata appositamente scritta così '
        'per verificare che nessun PC sia andato nella '
        'OU/dismessi/computer"\n'
    )

    warnings = []

    if len(matched) > 1:
        warnings.append(
            "Sono stati trovati più computer. "
            "È stato utilizzato il primo."
        )

    return (
        csv_content,
        row,
        computer,
        warnings
    )


# ============================================================
# TESTO DEPROVISIONING
# ============================================================

def genera_istruzioni(
    sam,
    dl_list,
    sm_list,
    entra_groups
):
    email = f"{sam}@consip.it"

    lines = [
        "Ciao,",
        f"per {email}:",
        ""
    ]

    step = 1

    fixed = [
        (
            "Disabilitare invio ad utente "
            "(Message Delivery Restrictions)"
        ),
        "Impostare Hide dalla Rubrica",
        (
            "Disabilitare accesso Mailbox "
            "(Mailbox features - Disable Protocolli/OWA)"
        ),
        (
            "Estrarre il PST (O365 eDiscovery) "
            "e archiviarlo nel percorso previsto "
            "per il backup delle email cancellate "
            "(in 7z con password condivisa)"
        ),
        "Rimuovere i Ruoli assegnati",
        (
            "Rimuovere le applicazioni "
            "dall'utenza Azure"
        )
    ]

    for description in fixed:
        lines.append(
            f"{step}. {description}"
        )
        step += 1

    # DL
    if dl_list:
        lines.append(
            f"{step}. Rimozione abilitazione dalle DL:"
        )

        for dl in dl_list:
            lines.append(
                f"   - {dl}"
            )

        step += 1

    else:
        lines.append(
            f"{step}. Nessuna DL rilevata "
            "per l'utente."
        )
        step += 1

    # Azure
    lines.append(
        f"{step}. Disabilitare l'account di Azure"
    )
    step += 1

    # Shared Mailbox
    if sm_list:
        lines.append(
            f"{step}. Rimozione abilitazione da SM:"
        )

        for sm in sm_list:
            lines.append(
                f"   - {sm}"
            )

        step += 1

    else:
        lines.append(
            f"{step}. Nessuna SM profilata "
            "all'utente."
        )
        step += 1

    # Entra
    if entra_groups:
        lines.append(
            f"{step}. Rimozione gruppi Azure/Entra ID:"
        )

        for group in entra_groups:
            lines.append(
                f"   - {group}"
            )

        step += 1

    else:
        lines.append(
            f"{step}. Nessun gruppo Azure/Entra ID "
            "aggiuntivo da rimuovere."
        )
        step += 1

    lines.append(
        f"{step}. Cancellare la foto "
        "da Azure (se applicabile)"
    )
    step += 1

    lines.append(
        f"{step}. Rimozione Wi-Fi"
    )

    return "\n".join(lines)


# ============================================================
# INTERFACCIA
# ============================================================

st.subheader(
    "Modulo Deprovisioning"
)

sam = st.text_input(
    "Nome utente (sAMAccountName)",
    "",
    help=(
        "Esempio: mario.rossi "
        "oppure mario.rossi.ext"
    )
).strip().lower()


# ============================================================
# UPLOAD FILE
# ============================================================

st.markdown(
    "### File di estrazione"
)

col1, col2 = st.columns(2)

with col1:
    dl_file = st.file_uploader(
        "Carica file DL (Excel)",
        type=["xlsx"],
        key="dl"
    )

    sm_file = st.file_uploader(
        "Carica file SM (Excel)",
        type=["xlsx"],
        key="sm"
    )

    mg_file = st.file_uploader(
        "Carica Estr_MembriGruppi (Excel)",
        type=["xlsx"],
        key="mg"
    )

with col2:
    entra_file = st.file_uploader(
        "Carica file Entra (Excel)",
        type=["xlsx"],
        key="entra"
    )

    device_file = st.file_uploader(
        "Carica Estr_Device (Excel)",
        type=["xlsx"],
        key="device"
    )


# ============================================================
# GENERAZIONE
# ============================================================

if st.button(
    "Genera Deprovisioning",
    type="primary"
):

    if not sam:
        st.error(
            "Inserire lo sAMAccountName."
        )
        st.stop()

    # --------------------------------------------------------
    # Lettura file
    # --------------------------------------------------------

    dl_df = read_excel(
        dl_file
    )

    sm_df = read_excel(
        sm_file
    )

    mg_df = read_excel(
        mg_file
    )

    entra_df = read_excel(
        entra_file
    )

    device_df = read_excel(
        device_file
    )

    # --------------------------------------------------------
    # Nome utente
    # --------------------------------------------------------

    display_name = build_display_name(
        sam
    )

    persona = build_filename_person(
        sam
    )

    persona_underscore = persona.replace(
        " ",
        "_"
    )

    # --------------------------------------------------------
    # Gruppi AD
    # --------------------------------------------------------

    gruppi_ad, warning_ad = extract_ad_groups(
        sam,
        mg_df
    )

    csv_utente, row_utente = genera_csv_utente(
        sam,
        gruppi_ad
    )

    # --------------------------------------------------------
    # DL
    # --------------------------------------------------------

    dl_list, warning_dl = extract_dl(
        sam,
        dl_df
    )

    # --------------------------------------------------------
    # SM
    # --------------------------------------------------------

    sm_list, warning_sm = extract_sm(
        sam,
        sm_df
    )

    # --------------------------------------------------------
    # Entra
    # --------------------------------------------------------

    entra_raw, warning_entra = extract_entra(
        sam,
        entra_df
    )

    entra_groups = filter_entra_groups(
        entra_raw,
        dl_df,
        sm_df,
        mg_df
    )

    # --------------------------------------------------------
    # Computer
    # --------------------------------------------------------

    (
        computer_csv,
        computer_row,
        computer_name,
        warning_device
    ) = genera_device_csv(
        sam,
        device_df
    )

    # --------------------------------------------------------
    # Istruzioni
    # --------------------------------------------------------

    istruzioni = genera_istruzioni(
        sam,
        dl_list,
        sm_list,
        entra_groups
    )

    # ========================================================
    # NOMI FILE
    # ========================================================

    filename_utente = (
        f"1_{persona}_deprovisioning_utenza.csv"
    )

    filename_istruzioni = (
        f"2_{persona}_istruzioni_deprovisioning.txt"
    )

    filename_computer = (
        f"3_{persona_underscore}_computer_remove.csv"
    )

    # ========================================================
    # GIIT
    # ========================================================

    subject_giit = (
        "01 - Deprovisioning Utente "
        "(Utenza/Casella/Profilazione) - "
        f"{display_name} → GIIT"
    )

    message_lines = [
        "Ciao,",
        (
            f"si richiede di effettuare il "
            f"deprovisioning dell'utenza {sam}."
        ),
        "",
        (
            "Elaborare i file contenuti nella "
            "cartella compressa .zip "
            "nel seguente ordine:"
        ),
        (
            f"1. {filename_utente} per la rimozione "
            "dei gruppi AD, disabilitazione dell'utenza "
            "e spostamento nella OU prevista."
        ),
        (
            f"2. {filename_istruzioni} per le attività "
            "di deprovisioning della casella, DL, "
            "Shared Mailbox ed Entra ID."
        )
    ]

    if computer_csv:
        message_lines.append(
            f"3. {filename_computer} per la rimozione "
            f"dei riferimenti utente dal computer "
            f"{computer_name}."
        )

    message_giit = "\n".join(
        message_lines
    )

    # ========================================================
    # GUIDA
    # ========================================================

    guida = (
        f"{subject_giit}\n\n"
        f"{message_giit}\n\n"
        "ISTRUZIONI DEPROVISIONING\n\n"
        f"{istruzioni}"
    )

    filename_guida = (
        f"Guida_Flusso_Deprovisioning_"
        f"{persona_underscore}.txt"
    )

    # ========================================================
    # ZIP
    # ========================================================

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
            istruzioni
        )

        if computer_csv:
            zip_file.writestr(
                filename_computer,
                computer_csv
            )

        zip_file.writestr(
            filename_guida,
            guida
        )

    zip_buffer.seek(0)

    bundle_name = (
        f"Deprovisioning_"
        f"{persona_underscore}.zip"
    )

    # ========================================================
    # VISUALIZZAZIONE
    # ========================================================

    st.divider()

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
    # 1 - Utente
    # --------------------------------------------------------

    st.markdown(
        f"### 1. {filename_utente}"
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
        key="download_utente"
    )

    # --------------------------------------------------------
    # Gruppi AD
    # --------------------------------------------------------

    with st.expander(
        "Gruppi AD da rimuovere"
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
    # 2 - Istruzioni
    # --------------------------------------------------------

    st.markdown(
        f"### 2. {filename_istruzioni}"
    )

    st.text(
        istruzioni
    )

    st.download_button(
        "Scarica 2 - Istruzioni Deprovisioning",
        data=istruzioni,
        file_name=filename_istruzioni,
        mime="text/plain",
        key="download_istruzioni"
    )

    # --------------------------------------------------------
    # 3 - Computer
    # --------------------------------------------------------

    if computer_csv:

        st.markdown(
            f"### 3. {filename_computer}"
        )

        st.dataframe(
            pd.DataFrame(
                [computer_row],
                columns=HEADER_COMPUTER
            ),
            use_container_width=True
        )

        st.download_button(
            "Scarica 3 - Computer",
            data=computer_csv,
            file_name=filename_computer,
            mime="text/csv",
            key="download_computer"
        )

    else:

        st.markdown(
            "### 3. Computer"
        )

        st.info(
            "Nessun CSV Computer da generare."
        )

    # --------------------------------------------------------
    # ZIP completo
    # --------------------------------------------------------

    st.divider()

    st.markdown(
        "### Bundle completo"
    )

    st.download_button(
        "📦 Scarica bundle completo ZIP",
        data=zip_buffer.getvalue(),
        file_name=bundle_name,
        mime="application/zip",
        key="download_bundle"
    )

    # ========================================================
    # AVVISI
    # ========================================================

    warnings = unique_in_order(
        warning_ad
        + warning_dl
        + warning_sm
        + warning_entra
        + warning_device
    )

    if warnings:

        with st.expander(
            "⚠️ Avvisi",
            expanded=True
        ):

            for warning in warnings:
                st.warning(
                    warning
                )

    # ========================================================
    # DIAGNOSTICA
    # ========================================================

    with st.expander(
        "Diagnostica file"
    ):

        st.write(
            "Colonne Estr_MembriGruppi:",
            mg_df.columns.tolist()
            if not mg_df.empty
            else []
        )

        st.write(
            "Colonne DL:",
            dl_df.columns.tolist()
            if not dl_df.empty
            else []
        )

        st.write(
            "Colonne SM:",
            sm_df.columns.tolist()
            if not sm_df.empty
            else []
        )

        st.write(
            "Colonne Entra:",
            entra_df.columns.tolist()
            if not entra_df.empty
            else []
        )

        st.write(
            "Colonne Estr_Device:",
            device_df.columns.tolist()
            if not device_df.empty
            else []
        )

    st.success(
        f"Deprovisioning generato "
        f"correttamente per '{sam}'."
    )
