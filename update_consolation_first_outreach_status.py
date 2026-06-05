import re
from datetime import datetime
from pathlib import Path

import pandas as pd

LOOKUP_FILE = Path('outputs/200_winners_psid_name_lookup.xlsx')
OUTREACH_FILE = Path('outputs/final_consolation_200_with_journal_replacements_manychat_enriched.xlsx')
OUTREACH_NAME_COL = 'outreach_name'
STATUS_COL = '1st outreach status'
SNIPPET_COL = 'snippet'
MATCH_NAME_COL = 'participant_name_from_graph'
LATEST_MESSAGE_FROM_NAME_COL = 'latest_message_from_name'
LATEST_MESSAGE_FROM_ID_COL = 'latest_message_from_id'
PAGE_NAME = 'Tropicana Malaysia'
PAGE_ID = '594456394024035'
FIRST_OUTREACH_PHRASE = (
    'Tahniah! Anda telah dipilih sebagai salah seorang pemenang kempen Gandakan Kebaikan '
    'Bersama Tropicana Twister. Sila hantar dalam tempoh 7 hari maklumat berikut bagi tujuan '
    'penebusan hadiah:'
)


def normalize_text(value):
    text = re.sub(r'\s+', ' ', str(value).strip().casefold())
    return re.sub(r'\s+\d+$', '', text)


def is_page_sender(sender_name, sender_id):
    normalized_sender_name = normalize_text(sender_name)
    normalized_page_name = normalize_text(PAGE_NAME)
    return bool(
        clean_text(sender_id) == PAGE_ID
        or (normalized_sender_name and normalized_sender_name == normalized_page_name)
    )


def clean_text(value):
    if pd.isna(value):
        return ''
    return str(value).strip()


def classify_outreach_status(snippet, latest_message_from_name, latest_message_from_id):
    normalized_snippet = normalize_text(snippet)
    normalized_phrase = normalize_text(FIRST_OUTREACH_PHRASE)
    if normalized_phrase in normalized_snippet:
        if is_page_sender(latest_message_from_name, latest_message_from_id):
            return 'No Reply'
        return 'Replied'
    return 'Replied'


def write_excel_with_fallback(df, path):
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    candidate_paths = [path, path.with_name(f'{path.stem}_{timestamp}{path.suffix}')]

    last_error = None
    for candidate_path in candidate_paths:
        try:
            df.to_excel(candidate_path, index=False)
            return candidate_path
        except PermissionError as exc:
            last_error = exc
            continue

    if last_error is not None:
        raise last_error
    return path


def main():
    lookup = pd.read_excel(LOOKUP_FILE, dtype=str).fillna('')
    outreach = pd.read_excel(OUTREACH_FILE, dtype=str).fillna('')

    if MATCH_NAME_COL not in lookup.columns:
        raise KeyError(f'Expected column not found in lookup file: {MATCH_NAME_COL}')
    if SNIPPET_COL not in lookup.columns:
        raise KeyError(f'Expected column not found in lookup file: {SNIPPET_COL}')
    if LATEST_MESSAGE_FROM_NAME_COL not in lookup.columns:
        raise KeyError(f'Expected column not found in lookup file: {LATEST_MESSAGE_FROM_NAME_COL}')
    if LATEST_MESSAGE_FROM_ID_COL not in lookup.columns:
        raise KeyError(f'Expected column not found in lookup file: {LATEST_MESSAGE_FROM_ID_COL}')
    if OUTREACH_NAME_COL not in outreach.columns:
        raise KeyError(f'Expected column not found in outreach file: {OUTREACH_NAME_COL}')

    lookup = lookup.copy()
    outreach = outreach.copy()

    lookup['normalized_match_name'] = lookup[MATCH_NAME_COL].map(normalize_text)
    outreach['normalized_outreach_name'] = outreach[OUTREACH_NAME_COL].map(normalize_text)

    if lookup['normalized_match_name'].duplicated().any():
        duplicates = lookup.loc[lookup['normalized_match_name'].duplicated(keep=False), MATCH_NAME_COL].tolist()
        raise RuntimeError(f'Duplicate matched names found in lookup file: {duplicates}')

    merged = outreach.merge(
        lookup[[MATCH_NAME_COL, SNIPPET_COL, LATEST_MESSAGE_FROM_NAME_COL, LATEST_MESSAGE_FROM_ID_COL, 'normalized_match_name']],
        left_on='normalized_outreach_name',
        right_on='normalized_match_name',
        how='left',
        validate='one_to_one',
    )

    unmatched = merged[merged[SNIPPET_COL].eq('')][OUTREACH_NAME_COL].tolist()
    if unmatched:
        raise RuntimeError(f'Unable to match lookup rows for: {unmatched}')

    merged[STATUS_COL] = merged.apply(
        lambda row: classify_outreach_status(
            row[SNIPPET_COL],
            row[LATEST_MESSAGE_FROM_NAME_COL],
            row[LATEST_MESSAGE_FROM_ID_COL],
        ),
        axis=1,
    )

    insert_after = 'FB Inbox' if 'FB Inbox' in merged.columns else OUTREACH_NAME_COL
    ordered_columns = []
    for column in merged.columns:
        if column in {
            'normalized_outreach_name',
            'normalized_match_name',
            MATCH_NAME_COL,
            SNIPPET_COL,
            LATEST_MESSAGE_FROM_NAME_COL,
            LATEST_MESSAGE_FROM_ID_COL,
        }:
            continue
        if column == STATUS_COL:
            continue
        ordered_columns.append(column)
        if column == insert_after and STATUS_COL not in ordered_columns:
            ordered_columns.append(STATUS_COL)

    if STATUS_COL not in ordered_columns:
        ordered_columns.append(STATUS_COL)

    output = merged[ordered_columns].copy()
    saved_path = write_excel_with_fallback(output, OUTREACH_FILE)

    print(f'Saved: {saved_path}')
    print('rows:', len(output))
    print('replied:', int(output[STATUS_COL].eq('Replied').sum()))
    print('no_reply:', int(output[STATUS_COL].eq('No Reply').sum()))
    print(output[[OUTREACH_NAME_COL, STATUS_COL]].head(20).to_string(index=False))


if __name__ == '__main__':
    main()
