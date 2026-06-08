import re
from datetime import datetime
from pathlib import Path

import pandas as pd

LOOKUP_FILE = Path('outputs/200_winners_psid_name_lookup.xlsx')
OUTREACH_FILE = Path('outputs/final_consolation_200_with_journal_replacements_manychat_enriched.xlsx')
OUTREACH_NAME_COL = 'outreach_name'
SECOND_OUTREACH_DATE_COL = '2nd_outreach_date'
STATUS_COL = '2nd outreach status'
SNIPPET_COL = 'snippet'
MATCH_NAME_COL = 'participant_name_from_graph'
LATEST_MESSAGE_FROM_NAME_COL = 'latest_message_from_name'
LATEST_MESSAGE_FROM_ID_COL = 'latest_message_from_id'
PAGE_NAME = 'Tropicana Malaysia'
PAGE_ID = '594456394024035'
SECOND_OUTREACH_MATCH_FRAGMENTS = [
    'Peringatan untuk hantar maklumat anda sebelum 10 Jun 2026, 8pm bagi meneruskan proses penebusan hadiah.',
    'Ini adalah peringatan terakhir untuk menghantar maklumat yang diperlukan sebelum',
    'Sekiranya maklumat tidak diterima sebelum tarikh dan masa tersebut, hadiah anda tidak dapat diproses.',
    'bagi tujuan penebusan hadiah anda.',
]


def clean_text(value):
    if pd.isna(value):
        return ''
    return str(value).strip()


def normalize_text(value):
    text = re.sub(r'\s+', ' ', clean_text(value).casefold())
    return re.sub(r'\s+\d+$', '', text)


def message_contains_second_outreach(text):
    normalized_text = normalize_text(text)
    return any(normalize_text(fragment) in normalized_text for fragment in SECOND_OUTREACH_MATCH_FRAGMENTS)


def is_page_sender(sender_name, sender_id):
    normalized_sender_name = normalize_text(sender_name)
    normalized_page_name = normalize_text(PAGE_NAME)
    return bool(
        clean_text(sender_id) == PAGE_ID
        or (normalized_sender_name and normalized_sender_name == normalized_page_name)
    )


def classify_outreach_status(snippet, latest_message_from_name, latest_message_from_id):
    if message_contains_second_outreach(snippet):
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

    required_lookup_cols = {MATCH_NAME_COL, SNIPPET_COL, LATEST_MESSAGE_FROM_NAME_COL, LATEST_MESSAGE_FROM_ID_COL}
    missing_lookup_cols = [column for column in required_lookup_cols if column not in lookup.columns]
    if missing_lookup_cols:
        raise KeyError(f'Expected columns not found in lookup file: {missing_lookup_cols}')
    for column in [OUTREACH_NAME_COL, SECOND_OUTREACH_DATE_COL]:
        if column not in outreach.columns:
            raise KeyError(f'Expected column not found in outreach file: {column}')

    lookup = lookup.copy()
    outreach = outreach.copy()

    lookup['normalized_match_name'] = lookup[MATCH_NAME_COL].map(normalize_text)
    outreach['normalized_outreach_name'] = outreach[OUTREACH_NAME_COL].map(normalize_text)

    if lookup['normalized_match_name'].duplicated().any():
        duplicates = lookup.loc[lookup['normalized_match_name'].duplicated(keep=False), MATCH_NAME_COL].tolist()
        raise RuntimeError(f'Duplicate matched names found in lookup file: {duplicates}')

    target_mask = outreach[SECOND_OUTREACH_DATE_COL].ne('')
    targets = outreach.loc[target_mask].copy()
    non_targets = outreach.loc[~target_mask].copy()

    merged = targets.merge(
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

    if STATUS_COL not in non_targets.columns:
        non_targets[STATUS_COL] = ''

    target_output = merged.drop(
        columns=[
            'normalized_outreach_name',
            'normalized_match_name',
            MATCH_NAME_COL,
            SNIPPET_COL,
            LATEST_MESSAGE_FROM_NAME_COL,
            LATEST_MESSAGE_FROM_ID_COL,
        ],
        errors='ignore',
    )

    combined = pd.concat([target_output, non_targets], ignore_index=True)
    combined['_original_order'] = combined.index
    order_lookup = outreach.reset_index()[[OUTREACH_NAME_COL, 'index']].rename(columns={'index': '_source_order'})
    combined = combined.merge(order_lookup, on=OUTREACH_NAME_COL, how='left')
    combined = combined.sort_values(['_source_order', '_original_order']).drop(columns=['_source_order', '_original_order'])

    insert_after = SECOND_OUTREACH_DATE_COL if SECOND_OUTREACH_DATE_COL in combined.columns else 'FB Inbox'
    ordered_columns = []
    for column in combined.columns:
        if column == STATUS_COL:
            continue
        ordered_columns.append(column)
        if column == insert_after and STATUS_COL not in ordered_columns:
            ordered_columns.append(STATUS_COL)

    if STATUS_COL not in ordered_columns:
        ordered_columns.append(STATUS_COL)

    output = combined[ordered_columns].copy()
    saved_path = write_excel_with_fallback(output, OUTREACH_FILE)

    print(f'Saved: {saved_path}')
    print('rows:', len(output))
    print('second_outreach_rows:', int(output[SECOND_OUTREACH_DATE_COL].ne('').sum()))
    print('replied:', int(output[STATUS_COL].eq('Replied').sum()))
    print('no_reply:', int(output[STATUS_COL].eq('No Reply').sum()))
    print(output.loc[output[SECOND_OUTREACH_DATE_COL].ne(''), [OUTREACH_NAME_COL, SECOND_OUTREACH_DATE_COL, STATUS_COL]].head(20).to_string(index=False))


if __name__ == '__main__':
    main()
