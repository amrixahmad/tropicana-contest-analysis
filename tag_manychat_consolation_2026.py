import os
import time
from pathlib import Path

import pandas as pd
import requests

from enrich_manychat_outreach import load_env_file

INPUT_FILE = Path('outputs/final_consolation_200_with_journal_replacements_manychat_enriched.xlsx')
OUTPUT_FILE = Path('outputs/consolation_200_2026_manychat_tagging_audit.xlsx')
TAG_NAME = '200 consolation 2026'
API_BASE = 'https://api.manychat.com'
REQUEST_DELAY_SECONDS = 0.12


def clean_subscriber_id(value):
    text = '' if pd.isna(value) else str(value).strip()
    if text.endswith('.0'):
        text = text[:-2]
    return text


def manychat_request(method, path, token, **kwargs):
    response = requests.request(
        method,
        f'{API_BASE}{path}',
        headers={'Authorization': f'Bearer {token}'},
        timeout=30,
        **kwargs,
    )
    try:
        payload = response.json()
    except ValueError:
        payload = {'raw_response': response.text}
    if not response.ok:
        raise RuntimeError(f'{response.status_code}: {payload}')
    return payload


def ensure_tag_exists(token, tag_name):
    tags_payload = manychat_request('GET', '/fb/page/getTags', token)
    tags = tags_payload.get('data', [])
    for tag in tags:
        if str(tag.get('name', '')).casefold() == tag_name.casefold():
            return tag

    create_payload = manychat_request(
        'POST',
        '/fb/page/createTag',
        token,
        json={'name': tag_name},
    )
    return create_payload.get('data', {}).get('tag', {'name': tag_name})


def add_tag_to_subscriber(token, subscriber_id, tag_name):
    return manychat_request(
        'POST',
        '/fb/subscriber/addTagByName',
        token,
        json={'subscriber_id': int(subscriber_id), 'tag_name': tag_name},
    )


def main():
    load_env_file()
    token = os.environ.get('MANYCHAT_API_KEY') or os.environ.get('MANYCHAT_API_TOKEN')
    if not token:
        raise SystemExit('Set MANYCHAT_API_KEY or MANYCHAT_API_TOKEN before running this script.')

    winners = pd.read_excel(INPUT_FILE, dtype=str).fillna('')
    if 'manychat_id' not in winners.columns:
        raise KeyError('Expected manychat_id column in the consolation workbook.')
    if len(winners) != 200:
        raise RuntimeError(f'Expected 200 rows, found {len(winners)}.')

    winners = winners.copy()
    winners['manychat_id'] = winners['manychat_id'].map(clean_subscriber_id)
    if winners['manychat_id'].eq('').any():
        raise RuntimeError('Workbook contains blank manychat_id values.')
    if winners['manychat_id'].duplicated().any():
        raise RuntimeError('Workbook contains duplicate manychat_id values.')

    tag = ensure_tag_exists(token, TAG_NAME)
    audit_rows = []
    total = len(winners)

    for index, row in enumerate(winners.itertuples(index=False), start=1):
        row_data = row._asdict()
        subscriber_id = row_data['manychat_id']
        audit = {
            'row_number': index,
            'outreach_name': row_data.get('outreach_name', ''),
            'manychat_id': subscriber_id,
            'tag_name': TAG_NAME,
            'tag_id': tag.get('id'),
        }
        try:
            payload = add_tag_to_subscriber(token, subscriber_id, TAG_NAME)
            audit['tagging_status'] = payload.get('status', 'success')
            audit['tagging_error'] = ''
        except Exception as exc:
            audit['tagging_status'] = 'error'
            audit['tagging_error'] = str(exc)
        audit_rows.append(audit)
        print(f'[{index}/{total}] {audit["outreach_name"]} {subscriber_id}: {audit["tagging_status"]} {audit["tagging_error"]}')
        if index < total and REQUEST_DELAY_SECONDS > 0:
            time.sleep(REQUEST_DELAY_SECONDS)

    audit_df = pd.DataFrame(audit_rows)
    OUTPUT_FILE.parent.mkdir(exist_ok=True)
    audit_df.to_excel(OUTPUT_FILE, index=False)

    print(f'Saved {OUTPUT_FILE}')
    print('rows:', len(audit_df))
    print('successes:', int(audit_df['tagging_status'].ne('error').sum()))
    print('errors:', int(audit_df['tagging_status'].eq('error').sum()))


if __name__ == '__main__':
    main()
