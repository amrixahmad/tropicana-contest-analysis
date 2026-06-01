import os
from pathlib import Path

import pandas as pd
import requests

from enrich_manychat_outreach import load_env_file

INPUT_FILE = Path('outputs/final_winners_outreach_manychat_enriched.xlsx')
OUTPUT_FILE = Path('outputs/grand_prize_winner_tagging_audit.xlsx')
TAG_NAME = '2026 grand prize winners'
API_BASE = 'https://api.manychat.com'


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
        raise SystemExit('Set MANYCHAT_API_KEY before running this script.')

    winners = pd.read_excel(INPUT_FILE, dtype={'manychat_id': str})
    tag = ensure_tag_exists(token, TAG_NAME)

    audit_rows = []
    for _, row in winners.iterrows():
        subscriber_id = row['manychat_id']
        audit = {
            'outreach_name': row.get('outreach_name'),
            'manychat_api_name': row.get('manychat_api_name'),
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

    audit_df = pd.DataFrame(audit_rows)
    audit_df.to_excel(OUTPUT_FILE, index=False)
    print(f'Saved {OUTPUT_FILE}')
    print(audit_df.to_string(index=False))


if __name__ == '__main__':
    main()
