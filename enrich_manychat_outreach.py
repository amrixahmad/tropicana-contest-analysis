import os
from pathlib import Path

import pandas as pd
import requests

INPUT_FILE = Path('outputs/final_winners_outreach_draft.xlsx')
OUTPUT_FILE = Path('outputs/final_winners_outreach_manychat_enriched.xlsx')
API_BASE_URL = 'https://api.manychat.com/fb/subscriber/getInfo'


def load_env_file(path=Path('.env')):
    if not path.exists():
        return
    for line in path.read_text(encoding='utf-8').splitlines():
        line = line.strip()
        if not line or line.startswith('#') or '=' not in line:
            continue
        key, value = line.split('=', 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        os.environ.setdefault(key, value)


def get_subscriber_info(subscriber_id, token):
    response = requests.get(
        API_BASE_URL,
        headers={'Authorization': f'Bearer {token}'},
        params={'subscriber_id': subscriber_id},
        timeout=30,
    )
    response.raise_for_status()
    return response.json()


def flatten_manychat_data(payload):
    data = payload.get('data', payload)
    return {
        'manychat_api_name': data.get('name'),
        'manychat_api_first_name': data.get('first_name'),
        'manychat_api_last_name': data.get('last_name'),
        'manychat_api_status': data.get('status'),
        'manychat_api_gender': data.get('gender'),
        'manychat_api_profile_pic': data.get('profile_pic'),
        'manychat_api_live_chat_url': data.get('live_chat_url') or data.get('profile_url') or data.get('url'),
        'manychat_api_raw_id': data.get('id'),
    }


def main():
    load_env_file()
    token = os.environ.get('MANYCHAT_API_KEY') or os.environ.get('MANYCHAT_API_TOKEN')
    if not token:
        raise SystemExit('Set MANYCHAT_API_KEY before running this script.')

    outreach = pd.read_excel(INPUT_FILE, dtype={'manychat_id': str})
    enriched_rows = []

    for _, row in outreach.iterrows():
        enriched = row.to_dict()
        subscriber_id = enriched.get('manychat_id')
        try:
            payload = get_subscriber_info(subscriber_id, token)
            enriched.update(flatten_manychat_data(payload))
            enriched['manychat_api_lookup_status'] = 'success'
            enriched['manychat_api_error'] = ''
        except Exception as exc:
            enriched['manychat_api_lookup_status'] = 'error'
            enriched['manychat_api_error'] = str(exc)
        enriched_rows.append(enriched)

    output = pd.DataFrame(enriched_rows)
    if 'manychat_api_live_chat_url' in output.columns:
        output['manychat_link'] = output['manychat_api_live_chat_url'].fillna(output.get('manychat_link', ''))
    output.to_excel(OUTPUT_FILE, index=False)
    print(f'Saved {OUTPUT_FILE}')
    print(output[['outreach_name', 'matched_full_name_from_workbook', 'manychat_id', 'manychat_api_name', 'manychat_link', 'manychat_api_lookup_status']].to_string(index=False))


if __name__ == '__main__':
    main()
