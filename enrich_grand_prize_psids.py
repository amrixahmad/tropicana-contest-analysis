import os
import re
import time
from pathlib import Path

import pandas as pd
import requests

from enrich_manychat_outreach import load_env_file

WINNERS_FILE = Path('outputs/final_winners_outreach_manychat_enriched.xlsx')
PSID_FILE = Path('6 grand prize psid.csv')
OUTPUT_FILE = Path('outputs/grand_prize_winners_with_psids_and_inbox_links.xlsx')
GRAPH_URL = 'https://graph.facebook.com/v25.0/me/conversations'
REQUEST_DELAY_SECONDS = 0.35
PAGE_NAME = 'Tropicana Malaysia'


def normalize_name(value):
    return re.sub(r'\s+', ' ', str(value).strip().casefold())


def lookup_conversation(token, user_id):
    response = requests.get(
        GRAPH_URL,
        params={
            'user_id': user_id,
            'fields': 'participants,link,updated_time',
            'access_token': token,
        },
        timeout=30,
    )
    try:
        payload = response.json()
    except ValueError:
        payload = {'raw_response': response.text}

    result = {
        'facebook_psid': user_id,
        'psid_lookup_http_status': response.status_code,
        'psid_conversation_count': None,
        'psid_conversation_id': '',
        'facebook_inbox_link_from_psid': '',
        'psid_conversation_updated_time': '',
        'psid_participant_name': '',
        'psid_participants': '',
        'psid_lookup_error': '',
    }

    if not response.ok:
        result['psid_lookup_error'] = payload.get('error', {}).get('message', str(payload))
        return result

    conversations = payload.get('data', [])
    result['psid_conversation_count'] = len(conversations)
    if not conversations:
        result['psid_lookup_error'] = 'no conversation returned'
        return result

    conversation = conversations[0]
    link = conversation.get('link', '')
    if link.startswith('/'):
        link = f'https://www.facebook.com{link}'

    participants = conversation.get('participants', {}).get('data', [])
    non_page_participants = [
        participant for participant in participants
        if normalize_name(participant.get('name')) != normalize_name(PAGE_NAME)
    ]
    participant_name = non_page_participants[0].get('name', '') if non_page_participants else ''

    result.update({
        'psid_conversation_id': conversation.get('id', ''),
        'facebook_inbox_link_from_psid': link,
        'psid_conversation_updated_time': conversation.get('updated_time', ''),
        'psid_participant_name': participant_name,
        'psid_participants': ' | '.join(f"{p.get('name', '')}:{p.get('id', '')}" for p in participants),
    })
    return result


def main():
    load_env_file()
    token = os.environ.get('TROPICANA_PAGE_API_KEY')
    if not token:
        raise SystemExit('Set TROPICANA_PAGE_API_KEY before running this script.')

    winners = pd.read_excel(WINNERS_FILE, dtype={'manychat_id': str})
    psids = pd.read_csv(PSID_FILE, dtype={'pageuid': str})['pageuid'].dropna().astype(str).str.strip()
    psids = psids[psids.ne('')].drop_duplicates().tolist()

    psid_lookup_rows = []
    for idx, psid in enumerate(psids, start=1):
        psid_lookup_rows.append(lookup_conversation(token, psid))
        if idx < len(psids):
            time.sleep(REQUEST_DELAY_SECONDS)

    psid_lookup = pd.DataFrame(psid_lookup_rows)
    psid_lookup['normalized_psid_participant_name'] = psid_lookup['psid_participant_name'].map(normalize_name)

    winners = winners.copy()
    winners['normalized_manychat_api_name'] = winners['manychat_api_name'].map(normalize_name)

    output = winners.merge(
        psid_lookup,
        left_on='normalized_manychat_api_name',
        right_on='normalized_psid_participant_name',
        how='left',
        validate='one_to_one',
    )
    output['psid_match_status'] = output['facebook_psid'].notna().map({True: 'matched_by_api_name', False: 'unmatched'})

    ordered_cols = [
        'outreach_name',
        'matched_full_name_from_workbook',
        'manychat_api_name',
        'manychat_id',
        'facebook_psid',
        'manychat_link',
        'facebook_inbox_link_from_psid',
        'psid_match_status',
        'psid_participant_name',
        'psid_conversation_id',
        'psid_conversation_updated_time',
        'psid_lookup_http_status',
        'psid_lookup_error',
        'psid_participants',
    ]
    remaining_cols = [col for col in output.columns if col not in ordered_cols and not col.startswith('normalized_')]
    output = output[ordered_cols + remaining_cols]

    OUTPUT_FILE.parent.mkdir(exist_ok=True)
    output.to_excel(OUTPUT_FILE, index=False)
    print(f'Saved {OUTPUT_FILE}')
    print(output[ordered_cols].to_string(index=False))


if __name__ == '__main__':
    main()
