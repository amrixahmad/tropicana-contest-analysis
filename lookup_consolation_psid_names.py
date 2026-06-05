import os
import re
import time
from pathlib import Path

import pandas as pd
import requests

from enrich_manychat_outreach import load_env_file

INPUT_FILE = Path('200 winners psid.csv')
OUTPUT_FILE = Path('outputs/200_winners_psid_name_lookup.xlsx')
GRAPH_BASE_URL = 'https://graph.facebook.com/v25.0'
GRAPH_CONVERSATIONS_URL = f'{GRAPH_BASE_URL}/me/conversations'
REQUEST_DELAY_SECONDS = 0.35
PAGE_NAME = 'Tropicana Malaysia'


def normalize_name(value):
    return re.sub(r'\s+', ' ', str(value).strip().casefold())


def clean_text(value):
    if pd.isna(value):
        return ''
    return str(value).strip()


def graph_get(token, path, params):
    response = requests.get(
        f'{GRAPH_BASE_URL}/{path}',
        params={**params, 'access_token': token},
        timeout=30,
    )
    try:
        payload = response.json()
    except ValueError:
        payload = {'raw_response': response.text}
    return response, payload


def lookup_conversation(token, user_id):
    response = requests.get(
        GRAPH_CONVERSATIONS_URL,
        params={
            'user_id': user_id,
            'fields': 'id,participants,link,updated_time,message_count,unread_count,snippet,can_reply,is_subscribed',
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
        'lookup_http_status': response.status_code,
        'conversation_count': None,
        'conversation_id': '',
        'facebook_inbox_link': '',
        'conversation_updated_time': '',
        'participant_name_from_graph': '',
        'participants': '',
        'message_count': '',
        'unread_count': '',
        'snippet': '',
        'can_reply': '',
        'is_subscribed': '',
        'lookup_error': '',
    }

    if not response.ok:
        result['lookup_error'] = payload.get('error', {}).get('message', str(payload))
        return result

    conversations = payload.get('data', [])
    result['conversation_count'] = len(conversations)
    if not conversations:
        result['lookup_error'] = 'no conversation returned'
        return result

    conversation = conversations[0]
    link = clean_text(conversation.get('link', ''))
    if link.startswith('/'):
        link = f'https://www.facebook.com{link}'

    participants = conversation.get('participants', {}).get('data', [])
    non_page_participants = [
        participant for participant in participants
        if normalize_name(participant.get('name')) != normalize_name(PAGE_NAME)
    ]
    participant_name = non_page_participants[0].get('name', '') if non_page_participants else ''

    snippet = clean_text(conversation.get('snippet', ''))
    if len(snippet) > 250:
        snippet = f'{snippet[:250]}...'

    result.update({
        'conversation_id': clean_text(conversation.get('id', '')),
        'facebook_inbox_link': link,
        'conversation_updated_time': clean_text(conversation.get('updated_time', '')),
        'participant_name_from_graph': clean_text(participant_name),
        'participants': ' | '.join(f"{clean_text(p.get('name', ''))}:{clean_text(p.get('id', ''))}" for p in participants),
        'message_count': clean_text(conversation.get('message_count', '')),
        'unread_count': clean_text(conversation.get('unread_count', '')),
        'snippet': snippet,
        'can_reply': clean_text(conversation.get('can_reply', '')),
        'is_subscribed': clean_text(conversation.get('is_subscribed', '')),
    })
    return result


def lookup_latest_message(token, conversation_id):
    if not conversation_id:
        return {
            'messages_http_status': '',
            'latest_message_id': '',
            'latest_message_created_time': '',
            'latest_message_from_name': '',
            'latest_message_from_id': '',
            'latest_message_text_length': '',
            'messages_error': '',
        }

    response, payload = graph_get(
        token,
        f'{conversation_id}/messages',
        {'fields': 'id,created_time,from,to,message', 'limit': 1},
    )
    result = {
        'messages_http_status': response.status_code,
        'latest_message_id': '',
        'latest_message_created_time': '',
        'latest_message_from_name': '',
        'latest_message_from_id': '',
        'latest_message_text_length': '',
        'messages_error': '',
    }

    if not response.ok:
        result['messages_error'] = payload.get('error', {}).get('message', str(payload))
        return result

    latest_message = (payload.get('data') or [{}])[0]
    sender = latest_message.get('from') or {}
    result.update({
        'latest_message_id': clean_text(latest_message.get('id', '')),
        'latest_message_created_time': clean_text(latest_message.get('created_time', '')),
        'latest_message_from_name': clean_text(sender.get('name', '')),
        'latest_message_from_id': clean_text(sender.get('id', '')),
        'latest_message_text_length': len(latest_message.get('message') or ''),
    })
    return result


def main():
    load_env_file()
    token = (os.environ.get('TROPICANA_API_TOKEN') or os.environ.get('TROPICANA_PAGE_API_KEY') or '').strip()
    if not token:
        raise SystemExit('Set TROPICANA_API_TOKEN or TROPICANA_PAGE_API_KEY before running this script.')

    psids = pd.read_csv(INPUT_FILE, dtype={'pageuid': str}).fillna('')
    if 'pageuid' not in psids.columns:
        raise KeyError('Expected pageuid column in 200 winners psid.csv.')

    psids = psids.copy()
    psids['facebook_psid'] = psids['pageuid'].map(clean_text)
    psids = psids[psids['facebook_psid'].ne('')].drop_duplicates(subset=['facebook_psid']).reset_index(drop=True)

    rows = []
    total = len(psids)
    for index, row in enumerate(psids.itertuples(index=False), start=1):
        facebook_psid = row.facebook_psid
        lookup = lookup_conversation(token, facebook_psid)
        lookup['row_number'] = index
        lookup.update(lookup_latest_message(token, lookup.get('conversation_id', '')))
        rows.append(lookup)
        print(f'[{index}/{total}] {facebook_psid}: {lookup.get("lookup_http_status", "")} {lookup.get("participant_name_from_graph", "")} {lookup.get("lookup_error", "")}')
        if index < total:
            time.sleep(REQUEST_DELAY_SECONDS)

    output = pd.DataFrame(rows)
    ordered_columns = [
        'row_number',
        'facebook_psid',
        'participant_name_from_graph',
        'lookup_http_status',
        'conversation_count',
        'conversation_id',
        'facebook_inbox_link',
        'conversation_updated_time',
        'message_count',
        'unread_count',
        'can_reply',
        'is_subscribed',
        'latest_message_from_name',
        'latest_message_from_id',
        'latest_message_created_time',
        'latest_message_text_length',
        'participants',
        'snippet',
        'lookup_error',
        'messages_http_status',
        'messages_error',
    ]
    output = output[ordered_columns]

    OUTPUT_FILE.parent.mkdir(exist_ok=True)
    output.to_excel(OUTPUT_FILE, index=False)
    print(f'Saved {OUTPUT_FILE}')
    print('rows:', len(output))
    print('names_found:', int(output['participant_name_from_graph'].ne('').sum()))
    print('lookup_errors:', int(output['lookup_error'].ne('').sum()))
    print(output[['row_number', 'facebook_psid', 'participant_name_from_graph', 'lookup_http_status', 'lookup_error']].to_string(index=False))


if __name__ == '__main__':
    main()
