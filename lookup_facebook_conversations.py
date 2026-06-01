import os
import time
from pathlib import Path

import pandas as pd
import requests

from enrich_manychat_outreach import load_env_file

INPUT_FILE = Path('outputs/final_winners_outreach_manychat_enriched.xlsx')
PSID_INPUT_FILE = Path('outputs/grand_prize_winners_with_psids_and_inbox_links.xlsx')
OUTPUT_FILE = Path('outputs/grand_prize_facebook_conversation_lookup_test.xlsx')
PROBE_OUTPUT_FILE = Path('outputs/grand_prize_facebook_endpoint_probe.xlsx')
GRAPH_URL = 'https://graph.facebook.com/v25.0/me/conversations'
GRAPH_BASE_URL = 'https://graph.facebook.com/v25.0'
PAGE_ID = '594456394024035'
REQUEST_DELAY_SECONDS = 0.35

CONVERSATION_FIELDS = [
    'id',
    'is_owner',
    'link',
    'updated_time',
    'participants',
    'senders',
    'messages',
    'message_count',
    'unread_count',
    'snippet',
    'can_reply',
    'is_subscribed',
    'folder',
    'name',
    'thread_key',
    'thread_id',
    'subject',
]

MESSAGE_FIELDS = [
    'id',
    'created_time',
    'from',
    'to',
    'message',
    'attachments',
    'shares',
    'story',
    'tags',
    'sticker',
    'application',
    'admin_creator',
    'link',
    'permalink_url',
    'url',
    'conversation',
    'thread_key',
    'thread_id',
    'parent',
    'recipient',
    'sender',
    'mid',
]


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
        'http_status': response.status_code,
        'conversation_count': None,
        'conversation_id': '',
        'facebook_inbox_link': '',
        'conversation_updated_time': '',
        'participants': '',
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
    link = conversation.get('link', '')
    if link.startswith('/'):
        link = f'https://www.facebook.com{link}'

    participants = conversation.get('participants', {}).get('data', [])
    result.update({
        'conversation_id': conversation.get('id', ''),
        'facebook_inbox_link': link,
        'conversation_updated_time': conversation.get('updated_time', ''),
        'participants': ' | '.join(f"{p.get('name', '')}:{p.get('id', '')}" for p in participants),
    })
    return result


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


def summarize_value(value):
    if value is None:
        return ''
    if isinstance(value, str):
        return f'<text length {len(value)}>' if len(value) > 120 else value
    if isinstance(value, dict):
        data = value.get('data')
        if isinstance(data, list):
            return f'<object data_count={len(data)}>'
        return '<object>'
    if isinstance(value, list):
        return f'<list count={len(value)}>'
    return value


def probe_single_field(token, path, field):
    response, payload = graph_get(token, path, {'fields': field})
    result = {
        'field': field,
        'http_status': response.status_code,
        'is_valid': response.ok,
        'value_summary': '',
        'error': '',
    }
    if response.ok:
        result['value_summary'] = summarize_value(payload.get(field))
    else:
        result['error'] = payload.get('error', {}).get('message', str(payload))
    return result


def load_probe_input():
    if PSID_INPUT_FILE.exists():
        return pd.read_excel(PSID_INPUT_FILE, dtype={'manychat_id': str, 'facebook_psid': str})

    winners = pd.read_excel(INPUT_FILE, dtype={'manychat_id': str})
    winners['facebook_psid'] = winners['manychat_id']
    return winners


def run_endpoint_probe(token):
    winners = load_probe_input()
    conversation_rows = []
    message_rows = []
    conversation_field_rows = []
    message_field_rows = []

    for idx, winner in winners.iterrows():
        user_id = winner.get('facebook_psid') or winner.get('manychat_id')
        lookup = lookup_conversation(token, user_id)
        conversation_id = lookup.get('conversation_id')
        base_row = {
            'outreach_name': winner.get('outreach_name'),
            'manychat_api_name': winner.get('manychat_api_name'),
            'manychat_id': winner.get('manychat_id'),
            'facebook_psid': user_id,
            'conversation_id': conversation_id,
            'conversation_link': lookup.get('facebook_inbox_link'),
            'conversation_updated_time': lookup.get('conversation_updated_time'),
            'conversation_count': lookup.get('conversation_count'),
            'lookup_http_status': lookup.get('http_status'),
            'lookup_error': lookup.get('lookup_error'),
            'participants': lookup.get('participants'),
        }

        if conversation_id:
            response, payload = graph_get(
                token,
                conversation_id,
                {
                    'fields': (
                        'id,link,updated_time,participants,senders,'
                        'message_count,unread_count,snippet,can_reply,is_subscribed,folder,name'
                    )
                },
            )
            base_row.update({
                'conversation_detail_http_status': response.status_code,
                'message_count': payload.get('message_count') if response.ok else '',
                'unread_count': payload.get('unread_count') if response.ok else '',
                'snippet_summary': summarize_value(payload.get('snippet')) if response.ok else '',
                'can_reply': payload.get('can_reply') if response.ok else '',
                'is_subscribed': payload.get('is_subscribed') if response.ok else '',
                'folder': payload.get('folder') if response.ok else '',
                'conversation_detail_error': '' if response.ok else payload.get('error', {}).get('message', str(payload)),
            })

            messages_response, messages_payload = graph_get(
                token,
                f'{conversation_id}/messages',
                {'fields': 'id,created_time,from,to,message', 'limit': 1},
            )
            latest_message = (messages_payload.get('data') or [{}])[0] if messages_response.ok else {}
            message_id = latest_message.get('id', '')
            message_rows.append({
                'outreach_name': winner.get('outreach_name'),
                'manychat_api_name': winner.get('manychat_api_name'),
                'manychat_id': winner.get('manychat_id'),
                'facebook_psid': user_id,
                'conversation_id': conversation_id,
                'messages_http_status': messages_response.status_code,
                'latest_message_id': message_id,
                'latest_message_created_time': latest_message.get('created_time', ''),
                'latest_message_from': (latest_message.get('from') or {}).get('name', ''),
                'latest_message_text_length': len(latest_message.get('message') or ''),
                'messages_error': '' if messages_response.ok else messages_payload.get('error', {}).get('message', str(messages_payload)),
            })

            if idx == 0:
                for field in CONVERSATION_FIELDS:
                    row = {'endpoint_type': 'conversation', 'sample_conversation_id': conversation_id}
                    row.update(probe_single_field(token, conversation_id, field))
                    conversation_field_rows.append(row)

                if message_id:
                    for field in MESSAGE_FIELDS:
                        row = {'endpoint_type': 'message', 'sample_message_id': message_id}
                        row.update(probe_single_field(token, message_id, field))
                        message_field_rows.append(row)

        conversation_rows.append(base_row)
        if idx < len(winners) - 1:
            time.sleep(REQUEST_DELAY_SECONDS)

    with pd.ExcelWriter(PROBE_OUTPUT_FILE) as writer:
        pd.DataFrame(conversation_rows).to_excel(writer, sheet_name='conversation_summary', index=False)
        pd.DataFrame(message_rows).to_excel(writer, sheet_name='latest_messages', index=False)
        pd.DataFrame(conversation_field_rows).to_excel(writer, sheet_name='conversation_fields', index=False)
        pd.DataFrame(message_field_rows).to_excel(writer, sheet_name='message_fields', index=False)

    print(f'Saved {PROBE_OUTPUT_FILE}')
    summary_cols = [
        'outreach_name',
        'manychat_id',
        'facebook_psid',
        'conversation_id',
        'message_count',
        'unread_count',
        'can_reply',
        'is_subscribed',
        'conversation_detail_error',
    ]
    print(pd.DataFrame(conversation_rows)[summary_cols].to_string(index=False))


def main():
    load_env_file()
    token = os.environ.get('TROPICANA_PAGE_API_KEY')
    if not token:
        raise SystemExit('Set TROPICANA_PAGE_API_KEY before running this script.')

    run_endpoint_probe(token)

    winners = pd.read_excel(INPUT_FILE, dtype={'manychat_id': str})
    rows = []
    for _, winner in winners.iterrows():
        row = {
            'outreach_name': winner.get('outreach_name'),
            'manychat_api_name': winner.get('manychat_api_name'),
            'manychat_id': winner.get('manychat_id'),
        }
        row.update(lookup_conversation(token, winner.get('manychat_id')))
        rows.append(row)

    output = pd.DataFrame(rows)
    output.to_excel(OUTPUT_FILE, index=False)
    print(f'Saved {OUTPUT_FILE}')
    print(output.to_string(index=False))


if __name__ == '__main__':
    main()
