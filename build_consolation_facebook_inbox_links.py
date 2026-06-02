import os
import re
import time
from pathlib import Path
from urllib.parse import quote

import pandas as pd
import requests

from enrich_manychat_outreach import load_env_file

INPUT_FILE = Path('outputs/final_consolation_200_with_journal_replacements_manychat_enriched.xlsx')
OUTPUT_FILE = Path('outputs/final_consolation_200_facebook_inbox_link_candidates.xlsx')
GRAPH_URL = 'https://graph.facebook.com/v25.0/me/conversations'
PAGE_ID = '594456394024035'
REQUEST_DELAY_SECONDS = 0.35


def clean_text(value):
    if pd.isna(value):
        return ''
    return str(value).strip()


def clean_id(value):
    text = clean_text(value)
    if text.endswith('.0'):
        text = text[:-2]
    return text


def extract_legacy_inbox_numeric(link):
    match = re.search(r'/inbox/(\d+)/', str(link))
    return match.group(1) if match else ''


def business_suite_url(selected_item_id):
    selected_item_id = clean_text(selected_item_id)
    if not selected_item_id:
        return ''
    return (
        'https://business.facebook.com/latest/inbox/all/'
        f'?asset_id={PAGE_ID}&mailbox_id={PAGE_ID}'
        f'&selected_item_id={quote(selected_item_id, safe="")}'
        '&thread_type=FB_MESSAGE'
    )


def lookup_conversation(token, user_id):
    response = requests.get(
        GRAPH_URL,
        params={
            'user_id': user_id,
            'fields': 'id,participants,link,updated_time,can_reply,is_subscribed,message_count,unread_count,snippet',
            'access_token': token,
        },
        timeout=30,
    )
    try:
        payload = response.json()
    except ValueError:
        payload = {'raw_response': response.text}

    result = {
        'lookup_user_id': user_id,
        'lookup_http_status': response.status_code,
        'conversation_count': None,
        'graph_conversation_id': '',
        'graph_conversation_link': '',
        'legacy_inbox_numeric_from_graph_link': '',
        'conversation_updated_time': '',
        'can_reply': '',
        'is_subscribed': '',
        'message_count': '',
        'unread_count': '',
        'snippet': '',
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
    link = clean_text(conversation.get('link', ''))
    if link.startswith('/'):
        link = f'https://www.facebook.com{link}'

    participants = conversation.get('participants', {}).get('data', [])
    snippet = conversation.get('snippet', '')
    if isinstance(snippet, str) and len(snippet) > 250:
        snippet = f'{snippet[:250]}...'

    result.update({
        'graph_conversation_id': clean_text(conversation.get('id', '')),
        'graph_conversation_link': link,
        'legacy_inbox_numeric_from_graph_link': extract_legacy_inbox_numeric(link),
        'conversation_updated_time': clean_text(conversation.get('updated_time', '')),
        'can_reply': clean_text(conversation.get('can_reply', '')),
        'is_subscribed': clean_text(conversation.get('is_subscribed', '')),
        'message_count': clean_text(conversation.get('message_count', '')),
        'unread_count': clean_text(conversation.get('unread_count', '')),
        'snippet': clean_text(snippet),
        'participants': ' | '.join(f"{p.get('name', '')}:{p.get('id', '')}" for p in participants),
    })
    return result


def main():
    load_env_file()
    token = os.environ.get('TROPICANA_PAGE_API_KEY', '').strip()
    if not token:
        raise SystemExit('Set TROPICANA_PAGE_API_KEY before running this script.')

    winners = pd.read_excel(INPUT_FILE, dtype=str).fillna('')
    if 'manychat_id' not in winners.columns:
        raise KeyError('Expected a manychat_id column in the consolation workbook.')

    rows = []
    total = len(winners)
    for index, winner in enumerate(winners.itertuples(index=False), start=1):
        winner_data = winner._asdict()
        manychat_id = clean_id(winner_data.get('manychat_id', ''))
        lookup = lookup_conversation(token, manychat_id) if manychat_id else {
            'lookup_user_id': '',
            'lookup_http_status': '',
            'conversation_count': '',
            'graph_conversation_id': '',
            'graph_conversation_link': '',
            'legacy_inbox_numeric_from_graph_link': '',
            'conversation_updated_time': '',
            'can_reply': '',
            'is_subscribed': '',
            'message_count': '',
            'unread_count': '',
            'snippet': '',
            'participants': '',
            'lookup_error': 'blank manychat_id',
        }

        row = dict(winner_data)
        row['manychat_id'] = manychat_id
        row.update(lookup)
        row['candidate_selected_item_id_graph_conversation_id'] = lookup.get('graph_conversation_id', '')
        row['candidate_url_graph_conversation_id'] = business_suite_url(lookup.get('graph_conversation_id', ''))
        row['candidate_selected_item_id_manychat_id'] = manychat_id
        row['candidate_url_manychat_id'] = business_suite_url(manychat_id)
        row['candidate_selected_item_id_legacy_inbox_numeric'] = lookup.get('legacy_inbox_numeric_from_graph_link', '')
        row['candidate_url_legacy_inbox_numeric'] = business_suite_url(lookup.get('legacy_inbox_numeric_from_graph_link', ''))
        rows.append(row)

        print(f'[{index}/{total}] {winner_data.get("outreach_name", "")} {manychat_id}: {lookup.get("lookup_http_status", "")} {lookup.get("graph_conversation_id", "")} {lookup.get("lookup_error", "")}')
        if index < total:
            time.sleep(REQUEST_DELAY_SECONDS)

    output = pd.DataFrame(rows)
    front_columns = [
        'outreach_name',
        'manychat_id',
        'lookup_http_status',
        'conversation_count',
        'lookup_error',
        'graph_conversation_id',
        'legacy_inbox_numeric_from_graph_link',
        'graph_conversation_link',
        'candidate_url_graph_conversation_id',
        'candidate_url_manychat_id',
        'candidate_url_legacy_inbox_numeric',
        'conversation_updated_time',
        'can_reply',
        'is_subscribed',
        'message_count',
        'unread_count',
        'participants',
        'snippet',
    ]
    ordered_columns = [column for column in front_columns if column in output.columns]
    ordered_columns += [column for column in output.columns if column not in ordered_columns]
    output = output[ordered_columns]

    OUTPUT_FILE.parent.mkdir(exist_ok=True)
    output.to_excel(OUTPUT_FILE, index=False)

    print(f'Saved {OUTPUT_FILE}')
    print('rows:', len(output))
    print('successful lookups:', int(output['graph_conversation_id'].ne('').sum()))
    print('lookup errors:', int(output['lookup_error'].ne('').sum()))
    print('unique graph conversation ids:', output['graph_conversation_id'].replace('', pd.NA).nunique())


if __name__ == '__main__':
    main()
