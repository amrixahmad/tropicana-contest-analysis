import os
import re
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse

import pandas as pd
import requests

from enrich_manychat_outreach import load_env_file

OUTREACH_FILE = Path('outputs/final_consolation_200_with_journal_replacements_manychat_enriched.xlsx')
LOOKUP_FILE = Path('outputs/200_winners_psid_name_lookup.xlsx')
OUTPUT_FILE = Path('outputs/consolation_response_validation.xlsx')
GRAPH_BASE_URL = 'https://graph.facebook.com/v25.0'
GRAPH_CONVERSATIONS_URL = f'{GRAPH_BASE_URL}/me/conversations'
REQUEST_DELAY_SECONDS = 0.35
PAGE_NAME = 'Tropicana Malaysia'
PAGE_ID = '594456394024035'
OUTREACH_STATUS_COL = '1st outreach status'
OUTREACH_NAME_COL = 'outreach_name'
MATCH_NAME_COL = 'participant_name_from_graph'
FIRST_OUTREACH_PHRASE = (
    'Tahniah! Anda telah dipilih sebagai salah seorang pemenang kempen Gandakan Kebaikan '
    'Bersama Tropicana Twister. Sila hantar dalam tempoh 7 hari maklumat berikut bagi tujuan '
    'penebusan hadiah:'
)
MESSAGE_FIELDS = (
    'id,created_time,from,to,message,tags,'
    'attachments{mime_type,file_url,image_data,animated_image_data,video_data,name,id,url,description}'
)


def clean_text(value):
    if pd.isna(value):
        return ''
    return str(value).strip()


def normalize_text(value):
    text = re.sub(r'\s+', ' ', clean_text(value).casefold())
    return re.sub(r'\s+\d+$', '', text)


def is_page_sender(sender_name, sender_id):
    return bool(
        clean_text(sender_id) == PAGE_ID
        or normalize_text(sender_name) == normalize_text(PAGE_NAME)
    )


def graph_get(token, path, params=None, absolute_url=''):
    request_url = absolute_url or f'{GRAPH_BASE_URL}/{path}'
    response = requests.get(
        request_url,
        params=None if absolute_url else {**(params or {}), 'access_token': token},
        timeout=30,
    )
    try:
        payload = response.json()
    except ValueError:
        payload = {'raw_response': response.text}
    return response, payload


def lookup_conversation(token, facebook_psid):
    response = requests.get(
        GRAPH_CONVERSATIONS_URL,
        params={
            'user_id': facebook_psid,
            'fields': 'id,link,updated_time,participants,message_count,unread_count,snippet,can_reply,is_subscribed',
            'access_token': token,
        },
        timeout=30,
    )
    try:
        payload = response.json()
    except ValueError:
        payload = {'raw_response': response.text}

    result = {
        'lookup_http_status': response.status_code,
        'conversation_id': '',
        'facebook_inbox_link': '',
        'conversation_updated_time': '',
        'lookup_error': '',
    }
    if not response.ok:
        result['lookup_error'] = payload.get('error', {}).get('message', str(payload))
        return result

    conversations = payload.get('data') or []
    if not conversations:
        result['lookup_error'] = 'no conversation returned'
        return result

    conversation = conversations[0]
    link = clean_text(conversation.get('link', ''))
    if link.startswith('/'):
        link = f'https://www.facebook.com{link}'
    result.update({
        'conversation_id': clean_text(conversation.get('id', '')),
        'facebook_inbox_link': link,
        'conversation_updated_time': clean_text(conversation.get('updated_time', '')),
    })
    return result


def is_outreach_boundary_message(message):
    message_text = clean_text(message.get('message', ''))
    sender = message.get('from') or {}
    sender_name = clean_text(sender.get('name', ''))
    sender_id = clean_text(sender.get('id', ''))
    return (
        normalize_text(FIRST_OUTREACH_PHRASE) in normalize_text(message_text)
        and is_page_sender(sender_name, sender_id)
    )


def fetch_messages_until_outreach(token, conversation_id):
    messages = []
    outreach_boundary = {}
    next_url = ''
    params = {'fields': MESSAGE_FIELDS, 'limit': 100}

    while True:
        response, payload = graph_get(token, f'{conversation_id}/messages', params=params, absolute_url=next_url)
        if not response.ok:
            error_text = payload.get('error', {}).get('message', str(payload))
            raise RuntimeError(f'messages lookup failed: {response.status_code} {error_text}')

        batch = payload.get('data') or []
        if not batch:
            break

        for message in batch:
            message_text = clean_text(message.get('message', ''))
            if is_outreach_boundary_message(message):
                outreach_boundary = {
                    'outreach_message_id': clean_text(message.get('id', '')),
                    'outreach_message_created_time': clean_text(message.get('created_time', '')),
                    'outreach_message_from_name': clean_text((message.get('from') or {}).get('name', '')),
                    'outreach_message_from_id': clean_text((message.get('from') or {}).get('id', '')),
                    'outreach_message_text_length': len(message_text),
                }
                return messages, outreach_boundary
            messages.append(message)

        next_url = clean_text((payload.get('paging') or {}).get('next', ''))
        if not next_url:
            break
        params = None

    return messages, outreach_boundary


def collect_tags(message):
    tag_rows = (message.get('tags') or {}).get('data') or []
    names = [clean_text(tag.get('name', '')) for tag in tag_rows if clean_text(tag.get('name', ''))]
    return ' | '.join(names)


def iter_urls(value, path=''):
    if isinstance(value, dict):
        for key, nested in value.items():
            nested_path = f'{path}.{key}' if path else str(key)
            yield from iter_urls(nested, nested_path)
    elif isinstance(value, list):
        for index, nested in enumerate(value):
            nested_path = f'{path}[{index}]'
            yield from iter_urls(nested, nested_path)
    elif isinstance(value, str):
        if value.startswith('http://') or value.startswith('https://'):
            yield path, value


def classify_image_url(url, path):
    lower_url = clean_text(url).casefold()
    lower_path = clean_text(path).casefold()
    image_extensions = ('.jpg', '.jpeg', '.png', '.webp', '.gif', '.bmp', '.heic', '.heif')
    return (
        lower_url.endswith(image_extensions)
        or any(token in lower_url for token in ['fbcdn.net', 'scontent', 'image', 'jpeg', 'jpg', 'png'])
        or any(token in lower_path for token in ['image_data', 'preview_url', 'image'])
    )


def extract_attachment_rows(message, base_row):
    attachment_rows = []
    attachments = (message.get('attachments') or {}).get('data') or []
    for attachment_index, attachment in enumerate(attachments, start=1):
        urls = list(iter_urls(attachment))
        image_urls = [url for path, url in urls if classify_image_url(url, path)]
        primary_url = image_urls[0] if image_urls else (urls[0][1] if urls else '')
        preview_url = image_urls[1] if len(image_urls) > 1 else ''
        image_url_count = len(image_urls)
        attachment_rows.append({
            **base_row,
            'attachment_index': attachment_index,
            'attachment_id': clean_text(attachment.get('id', '')),
            'attachment_name': clean_text(attachment.get('name', '')),
            'attachment_mime_type': clean_text(attachment.get('mime_type', '')),
            'primary_attachment_url': primary_url,
            'primary_attachment_url_host': urlparse(primary_url).netloc if primary_url else '',
            'preview_url': preview_url,
            'preview_url_host': urlparse(preview_url).netloc if preview_url else '',
            'image_url_count': image_url_count,
            'all_detected_urls': ' | '.join(url for _, url in urls),
            'message_tag_names': collect_tags(message),
            'is_likely_proof_of_purchase': bool(base_row['is_from_winner'] and image_url_count > 0),
        })
    return attachment_rows


def build_message_row(winner_row, message_rank_latest_first, message, facebook_psid):
    sender = message.get('from') or {}
    sender_id = clean_text(sender.get('id', ''))
    sender_name = clean_text(sender.get('name', ''))
    message_text = clean_text(message.get('message', ''))
    is_from_winner = sender_id == facebook_psid or normalize_text(sender_name) == normalize_text(winner_row.get(MATCH_NAME_COL, ''))
    attachments = (message.get('attachments') or {}).get('data') or []
    return {
        'outreach_name': winner_row.get(OUTREACH_NAME_COL, ''),
        'facebook_psid': facebook_psid,
        'conversation_id': winner_row.get('conversation_id', ''),
        'message_rank_latest_first': message_rank_latest_first,
        'message_id': clean_text(message.get('id', '')),
        'message_created_time': clean_text(message.get('created_time', '')),
        'message_from_name': sender_name,
        'message_from_id': sender_id,
        'is_from_winner': is_from_winner,
        'message_text': message_text,
        'message_text_length': len(message_text),
        'attachment_count': len(attachments),
        'message_tag_names': collect_tags(message),
    }


def write_excel_with_fallback(output_path, sheets):
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    candidate_paths = [output_path, output_path.with_name(f'{output_path.stem}_{timestamp}{output_path.suffix}')]
    last_error = None
    for candidate_path in candidate_paths:
        try:
            with pd.ExcelWriter(candidate_path, engine='openpyxl') as writer:
                for sheet_name, df in sheets.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            return candidate_path
        except PermissionError as exc:
            last_error = exc
            continue
    if last_error is not None:
        raise last_error
    return output_path


def main():
    load_env_file()
    token = (os.environ.get('TROPICANA_API_TOKEN') or os.environ.get('TROPICANA_PAGE_API_KEY') or '').strip()
    if not token:
        raise SystemExit('Set TROPICANA_API_TOKEN or TROPICANA_PAGE_API_KEY before running this script.')

    outreach = pd.read_excel(OUTREACH_FILE, dtype=str).fillna('')
    lookup = pd.read_excel(LOOKUP_FILE, dtype=str).fillna('')

    if OUTREACH_STATUS_COL not in outreach.columns:
        raise KeyError(f'Expected column not found in outreach file: {OUTREACH_STATUS_COL}')
    if MATCH_NAME_COL not in lookup.columns or 'facebook_psid' not in lookup.columns:
        raise KeyError('Expected facebook_psid and participant_name_from_graph in lookup file.')

    lookup = lookup.copy()
    outreach = outreach.copy()
    lookup['normalized_match_name'] = lookup[MATCH_NAME_COL].map(normalize_text)
    outreach['normalized_outreach_name'] = outreach[OUTREACH_NAME_COL].map(normalize_text)

    replied = outreach[outreach[OUTREACH_STATUS_COL].eq('Replied')].copy()
    merged = replied.merge(
        lookup[['facebook_psid', MATCH_NAME_COL, 'normalized_match_name']],
        left_on='normalized_outreach_name',
        right_on='normalized_match_name',
        how='left',
        validate='one_to_one',
    )

    summary_rows = []
    message_rows = []
    attachment_rows = []
    error_rows = []

    total = len(merged)
    for index, winner in enumerate(merged.to_dict(orient='records'), start=1):
        facebook_psid = clean_text(winner.get('facebook_psid', ''))
        winner_name = winner.get(OUTREACH_NAME_COL, '')
        try:
            if not facebook_psid:
                raise RuntimeError('missing facebook_psid after name matching')

            conversation_lookup = lookup_conversation(token, facebook_psid)
            if conversation_lookup.get('lookup_error'):
                raise RuntimeError(conversation_lookup['lookup_error'])

            winner['conversation_id'] = conversation_lookup.get('conversation_id', '')
            winner['facebook_inbox_link_from_psid'] = conversation_lookup.get('facebook_inbox_link', '')
            winner['conversation_updated_time'] = conversation_lookup.get('conversation_updated_time', '')

            response_messages, outreach_boundary = fetch_messages_until_outreach(token, winner['conversation_id'])
            if not outreach_boundary:
                raise RuntimeError('outreach boundary phrase not found in message history')

            collected_rows = []
            for rank, message in enumerate(response_messages, start=1):
                message_row = build_message_row(winner, rank, message, facebook_psid)
                message_rows.append(message_row)
                collected_rows.append(message_row)
                attachment_rows.extend(extract_attachment_rows(message, message_row))

            chronological_messages = sorted(collected_rows, key=lambda row: row['message_created_time'])
            winner_messages = [row['message_text'] for row in chronological_messages if row['is_from_winner'] and row['message_text']]
            all_messages = [row['message_text'] for row in chronological_messages if row['message_text']]
            winner_attachment_rows = [row for row in attachment_rows if row['outreach_name'] == winner_name and row['is_from_winner']]
            proof_rows = [row for row in winner_attachment_rows if row['is_likely_proof_of_purchase']]

            summary_rows.append({
                'outreach_name': winner_name,
                'manychat_id': winner.get('manychat_id', ''),
                'facebook_psid': facebook_psid,
                'FB Inbox': winner.get('FB Inbox', ''),
                'conversation_id': winner.get('conversation_id', ''),
                'conversation_updated_time': winner.get('conversation_updated_time', ''),
                'reply_window_message_count': len(collected_rows),
                'reply_window_winner_message_count': sum(1 for row in collected_rows if row['is_from_winner']),
                'reply_window_attachment_count': len(winner_attachment_rows),
                'proof_image_count': len(proof_rows),
                'latest_reply_time': chronological_messages[-1]['message_created_time'] if chronological_messages else '',
                'earliest_reply_time_after_outreach': chronological_messages[0]['message_created_time'] if chronological_messages else '',
                'outreach_message_id': outreach_boundary.get('outreach_message_id', ''),
                'outreach_message_created_time': outreach_boundary.get('outreach_message_created_time', ''),
                'winner_reply_text_combined': '\n\n'.join(winner_messages),
                'all_message_text_combined': '\n\n'.join(all_messages),
                'proof_image_urls': '\n'.join(row['primary_attachment_url'] for row in proof_rows if row['primary_attachment_url']),
            })
        except Exception as exc:
            error_rows.append({
                'outreach_name': winner_name,
                'facebook_psid': facebook_psid,
                'error': str(exc),
            })
        print(f'[{index}/{total}] {winner_name} {facebook_psid}: done')
        if index < total:
            time.sleep(REQUEST_DELAY_SECONDS)

    summary_df = pd.DataFrame(summary_rows)
    message_df = pd.DataFrame(message_rows)
    attachment_df = pd.DataFrame(attachment_rows)
    error_df = pd.DataFrame(error_rows)

    saved_path = write_excel_with_fallback(
        OUTPUT_FILE,
        {
            'validation_summary': summary_df,
            'response_messages': message_df,
            'attachment_url_details': attachment_df,
            'errors': error_df,
        },
    )

    print(f'Saved: {saved_path}')
    print('replied_rows_processed:', len(merged))
    print('summary_rows:', len(summary_df))
    print('message_rows:', len(message_df))
    print('attachment_rows:', len(attachment_df))
    print('errors:', len(error_df))


if __name__ == '__main__':
    main()
