from pathlib import Path
import re
from datetime import datetime

import numpy as np
import pandas as pd

DATA_FILE = Path('journal-report-all-20260426.xlsx')
OUTPUT_DIR = Path('outputs')
OUTPUT_DIR.mkdir(exist_ok=True)

SELECTED_NAMES = [
    'Mohd Ariff',
    'Jen',
    'LaiLa La',
    'Nur Syifa Syairah',
    'Sheikh',
    'Zahra',
    'Qaseh Maisara',
    'Hawatif Hanim',
    'Aini Nazlifa',
    'Kei Chin',
    'Nadiy',
    'Marc Arviey',
    'Mega',
    'Vincent Ling',
    'Nik Azuan Nik Azlan',
    'Ting Ngik Sieng',
    'Pei',
    'Harraz Mikael',
    'Aina Nadia',
    'Farhan',
]

INELIGIBLE_NAMES = [
    'Ida Rahayu Zainol',
    'Ida Rahayu',
    'Nurul Amri',
    'Nurul Amri Ahmad Hasir',
    'Vinesh Rao',
    'Thomas Chin',
    'Johan Rcmc',
    'Ned Murad',
    'Ah Zhong',
    'Kai Justin',
    'Nora Rara',
    'Byun Stella',
    'Asyraaf Hazim',
    'Cintaku Hati',
    'Melvin Chris',
    'Raimah Othman',
    'Dinie Gulam',
    'Ajjalil A Bakar',
    'Faizah Nor',
    'Fazie Badruddin',
    'Suri Diamond Love',
    'Guan Wen Tai',
    'Yed Rahim',
    'Fiqa Nasaruddin',
    'Nadhirah Karem',
    'Peter Ooi',
    'Syn Wai',
    'Leeya Arrianna',
    'Maliq No',
    'Yatie Afrina',
    'Mohamadi Madie',
    'Monkeys Nami',
    'Ahn Asri',
    'Syed Amir',
    'Mariam Mzn',
    'Mdm Nur',
    'Priya Nair',
    'Su Bee',
    'Bella Asri',
    'Azura AA',
    'Arifull Islam',
    'Asy Asya',
    'Auni Aqila',
    'Najwan Amir',
]

INELIGIBLE_IDS = [
    '26068260039473989',
    '685572017',
    '2058410152',
    '9349164381867289',
    '1459333776',
    '1558303682',
]

ID_COL = 'manychat_id'
NAME_COL = 'participant_name'
JOURNAL_COL = 'journal_text'
POINTS_COL = 'all_time_oranges'

journals = pd.read_excel(DATA_FILE, sheet_name='Raw Journals')
contributors = pd.read_excel(DATA_FILE, sheet_name='Contributors Summary')

participant_cols = [
    ID_COL,
    POINTS_COL,
    'all_time_journals',
    'all_time_pledges',
    'subscribed_at_gmt8',
]

analysis_df = journals.merge(
    contributors[participant_cols],
    on=ID_COL,
    how='left',
    validate='many_to_one',
)

analysis_df['journal_text_clean'] = analysis_df[JOURNAL_COL].fillna('').astype(str).str.strip()
analysis_df['journal_text_length'] = analysis_df['journal_text_clean'].str.len()
analysis_df[POINTS_COL] = pd.to_numeric(analysis_df[POINTS_COL], errors='coerce').fillna(0)
analysis_df['all_time_journals'] = pd.to_numeric(analysis_df['all_time_journals'], errors='coerce').fillna(0)
eligible_journals = analysis_df[analysis_df['journal_text_length'] > 0].copy()

personal_terms = [
    'saya', 'aku', 'kami', 'keluarga', 'ibu', 'mak', 'mama', 'ayah', 'bapa', 'papa',
    'anak', 'isteri', 'suami', 'adik', 'abang', 'kakak', 'nenek', 'datuk', 'rakan',
    'kawan', 'jiran', 'my', 'i ', 'me ', 'we ', 'family', 'mother', 'father', 'friend',
]

kindness_terms = [
    'baik', 'kebaikan', 'bantu', 'membantu', 'tolong', 'menolong', 'hulur', 'kongsi',
    'berkongsi', 'sedekah', 'ikhlas', 'prihatin', 'kasih', 'sayang', 'senyum', 'gembira',
    'terima kasih', 'appreciate', 'kindness', 'help', 'helped', 'share', 'shared', 'care',
]

emotion_terms = [
    'sedih', 'terharu', 'menangis', 'syukur', 'bersyukur', 'gembira', 'bahagia', 'susah',
    'sukar', 'cabaran', 'dugaan', 'letih', 'penat', 'ikhlas', 'tersentuh', 'harapan',
    'sad', 'touched', 'grateful', 'thankful', 'happy', 'struggle', 'challenge', 'hope',
]

story_terms = [
    'pada suatu', 'hari itu', 'ketika', 'semasa', 'waktu', 'masa tu', 'bermula', 'kisah',
    'cerita', 'pengalaman', 'akhirnya', 'selepas', 'sebelum', 'then', 'when', 'after',
    'before', 'story', 'experience',
]

promo_terms = [
    'tropicana twister', 'twister', 'rasa segar', 'hilang dahaga', 'orange', 'oren',
]

def normalize_text(value):
    text = str(value).lower().strip()
    text = re.sub(r'\s+', ' ', text)
    text = re.sub(r'[^\w\s]', '', text)
    return text

def normalize_name(value):
    return re.sub(r'\s+', ' ', str(value).strip().casefold())

def count_terms(text, terms):
    normalized = f' {normalize_text(text)} '
    return sum(1 for term in terms if f' {normalize_text(term)} ' in normalized)

shortlist_df = eligible_journals.copy()
shortlist_df['normalized_journal_text'] = shortlist_df[JOURNAL_COL].map(normalize_text)
shortlist_df['duplicate_text_count'] = shortlist_df.groupby('normalized_journal_text')['journal_id'].transform('count')
shortlist_df['participant_entry_rank_by_length'] = shortlist_df.groupby(ID_COL)['journal_text_length'].rank(method='first', ascending=False)

shortlist_df['personal_term_count'] = shortlist_df[JOURNAL_COL].map(lambda x: count_terms(x, personal_terms))
shortlist_df['kindness_term_count'] = shortlist_df[JOURNAL_COL].map(lambda x: count_terms(x, kindness_terms))
shortlist_df['emotion_term_count'] = shortlist_df[JOURNAL_COL].map(lambda x: count_terms(x, emotion_terms))
shortlist_df['story_term_count'] = shortlist_df[JOURNAL_COL].map(lambda x: count_terms(x, story_terms))
shortlist_df['promo_term_count'] = shortlist_df[JOURNAL_COL].map(lambda x: count_terms(x, promo_terms))

shortlist_df['length_score'] = np.clip(shortlist_df['journal_text_length'] / 1200, 0, 1) * 30
shortlist_df['personal_score'] = np.clip(shortlist_df['personal_term_count'], 0, 5) * 5
shortlist_df['kindness_score'] = np.clip(shortlist_df['kindness_term_count'], 0, 5) * 5
shortlist_df['emotion_score'] = np.clip(shortlist_df['emotion_term_count'], 0, 5) * 4
shortlist_df['story_score'] = np.clip(shortlist_df['story_term_count'], 0, 5) * 4
shortlist_df['duplicate_penalty'] = np.where(shortlist_df['duplicate_text_count'] > 1, 25, 0)
shortlist_df['short_generic_promo_penalty'] = np.where(
    (shortlist_df['journal_text_length'] < 120) & (shortlist_df['promo_term_count'] > 0),
    20,
    0,
)

shortlist_df['heuristic_content_score'] = (
    shortlist_df['length_score']
    + shortlist_df['personal_score']
    + shortlist_df['kindness_score']
    + shortlist_df['emotion_score']
    + shortlist_df['story_score']
    - shortlist_df['duplicate_penalty']
    - shortlist_df['short_generic_promo_penalty']
).round(2)

contributors = contributors.copy()
contributors[POINTS_COL] = pd.to_numeric(contributors[POINTS_COL], errors='coerce').fillna(0)
contributors['normalized_name'] = contributors[NAME_COL].map(normalize_name)
ineligible_normalized_names = {normalize_name(name) for name in INELIGIBLE_NAMES}
explicit_ineligible_ids = pd.to_numeric(pd.Series(INELIGIBLE_IDS), errors='coerce').dropna().astype('int64').tolist()
ineligible_ids = contributors.loc[
    contributors['normalized_name'].isin(ineligible_normalized_names),
    ID_COL,
].dropna().drop_duplicates().tolist()
ineligible_ids = sorted(set(ineligible_ids) | set(explicit_ineligible_ids))
shortlist_df = shortlist_df[~shortlist_df[ID_COL].isin(ineligible_ids)].copy()
contributors = contributors[~contributors[ID_COL].isin(ineligible_ids)].copy()

selection_rows = []
for idx, name in enumerate(SELECTED_NAMES, start=1):
    normalized = normalize_name(name)
    matches = contributors[contributors['normalized_name'] == normalized].copy()
    category = 'main_winner' if idx <= 6 else 'backup'
    if matches.empty:
        selection_rows.append({
            'selection_order': idx,
            'selection_category': category,
            'selected_name_input': name,
            'match_status': 'unmatched',
        })
    else:
        for _, match in matches.iterrows():
            selection_rows.append({
                'selection_order': idx,
                'selection_category': category,
                'selected_name_input': name,
                'match_status': 'matched' if len(matches) == 1 else 'multiple_exact_name_matches',
                NAME_COL: match[NAME_COL],
                ID_COL: match[ID_COL],
                POINTS_COL: match[POINTS_COL],
                'all_time_journals': match.get('all_time_journals'),
                'all_time_pledges': match.get('all_time_pledges'),
            })

selection_resolution = pd.DataFrame(selection_rows)
matched_ids = selection_resolution.loc[selection_resolution[ID_COL].notna(), ID_COL].drop_duplicates().tolist()

best_journal_cols = [
    ID_COL,
    'journal_id',
    'created_at_gmt8',
    JOURNAL_COL,
    'journal_text_length',
    'heuristic_content_score',
    'length_score',
    'personal_score',
    'kindness_score',
    'emotion_score',
    'story_score',
    'duplicate_penalty',
    'short_generic_promo_penalty',
    'personal_term_count',
    'kindness_term_count',
    'emotion_term_count',
    'story_term_count',
    'duplicate_text_count',
]

best_journals = (
    shortlist_df[shortlist_df[ID_COL].isin(matched_ids)]
    .sort_values(['heuristic_content_score', 'journal_text_length'], ascending=[False, False], kind='mergesort')
    .drop_duplicates(ID_COL)[best_journal_cols]
    .copy()
)

selected_with_best_journals = selection_resolution.merge(
    best_journals,
    on=ID_COL,
    how='left',
)

participant_points = contributors.drop(columns=['normalized_name']).sort_values(POINTS_COL, ascending=False, kind='mergesort').copy()
final_consolation_200 = (
    participant_points[~participant_points[ID_COL].isin(matched_ids)]
    .sort_values(POINTS_COL, ascending=False, kind='mergesort')
    .head(200)
    .copy()
)

consolation_best_journal_cols = [
    ID_COL,
    'journal_id',
    'created_at_gmt8',
    JOURNAL_COL,
    'journal_text_length',
    'heuristic_content_score',
    'length_score',
    'personal_score',
    'kindness_score',
    'emotion_score',
    'story_score',
    'duplicate_penalty',
    'short_generic_promo_penalty',
    'personal_term_count',
    'kindness_term_count',
    'emotion_term_count',
    'story_term_count',
    'duplicate_text_count',
]

consolation_best_journals = (
    shortlist_df[shortlist_df[ID_COL].isin(final_consolation_200[ID_COL])]
    .sort_values(['heuristic_content_score', 'journal_text_length'], ascending=[False, False], kind='mergesort')
    .drop_duplicates(ID_COL)[consolation_best_journal_cols]
    .rename(columns={
        'journal_id': 'selected_journal_id',
        'created_at_gmt8': 'selected_journal_created_at_gmt8',
        JOURNAL_COL: 'selected_journal_text',
        'journal_text_length': 'selected_journal_text_length',
        'heuristic_content_score': 'selected_journal_heuristic_content_score',
        'length_score': 'selected_journal_length_score',
        'personal_score': 'selected_journal_personal_score',
        'kindness_score': 'selected_journal_kindness_score',
        'emotion_score': 'selected_journal_emotion_score',
        'story_score': 'selected_journal_story_score',
        'duplicate_penalty': 'selected_journal_duplicate_penalty',
        'short_generic_promo_penalty': 'selected_journal_short_generic_promo_penalty',
        'personal_term_count': 'selected_journal_personal_term_count',
        'kindness_term_count': 'selected_journal_kindness_term_count',
        'emotion_term_count': 'selected_journal_emotion_term_count',
        'story_term_count': 'selected_journal_story_term_count',
        'duplicate_text_count': 'selected_journal_duplicate_text_count',
    })
    .copy()
)

final_consolation_200 = final_consolation_200.merge(
    consolation_best_journals,
    on=ID_COL,
    how='left',
    validate='one_to_one',
)

final_consolation_200['selected_journal_status'] = np.where(
    final_consolation_200['selected_journal_id'].isna(),
    'no_journal_available',
    'selected_highest_heuristic_score',
)
final_consolation_200['selected_journal_selection_rule'] = np.where(
    final_consolation_200['selected_journal_id'].isna(),
    'participant has no journal entries in Raw Journals',
    'highest heuristic_content_score; tie-breaker longest journal_text_length',
)
consolation_without_journals = final_consolation_200[
    final_consolation_200['selected_journal_id'].isna()
].copy()

replacement_candidate_pool = participant_points[
    ~participant_points[ID_COL].isin(matched_ids)
    & ~participant_points[ID_COL].isin(final_consolation_200[ID_COL])
    & participant_points[ID_COL].isin(shortlist_df[ID_COL])
].sort_values(POINTS_COL, ascending=False, kind='mergesort').copy()

replacement_candidate_best_journals = (
    shortlist_df[shortlist_df[ID_COL].isin(replacement_candidate_pool[ID_COL])]
    .sort_values(['heuristic_content_score', 'journal_text_length'], ascending=[False, False], kind='mergesort')
    .drop_duplicates(ID_COL)[consolation_best_journal_cols]
    .rename(columns={
        'journal_id': 'selected_journal_id',
        'created_at_gmt8': 'selected_journal_created_at_gmt8',
        JOURNAL_COL: 'selected_journal_text',
        'journal_text_length': 'selected_journal_text_length',
        'heuristic_content_score': 'selected_journal_heuristic_content_score',
        'length_score': 'selected_journal_length_score',
        'personal_score': 'selected_journal_personal_score',
        'kindness_score': 'selected_journal_kindness_score',
        'emotion_score': 'selected_journal_emotion_score',
        'story_score': 'selected_journal_story_score',
        'duplicate_penalty': 'selected_journal_duplicate_penalty',
        'short_generic_promo_penalty': 'selected_journal_short_generic_promo_penalty',
        'personal_term_count': 'selected_journal_personal_term_count',
        'kindness_term_count': 'selected_journal_kindness_term_count',
        'emotion_term_count': 'selected_journal_emotion_term_count',
        'story_term_count': 'selected_journal_story_term_count',
        'duplicate_text_count': 'selected_journal_duplicate_text_count',
    })
    .copy()
)

replacement_candidates_top_20 = replacement_candidate_pool.head(20).merge(
    replacement_candidate_best_journals,
    on=ID_COL,
    how='left',
    validate='one_to_one',
)
replacement_candidates_top_20['replacement_candidate_reason'] = (
    'next highest all_time_oranges outside current consolation 200, has at least one journal'
)

replacement_count_needed = len(consolation_without_journals)
replacement_candidates_selected = replacement_candidates_top_20.head(replacement_count_needed).copy()
replacement_candidates_selected['selected_journal_status'] = 'selected_highest_heuristic_score'
replacement_candidates_selected['selected_journal_selection_rule'] = (
    'highest heuristic_content_score; tie-breaker longest journal_text_length'
)
replacement_candidates_selected['replacement_status'] = 'replacement_for_no_journal_winner'

final_consolation_200_with_replacements = pd.concat(
    [
        final_consolation_200[final_consolation_200['selected_journal_id'].notna()].copy(),
        replacement_candidates_selected.copy(),
    ],
    ignore_index=True,
    sort=False,
).sort_values(POINTS_COL, ascending=False, kind='mergesort').reset_index(drop=True)

final_consolation_200_with_replacements['replacement_status'] = (
    final_consolation_200_with_replacements['replacement_status'].fillna('original_consolation_winner')
)

def prepare_excel_output(df):
    output = df.copy()
    if ID_COL in output.columns:
        output[ID_COL] = output[ID_COL].map(lambda value: '' if pd.isna(value) else str(int(value)) if isinstance(value, float) and value.is_integer() else str(value))
    return output

selected_output_path = OUTPUT_DIR / 'final_selected_winners_and_backups_best_journals.xlsx'
resolution_output_path = OUTPUT_DIR / 'final_selected_name_resolution.xlsx'
consolation_output_path = OUTPUT_DIR / 'final_consolation_200_excluding_selected.xlsx'
consolation_without_journals_output_path = OUTPUT_DIR / 'final_consolation_winners_without_journals.xlsx'
replacement_candidates_output_path = OUTPUT_DIR / 'replacement_candidates_next_20_with_journals.xlsx'
consolation_with_replacements_output_path = OUTPUT_DIR / 'final_consolation_200_with_journal_replacements.xlsx'

timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')

def write_excel_with_fallback(df, path):
    try:
        prepare_excel_output(df).to_excel(path, index=False)
        return path
    except PermissionError:
        fallback_path = path.with_name(f'{path.stem}_{timestamp}{path.suffix}')
        prepare_excel_output(df).to_excel(fallback_path, index=False)
        return fallback_path

selected_output_path = write_excel_with_fallback(selected_with_best_journals, selected_output_path)
resolution_output_path = write_excel_with_fallback(selection_resolution, resolution_output_path)
consolation_output_path = write_excel_with_fallback(final_consolation_200, consolation_output_path)
consolation_without_journals_output_path = write_excel_with_fallback(consolation_without_journals, consolation_without_journals_output_path)
replacement_candidates_output_path = write_excel_with_fallback(replacement_candidates_top_20, replacement_candidates_output_path)
consolation_with_replacements_output_path = write_excel_with_fallback(final_consolation_200_with_replacements, consolation_with_replacements_output_path)

print('Selected names provided:', len(SELECTED_NAMES))
print('Excluded ineligible names provided:', len(INELIGIBLE_NAMES))
print('Excluded ineligible participant IDs:', len(ineligible_ids))
print('Unique matched participant IDs excluded:', len(matched_ids))
print('Unmatched selected names:', int((selection_resolution['match_status'] == 'unmatched').sum()))
print('Multiple exact-name match rows:', int((selection_resolution['match_status'] == 'multiple_exact_name_matches').sum()))
print('Final consolation winners:', len(final_consolation_200))
print('Saved:', selected_output_path)
print('Saved:', resolution_output_path)
print('Saved:', consolation_output_path)
print('Saved:', consolation_without_journals_output_path)
print('Saved:', replacement_candidates_output_path)
print('Saved:', consolation_with_replacements_output_path)
print('Consolation winners without journals:', len(consolation_without_journals))
print('Replacement candidates with journals:', len(replacement_candidates_top_20))
print('Final consolation winners with replacements:', len(final_consolation_200_with_replacements))
print('\nTop 20 final consolation winners:')
print(final_consolation_200[[NAME_COL, ID_COL, POINTS_COL, 'all_time_journals', 'all_time_pledges']].head(20).to_string(index=False))
