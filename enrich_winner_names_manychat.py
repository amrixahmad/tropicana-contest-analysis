from __future__ import annotations

import argparse
import json
import os
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Any
from urllib import error, parse, request

import pandas as pd

DEFAULT_INPUT_PATH = Path("outputs") / "final_selected_winners_and_backups_best_journals.xlsx"
API_BASE_URL = "https://api.manychat.com"


def load_dotenv_file(env_path: Path) -> None:
    if not env_path.exists():
        return

    for raw_line in env_path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue

        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = value


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", default=str(DEFAULT_INPUT_PATH))
    parser.add_argument("--output", default="")
    parser.add_argument("--sheet-name", default="")
    parser.add_argument("--id-column", default="manychat_id")
    parser.add_argument("--name-column", default="participant_name")
    parser.add_argument("--token-env", default="MANYCHAT_API_TOKEN")
    parser.add_argument("--env-file", default=".env")
    parser.add_argument("--pause-seconds", type=float, default=0.12)
    parser.add_argument("--timeout-seconds", type=float, default=30.0)
    return parser.parse_args()


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    text = str(value).strip()
    if text.lower() == "nan":
        return ""
    return text


def normalize_name(value: Any) -> str:
    return " ".join(clean_text(value).casefold().split())


def coerce_subscriber_id(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float):
        if pd.isna(value):
            return ""
        if value.is_integer():
            return str(int(value))
        return clean_text(value)
    if isinstance(value, int):
        return str(value)
    text = clean_text(value)
    if not text:
        return ""
    normalized = text.replace(",", "")
    if normalized.isdigit():
        return normalized
    if normalized.endswith(".0") and normalized[:-2].isdigit():
        return normalized[:-2]
    return normalized


def default_output_path(input_path: Path) -> Path:
    return input_path.with_name(f"{input_path.stem}_manychat_enriched{input_path.suffix}")


def prepare_excel_output(df: pd.DataFrame, id_columns: list[str]) -> pd.DataFrame:
    output = df.copy()
    for column in id_columns:
        if column in output.columns:
            output[column] = output[column].map(coerce_subscriber_id)
    return output


def parse_json_response(raw_bytes: bytes) -> dict[str, Any]:
    if not raw_bytes:
        return {}
    try:
        return json.loads(raw_bytes.decode("utf-8"))
    except json.JSONDecodeError:
        return {"status": "error", "message": raw_bytes.decode("utf-8", errors="replace")}


def fetch_manychat_subscriber(subscriber_id: str, token: str, timeout_seconds: float) -> dict[str, Any]:
    url = f"{API_BASE_URL}/fb/subscriber/getInfo?{parse.urlencode({'subscriber_id': subscriber_id})}"
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
    }
    api_request = request.Request(url, headers=headers, method="GET")
    try:
        with request.urlopen(api_request, timeout=timeout_seconds) as response:
            payload = parse_json_response(response.read())
            return {
                "http_status": response.status,
                "payload": payload,
            }
    except error.HTTPError as exc:
        payload = parse_json_response(exc.read())
        return {
            "http_status": exc.code,
            "payload": payload,
        }
    except error.URLError as exc:
        return {
            "http_status": None,
            "payload": {
                "status": "error",
                "message": str(exc.reason),
            },
        }


def choose_best_name(existing_name: str, manychat_full_name: str) -> tuple[str, str]:
    existing_tokens = [token for token in existing_name.split() if token]
    manychat_tokens = [token for token in manychat_full_name.split() if token]

    if manychat_full_name and not existing_name:
        return manychat_full_name, "filled_from_manychat"
    if existing_name and not manychat_full_name:
        return existing_name, "manychat_name_missing"
    if not existing_name and not manychat_full_name:
        return "", "no_name_available"
    if normalize_name(existing_name) == normalize_name(manychat_full_name):
        return manychat_full_name or existing_name, "matched"
    if len(manychat_tokens) > len(existing_tokens):
        return manychat_full_name, "manychat_has_more_complete_name"
    if len(existing_tokens) > len(manychat_tokens):
        return existing_name, "existing_name_more_complete"
    return manychat_full_name, "name_conflict_review_needed"


def flatten_subscriber(subscriber_id: str, existing_name: str, api_result: dict[str, Any]) -> dict[str, Any]:
    payload = api_result.get("payload") or {}
    data = payload.get("data") or {}
    first_name = clean_text(data.get("first_name"))
    last_name = clean_text(data.get("last_name"))
    full_name = clean_text(data.get("name"))
    if not full_name:
        full_name = " ".join(part for part in [first_name, last_name] if part).strip()

    best_name, best_name_status = choose_best_name(clean_text(existing_name), full_name)

    custom_fields = data.get("custom_fields") or []
    tags = data.get("tags") or []

    return {
        "manychat_id": subscriber_id,
        "manychat_http_status": api_result.get("http_status"),
        "manychat_api_status": clean_text(payload.get("status")),
        "manychat_api_message": clean_text(payload.get("message")),
        "manychat_first_name": first_name,
        "manychat_last_name": last_name,
        "manychat_full_name": full_name,
        "best_outreach_name": best_name,
        "best_outreach_name_status": best_name_status,
        "manychat_phone": clean_text(data.get("phone")),
        "manychat_email": clean_text(data.get("email")),
        "manychat_profile_pic": clean_text(data.get("profile_pic")),
        "manychat_locale": clean_text(data.get("locale")),
        "manychat_language": clean_text(data.get("language")),
        "manychat_timezone": clean_text(data.get("timezone")),
        "manychat_page_id": clean_text(data.get("page_id")),
        "manychat_ig_username": clean_text(data.get("ig_username")),
        "manychat_ig_id": clean_text(data.get("ig_id")),
        "manychat_whatsapp_phone": clean_text(data.get("whatsapp_phone")),
        "manychat_last_input_text": clean_text(data.get("last_input_text")),
        "manychat_live_chat_url": clean_text(data.get("live_chat_url")),
        "manychat_subscribed": clean_text(data.get("subscribed")),
        "manychat_last_interaction": clean_text(data.get("last_interaction")),
        "manychat_last_seen": clean_text(data.get("last_seen")),
        "manychat_custom_fields_json": json.dumps(custom_fields, ensure_ascii=False),
        "manychat_tags_json": json.dumps(tags, ensure_ascii=False),
    }


def load_input_dataframe(input_path: Path, sheet_name: str) -> pd.DataFrame:
    if sheet_name:
        return pd.read_excel(input_path, sheet_name=sheet_name)
    return pd.read_excel(input_path)


def build_outreach_df(enriched_df: pd.DataFrame) -> pd.DataFrame:
    preferred_columns = [
        "best_outreach_name",
        "best_outreach_name_status",
        "participant_name",
        "manychat_full_name",
        "manychat_first_name",
        "manychat_last_name",
        "manychat_id",
        "manychat_phone",
        "manychat_email",
        "manychat_last_interaction",
        "manychat_last_seen",
        "manychat_live_chat_url",
        "replacement_status",
        "selected_journal_status",
        "selected_journal_text",
        "all_time_oranges",
    ]
    available_columns = [column for column in preferred_columns if column in enriched_df.columns]
    outreach_df = enriched_df[available_columns].copy()
    rename_map = {
        "best_outreach_name": "outreach_name",
        "participant_name": "original_participant_name",
    }
    outreach_df = outreach_df.rename(columns={key: value for key, value in rename_map.items() if key in outreach_df.columns})
    return outreach_df


def build_detailed_records_df(enriched_df: pd.DataFrame) -> pd.DataFrame:
    preferred_front_columns = [
        "best_outreach_name",
        "best_outreach_name_status",
        "participant_name",
        "manychat_full_name",
        "manychat_first_name",
        "manychat_last_name",
        "manychat_id",
        "replacement_status",
        "manychat_phone",
        "manychat_email",
        "manychat_last_interaction",
        "manychat_last_seen",
        "manychat_live_chat_url",
    ]
    remaining_columns = [column for column in enriched_df.columns if column not in preferred_front_columns]
    ordered_columns = [column for column in preferred_front_columns if column in enriched_df.columns] + remaining_columns
    return enriched_df[ordered_columns].copy()


def write_outputs(enriched_df: pd.DataFrame, lookup_df: pd.DataFrame, output_path: Path) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    candidate_paths = [output_path, output_path.with_name(f"{output_path.stem}_{timestamp}{output_path.suffix}")]

    last_error: Exception | None = None
    for candidate_path in candidate_paths:
        try:
            outreach_df = build_outreach_df(enriched_df)
            detailed_records_df = build_detailed_records_df(enriched_df)
            with pd.ExcelWriter(candidate_path, engine="openpyxl") as writer:
                prepare_excel_output(outreach_df, ["manychat_id"]).to_excel(writer, sheet_name="outreach_list", index=False)
                prepare_excel_output(detailed_records_df, ["manychat_id"]).to_excel(writer, sheet_name="detailed_records", index=False)
            return candidate_path
        except PermissionError as exc:
            last_error = exc
            continue

    if last_error is not None:
        raise last_error
    return output_path


def main() -> int:
    args = parse_args()
    input_path = Path(args.input)
    output_path = Path(args.output) if args.output else default_output_path(input_path)
    env_file_path = Path(args.env_file)

    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    load_dotenv_file(env_file_path)
    token = os.getenv(args.token_env, "").strip()
    if not token:
        raise RuntimeError(
            f"Set the {args.token_env} environment variable or add it to {env_file_path}."
        )

    source_df = load_input_dataframe(input_path, args.sheet_name)
    if args.id_column not in source_df.columns:
        raise KeyError(f"Column not found: {args.id_column}")
    if args.name_column not in source_df.columns:
        raise KeyError(f"Column not found: {args.name_column}")

    working_df = source_df.copy()
    working_df["manychat_id"] = working_df[args.id_column].map(coerce_subscriber_id)
    working_df["existing_name"] = working_df[args.name_column].map(clean_text)

    unique_lookups = (
        working_df[["manychat_id", "existing_name"]]
        .drop_duplicates(subset=["manychat_id"])
        .copy()
    )
    unique_lookups = unique_lookups[unique_lookups["manychat_id"] != ""]

    lookup_rows: list[dict[str, Any]] = []
    total = len(unique_lookups)

    for index, row in enumerate(unique_lookups.itertuples(index=False), start=1):
        subscriber_id = row.manychat_id
        existing_name = row.existing_name
        api_result = fetch_manychat_subscriber(subscriber_id, token, args.timeout_seconds)
        lookup_rows.append(flatten_subscriber(subscriber_id, existing_name, api_result))
        print(f"[{index}/{total}] fetched ManyChat subscriber {subscriber_id}")
        if index < total and args.pause_seconds > 0:
            time.sleep(args.pause_seconds)

    lookup_df = pd.DataFrame(lookup_rows)
    enriched_df = working_df.merge(lookup_df, on="manychat_id", how="left", validate="many_to_one")

    saved_path = write_outputs(enriched_df, lookup_df, output_path)

    matched_count = int((enriched_df.get("best_outreach_name_status") == "matched").sum()) if "best_outreach_name_status" in enriched_df.columns else 0
    improved_count = int((enriched_df.get("best_outreach_name_status") == "manychat_has_more_complete_name").sum()) if "best_outreach_name_status" in enriched_df.columns else 0
    filled_count = int((enriched_df.get("best_outreach_name_status") == "filled_from_manychat").sum()) if "best_outreach_name_status" in enriched_df.columns else 0

    print(f"Saved: {saved_path}")
    print(f"Rows processed: {len(enriched_df)}")
    print(f"Unique subscriber lookups: {len(lookup_df)}")
    print(f"Matched existing names: {matched_count}")
    print(f"More complete ManyChat names: {improved_count}")
    print(f"Filled missing names from ManyChat: {filled_count}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(str(exc), file=sys.stderr)
        raise SystemExit(1)
