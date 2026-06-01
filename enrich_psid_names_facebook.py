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

DEFAULT_INPUT_PATH = Path("tropicana contest psid.csv")
DEFAULT_OUTPUT_DIR = Path("outputs")
GRAPH_API_BASE_URL = "https://graph.facebook.com/v20.0"


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
    parser.add_argument("--id-column", default="pageuid")
    parser.add_argument("--token-env", default="TROPICANA_API_TOKEN")
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


def coerce_psid(value: Any) -> str:
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
    DEFAULT_OUTPUT_DIR.mkdir(exist_ok=True)
    return DEFAULT_OUTPUT_DIR / f"{input_path.stem}_facebook_enriched.xlsx"


def parse_json_response(raw_bytes: bytes) -> dict[str, Any]:
    if not raw_bytes:
        return {}
    try:
        return json.loads(raw_bytes.decode("utf-8"))
    except json.JSONDecodeError:
        return {"message": raw_bytes.decode("utf-8", errors="replace")}


def fetch_facebook_profile(psid: str, token: str, timeout_seconds: float) -> dict[str, Any]:
    query = parse.urlencode({
        "fields": "first_name,last_name,name,profile_pic",
        "access_token": token,
    })
    url = f"{GRAPH_API_BASE_URL}/{parse.quote(psid)}?{query}"
    api_request = request.Request(url, headers={"Accept": "application/json"}, method="GET")
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
                "error": {
                    "message": str(exc.reason),
                }
            },
        }


def flatten_profile(psid: str, api_result: dict[str, Any]) -> dict[str, Any]:
    payload = api_result.get("payload") or {}
    error_payload = payload.get("error") or {}
    first_name = clean_text(payload.get("first_name"))
    last_name = clean_text(payload.get("last_name"))
    full_name = clean_text(payload.get("name"))
    if not full_name:
        full_name = " ".join(part for part in [first_name, last_name] if part).strip()

    return {
        "facebook_psid": psid,
        "facebook_http_status": api_result.get("http_status"),
        "facebook_first_name": first_name,
        "facebook_last_name": last_name,
        "facebook_full_name": full_name,
        "facebook_profile_pic": clean_text(payload.get("profile_pic")),
        "facebook_error_message": clean_text(error_payload.get("message")),
        "facebook_error_type": clean_text(error_payload.get("type")),
        "facebook_error_code": clean_text(error_payload.get("code")),
        "facebook_error_subcode": clean_text(error_payload.get("error_subcode")),
        "facebook_trace_id": clean_text(error_payload.get("fbtrace_id")),
        "facebook_lookup_status": "success" if full_name else "error",
    }


def prepare_excel_output(df: pd.DataFrame, id_columns: list[str]) -> pd.DataFrame:
    output = df.copy()
    for column in id_columns:
        if column in output.columns:
            output[column] = output[column].map(coerce_psid)
    return output


def load_input_dataframe(input_path: Path, sheet_name: str) -> pd.DataFrame:
    if input_path.suffix.lower() == ".csv":
        return pd.read_csv(input_path)
    if sheet_name:
        return pd.read_excel(input_path, sheet_name=sheet_name)
    return pd.read_excel(input_path)


def write_outputs(enriched_df: pd.DataFrame, output_path: Path) -> Path:
    output_path.parent.mkdir(exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    candidate_paths = [output_path, output_path.with_name(f"{output_path.stem}_{timestamp}{output_path.suffix}")]

    last_error: Exception | None = None
    for candidate_path in candidate_paths:
        try:
            with pd.ExcelWriter(candidate_path, engine="openpyxl") as writer:
                prepare_excel_output(enriched_df, ["facebook_psid"]).to_excel(writer, sheet_name="psid_profiles", index=False)
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
        raise RuntimeError(f"Set the {args.token_env} environment variable or add it to {env_file_path}.")

    source_df = load_input_dataframe(input_path, args.sheet_name)
    if args.id_column not in source_df.columns:
        raise KeyError(f"Column not found: {args.id_column}")

    working_df = source_df.copy()
    working_df["facebook_psid"] = working_df[args.id_column].map(coerce_psid)
    unique_psids = working_df[["facebook_psid"]].drop_duplicates().copy()
    unique_psids = unique_psids[unique_psids["facebook_psid"] != ""]

    lookup_rows: list[dict[str, Any]] = []
    total = len(unique_psids)

    for index, row in enumerate(unique_psids.itertuples(index=False), start=1):
        psid = row.facebook_psid
        api_result = fetch_facebook_profile(psid, token, args.timeout_seconds)
        lookup_rows.append(flatten_profile(psid, api_result))
        print(f"[{index}/{total}] fetched Facebook profile {psid}")
        if index < total and args.pause_seconds > 0:
            time.sleep(args.pause_seconds)

    lookup_df = pd.DataFrame(lookup_rows)
    enriched_df = working_df.merge(lookup_df, on="facebook_psid", how="left", validate="many_to_one")
    saved_path = write_outputs(enriched_df, output_path)

    success_count = int((enriched_df.get("facebook_lookup_status") == "success").sum()) if "facebook_lookup_status" in enriched_df.columns else 0
    error_count = int((enriched_df.get("facebook_lookup_status") == "error").sum()) if "facebook_lookup_status" in enriched_df.columns else 0

    print(f"Saved: {saved_path}")
    print(f"Rows processed: {len(enriched_df)}")
    print(f"Unique PSIDs looked up: {len(lookup_df)}")
    print(f"Successful profile lookups: {success_count}")
    print(f"Errored profile lookups: {error_count}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(str(exc), file=sys.stderr)
        raise SystemExit(1)
