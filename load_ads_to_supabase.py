import os
import re
import hashlib
from datetime import datetime, timezone

import google.auth
import pandas as pd
from googleapiclient.discovery import build
from supabase import create_client, Client


SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
]

# Column F = index 5 (zero-based index: A=0, B=1, C=2, D=3, E=4, F=5)
AD_TYPE_COLUMN_INDEX = 5

# Image ads source tab
IMAGE_ADS_SOURCE_TAB = "Clean Data 2"

TABLE_BY_AD_TYPE = {
    "image": "image_ads",
    "video": "video_ads",
    "text": "text_ads",
}

SELECTED_COLUMNS_BY_AD_TYPE = {
    "image": [
        "advertiser",
        "name",
        "app_link",
        "image_url",
        "claim_time",
    ],
    "video": [
        "advertiser",
        "name",
        "app_link",
        "youtube_url",
        "claim_time",
        "views",
    ],
    "text": [
        "advertiser",
        "name",
        "app_link",
        "headline",
        "claim_time",
        "description",
    ],
}

IMAGE_HASH_COLUMNS = [
    "image_url",
]

VIDEO_HASH_COLUMNS = [
    "advertiser",
    "_ad_type_raw",
]

TEXT_HASH_COLUMNS = [
    "_source_spreadsheet_id",
    "_source_tab",
    "_source_row_number",
    "_ad_type_raw",
    "_ad_type",
]

METADATA_COLUMNS = [
    "_row_hash",
    "_source_spreadsheet_id",
    "_source_tab",
    "_source_row_number",
    "_ad_type_raw",
    "_ad_type",
    "_loaded_at_utc",
]


def required_env(name: str) -> str:
    value = os.getenv(name)
    if not value:
        raise RuntimeError(f"Missing required environment variable: {name}")
    return value.strip()


def normalize_spreadsheet_id(value: str) -> str:
    value = str(value or "").strip()
    value = value.strip('"').strip("'")
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", value)
    if match:
        return match.group(1).strip()
    return value


def normalize_compare_text(value: str) -> str:
    return re.sub(
        r"\s+",
        " ",
        str(value or "").replace("\u00a0", " ")
    ).strip().lower()


def get_spreadsheet_ids() -> list[str]:
    value = required_env("SPREADSHEET_IDS")
    raw_items = re.split(r"[,;\n\r\t ]+", value)
    spreadsheet_ids = []

    for item in raw_items:
        spreadsheet_id = normalize_spreadsheet_id(item)
        if spreadsheet_id:
            spreadsheet_ids.append(spreadsheet_id)

    spreadsheet_ids = list(dict.fromkeys(spreadsheet_ids))
    if not spreadsheet_ids:
        raise RuntimeError("SPREADSHEET_IDS is empty.")
    return spreadsheet_ids


def quote_sheet_range(sheet_name: str) -> str:
    safe_name = sheet_name.replace("'", "''")
    return f"'{safe_name}'!A:ZZ"


def sanitize_column_name(name: str, fallback: str) -> str:
    name = str(name or "").strip().lower()
    name = re.sub(r"[^a-z0-9_]+", "_", name)
    name = re.sub(r"_+", "_", name).strip("_")

    if not name:
        name = fallback
    if not re.match(r"^[a-z_]", name):
        name = f"_{name}"
    return name[:300]


def make_unique_columns(columns):
    result = []
    seen = {}

    for i, col in enumerate(columns):
        base = sanitize_column_name(col, f"col_{i + 1}")
        count = seen.get(base, 0) + 1
        seen[base] = count

        if count == 1:
            result.append(base)
        else:
            suffix = f"_{count}"
            result.append(f"{base[:300 - len(suffix)]}{suffix}")

    return result


def is_valid_video_id(value: str) -> bool:
    value = str(value or "").strip().lower()
    return bool(re.fullmatch(r"[a-f0-9]{16}", value))


def normalize_ad_type(value: str) -> str:
    raw_value = str(value or "").strip()
    lower_value = raw_value.lower()
    lower_value = lower_value.replace("_", " ").replace("-", " ")
    lower_value = re.sub(r"\s+", " ", lower_value).strip()

    skip_values = {
        "", "n/a", "na", "n.a", "n.a.", "none", "null", "-", "--", "error", "err", "failed", "failure",
    }

    if lower_value in skip_values:
        return "skip"
    if lower_value in {"text", "text ad", "text ads", "copy"}:
        return "text"
    if lower_value in {"image", "image ad", "image ads", "img"}:
        return "image"
    if is_valid_video_id(raw_value):
        return "video"
    return "skip"


def clean_key_value(value) -> str:
    return str(value or "").strip().lower()


def get_all_sheet_names(sheets_service, spreadsheet_id: str):
    spreadsheet = (
        sheets_service.spreadsheets()
        .get(
            spreadsheetId=spreadsheet_id,
            fields="sheets.properties.title",
        )
        .execute()
    )
    return [sheet["properties"]["title"] for sheet in spreadsheet.get("sheets", [])]


def find_matching_tab_name(sheets_service, spreadsheet_id: str, expected_tab_name: str) -> str:
    sheet_names = get_all_sheet_names(sheets_service=sheets_service, spreadsheet_id=spreadsheet_id)
    expected_norm = normalize_compare_text(expected_tab_name)

    for sheet_name in sheet_names:
        if normalize_compare_text(sheet_name) == expected_norm:
            return sheet_name

    raise RuntimeError(
        f"Could not find tab '{expected_tab_name}' in spreadsheet '{spreadsheet_id}'. "
        f"Available tabs: {sheet_names}"
    )


def read_sheet_tab(
    sheets_service,
    spreadsheet_id: str,
    tab_name: str,
    image_only_extra_read: bool = False,
) -> pd.DataFrame:
    response = (
        sheets_service.spreadsheets()
        .values()
        .get(
            spreadsheetId=spreadsheet_id,
            range=quote_sheet_range(tab_name),
        )
        .execute()
    )

    values = response.get("values", [])
    if not values:
        print(f"Skipping empty tab: {tab_name}")
        return pd.DataFrame()

    width = max(len(row) for row in values)
    if width <= AD_TYPE_COLUMN_INDEX:
        raise RuntimeError(
            f"Spreadsheet '{spreadsheet_id}', tab '{tab_name}' does not have column F. "
            "Column F is required to identify image/text/video ID/N/A."
        )

    raw_header = values[0] + [f"col_{i + 1}" for i in range(len(values[0]), width)]
    header = make_unique_columns(raw_header)
    ad_type_column_name = header[AD_TYPE_COLUMN_INDEX]

    rows = []
    row_numbers = []

    for sheet_row_number, row in enumerate(values[1:], start=2):
        padded = row + [""] * (width - len(row))
        if not any(str(cell).strip() for cell in padded):
            continue
        rows.append(padded)
        row_numbers.append(sheet_row_number)

    if not rows:
        print(f"No data rows found in spreadsheet '{spreadsheet_id}', tab '{tab_name}'")
        return pd.DataFrame()

    df = pd.DataFrame(rows, columns=header)

    for col in df.columns:
        df[col] = df[col].astype("string").fillna("")

    df["_source_spreadsheet_id"] = normalize_spreadsheet_id(spreadsheet_id)
    df["_source_tab"] = tab_name
    df["_source_row_number"] = row_numbers
    df["_ad_type_raw"] = df[ad_type_column_name].astype("string").fillna("")
    df["_ad_type"] = df["_ad_type_raw"].apply(normalize_ad_type)
    df["_loaded_at_utc"] = datetime.now(timezone.utc)
    df["_image_only_extra_read"] = bool(image_only_extra_read)

    print(
        f"Read {len(df)} rows from spreadsheet '{spreadsheet_id}', tab '{tab_name}'. "
        f"Column F header detected as '{ad_type_column_name}'."
    )
    return df


def get_selected_columns(ad_type: str) -> list[str]:
    selected_columns = SELECTED_COLUMNS_BY_AD_TYPE.get(ad_type, [])
    if not selected_columns:
        raise RuntimeError(f"No selected columns configured for ad type: {ad_type}")
    return [sanitize_column_name(col, col) for col in selected_columns]


def get_hash_columns(ad_type: str) -> list[str]:
    if ad_type == "image":
        return IMAGE_HASH_COLUMNS
    if ad_type == "video":
        return VIDEO_HASH_COLUMNS
    if ad_type == "text":
        return TEXT_HASH_COLUMNS
    return []


def make_row_hash_from_values(values: list[str]) -> str:
    raw = "|".join(str(value or "").strip().lower() for value in values)
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


def make_row_hash(row, hash_columns: list[str]) -> str:
    values = [row.get(col, "") for col in hash_columns]
    return make_row_hash_from_values(values)


def select_columns_for_storage(df: pd.DataFrame, ad_type: str) -> pd.DataFrame:
    selected_columns = get_selected_columns(ad_type)

    if ad_type == "image" and "image_url" not in df.columns:
        raise RuntimeError(
            "Required column 'image_url' was not found for image ads. "
            f"Available columns: {list(df.columns)}"
        )

    metadata_without_hash = [col for col in METADATA_COLUMNS if col != "_row_hash"]
    base_columns = selected_columns + metadata_without_hash
    base_columns = list(dict.fromkeys(base_columns))

    for col in base_columns:
        if col not in df.columns:
            print(f"Warning: column '{col}' not found for ad type '{ad_type}'. Creating blank column.")
            df[col] = ""

    result = df[base_columns].copy()

    for col in selected_columns:
        result[col] = result[col].astype("string").fillna("")

    result["_source_spreadsheet_id"] = result["_source_spreadsheet_id"].astype("string").fillna("")
    result["_source_tab"] = result["_source_tab"].astype("string").fillna("")
    result["_source_row_number"] = pd.to_numeric(
        result["_source_row_number"],
        errors="coerce",
    ).astype("Int64")

    result["_ad_type_raw"] = result["_ad_type_raw"].astype("string").fillna("")
    result["_ad_type"] = result["_ad_type"].astype("string").fillna("")
    result["_loaded_at_utc"] = pd.to_datetime(
        result["_loaded_at_utc"],
        errors="coerce",
        utc=True,
    )

    if ad_type == "image":
        before_blank_url_filter = len(result)
        result["image_url"] = result["image_url"].astype("string").fillna("")
        result = result[result["image_url"].astype(str).str.strip() != ""].copy()
        after_blank_url_filter = len(result)
        print(
            f"Removed {before_blank_url_filter - after_blank_url_filter} image rows "
            "because image_url is blank."
        )

    hash_columns = get_hash_columns(ad_type)
    result["_row_hash"] = result.apply(lambda row: make_row_hash(row, hash_columns), axis=1)
    result["_row_hash"] = result["_row_hash"].astype("string").fillna("")

    final_columns = selected_columns + METADATA_COLUMNS
    final_columns = list(dict.fromkeys(final_columns))
    return result[final_columns].copy()


def get_existing_duplicate_keys_from_supabase(
    supabase_client: Client,
    table_name: str,
    ad_type: str,
) -> set:
    if ad_type == "image":
        select_cols = "image_url"
        key_columns = ["image_url"]
    elif ad_type == "video":
        select_cols = "advertiser,_ad_type_raw"
        key_columns = ["advertiser", "_ad_type_raw"]
    else:
        select_cols = "_row_hash"
        key_columns = ["_row_hash"]

    existing_keys = set()
    page_size = 1000
    start = 0

    while True:
        try:
            response = (
                supabase_client.table(table_name)
                .select(select_cols)
                .range(start, start + page_size - 1)
                .execute()
            )
            # Handle different versions of the supabase-py response structure safely
            rows = getattr(response, "data", None)
            if rows is None and isinstance(response, dict):
                rows = response.get("data", [])
            
            if not rows:
                break

            for row in rows:
                values = [clean_key_value(row.get(col, "")) for col in key_columns]
                if any(values):
                    existing_keys.add(tuple(values))

            if len(rows) < page_size:
                break
            start += page_size
        except Exception as e:
            print(f"Error reading existing keys from Supabase table {table_name}: {e}")
            break

    print(f"Found {len(existing_keys)} existing duplicate keys in Supabase {table_name}")
    return existing_keys


def make_current_duplicate_key(row, ad_type: str):
    if ad_type == "image":
        return (clean_key_value(row.get("image_url", "")),)
    if ad_type == "video":
        return (
            clean_key_value(row.get("advertiser", "")),
            clean_key_value(row.get("_ad_type_raw", "")),
        )
    return (clean_key_value(row.get("_row_hash", "")),)


def is_blank_duplicate_key(key) -> bool:
    try:
        return not any(str(value or "").strip() for value in key)
    except Exception:
        return True


def load_dataframe_to_supabase(
    supabase_client: Client,
    df: pd.DataFrame,
    table_name: str,
    ad_type: str,
):
    if df.empty:
        print(f"No rows to load for {table_name}")
        return

    # Build normalized duplicate key for current batch
    df["_dedupe_key"] = df.apply(lambda row: make_current_duplicate_key(row, ad_type), axis=1)

    before_blank_key_filter = len(df)
    df = df[~df["_dedupe_key"].apply(is_blank_duplicate_key)].copy()
    after_blank_key_filter = len(df)
    print(
        f"Removed {before_blank_key_filter - after_blank_key_filter} rows "
        f"with blank duplicate key for {table_name}."
    )

    before_batch_dedupe = len(df)
    df = df.drop_duplicates(subset=["_dedupe_key"]).copy()
    after_batch_dedupe = len(df)
    print(
        f"Removed {before_batch_dedupe - after_batch_dedupe} duplicate rows "
        f"inside current batch for {table_name}."
    )

    # Check against existing rows stored in Supabase
    existing_keys = get_existing_duplicate_keys_from_supabase(
        supabase_client=supabase_client,
        table_name=table_name,
        ad_type=ad_type,
    )

    if existing_keys:
        before_existing_filter = len(df)
        df = df[~df["_dedupe_key"].isin(existing_keys)].copy()
        after_existing_filter = len(df)
        print(
            f"Skipped {before_existing_filter - after_existing_filter} rows "
            f"already existing in {table_name}."
        )
    else:
        print(f"No existing rows found in {table_name}. All current unique rows are new.")

    df = df.drop(columns=["_dedupe_key"], errors="ignore")

    if df.empty:
        print(f"No new rows to append into {table_name}")
        return

    df_copy = df.copy()
    if "_loaded_at_utc" in df_copy.columns:
        df_copy["_loaded_at_utc"] = df_copy["_loaded_at_utc"].astype(str)

    # Map pandas NaN/NaT to None for proper SQL NULL values
    df_copy = df_copy.where(pd.notnull(df_copy), None)
    records = df_copy.to_dict(orient="records")

    chunk_size = 500
    total_inserted = 0

    for i in range(0, len(records), chunk_size):
        chunk = records[i: i + chunk_size]
        try:
            supabase_client.table(table_name).insert(chunk).execute()
            total_inserted += len(chunk)
        except Exception as e:
            print(f"Error inserting batch into Supabase {table_name}: {e}")

    print(f"Appended {total_inserted} new rows into Supabase {table_name}")


def main():
    supabase_url = required_env("SUPABASE_URL")
    supabase_key = required_env("SUPABASE_KEY")
    supabase_client: Client = create_client(supabase_url, supabase_key)
    print("Supabase connection initialized.")

    spreadsheet_ids = get_spreadsheet_ids()
    image_ads_source_spreadsheet_id = normalize_spreadsheet_id(
        required_env("IMAGE_ADS_SOURCE_SPREADSHEET_ID")
    )

    print("Loaded spreadsheet IDs from SPREADSHEET_IDS:")
    for sid in spreadsheet_ids:
        print(f"- {sid[:6]}...{sid[-6:]} length={len(sid)}")

    print(
        "Image source spreadsheet ID:",
        f"{image_ads_source_spreadsheet_id[:6]}...{image_ads_source_spreadsheet_id[-6:]}",
        f"length={len(image_ads_source_spreadsheet_id)}",
    )

    credentials, _ = google.auth.default(scopes=SCOPES)
    sheets_service = build(
        "sheets",
        "v4",
        credentials=credentials,
        cache_discovery=False,
    )

    frames = []

    for spreadsheet_id in spreadsheet_ids:
        print(f"Reading spreadsheet file: {spreadsheet_id}")
        sheet_names = get_all_sheet_names(sheets_service=sheets_service, spreadsheet_id=spreadsheet_id)

        print(f"Found {len(sheet_names)} internal sheet/tab(s) in spreadsheet {spreadsheet_id}:")
        for sheet_name in sheet_names:
            print(f"- {repr(sheet_name)}")

        for sheet_name in sheet_names:
            df = read_sheet_tab(
                sheets_service=sheets_service,
                spreadsheet_id=spreadsheet_id,
                tab_name=sheet_name,
                image_only_extra_read=False,
            )
            if not df.empty:
                frames.append(df)

    if image_ads_source_spreadsheet_id not in spreadsheet_ids:
        print(
            "IMAGE_ADS_SOURCE_SPREADSHEET_ID is not inside SPREADSHEET_IDS. "
            "Reading Clean Data 2 separately for image_ads only."
        )
        actual_image_tab_name = find_matching_tab_name(
            sheets_service=sheets_service,
            spreadsheet_id=image_ads_source_spreadsheet_id,
            expected_tab_name=IMAGE_ADS_SOURCE_TAB,
        )
        image_extra_df = read_sheet_tab(
            sheets_service=sheets_service,
            spreadsheet_id=image_ads_source_spreadsheet_id,
            tab_name=actual_image_tab_name,
            image_only_extra_read=True,
        )
        if not image_extra_df.empty:
            frames.append(image_extra_df)
    else:
        print(
            "IMAGE_ADS_SOURCE_SPREADSHEET_ID is already inside SPREADSHEET_IDS. "
            "Image rows will be filtered from that file only."
        )

    if not frames:
        raise RuntimeError("No data found in any spreadsheet file.")

    all_ads_df = pd.concat(frames, ignore_index=True, sort=False)
    print(f"Total rows read from all source tabs: {len(all_ads_df)}")

    print("Detected ad type counts before skip filter:")
    print(all_ads_df["_ad_type"].value_counts(dropna=False))

    skip_rows = all_ads_df[all_ads_df["_ad_type"] == "skip"]
    if not skip_rows.empty:
        print(f"Skipping {len(skip_rows)} rows where column F is N/A/blank/error/invalid video ID.")

    all_ads_df = all_ads_df[all_ads_df["_ad_type"].isin(["image", "video", "text"])].copy()

    error_like_values = {
        "error",
        "failed",
        "failure",
        "keyboard_arrow_right",
    }

    normal_cols = [col for col in all_ads_df.columns if not col.startswith("_")]
    if normal_cols:
        error_mask = pd.Series(False, index=all_ads_df.index)
        for col in normal_cols:
            error_mask = error_mask | all_ads_df[col].astype(str).str.strip().str.lower().isin(error_like_values)

        bad_count = int(error_mask.sum())
        if bad_count:
            print(f"Removing {bad_count} scraper error rows before Supabase load.")
        all_ads_df = all_ads_df[~error_mask].copy()

    for ad_type, table_name in TABLE_BY_AD_TYPE.items():
        filtered_df = all_ads_df[all_ads_df["_ad_type"] == ad_type].copy()

        if ad_type in {"video", "text"} and "_image_only_extra_read" in filtered_df.columns:
            before_extra_filter = len(filtered_df)
            filtered_df = filtered_df[
                ~filtered_df["_image_only_extra_read"].fillna(False).astype(bool)
            ].copy()
            after_extra_filter = len(filtered_df)

            if before_extra_filter != after_extra_filter:
                print(
                    f"Removed {before_extra_filter - after_extra_filter} image-only extra-read rows "
                    f"from {table_name}."
                )

        if ad_type == "video":
            before_video_valid_filter = len(filtered_df)
            filtered_df = filtered_df[filtered_df["_ad_type_raw"].apply(is_valid_video_id)].copy()
            after_video_valid_filter = len(filtered_df)
            print(
                f"Video valid ID filter removed "
                f"{before_video_valid_filter - after_video_valid_filter} rows."
            )

        if ad_type == "image":
            before_image_source_filter = len(filtered_df)
            filtered_df["_source_spreadsheet_id_norm"] = filtered_df["_source_spreadsheet_id"].apply(
                normalize_spreadsheet_id
            )
            filtered_df["_source_tab_norm"] = filtered_df["_source_tab"].apply(
                normalize_compare_text
            )

            filtered_df = filtered_df[
                (filtered_df["_source_spreadsheet_id_norm"] == image_ads_source_spreadsheet_id)
                & (filtered_df["_source_tab_norm"] == normalize_compare_text(IMAGE_ADS_SOURCE_TAB))
            ].copy()

            filtered_df = filtered_df.drop(
                columns=["_source_spreadsheet_id_norm", "_source_tab_norm"],
                errors="ignore",
            )
            after_image_source_filter = len(filtered_df)
            print(
                f"Image source/tab filter removed "
                f"{before_image_source_filter - after_image_source_filter} rows."
            )

        print(
            f"Preparing {table_name}: "
            f"{len(filtered_df)} rows from "
            f"{filtered_df['_source_spreadsheet_id'].nunique() if not filtered_df.empty else 0} spreadsheet file(s), "
            f"{filtered_df['_source_tab'].nunique() if not filtered_df.empty else 0} tab(s)"
        )

        if ad_type == "image" and not filtered_df.empty:
            print("Image rows by source tab:")
            print(filtered_df["_source_tab"].value_counts(dropna=False))
            print("Image URL count before selection:")
            print(filtered_df["image_url"].astype(str).str.strip().ne("").sum() if "image_url" in filtered_df.columns else 0)

        filtered_df = select_columns_for_storage(df=filtered_df, ad_type=ad_type)

        load_dataframe_to_supabase(
            supabase_client=supabase_client,
            df=filtered_df,
            table_name=table_name,
            ad_type=ad_type,
        )


if __name__ == "__main__":
    main()
