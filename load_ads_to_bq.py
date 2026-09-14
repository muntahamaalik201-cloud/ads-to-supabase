import os
import re
import hashlib
from datetime import datetime, timezone

import google.auth
import pandas as pd
from googleapiclient.discovery import build
from google.cloud import bigquery
from google.api_core.exceptions import NotFound
from supabase import create_client, Client


SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/cloud-platform",
]

# Column F = index 5 because Python uses zero-based indexes:
# A=0, B=1, C=2, D=3, E=4, F=5
AD_TYPE_COLUMN_INDEX = 5

# Image ads should only come from this internal Google Sheet tab.
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


# Image duplicate prevention:
# Same image_url will not append again.
IMAGE_HASH_COLUMNS = [
    "image_url",
]

# Video duplicate prevention:
# Same advertiser + same Column F video ID will not append again.
# Column F value is stored as _ad_type_raw.
VIDEO_HASH_COLUMNS = [
    "advertiser",
    "_ad_type_raw",
]

# Text logic unchanged from your source-row based setup.
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

    return [
        sheet["properties"]["title"]
        for sheet in spreadsheet.get("sheets", [])
    ]


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


def read_sheet_tab(sheets_service, spreadsheet_id: str, tab_name: str, image_only_extra_read: bool = False) -> pd.DataFrame:
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

    return [
        sanitize_column_name(col, col)
        for col in selected_columns
    ]


def get_hash_columns(ad_type: str) -> list[str]:
    if ad_type == "image":
        return IMAGE_HASH_COLUMNS
    if ad_type == "video":
        return VIDEO_HASH_COLUMNS
    if ad_type == "text":
        return TEXT_HASH_COLUMNS
    return []


def make_row_hash_from_values(values: list[str]) -> str:
    raw = "|".join(
        str(value or "").strip().lower()
        for value in values
    )
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


def make_row_hash(row, hash_columns: list[str]) -> str:
    values = [row.get(col, "") for col in hash_columns]
    return make_row_hash_from_values(values)


def select_columns_for_bigquery(df: pd.DataFrame, ad_type: str) -> pd.DataFrame:
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
        print(f"Removed {before_blank_url_filter - after_blank_url_filter} image rows because image_url is blank.")

    hash_columns = get_hash_columns(ad_type)
    result["_row_hash"] = result.apply(lambda row: make_row_hash(row, hash_columns), axis=1)
    result["_row_hash"] = result["_row_hash"].astype("string").fillna("")

    final_columns = selected_columns + METADATA_COLUMNS
    final_columns = list(dict.fromkeys(final_columns))

    return result[final_columns].copy()


def get_bigquery_schema(df: pd.DataFrame):
    schema = []
    for col in df.columns:
        if col == "_loaded_at_utc":
            schema.append(bigquery.SchemaField(col, "TIMESTAMP"))
        elif col == "_source_row_number":
            schema.append(bigquery.SchemaField(col, "INTEGER"))
        else:
            schema.append(bigquery.SchemaField(col, "STRING"))
    return schema


def ensure_table_exists_and_schema(bq_client: bigquery.Client, table_id: str, schema):
    try:
        table = bq_client.get_table(table_id)
        existing_field_names = {field.name for field in table.schema}
        missing_fields = [
            field for field in schema
            if field.name not in existing_field_names
        ]
        if missing_fields:
            table.schema = list(table.schema) + missing_fields
            bq_client.update_table(table, ["schema"])
            print(f"Updated schema for {table_id}. Added columns: {[field.name for field in missing_fields]}")
        else:
            print(f"Table already exists with correct schema: {table_id}")
    except NotFound:
        table = bigquery.Table(table_id, schema=schema)
        bq_client.create_table(table)
        print(f"Created table: {table_id}")


def get_existing_duplicate_keys(bq_client: bigquery.Client, table_id: str, ad_type: str) -> set:
    try:
        table = bq_client.get_table(table_id)
    except NotFound:
        return set()

    existing_schema_by_name = {field.name: field for field in table.schema}

    if ad_type == "image":
        key_columns = ["image_url"]
    elif ad_type == "video":
        key_columns = ["advertiser", "_ad_type_raw"]
    else:
        key_columns = ["_row_hash"]

    selected_field_names = [col for col in key_columns if col in existing_schema_by_name]

    if len(selected_field_names) != len(key_columns):
        print(f"Could not find all duplicate key columns in {table_id}. Needed: {key_columns}, Found: {selected_field_names}")
        return set()

    selected_fields = [existing_schema_by_name[col] for col in selected_field_names]
    existing_keys = set()
    rows = bq_client.list_rows(table, selected_fields=selected_fields, page_size=10000)

    for row in rows:
        values = []
        for col in key_columns:
            try:
                value = row[col]
            except Exception:
                value = ""
            values.append(clean_key_value(value))
            
        if not any(values):
            continue
            
        existing_keys.add(tuple(values))

    print(f"Found {len(existing_keys)} existing duplicate keys in {table_id}")
    return existing_keys


def make_current_duplicate_key(row, ad_type: str):
    if ad_type == "image":
        return (clean_key_value(row.get("image_url", "")),)
    if ad_type == "video":
        return (clean_key_value(row.get("advertiser", "")), clean_key_value(row.get("_ad_type_raw", "")),)
    return (clean_key_value(row.get("_row_hash", "")),)


def is_blank_duplicate_key(key) -> bool:
    try:
        return not any(str(value or "").strip() for value in key)
    except Exception:
        return True


def process_deduplication(bq_client: bigquery.Client, df: pd.DataFrame, project_id: str, dataset_id: str, table_name: str, ad_type: str) -> pd.DataFrame:
    """Handles the deduplication logic against BigQuery and returns the filtered DataFrame containing only new rows."""
    table_id = f"{project_id}.{dataset_id}.{table_name}"
    
    if df.empty:
        return df

    df["_dedupe_key"] = df.apply(lambda row: make_current_duplicate_key(row, ad_type), axis=1)

    before_blank_key_filter = len(df)
    df = df[~df["_dedupe_key"].apply(is_blank_duplicate_key)].copy()
    print(f"Removed {before_blank_key_filter - len(df)} rows with blank duplicate key for {table_id}.")

    before_batch_dedupe = len(df)
    df = df.drop_duplicates(subset=["_dedupe_key"]).copy()
    print(f"Removed {before_batch_dedupe - len(df)} duplicate rows inside current batch for {table_id}.")

    existing_keys = get_existing_duplicate_keys(bq_client=bq_client, table_id=table_id, ad_type=ad_type)

    if existing_keys:
        before_existing_filter = len(df)
        df = df[~df["_dedupe_key"].isin(existing_keys)].copy()
        print(f"Skipped {before_existing_filter - len(df)} rows already existing in {table_id}.")
    else:
        print(f"No existing rows found in {table_id}. All current unique rows are new.")

    df = df.drop(columns=["_dedupe_key"], errors="ignore")
    return df


def load_dataframe_to_bigquery(bq_client: bigquery.Client, df: pd.DataFrame, project_id: str, dataset_id: str, table_name: str, schema):
    table_id = f"{project_id}.{dataset_id}.{table_name}"

    ensure_table_exists_and_schema(bq_client=bq_client, table_id=table_id, schema=schema)

    if df.empty:
        print(f"No new rows to append into BigQuery table {table_id}")
        return

    job_config = bigquery.LoadJobConfig(
        schema=schema,
        write_disposition=bigquery.WriteDisposition.WRITE_APPEND,
        create_disposition=bigquery.CreateDisposition.CREATE_IF_NEEDED,
    )

    job = bq_client.load_table_from_dataframe(df, table_id, job_config=job_config)
    job.result()
    print(f"Appended {len(df)} new rows into BigQuery {table_id}")


def load_dataframe_to_supabase(supabase_client: Client, df: pd.DataFrame, table_name: str):
    """Pushes a dataframe directly to Supabase via its REST API."""
    if not supabase_client:
        return
        
    if df.empty:
        print(f"No new rows to append into Supabase table {table_name}")
        return

    # Create a copy so we don't mutate the original DataFrame
    df_copy = df.copy()

    # Convert pandas Timestamp objects to strings so JSON can read them correctly
    if "_loaded_at_utc" in df_copy.columns:
        df_copy["_loaded_at_utc"] = df_copy["_loaded_at_utc"].astype(str)

    # Convert pandas NaN/NaT to python None (which translates to NULL in Supabase)
    df_copy = df_copy.where(pd.notnull(df_copy), None)

    records = df_copy.to_dict(orient="records")

    # Supabase handles large uploads best when split into chunks
    chunk_size = 500
    total_inserted = 0

    for i in range(0, len(records), chunk_size):
        chunk = records[i:i + chunk_size]
        try:
            supabase_client.table(table_name).insert(chunk).execute()
            total_inserted += len(chunk)
        except Exception as e:
            print(f"Error inserting chunk into Supabase {table_name}: {e}")

    print(f"Appended {total_inserted} new rows into Supabase {table_name}")


def main():
    project_id = required_env("GCP_PROJECT_ID")
    dataset_id = required_env("BQ_DATASET")

    # Initialize Supabase client
    supabase_url = os.getenv("SUPABASE_URL")
    supabase_key = os.getenv("SUPABASE_KEY")
    supabase_client = None
    if supabase_url and supabase_key:
        supabase_client = create_client(supabase_url, supabase_key)
        print("Supabase connection initialized.")
    else:
        print("Warning: SUPABASE_URL or SUPABASE_KEY is missing. Supabase upload will be skipped.")

    spreadsheet_ids = get_spreadsheet_ids()

    image_ads_source_spreadsheet_id = normalize_spreadsheet_id(
        required_env("IMAGE_ADS_SOURCE_SPREADSHEET_ID")
    )

    credentials, _ = google.auth.default(scopes=SCOPES)

    sheets_service = build(
        "sheets",
        "v4",
        credentials=credentials,
        cache_discovery=False,
    )

    bq_client = bigquery.Client(
        project=project_id,
        credentials=credentials,
    )

    frames = []

    # Read normal spreadsheet IDs for all tables.
    for spreadsheet_id in spreadsheet_ids:
        print(f"Reading spreadsheet file: {spreadsheet_id}")
        sheet_names = get_all_sheet_names(sheets_service=sheets_service, spreadsheet_id=spreadsheet_id)

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

    if not frames:
        raise RuntimeError("No data found in any spreadsheet file.")

    all_ads_df = pd.concat(frames, ignore_index=True, sort=False)

    skip_rows = all_ads_df[all_ads_df["_ad_type"] == "skip"]
    if not skip_rows.empty:
        print(f"Skipping {len(skip_rows)} rows where column F is N/A/blank/error/invalid.")

    all_ads_df = all_ads_df[all_ads_df["_ad_type"].isin(["image", "video", "text"])].copy()

    error_like_values = {"error", "failed", "failure", "keyboard_arrow_right"}
    normal_cols = [col for col in all_ads_df.columns if not col.startswith("_")]

    if normal_cols:
        error_mask = pd.Series(False, index=all_ads_df.index)
        for col in normal_cols:
            error_mask = error_mask | all_ads_df[col].astype(str).str.strip().str.lower().isin(error_like_values)

        bad_count = int(error_mask.sum())
        if bad_count:
            print(f"Removing {bad_count} scraper error rows before load.")
        all_ads_df = all_ads_df[~error_mask].copy()

    # Final routing to BigQuery and Supabase
    for ad_type, table_name in TABLE_BY_AD_TYPE.items():
        filtered_df = all_ads_df[all_ads_df["_ad_type"] == ad_type].copy()

        if ad_type in {"video", "text"} and "_image_only_extra_read" in filtered_df.columns:
            filtered_df = filtered_df[~filtered_df["_image_only_extra_read"].fillna(False).astype(bool)].copy()

        if ad_type == "video":
            filtered_df = filtered_df[filtered_df["_ad_type_raw"].apply(is_valid_video_id)].copy()

        if ad_type == "image":
            filtered_df["_source_spreadsheet_id_norm"] = filtered_df["_source_spreadsheet_id"].apply(normalize_spreadsheet_id)
            filtered_df["_source_tab_norm"] = filtered_df["_source_tab"].apply(normalize_compare_text)
            filtered_df = filtered_df[
                (filtered_df["_source_spreadsheet_id_norm"] == image_ads_source_spreadsheet_id)
                & (filtered_df["_source_tab_norm"] == normalize_compare_text(IMAGE_ADS_SOURCE_TAB))
            ].copy()
            filtered_df = filtered_df.drop(columns=["_source_spreadsheet_id_norm", "_source_tab_norm"], errors="ignore")

        filtered_df = select_columns_for_bigquery(df=filtered_df, ad_type=ad_type)
        schema = get_bigquery_schema(filtered_df)

        # 1. First, check BigQuery to remove rows we've already synced
        new_rows_df = process_deduplication(
            bq_client=bq_client,
            df=filtered_df,
            project_id=project_id,
            dataset_id=dataset_id,
            table_name=table_name,
            ad_type=ad_type,
        )

        # 2. Upload ONLY the newly discovered rows to BigQuery
        load_dataframe_to_bigquery(
            bq_client=bq_client,
            df=new_rows_df,
            project_id=project_id,
            dataset_id=dataset_id,
            table_name=table_name,
            schema=schema,
        )

        # 3. Upload exactly the same new rows to Supabase
        load_dataframe_to_supabase(
            supabase_client=supabase_client,
            df=new_rows_df,
            table_name=table_name,
        )


if __name__ == "__main__":
    main()
