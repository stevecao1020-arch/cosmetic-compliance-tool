from __future__ import annotations

import html
import io
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable
from xml.etree import ElementTree as ET

import pandas as pd
import requests


FDA_INVENTORY_URL = "https://hfpappexternal.fda.gov/scripts/fdcc/index.cfm?set=ColorAdditives"
FDA_INVENTORY_DOWNLOAD_URL = "https://hfpappexternal.fda.gov/scripts/fdcc/cfc/XMLService.cfm?method=downloadxls&set=ColorAdditives"
FDA_SUMMARY_URL = "https://www.fda.gov/industry/color-additives/summary-color-additives-use-united-states-foods-drugs-cosmetics-and-medical-devices?os=apprefapp"
GOVINFO_XML_TEMPLATE = "https://www.govinfo.gov/content/pkg/CFR-2025-title21-vol1/xml/CFR-2025-title21-vol1-sec{section_key}.xml"
LOCAL_FDA_CSV_PATH = Path(r"C:\Users\Administrator\Downloads\ColorAdditives.csv")

OUTPUT_FILE = "US Color Additives Database.xlsx"


@dataclass
class ColorRecord:
    name: str
    identity_structure: str
    cas_no: str
    uses_and_restrictions: str
    part81_classification: str = ""
    part82_classification: str = ""


def project_root() -> Path:
    return Path(__file__).resolve().parent


def database_dir() -> Path:
    path = project_root() / "database"
    path.mkdir(parents=True, exist_ok=True)
    return path


def normalize_space(text: str) -> str:
    text = html.unescape(text or "")
    text = text.replace("\xa0", " ").replace("\u2009", " ").replace("\u2013", "-").replace("\u2014", "-")
    return re.sub(r"\s+", " ", text).strip()


def parse_csv_text(raw_text: str) -> pd.DataFrame:
    lines = raw_text.splitlines()
    data_text = "\n".join(lines[4:])
    df = pd.read_csv(io.StringIO(data_text))
    df.columns = [normalize_space(str(col)) for col in df.columns]
    return df


def load_inventory_from_local_csv(csv_path: Path) -> pd.DataFrame:
    raw = csv_path.read_bytes()
    for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            return parse_csv_text(raw.decode(encoding))
        except UnicodeDecodeError:
            continue
    return parse_csv_text(raw.decode("latin-1", errors="replace"))


def fetch_fda_inventory() -> pd.DataFrame:
    if LOCAL_FDA_CSV_PATH.exists():
        return load_inventory_from_local_csv(LOCAL_FDA_CSV_PATH)
    session = requests.Session()
    session.get(FDA_INVENTORY_URL, timeout=30)
    raw = session.get(FDA_INVENTORY_DOWNLOAD_URL, timeout=60)
    raw.raise_for_status()
    return parse_csv_text(raw.text)


def extract_sections(row: pd.Series) -> list[str]:
    sections: list[str] = []
    for col in row.index:
        if not str(col).startswith("Regnum"):
            continue
        val = row.get(col)
        if pd.isna(val):
            continue
        match = re.search(r"([0-9]+)\.([0-9A-Za-z]+)", str(val))
        if match:
            sections.append(f"{match.group(1)}.{match.group(2)}")
    return sections


def section_numeric(section: str) -> int | None:
    match = re.match(r"^\d+\.(\d+)", section)
    return int(match.group(1)) if match else None


def classify_part81(sections: list[str]) -> str:
    labels: list[str] = []
    if "81.1" in sections:
        labels.append("81.1 - Provisional lists of color additives")
    if "81.10" in sections:
        labels.append("81.10 - Termination of provisional listings of color additives")
    if "81.30" in sections:
        labels.append("81.30 - Cancellation of certificates")
    return "; ".join(labels)


def classify_part82(sections: list[str]) -> str:
    labels: list[str] = []
    for section in sections:
        if not section.startswith("82."):
            continue
        num = section_numeric(section)
        if num is None:
            continue
        if 3 <= num <= 6:
            label = f"{section} - Part 82 Subpart A (General Provisions)"
        elif 50 <= num <= 706:
            label = f"{section} - Part 82 Subpart B (Foods, Drugs, and Cosmetics)"
        elif 1050 <= num <= 1710:
            label = f"{section} - Part 82 Subpart C (Drugs and Cosmetics)"
        elif 2050 <= num <= 2707:
            label = f"{section} - Part 82 Subpart D (Externally Applied Drugs and Cosmetics)"
        else:
            label = f"{section} - Part 82"
        if label not in labels:
            labels.append(label)
    return "; ".join(labels)


def select_section(row: pd.Series, part: str, min_value: int, max_value: int) -> str | None:
    for section in extract_sections(row):
        if not section.startswith(f"{part}."):
            continue
        num = section_numeric(section)
        if num is not None and min_value <= num <= max_value:
            return section
    return None


def get_text(element: ET.Element) -> str:
    return normalize_space("".join(element.itertext()))


def strip_heading(text: str) -> str:
    return re.sub(r"^\([a-z]\)\s*[A-Za-z /,-]+?\.\s*", "", text).strip()


def fetch_section_xml(section: str) -> ET.Element | None:
    key = section.replace(".", "-")
    url = GOVINFO_XML_TEMPLATE.format(section_key=key)
    response = requests.get(url, timeout=30)
    content_type = response.headers.get("content-type", "")
    if response.status_code != 200 or "xml" not in content_type:
        return None
    return ET.fromstring(response.content)


def extract_identity_and_uses(section: str) -> tuple[str, str, str]:
    root = fetch_section_xml(section)
    if root is None:
        return "", "", ""

    sec = root.find("SECTION")
    if sec is None:
        return "", "", ""

    subject = normalize_space(sec.findtext("SUBJECT", default="")).rstrip(".")
    identity_parts: list[str] = []
    uses_parts: list[str] = []
    mode: str | None = None

    for child in sec:
        if child.tag != "P":
            continue
        text = get_text(child)
        if not text:
            continue
        if re.match(r"^\(a\)", text) and "Identity" in text:
            mode = "identity"
            cleaned = strip_heading(text)
            if cleaned:
                identity_parts.append(cleaned)
            continue
        if re.match(r"^\(b\)", text) and "restriction" in text.lower():
            mode = "uses"
            cleaned = strip_heading(text)
            if cleaned:
                uses_parts.append(cleaned)
            continue
        if re.match(r"^\([c-z]\)", text):
            mode = None
            continue
        if mode == "identity":
            identity_parts.append(text)
        elif mode == "uses":
            uses_parts.append(text)

    return subject, normalize_space(" ".join(identity_parts)), normalize_space(" ".join(uses_parts))


def build_allowed_records(df: pd.DataFrame, part: str, min_value: int, max_value: int) -> list[ColorRecord]:
    records: list[ColorRecord] = []
    use_cosmetics = df["Use"].fillna("").str.contains("Cosmetics", case=False)
    filtered = df[use_cosmetics].copy()

    for _, row in filtered.iterrows():
        sections = extract_sections(row)
        section = select_section(row, part=part, min_value=min_value, max_value=max_value)
        if not section:
            continue
        subject, identity, uses = extract_identity_and_uses(section)
        name = subject or normalize_space(str(row.get("Color", "")))
        cas_no = normalize_space(str(row.get("CAS Reg No or other ID code", ""))).strip()
        records.append(
            ColorRecord(
                name=name,
                identity_structure=identity,
                cas_no="" if cas_no.lower() == "nan" else cas_no,
                uses_and_restrictions=uses or normalize_space(str(row.get("RESTRICTIONS", ""))),
                part81_classification=classify_part81(sections),
                part82_classification=classify_part82(sections),
            )
        )

    records.sort(key=lambda r: r.name.upper())
    return dedupe_records(records)


def build_non_cosmetic_certified_records(df: pd.DataFrame) -> list[ColorRecord]:
    records: list[ColorRecord] = []
    use_cosmetics = df["Use"].fillna("").str.contains("Cosmetics", case=False)
    status = df["Status"].fillna("")

    filtered = df[(~use_cosmetics) & status.str.contains("certification required", case=False)].copy()
    for _, row in filtered.iterrows():
        sections = extract_sections(row)
        if not any(sec.startswith("74.") for sec in sections):
            continue
        section = None
        for sec in sections:
            if sec.startswith("74."):
                section = sec
                break
        if not section:
            continue
        if select_section(row, part="74", min_value=2000, max_value=2999):
            continue

        subject, identity, _ = extract_identity_and_uses(section)
        name = subject or normalize_space(str(row.get("Color", "")))
        cas_no = normalize_space(str(row.get("CAS Reg No or other ID code", ""))).strip()
        use_text = normalize_space(str(row.get("Use", "")))
        restrictions = normalize_space(str(row.get("RESTRICTIONS", "")))
        if use_text:
            restrictions = f"Current FDA use: {use_text}. {restrictions}".strip()
        restrictions = normalize_space(f"{restrictions} Not listed for current cosmetic use.")
        records.append(
            ColorRecord(
                name=name,
                identity_structure=identity,
                cas_no="" if cas_no.lower() == "nan" else cas_no,
                uses_and_restrictions=restrictions,
                part81_classification=classify_part81(sections),
                part82_classification=classify_part82(sections),
            )
        )

    lead_row = df[df["Color"].fillna("").str.strip().eq("Lead Acetate")]
    if not lead_row.empty:
        row = lead_row.iloc[0]
        lead_identity = ""
        lead_restrictions = "No longer permitted for coloring hair on the scalp. Termination of listing."
        cas_no = normalize_space(str(row.get("CAS Reg No or other ID code", ""))).strip()
        records.append(
            ColorRecord(
                name="Lead acetate",
                identity_structure=lead_identity,
                cas_no="" if cas_no.lower() == "nan" else cas_no,
                uses_and_restrictions=lead_restrictions,
                part81_classification=classify_part81(extract_sections(row)),
                part82_classification=classify_part82(extract_sections(row)),
            )
        )

    records.sort(key=lambda r: r.name.upper())
    return dedupe_records(records)


def build_non_cosmetic_part81_records(df: pd.DataFrame) -> list[ColorRecord]:
    records: list[ColorRecord] = []
    for _, row in df.iterrows():
        sections = extract_sections(row)
        has_part81_negative = "81.10" in sections or "81.30" in sections
        if not has_part81_negative:
            continue

        has_73_cosmetics = select_section(row, part="73", min_value=2000, max_value=2999) is not None
        has_74_cosmetics = select_section(row, part="74", min_value=2000, max_value=2999) is not None
        if has_73_cosmetics or has_74_cosmetics:
            continue

        name = normalize_space(str(row.get("Color", "")))
        cas_no = normalize_space(str(row.get("CAS Reg No or other ID code", ""))).strip()
        restrictions = normalize_space(str(row.get("RESTRICTIONS", "")))
        if not restrictions:
            restrictions = "Part 81 indicates termination of provisional listing and/or cancellation of certificates."
        restrictions = normalize_space(f"{restrictions} Not listed for current cosmetic use.")

        section = None
        for sec in sections:
            if sec.startswith("73.") or sec.startswith("74."):
                section = sec
                break
        identity = ""
        subject = ""
        if section:
            subject, identity, _ = extract_identity_and_uses(section)
        records.append(
            ColorRecord(
                name=subject or name,
                identity_structure=identity,
                cas_no="" if cas_no.lower() == "nan" else cas_no,
                uses_and_restrictions=restrictions,
                part81_classification=classify_part81(sections),
                part82_classification=classify_part82(sections),
            )
        )

    records.sort(key=lambda r: r.name.upper())
    return dedupe_records(records)


def dedupe_records(records: Iterable[ColorRecord]) -> list[ColorRecord]:
    seen: set[tuple[str, str, str, str]] = set()
    deduped: list[ColorRecord] = []
    for rec in records:
        key = (rec.name, rec.identity_structure, rec.cas_no, rec.uses_and_restrictions)
        if key in seen:
            continue
        seen.add(key)
        deduped.append(rec)
    return deduped


def records_to_df(records: list[ColorRecord]) -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "Name": rec.name,
                "Identity / structure": rec.identity_structure,
                "CAS No.": rec.cas_no,
                "Uses and restrictions": rec.uses_and_restrictions,
                "Historical Part 81 references": rec.part81_classification,
                "Part 82 classification": rec.part82_classification,
            }
            for rec in records
        ]
    )


def build_rules_df() -> pd.DataFrame:
    rules = [
        (
            "Primary source",
            f"Local FDA export used as primary input: {LOCAL_FDA_CSV_PATH}" if LOCAL_FDA_CSV_PATH.exists() else "Local FDA export not found; script fell back to FDA live download.",
        ),
        (
            "Rule",
            "Part 73 Subpart C color additives are the current exempt-from-certification cosmetics color additives listed by FDA.",
        ),
        (
            "Rule",
            "Part 74 Subpart C color additives are the current certification-required cosmetics color additives listed by FDA.",
        ),
        (
            "Rule",
            "Any Part 74 certified color additive not currently listed in Subpart C should be treated as not listed for current cosmetic use.",
        ),
        (
            "Rule",
            "Part 73 entries not listed in Subpart C should not be treated as current cosmetics color additives under this database design.",
        ),
        (
            "Rule",
            "Part 81 references are treated as historical status flags. They do not override a current 73/74 cosmetics listing.",
        ),
        (
            "Rule",
            "Part 82 references are treated as supplemental category flags for certified provisionally listed colors and their subparts (B/C/D).",
        ),
        (
            "Rule",
            "A Part 81 termination/cancellation reference is used as a non-cosmetic trigger only when the same ingredient does not also have a current 73/74 cosmetics citation.",
        ),
        (
            "Source",
            f"FDA color additive inventory: {FDA_INVENTORY_URL}",
        ),
        (
            "Source",
            f"FDA summary page: {FDA_SUMMARY_URL}",
        ),
        (
            "Source",
            "Identity and cosmetics use/restriction text were pulled from official govinfo CFR XML for the cited 2025 Title 21 sections when available.",
        ),
        (
            "Note",
            "Lead acetate is included in the non-cosmetic sheet because the FDA summary indicates the listing was repealed and the additive is no longer permitted.",
        ),
    ]
    return pd.DataFrame(rules, columns=["Type", "Value"])


def write_workbook(
    allowed_73: list[ColorRecord],
    allowed_74: list[ColorRecord],
    not_allowed: list[ColorRecord],
    output_path: Path,
) -> Path:
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        records_to_df(allowed_73).to_excel(writer, sheet_name="73_SubpartC_Allowed", index=False)
        records_to_df(allowed_74).to_excel(writer, sheet_name="74_SubpartC_Certified", index=False)
        records_to_df(not_allowed).to_excel(writer, sheet_name="Not_For_Cosmetics", index=False)
        build_rules_df().to_excel(writer, sheet_name="Rules_And_Sources", index=False)
    return output_path


def write_workbook_with_fallback(
    allowed_73: list[ColorRecord],
    allowed_74: list[ColorRecord],
    not_allowed: list[ColorRecord],
    output_path: Path,
) -> Path:
    try:
        return write_workbook(allowed_73, allowed_74, not_allowed, output_path)
    except PermissionError:
        alt_path = output_path.with_name(
            f"{output_path.stem}_{datetime.now().strftime('%Y%m%d_%H%M%S')}{output_path.suffix}"
        )
        return write_workbook(allowed_73, allowed_74, not_allowed, alt_path)


def main() -> None:
    df = fetch_fda_inventory()
    allowed_73 = build_allowed_records(df, part="73", min_value=2000, max_value=2999)
    allowed_74 = build_allowed_records(df, part="74", min_value=2000, max_value=2999)
    not_allowed = dedupe_records(
        build_non_cosmetic_certified_records(df) + build_non_cosmetic_part81_records(df)
    )

    output_path = database_dir() / OUTPUT_FILE
    written_path = write_workbook_with_fallback(allowed_73, allowed_74, not_allowed, output_path)

    print(f"Output: {written_path}")
    print(f"73_SubpartC_Allowed rows: {len(allowed_73)}")
    print(f"74_SubpartC_Certified rows: {len(allowed_74)}")
    print(f"Not_For_Cosmetics rows: {len(not_allowed)}")


if __name__ == "__main__":
    main()
