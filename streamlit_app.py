

import pdfplumber
import pandas as pd
import streamlit as st


# ---------- CONFIG / TARGET COLUMNS ----------

TARGET_COLUMNS = [
    "ReferralManager",
    "ReferralEmail",
    "Brand",
    "QuoteNumber",
    "QuoteDate",
    "Company",
    "FirstName",
    "LastName",
    "ContactEmail",
    "ContactPhone",
    "Address",
    "County",
    "City",
    "State",
    "ZipCode",
    "Country",
    "item_id",
    "item_desc",
    "UnitPrice",
    "TotalSales",
    "QuoteValidDate",
    "CustomerNumber",
    "manufacturer_Name",
    "PDF",
    "DemoQuote",
]


# ---------- PDF PARSING HELPERS ----------

def extract_full_text(pdf_bytes: bytes) -> str:
    """Extract text from all pages of the PDF."""
    full_text = ""
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            txt = page.extract_text() or ""
            full_text += txt + "\n"
    return full_text


def extract_header_info(full_text: str) -> Dict[str, Optional[str]]:
    """
    Extract quote-level fields from the PDF header.
    Designed for Cadre Wire quote layout.
    """
    header: Dict[str, Optional[str]] = {}

    # Quote number + date: "Quote 120987 Date 11/24/2025"
    m_quote = re.search(r"Quote\s+(\d+)\s+Date\s+(\d{1,2}/\d{1,2}/\d{4})", full_text)
    if m_quote:
        header["QuoteNumber"] = m_quote.group(1)
        header["QuoteDate"] = m_quote.group(2)

    # Customer number: "Customer 100725"
    m_cust = re.search(r"Customer\s+(\d+)", full_text)
    if m_cust:
        header["CustomerNumber"] = m_cust.group(1)

    # Contact: "Contact Mike Shafer"
    m_contact = re.search(r"Contact\s+([A-Za-z .'-]+)", full_text)
    first_name = last_name = None
    if m_contact:
        name = m_contact.group(1).strip()
        parts = name.split()
        if len(parts) >= 2:
            first_name = parts[0]
            last_name = " ".join(parts[1:])
        elif parts:
            first_name = parts[0]
    header["FirstName"] = first_name
    header["LastName"] = last_name

    # Block between "Quoted For:" and "Quote Good Through"
    if "Quoted For:" in full_text and "Quote Good Through" in full_text:
        start = full_text.index("Quoted For:")
        end = full_text.index("Quote Good Through")
        addr_block = full_text[start:end]

        # Company name between "Quoted For:" and "Ship To:"
        m_company = re.search(r"Quoted For:\s*(.+?)\s+Ship To:", addr_block)
        if m_company:
            header["Company"] = m_company.group(1).strip()

        # Street address: first address before the second one
        m_addr = re.search(
            r"(\d{3,6}\s+[A-Za-z0-9 .]+?)\s+\d{3,6}\s+[A-Za-z0-9 ]+",
            addr_block,
        )
        if m_addr:
            header["Address"] = m_addr.group(1).strip()

        # City, state, zip: e.g. "Akron, OH 44306 Akron, OH 44306"
        m_city = re.search(
            r"([A-Za-z .]+),\s*([A-Z]{2})\s+(\d{5})(?:-\d{4})?",
            addr_block,
        )
        if m_city:
            header["City"] = m_city.group(1).strip()
            header["State"] = m_city.group(2)
            header["ZipCode"] = m_city.group(3)

        # Country
        if "United States of America" in addr_block:
            header["Country"] = "USA"

    # Quote valid date: "Quote Good Through 12/09/2025"
    m_valid = re.search(r"Quote Good Through\s+(\d{1,2}/\d{1,2}/\d{4})", full_text)
    if m_valid:
        header["QuoteValidDate"] = m_valid.group(1)

    return header


def extract_line_items(full_text: str) -> List[Dict[str, str]]:
    """
    Extract all line items from the body.
    Matches lines like:
    '1 HW.MAGFOOT-170 27 EAC 3,600.00000EAC 97,200.00'
    and pulls description text until the next item.
    """
    pattern = re.compile(
        r"(?m)^(\d+)\s+([A-Z0-9.\-]+)\s+(\d+)\s+EAC\s+([\d,]+\.\d+)\s*EAC\s+([\d,]+\.\d{2})"
    )
    matches = list(pattern.finditer(full_text))

    items: List[Dict[str, str]] = []

    for i, m in enumerate(matches):
        line_no = m.group(1)
        item_id = m.group(2)
        qty = m.group(3)
        unit_price = m.group(4)
        total = m.group(5)

        # Description = text between this match and the next match
        start_desc = m.end()
        if i + 1 < len(matches):
            end_desc = matches[i + 1].start()
        else:
            # Stop before summary like "Product 123,390.00"
            stop_word = "Product"
            end_desc = full_text.find(stop_word, start_desc)
            if end_desc == -1:
                end_desc = len(full_text)

        desc_raw = full_text[start_desc:end_desc].strip()
        desc_clean = " ".join(desc_raw.split())  # collapse whitespace

        items.append(
            {
                "line_no": line_no,
                "item_id": item_id,
                "qty": qty,
                "unit_price": unit_price,
                "total": total,
                "description": desc_clean,
            }
        )

    return items


def build_rows_for_pdf(
    pdf_bytes: bytes,
    filename: str,
    referral_manager: Optional[str],
    referral_email: str,
    brand: str,
) -> List[Dict]:
    """
    Parse one PDF and return a list of row dicts following TARGET_COLUMNS.
    """
    full_text = extract_full_text(pdf_bytes)
    header = extract_header_info(full_text)
    items = extract_line_items(full_text)

    rows: List[Dict] = []

    for it in items:
        # Convert numeric strings to floats for Excel
        try:
            unit_price_val = float(it["unit_price"].replace(",", ""))
        except Exception:
            unit_price_val = None

        try:
            total_val = float(it["total"].replace(",", ""))
        except Exception:
            total_val = None

        row = {
            "ReferralManager": referral_manager or None,
            "ReferralEmail": referral_email or None,
            "Brand": brand or None,
            "QuoteNumber": header.get("QuoteNumber"),
            "QuoteDate": header.get("QuoteDate"),
            "Company": header.get("Company"),
            "FirstName": header.get("FirstName"),
            "LastName": header.get("LastName"),
            "ContactEmail": None,
            "ContactPhone": None,
            "Address": header.get("Address"),
            "County": None,
            "City": header.get("City"),
            "State": header.get("State"),
            "ZipCode": header.get("ZipCode"),
            "Country": header.get("Country"),
            "item_id": it["item_id"],
            "item_desc": it["description"],
            "UnitPrice": unit_price_val,
            "TotalSales": total_val,
            "QuoteValidDate": header.get("QuoteValidDate"),
            "CustomerNumber": header.get("CustomerNumber"),
            "manufacturer_Name": None,
            "PDF": filename,
            "DemoQuote": None,
        }
        rows.append(row)

    return rows


# ---------- STREAMLIT APP ----------

st.set_page_config(page_title="Cadre Quote PDF → Excel", layout="wide")

st.title("Cadre Quote PDF → Excel Extractor")

st.markdown(
    """
Upload Cadre Wire quote PDFs and download all line items in a single Excel file.

- Supports up to **100 PDFs** per run.  
- Designed for the **same layout and alignment** as your sample Cadre quote.  
- Each product line item becomes its own row in Excel.
"""
)

with st.sidebar:
    st.header("Defaults / Mapping")
    referral_manager = st.text_input("Referral Manager (optional)", value="")
    referral_email = st.text_input(
        "Referral Email",
        value="plawruk@cadrewire.com",
        help="Used in the ReferralEmail column of the Excel output.",
    )
    brand = st.text_input("Brand", value="Cadre Wire Group")

uploaded_files = st.file_uploader(
    "Upload up to 100 quote PDFs",
    type=["pdf"],
    accept_multiple_files=True,
    help="Files must follow the Cadre quote layout (same format / alignment).",
)

process = st.button("Process PDFs")

if process:
    if not uploaded_files:
        st.error("Please upload at least one PDF.")
    elif len(uploaded_files) > 100:
        st.error("Please upload 100 PDFs or fewer at a time.")
    else:
        all_rows: List[Dict] = []
        progress = st.progress(0.0)
        status = st.empty()

        for idx, f in enumerate(uploaded_files, start=1):
            status.text(f"Processing {idx}/{len(uploaded_files)}: {f.name}")
            pdf_bytes = f.read()
            try:
                rows = build_rows_for_pdf(
                    pdf_bytes=pdf_bytes,
                    filename=f.name,
                    referral_manager=referral_manager,
                    referral_email=referral_email,
                    brand=brand,
                )
                all_rows.extend(rows)
            except Exception as e:
                st.warning(f"Error processing {f.name}: {e}")
            progress.progress(idx / len(uploaded_files))

        if not all_rows:
            st.error("No line items were found in the uploaded PDFs.")
        else:
            df = pd.DataFrame(all_rows, columns=TARGET_COLUMNS)
            st.success(
                f"Parsed {len(uploaded_files)} PDF(s) with {len(df)} total line items."
            )

            st.subheader("Preview (first 50 rows)")
            st.dataframe(df.head(50), use_container_width=True)

            # Excel download
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="Quotes")
            output.seek(0)

            st.download_button(
                label="Download Excel Spreadsheet",
                data=output,
                file_name="quotes_extracted.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
            )
else:
    st.info("Upload PDFs and click **Process PDFs** to start.")
