import streamlit as st
from openpyxl import Workbook
import io
import re

print("✅ Running latest version")

st.set_page_config(page_title="ISMR to Excel", layout="centered")
st.title("📄 ISMR to Excel Converter")
st.write("This version avoids pandas and handles all inconsistent `.ismr` files safely.")

uploaded_files = st.file_uploader(
    "📂 Choose ISMR files", type=["ismr", "txt"], accept_multiple_files=True
)

use_header = st.checkbox("Use first non-comment row as headers", value=False)

if uploaded_files:
    output = io.BytesIO()
    wb = Workbook()
    wb.remove(wb.active)  # Remove default sheet

    for file in uploaded_files:
        st.markdown(f"---\n### 📂 Processing: `{file.name}`")

        try:
            # Decode and split lines
            content = file.read().decode("utf-8", errors="ignore")
            lines = content.strip().splitlines()
            data_lines = [line.strip() for line in lines if line.strip() and not line.startswith("#")]
            if not data_lines:
                st.warning(f"⚠️ `{file.name}` is empty or contains only comments.")
                continue

            parsed = [line.split(',') for line in data_lines]
            max_len = max(len(row) for row in parsed)
            normalized = [row + [''] * (max_len - len(row)) for row in parsed]

            # Add a new sheet
            sheet_name = re.sub(r'[^A-Za-z0-9]', '_', file.name.rsplit('.', 1)[0])[:31]
            ws = wb.create_sheet(title=sheet_name)

            # Optional header
            if use_header:
                ws.append(normalized[0])  # First row as header
                for row in normalized[1:]:
                    ws.append(row)
            else:
                for row in normalized:
                    ws.append(row)

            st.success(f"✅ Added sheet: `{sheet_name}` with {len(normalized)} rows")

        except Exception as e:
            st.error(f"❌ Error processing `{file.name}`: {e}")

    # Save workbook
    wb.save(output)
    output.seek(0)

    st.download_button(
        label="⬇️ Download Excel File with All Sheets",
        data=output,
        file_name="ismr_merged_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )