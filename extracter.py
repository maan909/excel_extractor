import os
import re
import xlrd
import pandas as pd
from datetime import datetime

# ---------- HELPERS ----------

def atoi(text):
    return int(text) if text.isdigit() else text.lower()

def natural_keys(text):
    return [atoi(c) for c in re.split(r"(\d+)", text)]

def clean_number(value):
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return value

    s = str(value).strip()
    if s == "":
        return None

    s_clean = re.sub(r"[^\d.\-]", "", s)

    if s_clean in ("", ".", "-", "-.", "-0"):
        return None

    try:
        return int(s_clean) if "." not in s_clean else float(s_clean)
    except:
        return None


# ---------- MAIN FUNCTION ----------

def extract_xls_data(folder_path, output_file="extracted_output.xlsx"):
    extracted_rows = []
    sr_no = 1

    # Scan for XLS files only
    all_files = [os.path.join(folder_path, f)
                 for f in os.listdir(folder_path)
                 if f.lower().endswith(".xls")]

    all_files = sorted(all_files, key=natural_keys)

    if not all_files:
        print("❌ No XLS files found.")
        df = pd.DataFrame(columns=["Sr No", "Bill No", "Date", "Description", "Section", "Amount"])
        df.to_excel(output_file, index=False)
        return

    print("📁 XLS files found:", len(all_files))

    for file_path in all_files:
        print(f"\n📄 Reading: {os.path.basename(file_path)}")

        try:
            wb = xlrd.open_workbook(file_path)
            sheet = wb.sheet_by_index(0)

            def get_cell(r, c):
                try:
                    cell = sheet.cell(r, c)
                    return cell.value, cell.ctype
                except:
                    return None, None

            bill_val, bill_type = get_cell(1, 8)
            date_val, date_type = get_cell(10, 8)
            section_val, section_type = get_cell(17, 1)

            description_list = []
            start_row = 19
            col = 1

            try:
                val = sheet.cell_value(start_row, col)
                if val is None or str(val).strip() == "":
                    start_row = 20
            except:
                start_row = 20

            r = start_row
            while True:
                try:
                    val = sheet.cell_value(r, col)
                except:
                    break

                if val is None or str(val).strip() == "":
                    break

                description_list.append(str(val).strip())
                r += 1

            desc_out = ", ".join(description_list) if description_list else None

            date_out = None
            if date_val not in (None, ""):
                if date_type == xlrd.XL_CELL_DATE:
                    try:
                        dt = xlrd.xldate_as_datetime(date_val, wb.datemode)
                        date_out = dt.date().isoformat()
                    except:
                        date_out = str(date_val)
                else:
                    date_out = str(date_val).strip()

            amt_val, amt_type = get_cell(36, 8)
            amount_out = clean_number(amt_val)

            if amount_out is None:
                for r in range(19, sheet.nrows):
                    num = clean_number(sheet.cell_value(r, 8))
                    if num is not None:
                        amount_out = num
                        break

            bill_out = bill_val if bill_val not in (None, "") else None
            section_out = section_val if section_val not in (None, "") else None

            extracted_rows.append([
                sr_no, bill_out, date_out, desc_out, section_out, amount_out
            ])
            sr_no += 1

        except Exception as exc:
            print("⚠ Error:", exc)

    df = pd.DataFrame(
        extracted_rows,
        columns=["Sr No", "Bill No", "Date", "Description", "Section", "Amount"]
    )

    df.to_excel(output_file, index=False)
    print("\n✔ Extraction completed! Saved to:", output_file)
