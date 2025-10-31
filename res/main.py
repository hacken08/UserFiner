import pandas as pd
import re
from openpyxl.utils import get_column_letter

# ---------- Utility: Normalize Mobile Numbers ----------
def normalize_mobile(num):
    """Clean and standardize mobile numbers (keep last 10 digits)"""
    if pd.isna(num):
        return None
    num = str(num)
    num = re.sub(r'\D', '', num)  # remove all non-digits
    num = re.sub(r'^(91|0|91)?', '', num)  # remove leading +91, 91, or 0
    if len(num) > 10:
        num = num[-10:]  # keep last 10 digits
    return num if len(num) == 10 else None


# ---------- Utility: Detect Possible Phone Columns ----------
def get_phone_columns(df):
    """Find columns that likely contain phone numbers"""
    keywords = ['phone', 'mobile', 'ph', 'contact', 'number']
    cols = [col for col in df.columns if any(k in col.lower() for k in keywords)]
    return cols


# ---------- Load Files ----------
user_file = input(' Give user file(.csv): ')
elec_file = input(" Give electercian excel(.xls): ")


user_file = user_file.replace("\"", "")
elec_file = elec_file.replace("\"", "")

try:
    user_df = pd.read_csv(user_file, dtype=str)
except Exception:
    user_df = pd.read_excel(user_file, dtype=str)

# 👇 tell pandas that the actual headers in electrician.xlsx start from 2nd row
elec_df = pd.read_excel(elec_file, dtype=str, header=1)


# ---------- Extract and Normalize Numbers ----------
user_cols = get_phone_columns(user_df)
elec_cols = get_phone_columns(elec_df)


user_numbers = set()
for col in user_cols:
    user_numbers.update(user_df[col].map(normalize_mobile).dropna())

print(f"✅ Found {len(user_numbers)} unique mobile numbers in user file.")

from openpyxl.utils import get_column_letter

# ---------- Build lookup for user.csv numbers with their cell address ----------
user_number_map = {}
for col_name in user_cols:
    col_index = user_df.columns.get_loc(col_name)
    for row_idx, val in enumerate(user_df[col_name]):
        num = normalize_mobile(val)
        if num:
            user_number_map[num] = f"{get_column_letter(col_index + 1)}{row_idx + 2}"

# ---------- Match Numbers in Electrician File ----------
matches = []

print(f"{'Sr.':<5} | {'NAME':<25} | {'MOBILE':<12} | {'ELECTERCIAN':<12} | {'USER EXCEL':<8}")
print("-" * 80)
for row_idx, row in elec_df.iterrows():
    for col_name in elec_cols:
        num = normalize_mobile(row[col_name])
        if num and num in user_number_map:
            elec_cell = f"{get_column_letter(elec_df.columns.get_loc(col_name)+1)}{row_idx+3}"
            user_cell = user_number_map[num]

            matches.append({
                "Sr No": row.get("Sr No", row_idx + 1),
                "Name": row.get("Name", ""),
                "Matched_Number": num,
                "Electrician_Cell": elec_cell,
                "User_Cell": user_cell
            })

            # 💡 Clean, aligned print
            print(f"{row.get('Sr No', row_idx + 1):<5} | {row.get('NAME', ''):<25} | {num:<12} | "
                  f"elec → {elec_cell:<8} | user → {user_cell}")
            print("-" * 80)
            break  # stop after first match in that row

# ---------- Export Results ----------
if matches:
    result_df = pd.DataFrame(matches)
    result_df.to_excel("matched_numbers.xlsx", index=False)
    print(f"\n✅ Matching complete! {len(matches)} matches found.")
    print("📄 Results saved to matched_numbers.xlsx")
else:
    print("\n❌ No matching numbers found.")
