
import re

# Mock GP_MAP and clean_grade
GP_MAP = {
    '10': 'O', '09': 'A+', '9': 'A+', '08': 'A', '8': 'A',
    '07': 'B+', '7': 'B+', '06': 'B', '6': 'B', '05': 'C', '5': 'C',
    '04': 'P', '4': 'P', '00': 'F', '0': 'F'
}

def clean_grade(grade, gp=None):
    if not grade or grade == "NA": return "NA"
    if gp and gp in GP_MAP:
        print(f"      [DEBUG] Using GP '{gp}' -> Grade '{GP_MAP[gp]}'")
        return GP_MAP[gp]
    grade = grade.upper().strip()
    grade = grade.replace("BT", "B+").replace("CT", "C+").replace("A1", "A+")
    if grade == '0': grade = 'O'
    return grade

# Pattern skipping Tot% and Crd, capturing Grade, GP, CP
# Grade: ([A-Z][\w\+]*|0)
# GP: (10|0?[0-9])
# CP: (\d+)
regex_pattern = r"410243.*?(?:\d{2,3}/\d{2,3}\s+){2}(\d{2,3})/100.*?([A-Z][\w\+]*|0)\s+((?:10|0?[0-9]))\s+(\d+)"

print(f"Regex: {regex_pattern}")

test_texts = [
    "410243 BLOCKCHAIN TECHNOLOGY * 020/030 041/070 061/100 -- -- 61 03 A 08 24",
    "410243 BLOCKCHAIN TECHNOLOGY * 020/030 041/070 061/100 -- -- 61 03 At 08 24", # OCR Error At
    "410243 BLOCKCHAIN TECHNOLOGY * 020/030 041/070 061/100 -- -- 61 03 A+ 09 27",
    "410243 BLOCKCHAIN TECHNOLOGY * 020/030 041/070 061/100       61 03 A 08 24", # No dashes
    "410243 BLOCKCHAIN TECHNOLOGY * 020/030 041/070 061/100       61 03 0 10 30"  # Grade O (read as 0)
]

for i, text in enumerate(test_texts):
    print(f"\nTest {i+1}: {text}")
    match = re.search(regex_pattern, text, re.DOTALL | re.IGNORECASE)
    if match:
        groups = match.groups()
        print(f"  Groups: {groups}")
        # Groups: 1=Marks, 2=Grade, 3=GP, 4=CP
        raw_grd = groups[1]
        raw_gp = groups[2]
        final = clean_grade(raw_grd, gp=raw_gp)
        print(f"  Result: '{final}'")
    else:
        print("  NO MATCH")
