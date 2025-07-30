import os
import pandas as pd
import re

def hankaku_to_zenkaku_kana(text):
    if pd.isnull(text):
        return text
    half = "ｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜｦﾝｰｧｨｩｪｫｬｭｮｯ｡｢｣､･"
    full = "アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワヲンーァィゥェォャュョッ。「」、・"
    return ''.join(full[half.index(ch)] if ch in half else ch for ch in str(text))

def normalize(text):
    return str(text).strip().lower().replace(" ", "").replace("　", "")

def find_kchbn_blocks(row):
    norm_row = [normalize(cell) for cell in row]
    blocks = []
    for i in range(len(norm_row)):
        if "種類" in norm_row[i]:
            try:
                idx_k = i
                idx_c = next(j for j in range(idx_k + 1, len(norm_row)) if "ch" in norm_row[j])
                idx_b = next(j for j in range(idx_c + 1, len(norm_row)) if "bit" in norm_row[j])
                idx_n = next(j for j in range(idx_b + 1, len(norm_row)) if "名称" in norm_row[j])
                idx_m = next(j for j in range(idx_n + 1, len(norm_row)) if "機器" in norm_row[j])
                blocks.append((idx_k, idx_c, idx_b, idx_n, idx_m))
            except StopIteration:
                continue
    return blocks

def find_io_comment_blocks(row):
    norm_row = [normalize(cell) for cell in row]
    blocks = []
    for i in range(len(norm_row)):
        if "i/o" in norm_row[i]:
            try:
                idx_io = i
                idx_com = next(j for j in range(idx_io + 1, len(norm_row)) if "ｺﾒﾝﾄ" in norm_row[j] or "コメント" in norm_row[j])
                idx_ext = next(j for j in range(idx_com + 1, len(norm_row)) if "抽出用" in norm_row[j])
                blocks.append((idx_io, idx_com, idx_ext))
            except StopIteration:
                continue
    return blocks

def extract_address_number(addr):
    match = re.search(r"(\d+)", addr)
    return int(match.group(1)) if match else float('inf')

def main():
    cwd = os.getcwd()
    excel_file = next((f for f in os.listdir(cwd) if f.endswith(".xlsx") and f.startswith("01_入出力")), None)
    if not excel_file:
        print("❌ Excelファイルが見つかりません。")
        input("Enterキーで終了")
        return

    print(f"▶ 対象Excelファイル：{excel_file}")

    match = re.search(r"01_入出力(.*?)デバイスリスト", excel_file)
    middle = match.group(1).replace(" ", "") if match else ""
    output_name = f"{middle}_ECAD IOアドレス転記表.csv" if middle else "ECAD IOアドレス転記表.csv"
    output_csv = os.path.join(cwd, output_name)
    excel_path = os.path.join(cwd, excel_file)

    print("▶ Excelファイル読み込み中...")
    xls = pd.ExcelFile(excel_path)
    results = []

    for sheet in [s for s in xls.sheet_names if s.startswith("★")]:
        print(f"▶ シート処理中：{sheet}")
        df = pd.read_excel(excel_path, sheet_name=sheet, header=None, dtype=str)
        df.fillna("", inplace=True)

        for row_idx in range(len(df) - 1):
            row = df.iloc[row_idx]

            for (col_k, col_c, col_b, col_n, col_m) in find_kchbn_blocks(row):
                for i in range(row_idx + 1, min(row_idx + 17, len(df))):
                    kind = str(df.iat[i, col_k])
                    ch = str(df.iat[i, col_c])
                    bit = str(df.iat[i, col_b])
                    name = hankaku_to_zenkaku_kana(df.iat[i, col_n])
                    mach = str(df.iat[i, col_m])
                    addr = kind + ch + bit
                    if addr.strip():
                        results.append([name, mach, addr])

            for (col_io, col_com, col_ext) in find_io_comment_blocks(row):
                start_row = row_idx + 1
                if start_row >= len(df): continue
                first_val = str(df.iat[start_row, col_io]).upper()
                if "COM" in first_val:
                    start_row += 1
                for i in range(start_row, min(start_row + 16, len(df))):
                    part1 = str(df.iat[i, col_io])
                    part2 = str(df.iat[i, col_io + 1])
                    addr = part1 + part2
                    name = hankaku_to_zenkaku_kana(df.iat[i, col_com])
                    mach = str(df.iat[i, col_ext])
                    if addr.strip() and not "COM" in addr.upper():
                        results.append([name, mach, addr])

    results.sort(key=lambda x: extract_address_number(x[2]))
    pd.DataFrame(results).to_csv(output_csv, index=False, header=False, encoding="utf-8-sig")

    print("✅ 抽出完了！出力ファイル：", output_csv)
    input("Enterキーで終了")

if __name__ == "__main__":
    main()
