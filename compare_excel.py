"""
Excel Comparison Tool - เครื่องมือเปรียบเทียบไฟล์ Excel
ใช้สำหรับเปรียบเทียบไฟล์ Excel 2 ไฟล์แบบ Cell-by-Cell
รองรับการตรวจสอบค่าและสีพื้นหลังของ Cell
"""

import pandas as pd
import numpy as np
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles import PatternFill


def normalize_value(value):
    """
    ทำความสะอาดค่าโดยลบช่องว่าง (Whitespace) ออก
    """
    if pd.isna(value):
        return None
    if isinstance(value, str):
        return value.strip()
    return value


def get_cell_background_color(cell):
    """
    ดึงสีพื้นหลังของ Cell
    """
    try:
        fill = cell.fill
        if fill.fill_type == "solid" or fill.patternType == "solid":
            fg_color = fill.fgColor
            if fg_color.type == "rgb" and fg_color.rgb:
                return fg_color.rgb
            elif fg_color.type == "indexed":
                return f"indexed_{fg_color.indexed}"
            elif fg_color.type == "theme":
                return f"theme_{fg_color.theme}"
        return "no_fill"
    except:
        return "no_fill"


def color_to_readable(color_code):
    """
    แปลงรหัสสีเป็นข้อความที่อ่านง่าย
    """
    if color_code == "no_fill" or color_code == "00000000":
        return "(ไม่มีสี)"
    if color_code.startswith("FF") and len(color_code) == 8:
        return f"#{color_code[2:]}"  # ตัด FF ออก แสดงเป็น #RRGGBB
    return f"#{color_code}"


def compare_excels(file1_path: str, file2_path: str, output_path: str = "comparison_report.xlsx", check_background: bool = True):
    """
    เปรียบเทียบไฟล์ Excel 2 ไฟล์แบบ Cell-by-Cell
    
    Parameters:
    -----------
    file1_path : str
        พาธของไฟล์ Excel ไฟล์ที่ 1
    file2_path : str
        พาธของไฟล์ Excel ไฟล์ที่ 2
    output_path : str
        พาธของไฟล์รายงานผลลัพธ์ (default: comparison_report.xlsx)
    check_background : bool
        ตรวจสอบสีพื้นหลังด้วยหรือไม่ (default: True)
    
    Returns:
    --------
    list : รายการความแตกต่างที่พบ
    """
    
    print("=" * 60)
    print("🔍 Excel Comparison Tool - เริ่มเปรียบเทียบไฟล์")
    print("=" * 60)
    
    # ตรวจสอบว่าไฟล์มีอยู่หรือไม่
    if not Path(file1_path).exists():
        print(f"❌ ไม่พบไฟล์: {file1_path}")
        return []
    
    if not Path(file2_path).exists():
        print(f"❌ ไม่พบไฟล์: {file2_path}")
        return []
    
    # 1. โหลดไฟล์ Excel
    print(f"\n📂 กำลังโหลดไฟล์...")
    print(f"   ไฟล์ที่ 1: {file1_path}")
    print(f"   ไฟล์ที่ 2: {file2_path}")
    
    try:
        df1 = pd.read_excel(file1_path, engine='openpyxl')
        df2 = pd.read_excel(file2_path, engine='openpyxl')
        
        # โหลด workbook สำหรับตรวจสอบสีพื้นหลัง
        if check_background:
            wb1 = load_workbook(file1_path)
            wb2 = load_workbook(file2_path)
            ws1 = wb1.active
            ws2 = wb2.active
    except Exception as e:
        print(f"❌ เกิดข้อผิดพลาดในการอ่านไฟล์: {e}")
        return []
    
    print(f"\n📊 ข้อมูลไฟล์:")
    print(f"   ไฟล์ที่ 1: {df1.shape[0]} แถว, {df1.shape[1]} คอลัมน์")
    print(f"   ไฟล์ที่ 2: {df2.shape[0]} แถว, {df2.shape[1]} คอลัมน์")
    
    # 2. ตรวจสอบขนาดของไฟล์
    size_mismatch = False
    if df1.shape != df2.shape:
        size_mismatch = True
        print(f"\n⚠️  คำเตือน: ขนาดไฟล์ไม่เท่ากัน!")
        
        if df1.shape[0] != df2.shape[0]:
            print(f"   - จำนวนแถวต่างกัน: ไฟล์ 1 มี {df1.shape[0]} แถว, ไฟล์ 2 มี {df2.shape[0]} แถว")
        
        if df1.shape[1] != df2.shape[1]:
            print(f"   - จำนวนคอลัมน์ต่างกัน: ไฟล์ 1 มี {df1.shape[1]} คอลัมน์, ไฟล์ 2 มี {df2.shape[1]} คอลัมน์")
    else:
        print(f"\n✅ ขนาดไฟล์เท่ากัน")
    
    # 3. เตรียมข้อมูลสำหรับเปรียบเทียบ
    # หาขนาดที่ใช้ในการเปรียบเทียบ (ใช้ขนาดที่เล็กกว่า)
    max_rows = min(df1.shape[0], df2.shape[0])
    max_cols = min(df1.shape[1], df2.shape[1])
    
    # รวมคอลัมน์ทั้งหมดจากทั้ง 2 ไฟล์
    all_columns = list(df1.columns[:max_cols])
    
    # 4. เปรียบเทียบและหาจุดต่าง
    print(f"\n🔄 กำลังเปรียบเทียบข้อมูล...")
    if check_background:
        print(f"   (รวมถึงสีพื้นหลังของ Cell)")
    diff_locations = []
    
    for row in range(max_rows):
        for col in range(max_cols):
            val1 = normalize_value(df1.iloc[row, col])
            val2 = normalize_value(df2.iloc[row, col])
            
            col_name = all_columns[col]
            cell_address = f"{get_excel_column_letter(col + 1)}{row + 2}"  # +2 เพราะ Excel เริ่มที่ 1 และมี Header
            
            # เปรียบเทียบค่า (รองรับ None/NaN)
            values_equal = False
            if val1 is None and val2 is None:
                values_equal = True
            elif val1 is None or val2 is None:
                values_equal = False
            else:
                # เปรียบเทียบค่าปกติ
                try:
                    if isinstance(val1, (int, float)) and isinstance(val2, (int, float)):
                        # สำหรับตัวเลข ใช้การเปรียบเทียบแบบ tolerance
                        values_equal = np.isclose(val1, val2, equal_nan=True)
                    else:
                        values_equal = (str(val1) == str(val2))
                except:
                    values_equal = (val1 == val2)
            
            # ตรวจสอบสีพื้นหลัง
            colors_equal = True
            color1 = None
            color2 = None
            if check_background:
                cell1 = ws1.cell(row=row + 2, column=col + 1)  # +2 เพราะ Excel เริ่มที่ 1 และมี Header
                cell2 = ws2.cell(row=row + 2, column=col + 1)
                color1 = get_cell_background_color(cell1)
                color2 = get_cell_background_color(cell2)
                colors_equal = (color1 == color2)
            
            # บันทึกความแตกต่าง
            if not values_equal and not colors_equal:
                # ทั้งค่าและสีต่างกัน
                diff_locations.append({
                    "Cell": cell_address,
                    "Row": row + 2,
                    "Column": col_name,
                    "Diff_Type": "ค่า + สีพื้นหลัง",
                    "Value_File_1": df1.iloc[row, col] if pd.notna(df1.iloc[row, col]) else "(ว่าง)",
                    "Value_File_2": df2.iloc[row, col] if pd.notna(df2.iloc[row, col]) else "(ว่าง)",
                    "BgColor_File_1": color_to_readable(color1) if color1 else "",
                    "BgColor_File_2": color_to_readable(color2) if color2 else ""
                })
            elif not values_equal:
                # เฉพาะค่าต่างกัน
                diff_locations.append({
                    "Cell": cell_address,
                    "Row": row + 2,
                    "Column": col_name,
                    "Diff_Type": "ค่า",
                    "Value_File_1": df1.iloc[row, col] if pd.notna(df1.iloc[row, col]) else "(ว่าง)",
                    "Value_File_2": df2.iloc[row, col] if pd.notna(df2.iloc[row, col]) else "(ว่าง)",
                    "BgColor_File_1": "",
                    "BgColor_File_2": ""
                })
            elif not colors_equal:
                # เฉพาะสีต่างกัน
                diff_locations.append({
                    "Cell": cell_address,
                    "Row": row + 2,
                    "Column": col_name,
                    "Diff_Type": "สีพื้นหลัง",
                    "Value_File_1": df1.iloc[row, col] if pd.notna(df1.iloc[row, col]) else "(เหมือนกัน)",
                    "Value_File_2": df2.iloc[row, col] if pd.notna(df2.iloc[row, col]) else "(เหมือนกัน)",
                    "BgColor_File_1": color_to_readable(color1) if color1 else "",
                    "BgColor_File_2": color_to_readable(color2) if color2 else ""
                })
    
    # 5. ตรวจสอบแถว/คอลัมน์ที่เกินมา
    if size_mismatch:
        # แถวที่เกินมาในไฟล์ที่ 1
        if df1.shape[0] > df2.shape[0]:
            for row in range(df2.shape[0], df1.shape[0]):
                diff_locations.append({
                    "Cell": f"A{row + 2}",
                    "Row": row + 2,
                    "Column": "(แถวทั้งแถว)",
                    "Diff_Type": "แถวเกิน (ไฟล์ 1)",
                    "Value_File_1": "(มีข้อมูล)",
                    "Value_File_2": "(ไม่มีแถวนี้)",
                    "BgColor_File_1": "",
                    "BgColor_File_2": ""
                })
        
        # แถวที่เกินมาในไฟล์ที่ 2
        if df2.shape[0] > df1.shape[0]:
            for row in range(df1.shape[0], df2.shape[0]):
                diff_locations.append({
                    "Cell": f"A{row + 2}",
                    "Row": row + 2,
                    "Column": "(แถวทั้งแถว)",
                    "Diff_Type": "แถวเกิน (ไฟล์ 2)",
                    "Value_File_1": "(ไม่มีแถวนี้)",
                    "Value_File_2": "(มีข้อมูล)",
                    "BgColor_File_1": "",
                    "BgColor_File_2": ""
                })
        
        # คอลัมน์ที่เกินมาในไฟล์ที่ 1 - ตรวจสอบทุก Cell ในคอลัมน์ที่เกินมา
        if df1.shape[1] > df2.shape[1]:
            extra_cols = list(df1.columns[df2.shape[1]:])
            for col_idx, col_name in enumerate(extra_cols):
                actual_col_idx = df2.shape[1] + col_idx
                col_letter = get_excel_column_letter(actual_col_idx + 1)
                
                for row in range(df1.shape[0]):
                    cell_value = df1.iloc[row, actual_col_idx]
                    # ตรวจสอบเฉพาะ Cell ที่มีค่า (ไม่ว่าง)
                    if pd.notna(cell_value) and str(cell_value).strip() != '':
                        cell_address = f"{col_letter}{row + 2}"
                        
                        # ตรวจสอบสีพื้นหลังด้วย (ถ้าเปิดใช้งาน)
                        bg_color = ""
                        if check_background:
                            cell1 = ws1.cell(row=row + 2, column=actual_col_idx + 1)
                            bg_color = color_to_readable(get_cell_background_color(cell1))
                        
                        diff_locations.append({
                            "Cell": cell_address,
                            "Row": row + 2,
                            "Column": col_name,
                            "Diff_Type": "คอลัมน์เกิน (ไฟล์ 1)",
                            "Value_File_1": cell_value,
                            "Value_File_2": "(ไม่มีคอลัมน์นี้)",
                            "BgColor_File_1": bg_color,
                            "BgColor_File_2": ""
                        })
        
        # คอลัมน์ที่เกินมาในไฟล์ที่ 2 - ตรวจสอบทุก Cell ในคอลัมน์ที่เกินมา
        if df2.shape[1] > df1.shape[1]:
            extra_cols = list(df2.columns[df1.shape[1]:])
            for col_idx, col_name in enumerate(extra_cols):
                actual_col_idx = df1.shape[1] + col_idx
                col_letter = get_excel_column_letter(actual_col_idx + 1)
                
                for row in range(df2.shape[0]):
                    cell_value = df2.iloc[row, actual_col_idx]
                    # ตรวจสอบเฉพาะ Cell ที่มีค่า (ไม่ว่าง)
                    if pd.notna(cell_value) and str(cell_value).strip() != '':
                        cell_address = f"{col_letter}{row + 2}"
                        
                        # ตรวจสอบสีพื้นหลังด้วย (ถ้าเปิดใช้งาน)
                        bg_color = ""
                        if check_background:
                            cell2 = ws2.cell(row=row + 2, column=actual_col_idx + 1)
                            bg_color = color_to_readable(get_cell_background_color(cell2))
                        
                        diff_locations.append({
                            "Cell": cell_address,
                            "Row": row + 2,
                            "Column": col_name,
                            "Diff_Type": "คอลัมน์เกิน (ไฟล์ 2)",
                            "Value_File_1": "(ไม่มีคอลัมน์นี้)",
                            "Value_File_2": cell_value,
                            "BgColor_File_1": "",
                            "BgColor_File_2": bg_color
                        })
    
    # 6. แสดงผลและบันทึกไฟล์
    print("\n" + "=" * 60)
    print("📋 ผลลัพธ์การเปรียบเทียบ")
    print("=" * 60)
    
    if diff_locations:
        # นับจำนวนแต่ละประเภท
        value_diffs = sum(1 for d in diff_locations if d.get('Diff_Type') == 'ค่า')
        color_diffs = sum(1 for d in diff_locations if d.get('Diff_Type') == 'สีพื้นหลัง')
        both_diffs = sum(1 for d in diff_locations if d.get('Diff_Type') == 'ค่า + สีพื้นหลัง')
        row_diffs_f1 = sum(1 for d in diff_locations if d.get('Diff_Type') == 'แถวเกิน (ไฟล์ 1)')
        row_diffs_f2 = sum(1 for d in diff_locations if d.get('Diff_Type') == 'แถวเกิน (ไฟล์ 2)')
        col_diffs_f1 = sum(1 for d in diff_locations if d.get('Diff_Type') == 'คอลัมน์เกิน (ไฟล์ 1)')
        col_diffs_f2 = sum(1 for d in diff_locations if d.get('Diff_Type') == 'คอลัมน์เกิน (ไฟล์ 2)')
        
        print(f"\n❌ พบจุดต่างทั้งหมด {len(diff_locations)} จุด:")
        if value_diffs > 0:
            print(f"   - ค่าต่างกัน: {value_diffs} จุด")
        if color_diffs > 0:
            print(f"   - สีพื้นหลังต่างกัน: {color_diffs} จุด")
        if both_diffs > 0:
            print(f"   - ค่า + สีต่างกัน: {both_diffs} จุด")
        if row_diffs_f1 > 0:
            print(f"   - แถวเกินในไฟล์ 1: {row_diffs_f1} จุด")
        if row_diffs_f2 > 0:
            print(f"   - แถวเกินในไฟล์ 2: {row_diffs_f2} จุด")
        if col_diffs_f1 > 0:
            print(f"   - คอลัมน์เกินในไฟล์ 1: {col_diffs_f1} Cell (มีข้อมูล)")
        if col_diffs_f2 > 0:
            print(f"   - คอลัมน์เกินในไฟล์ 2: {col_diffs_f2} Cell (มีข้อมูล)")
        print()
        
        # แสดงรายละเอียดความแตกต่าง (จำกัดที่ 20 รายการแรก)
        display_count = min(20, len(diff_locations))
        for i, diff in enumerate(diff_locations[:display_count]):
            diff_type = diff.get('Diff_Type', 'ค่า')
            print(f"   {i+1}. Cell {diff['Cell']} (คอลัมน์: {diff['Column']}) [{diff_type}]")
            if diff_type in ['ค่า', 'ค่า + สีพื้นหลัง', 'แถวเกิน']:
                print(f"      ค่า ไฟล์ 1: {diff['Value_File_1']}")
                print(f"      ค่า ไฟล์ 2: {diff['Value_File_2']}")
            if diff_type in ['สีพื้นหลัง', 'ค่า + สีพื้นหลัง']:
                print(f"      สี ไฟล์ 1: {diff.get('BgColor_File_1', '')}")
                print(f"      สี ไฟล์ 2: {diff.get('BgColor_File_2', '')}")
            print()
        
        if len(diff_locations) > display_count:
            print(f"   ... และอีก {len(diff_locations) - display_count} จุด (ดูรายละเอียดทั้งหมดในไฟล์รายงาน)")
        
        # บันทึกรายงานลงไฟล์ Excel
        diff_df = pd.DataFrame(diff_locations)
        diff_df.to_excel(output_path, index=False, engine='openpyxl')
        print(f"\n💾 บันทึกรายงานความต่างไว้ที่: {output_path}")
        
    else:
        print(f"\n✅ ยินดีด้วย! ทั้ง 2 ไฟล์มีข้อมูลเหมือนกันทุกประการ")
    
    print("\n" + "=" * 60)
    print("🏁 เสร็จสิ้นการเปรียบเทียบ")
    print("=" * 60)
    
    return diff_locations


def get_excel_column_letter(col_num: int) -> str:
    """
    แปลงหมายเลขคอลัมน์เป็นตัวอักษร Excel (เช่น 1 -> A, 27 -> AA)
    """
    result = ""
    while col_num > 0:
        col_num, remainder = divmod(col_num - 1, 26)
        result = chr(65 + remainder) + result
    return result


def main():
    """
    ฟังก์ชันหลักสำหรับรันโปรแกรม
    """
    import argparse
    
    parser = argparse.ArgumentParser(
        description='Excel Comparison Tool - เปรียบเทียบไฟล์ Excel 2 ไฟล์',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
ตัวอย่างการใช้งาน:
  python compare_excel.py
  python compare_excel.py -f1 data1.xlsx -f2 data2.xlsx
  python compare_excel.py -f1 data1.xlsx -f2 data2.xlsx -o result.xlsx
  python compare_excel.py --no-color (ไม่ตรวจสอบสีพื้นหลัง)
        """
    )
    
    parser.add_argument('-f1', '--file1', 
                        default='file1.xlsx',
                        help='พาธไฟล์ Excel ไฟล์ที่ 1 (default: file1.xlsx)')
    
    parser.add_argument('-f2', '--file2',
                        default='file2.xlsx', 
                        help='พาธไฟล์ Excel ไฟล์ที่ 2 (default: file2.xlsx)')
    
    parser.add_argument('-o', '--output',
                        default='comparison_report.xlsx',
                        help='พาธไฟล์รายงานผลลัพธ์ (default: comparison_report.xlsx)')
    
    parser.add_argument('--no-color', 
                        action='store_true',
                        help='ไม่ตรวจสอบสีพื้นหลังของ Cell')
    
    args = parser.parse_args()
    
    # รันการเปรียบเทียบ
    compare_excels(args.file1, args.file2, args.output, check_background=not args.no_color)


if __name__ == "__main__":
    main()
