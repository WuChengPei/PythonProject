import pdfplumber
import pandas as pd
from tkinter import Tk, filedialog

# 獲得輸入的頁碼和表格索引

def browse_file():
    Tk().withdraw()
    file_path = filedialog.askopenfilename(
        title="選擇 PDF 檔案",
        filetypes=[("PDF files", "*.pdf")]
    )
    return file_path

def extract_table_from_pdf(pdf_path, page_number, table_index):
    """從指定的 PDF 文件中提取對應頁碼和索引的表格。"""
    with pdfplumber.open(pdf_path) as pdf:
        # 檢查頁碼是否有效
        if page_number > len(pdf.pages) or page_number < 1:
            raise ValueError("無效的頁碼")

        # 獲得頁碼對應的頁面
        page = pdf.pages[page_number - 1]
        tables = page.extract_tables()

        # 檢查表格索引是否有效
        if not tables or table_index > len(tables) or table_index < 1:
            raise ValueError("無效的表格索引")

        # 抓取指定的表格
        specific_table = tables[table_index - 1]

        # 將表格轉換為 DataFrame
        df = pd.DataFrame(specific_table[1:], columns=specific_table[0])
        return df

def main():
    # 瀏覽 PDF 檔案
    pdf_path = browse_file()
    if not pdf_path:
        print("未選擇任何文件")
        return

    # 輸入頁碼和表格索引
    page_number = int(input("請輸入要提取表格的頁碼（從1開始）："))
    table_index = int(input("請輸入要提取表格的索引（從1開始）："))

    # 提取表格並保存為 Excel
    try:
        df = extract_table_from_pdf(pdf_path, page_number, table_index)
        output_path = "output.xlsx"
        df.to_excel(output_path, index=False)
        print(f"表格已成功提取並保存到 {output_path}")
    except Exception as e:
        print(f"出現錯誤：{e}")

if __name__ == "__main__":
    main()

