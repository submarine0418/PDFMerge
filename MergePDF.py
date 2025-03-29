
"""

這個是合併pdf文件的程式，這邊是使用說明:
!!首先請注意"成績單"檔案命名規則 ! 一定要是 學校_姓名(如UM_王大頭)!!
而公文和學生資料裡面的檔案也是要類似的格式 需要->底線後_接姓名 像我是用(合併_王大頭)

1.先將有「公文與學生資料」合併檔案的資料夾的"路徑"放在folder1
2.再來將學生「成績單」的資料夾"路徑"放在folder2
3.將"合併好的公文"資料夾路徑"放在output_folder(請自行建一個放合併檔案的資料夾)

提醒一下：合併完的公文還沒蓋章，要找 Cola 蓋完章再給學生。還有，這程式一次會把資料夾裡所有檔案合併，所以記得檢查一下是誰的檔案！

"""
import os
import re
import PyPDF2
from collections import defaultdict

'''只需要修改這裡資料夾路徑即可，下面的程式不用動'''

folder1 = "C:/Users/User/Downloads/公文與學生資料"      # 「公文與學生資料」的資料夾路徑 (檔名如"合併_王大頭.pdf")
folder2 = "C:/Users/User/Downloads/成績單"              # 「成績單」資料夾 (檔名如"UM_王大頭.pdf")
output_folder = r"C:/Users/User\Downloads/合併"         # 合併後的輸出資料夾  

def extract_name(filename):
    """從檔案名字裡抓出姓名，比如 'UM_王大頭.pdf' 會抓到 '王大頭'"""
    match = re.search(r'_(\S+)\.pdf$', filename)
    return match.group(1) if match else None

def find_matching_pdfs(folder1, folder2):
    """在兩個資料夾裡找名字一樣的 PDF 檔案，準備合併"""
    pdf_dict = {}

    # 先看看成績單資料夾裡有哪些人
    folder2_files = {extract_name(file): file for file in os.listdir(folder2) if file.endswith(".pdf")}

    # 再去公文資料夾找名字一樣的檔案
    for file in os.listdir(folder1):
        if file.endswith(".pdf"):
            name = extract_name(file)
            if name and name in folder2_files:  # 如果公文和成績單都有這個人
                pdf_dict[folder2_files[name]] = [os.path.join(folder1, file), os.path.join(folder2, folder2_files[name])]

    return pdf_dict #回傳一個清單，告訴我們誰的檔案要合併

def merge_pdfs(pdf_dict, output_folder):
    """ 合併 PDF 並儲存到指定的輸出資料夾，檔名以 folder2 為準 """
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    for output_filename, pdf_list in pdf_dict.items():
        output_path = os.path.join(output_folder, output_filename)  # 合併後檔案的名字和位置
        merger = PyPDF2.PdfMerger()

        for pdf in sorted(pdf_list):  # 依照檔名排序，確保順序正確
            merger.append(pdf)

        merger.write(output_path)
        merger.close()
        print(f"已合併 {pdf_list}，輸出：{output_path}") #輸出完成會出現"已合併"

if __name__ == "__main__":
    pdf_dict = find_matching_pdfs(folder1, folder2)
    merge_pdfs(pdf_dict, output_folder)
