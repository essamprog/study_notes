import os
import comtypes.client

def pptx_to_pdf(pptx_path, pdf_path):
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1

    presentation = powerpoint.Presentations.Open(pptx_path)
    presentation.SaveAs(pdf_path, 32)  # 32 = PDF
    presentation.Close()
    powerpoint.Quit()

# مسار مجلد الملفات
source_folder = r"C:\Users\Essam Mohamed\Desktop\Web Task\Sub Term\Web Prog"
# مسار مجلد الحفظ
output_folder = r"C:\Users\Essam Mohamed\Desktop\Web Task\Sub Term\Operating System\Lec\PDF"

# التأكد أن مجلد الحفظ موجود
os.makedirs(output_folder, exist_ok=True)

# تحويل كل ملفات PPTX في مجلد المصدر
for filename in os.listdir(source_folder):
    if filename.endswith(".pptx"):
        pptx_file = os.path.join(source_folder, filename)
        pdf_file = os.path.join(output_folder, filename.replace(".pptx", ".pdf"))

        pptx_to_pdf(pptx_file, pdf_file)
        print(f"{filename} تم تحويله إلى PDF بنجاح.")
