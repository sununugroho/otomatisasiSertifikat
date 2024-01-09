import openpyxl
from docxtpl import DocxTemplate

excel="sertifikat_terbaik.xlsx"
load=openpyxl.load_workbook(excel)
sheet=load.active

get_value=list(sheet.values)
print(get_value)

doc=DocxTemplate("Sertifikat_Terbaik.docx")
for value_tuple in get_value[1:]:
    doc.render({ "JUDUL": value_tuple[0],
                 "TAHUN": value_tuple[1],
                 "NAMA": value_tuple[2], 
                 "TANGGAL": value_tuple[3], 
                 "PENYELENGGARA": value_tuple[4]  
    })

    doc.name="Sertifikat" + value_tuple[0] + value_tuple [2] + ".docx"
    doc.save(doc.name)
