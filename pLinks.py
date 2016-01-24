from selenium import webdriver
import sys,os,RdExc,xlsxwriter

def generate_counter():
    CNT = [0.5]
    def add_one():
        CNT[0] = CNT[0] + 0.5
        return CNT[0]
    return add_one

# Create a new workbook and add a worksheet

counter = generate_counter()
tables = RdExc.excel_table_byname()
workbook = xlsxwriter.Workbook('hyperlink.xlsx')
worksheet = workbook.add_worksheet('Hyperlinks')
for row in tables:
    name = row["Item"]
    url  = row["Link"]
    sFile= row["Type"]
    pLink= sFile+"_"+name+".jpg"
 

    if os.path.exists(sFile+"_"+name+".jpg"):
        try:
            print counter()
            # Write some hyperlinks
            worksheet.write_url("A"+str(counter()), pLink , "", name)
            workbook.save()

        except Exception:
            pass
workbook.close()
