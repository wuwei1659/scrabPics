from selenium import webdriver
import sys,os,RdExc,xlsxwriter

tables = RdExc.excel_table_byname()
for row in tables:
    name = row["Item"]
    url  = row["Link"]
    sFile= row["Type"]

    if os.path.exists(sFile+"_"+name+".jpg"):
        print "we have" + sFile+"_"+name+".jpg"
    else:
        try:
            browser=webdriver.Firefox()
            print "Let's go " + url
            browser.get(url)
            print "now we open it"
            browser.save_screenshot(sFile+"_"+name+".jpg")
            print "saving Img"
            browser.quit()
            print "Ok-------------------------------------"
        except Exception:
            print sFile+"_"+name+".jpg"+" exception haddpend ---Fail---Here!!!!!"
