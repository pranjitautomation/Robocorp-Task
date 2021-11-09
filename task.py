import os
import time as t


from Browser import Browser
from Browser.utils.data_types import SelectAttribute
from RPA.Excel.Files import Files
from RPA.FileSystem  import FileSystem
from RPA.HTTP import HTTP
from RPA.PDF import PDF



browser=Browser()


def open_website():
    browser.open_browser("https://robotsparebinindustries.com/")

def log_in():
    browser.type_text("css=#username", "maria")
    browser.type_secret("css=#password", "thoushallnotpass")
    browser.click("text=Log in")


def download_excel():
    http=HTTP()
    http.download(
        url="https://robotsparebinindustries.com/SalesData.xlsx",
        overwrite=True
    )



def fill_for_one(rep):
    browser.type_text("css=#firstname", rep["First Name"])
    browser.type_text("css=#lastname", rep["Last Name"])
    browser.type_text("css=#salesresult", str(rep["Sales"]))
    browser.select_options_by("css=#salestarget", SelectAttribute["value"], str(rep["Sales Target"]))
    browser.click("text=Submit")
    t.sleep(3)


def fill_all_excel():
    excel=Files()
    excel.open_workbook("SalesData.xlsx")
    sales=excel.read_worksheet_as_table(header=True)
    excel.close_workbook()

    for rep in sales:
        fill_for_one(rep)


def export_the_pdf():
    pran_html=browser.get_property(
        selector="css=#sales-results" ,property="outerHTML"
    )


    pdf=PDF()
    
    pdf.html_to_pdf(pran_html,"output/pranjit.results.pdf")


def log_out():
    browser.click("text=Log out")

def main():
    try:
        open_website()
        t.sleep(10)
        
        log_in()
        t.sleep(3)
        
        download_excel()
        t.sleep(5)

        fill_all_excel()

        export_the_pdf()

    finally:
        log_out()
        browser.playwright.close()



if __name__ == "__main__":
    main()