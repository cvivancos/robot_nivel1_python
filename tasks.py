from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.PDF import PDF


@task
def minimal_task():
    "Insert the sales data for the week and export it as a PDF."
    # message = message + " World!"
    browser.configure(
        slowmo=100,
    )
    open_the_intranet_website()
    download_excel_file()
    log_in()
    fill_and_submit_sales_form()
    collect_results()



def open_the_intranet_website():
    """Navegar a la URL"""
    browser.goto("https://robotsparebinindustries.com/")


def download_excel_file():
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/SalesData.xlsx", overwrite=True)


def log_in():
    """Rellenar el formulario de inicio de sesión."""
    page=browser.page()
    page.fill("#username", "maria")
    page.fill("#password", "thoushallnotpass")
    page.click("button:text('Log in')")


def fill_and_submit_sales_form(sales_rep):
    """Rellenar los datos de ventas y clicar en Submit"""
    page=browser.page()

    page.fill("#firstname", sales_rep["First Name"])
    page.fill("#lastname", sales_rep["Last Name"])
    page.select_option("#salestarget", str(sales_rep["Sales Target"]))
    page.fill("#salesresult", str(sales_rep["Sales"]))
    page.click("text=Submit")


def fill_form_with_excel_data():
    """Read data from excel and fill in the sales form"""
    excel = Files()
    excel.open_workbook("SalesData.xlsx")
    worksheet = excel.read_worksheet_as_table("data", header=True)
    excel.close_workbook()

    for row in worksheet:
        fill_and_submit_sales_form(row)

def collect_results():
    """Hacer captura de pantalla"""
    page = browser.page()
    page.screenshot(path="output/sales_summary.png")
    

def log_out():
    """Presionar el boton de cerrar sesion"""
    page = browser.page()  
    page.click("text=Log out")


def export_as_pdf():
    """Export the data to a pdf file"""
    page = browser.page()
    sales_results_html = page.locator("#sales-results").inner_html()

    pdf = PDF()
    pdf.html_to_pdf(sales_results_html, "output/sales_results.pdf")