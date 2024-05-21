from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive

@task
def order_robots_from_RobotSpareBin():
    """Insert orders into website"""
    browser.configure(slowmo=200)

    open_robot_order_website()
    download_orders()
    fill_form_file_data()
    archive_receipts()

def open_robot_order_website():
    """Navigates to the given url from the website"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")
    page = browser.page()
    page.click("text=OK")
    
def download_orders():
    """download excel sheet with orders"""
    http = HTTP()
    local_file_path = "orders.csv"  # Specify a local file path to save the downloaded file
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True, target_file=local_file_path)
    return local_file_path  # Return the local file path for further use


def fill_and_submit_orders_form(order):
    """Insert order details and submit"""
    page = browser.page()
    
    names_for_heads = {
        "1":"Roll-a-thor head",
        "2":"Peanut crusher head",
        "3":"D.A.V.E head",
        "4":"Andy Roid head",
        "5":"Spanner mate head",
        "6":"Drillbit 2000 head"}

    head_number = order["Head"]
    page.select_option("#head", names_for_heads.get(head_number))
    page.click('//*[@id="root"]/div/div[1]/div/div[1]/form/div[2]/div/div[{0}]/label'.format(order["Body"]))
    page.fill("input[placeholder='Enter the part number for the legs']", order["Legs"])
    page.fill("#address", order["Address"])
    while True: #start the loop
       page.click("#order")
       order_another = page.query_selector("#order-another")
       if order_another:
            pdf_path = store_receipt_as_pdf(int(order["Order number"]))
            screenshot_path = screenshot_robot(int(order["Order number"]))
            embed_screenshot_to_receipt(screenshot_path, pdf_path)
            order_another_bot()
            clicks_ok()
            break

def store_receipt_as_pdf(order_number):
    """Takes screenshot of the ordered bot image"""""
    page = browser.page()
    order_receipt_html = page.locator("#receipt").inner_html()
    pdf = PDF()
    pdf_path = "output/receipts/{0}.pdf".format(order_number)
    pdf.html_to_pdf(order_receipt_html, pdf_path)
    return pdf_path

def order_another_bot():
    """Clicks on button to order a new robot"""
    page = browser.page()
    page.click("#order-another")

def embed_screenshot_to_receipt(screenshot_path, pdf_path):
    """Embeds the screenshot to the bot receipt"""
    pdf = PDF()
    pdf.add_watermark_image_to_pdf(image_path=screenshot_path, 
                                   source_path=pdf_path, 
                                   output_path=pdf_path)

def clicks_ok():
    """Clicks on ok whenever a new order is made for bots"""
    page = browser.page()
    page.click('text=OK')

def collect_results():
    """Take screenshot of receipt"""
    page = browser.page()
    page.screenshot(path="output/sales_summary.png")

def screenshot_robot(order_number):
    """Takes screenshot of the ordered bot image"""
    page = browser.page()
    screenshot_receipt = "output/screenshots/{0}.png".format(order_number)
    page.locator("#robot-preview-image").screenshot(path=screenshot_receipt)
    return screenshot_receipt

def fill_form_file_data():
    """Read data from csv and fill in the robot order form"""
    csv_file = Tables()
    robot_orders = csv_file.read_table_from_csv("orders.csv")
    for order in robot_orders:
        fill_and_submit_orders_form(order)

def archive_receipts():
    """Saves all receipts in a zip folder"""
    save = Archive()
    save.archive_folder_with_zip("./output/receipts", "./output/receipts.zip")


