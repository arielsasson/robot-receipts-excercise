from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    browser.configure(
        slowmo=100,
    )
    download_csv_file()
    orders = get_orders()
    open_robot_order_website()
    fill_form_with_csv_data(orders)
    archive_receipts()
    
def download_csv_file():
    """Downloads excel file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

def get_orders():
    lib = Tables()
    return lib.read_table_from_csv("orders.csv")

def open_robot_order_website():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")


def fill_form_with_csv_data(orders):
    """Loads the orders data into the webpage"""
    for row in orders:
        close_annoying_modal()
        fill_and_submit_orders(row)

def close_annoying_modal():
    page = browser.page()
    page.click("button:text('OK')")

def fill_and_submit_orders(order):
    """Fills in the orders data and clicks the ORDER button"""
    page = browser.page()
    page.select_option("#head", order["Head"])
    page.click(f'#id-body-{order["Body"]}')
    page.fill("input[placeholder='Enter the part number for the legs']", order["Legs"])
    page.fill("#address", str(order["Address"]))
    page.click("#order")
    while page.is_visible("div.alert-danger"):
        page.click("#order")
    pdf_file = store_receipt_as_pdf(order["Order number"])
    screenshot = screenshot_robot(order["Order number"])
    embed_screenshot_to_receipt(screenshot, pdf_file)
    page.click("#order-another")

def store_receipt_as_pdf(order_number):
    """Prints receipt and saves it to pdf"""
    page = browser.page()
    order_receipt_html = page.locator("#order-completion").inner_html()
    pdf = PDF()
    filename = f'output/receipts/order-receipt{order_number}.pdf'
    pdf.html_to_pdf(order_receipt_html, filename)
    return filename

def screenshot_robot(order_number):
    """Takes a screenshot of the robot """
    page = browser.page()
    filename = f'output/robots/{order_number}.png'
    robot_image_html = page.locator("#robot-preview-image").screenshot(type='png', path=filename)
    return filename

def embed_screenshot_to_receipt(screenshot, pdf_file):
    """Adds the robot screenshot to the receipt pdf"""
    files_list = [
        pdf_file,
        screenshot
    ]
    pdf = PDF()
    pdf.add_files_to_pdf(files=files_list, target_document=pdf_file)

def archive_receipts():
    """Adds all the generated receipts to a ZIP file"""
    lib = Archive()
    lib.archive_folder_with_zip("output/receipts", "output/receipts.zip")