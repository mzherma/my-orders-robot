from robocorp.tasks import task
from robocorp import browser
from pathlib import Path

from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive

order_url="https://robotsparebinindustries.com/#/robot-order"
orders_file="https://robotsparebinindustries.com/orders.csv"
img_file_path = "output/images"
rcpt_file_path = "output/receipts"

Path(img_file_path).mkdir(exist_ok=True)
Path(rcpt_file_path).mkdir(exist_ok=True)

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
        slowmo=3000,
        screenshot="only-on-failure"
    )
    
    open_robot_order_website()
    
    orders = get_orders()

    for row in orders:
        close_annoying_modal()
        fill_the_form(row)
    
    archive_receipts()

def open_robot_order_website():
    """Navigates to the given URL"""
    browser.goto(order_url)

def download_orders_file():
    """Downloads orders file from the given URL"""
    http = HTTP()
    http.download(orders_file, overwrite=True)

def get_orders():
    """Read data from orders file and fill in the orders website"""

    download_orders_file()

    table = Tables()
    orders = table.read_table_from_csv("orders.csv", header=True)

    return orders

def close_annoying_modal():
    page = browser.page()

    ok_btn = '//*[@id="root"]/div/div[2]/div/div/div/div/div/button[1]'
    sure_btn = '//*[@id="root"]/div/div[2]/div/div/div/div/div/button[3]'
    no_btn = '//*[@id="root"]/div/div[2]/div/div/div/div/div/button[4]'

    page.click(ok_btn)

def fill_the_form(orders):
    """Fills in the sales data and clicks the 'Submit' button"""
    page = browser.page()
    retry_count = 0

    head = "#head"
    body = "#id-body-"
    legs = "//div[@id='root']/div[1]/div[1]/div[1]/div[1]/form[1]/div[3]/input[1]"
    address = "#address"
    preview_btn = "#preview"
    order_btn = "#order"
    new_order_btn = "#order-another"

    body = body + orders["Body"]

    page.select_option(head, orders["Head"])
    page.click(body)

    page.fill(legs, orders["Legs"])
    page.fill(address, orders["Address"])
    page.click(preview_btn)

    robot_img_file = screenshot_robot(orders["Order number"])
    # print("robot image file path: " + robot_img_file)
    
    try:
        page.click(order_btn)

        alert_found = page.locator(".alert-danger").is_visible()

        if alert_found and retry_count < 3:
            for retry_count in range(0,3):
                page.click(order_btn)
                retry_count += 1
    except:
        print("Fatal Error Order button not found after 3 retries")

    robot_pdf_file = store_receipt_as_pdf(orders["Order number"])
    embed_screenshot_to_receipt(robot_img_file, robot_pdf_file)
    # print("robot pdf file path: " + robot_pdf_file)

    ready_for_new_order = page.get_by_text("Order another robot").is_visible()
    # print("Ready for New Order")
    if ready_for_new_order:
        page.click(new_order_btn)
    else:
        print("Order Another Robot Button Not Found")


def store_receipt_as_pdf(order_number):
    """Export the data to a pdf file"""
    page = browser.page()

    receipt = '//*[@id="receipt"]'
    order_results_html = page.locator(receipt).inner_html()

    pdf = PDF()
    receipt_file_path = "output/receipts/order_results_" + order_number + ".pdf"

    pdf.html_to_pdf(order_results_html, receipt_file_path)

    return receipt_file_path

    
def screenshot_robot(order_number):
    """Fills in the sales data and clicks the 'Submit' button"""
    page = browser.page()
    download_img = '//*[@id="robot-preview-image"]'

    img_file_path = "output/images/robot_img_" + order_number + ".png"

    robot_img = page.locator(download_img)
    img_file = browser.screenshot(element=robot_img)

    Path(img_file_path).write_bytes(img_file)

    return img_file_path

def embed_screenshot_to_receipt(screenshot, pdf_file):
    """Export the data to a pdf file"""

    pdf = PDF()
    
    # print('Append img file: ' + screenshot + 'to pdf file: ' + pdf_file)
    file_properties = ':align=center'
    img_file_name = screenshot + file_properties

    list_of_files = [ img_file_name, ]
    pdf.add_files_to_pdf(
        files=list_of_files,
        target_document=pdf_file,
        append=True
        )

def archive_receipts():
    lib = Archive()

    lib.archive_folder_with_zip(rcpt_file_path, 'output/receipts.zip', recursive=True)
