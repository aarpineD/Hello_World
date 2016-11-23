#Owner:Arpine
#
# Steps to reproduce:
# 1. Create an order Plats
# 2. Go to Attachments tab
# 3. Attach all kind of files
# 4. Click on "Add to Order"
# 5. Click on order Edit icon
#
# Expected result:
# 1. Verify that all attachments are attached
# 2. Remove all attachments
# 3. Make sure that all attachments are removed successfully
#
import os
import time
from CRS_Lib import Actions, Explicit_Wait, Config, Login, Collect_log, Validate, DB_connect, OrderSearch
import shutil                                                               # for copy file
from datetime import datetime
import re
import sys
import autoit
from Required_fileds import Fill_in_required_filds

test_case = os.path.basename(__file__).replace("py","tc")

obj = Actions()
config = Config()
config.parse_config()

pwd = os.getcwd()
if len(sys.argv) > 1:
    log_file = sys.argv[1]
else:
    test_case_dir = os.path.basename(__file__).replace("py","tc")  # Read Script file name and change scriptName.py to scriptName.tc
    if not os.path.exists(test_case_dir):
        os.mkdir(test_case_dir)  # Make test case dir scriptName.tc

    log_file = str(pwd) + "\\" + test_case_dir + "\\run_out.log"

    if os.path.exists(log_file):
        shutil.copy2(log_file, log_file.replace(".log","") + "_copy.log")
        os.remove(log_file)

startTime = datetime.now()
Collect_log.start_time(log_file)
Collect_log.simple_log(log_file, "Test Case: " + os.path.basename(__file__) + "\n" )

obj.driver.maximize_window()
obj.login_to_crs( config.crs_urls[0] )  #0-local, 1-crs.qa, ...

Collect_log.simple_log(log_file,"--------------------- Run Part -------------------")
created_order_number=''
try:
    OIT = "Plats"
    account_name = "4815 - mher.simonyan@volo.global"
    symbol = "4815"
    # NumOfPages = "1"
    # payment_method = "Cash"

    obj.AddNewOrder(log_file)
    obj.Initialize(log_file)
    obj.select_OIT(OIT, log_file)
    # obj.Click_OrderItem_Tab(log_file)

    required_obj = Fill_in_required_filds(obj.driver)
    required_obj.crs_fill("CRS",log_file)
    obj.InputAccountName(account_name, log_file, symbol)

    # Open attachment tab
    obj.open_attachment_tab(log_file)

    # Upload an image
    obj.locate_attchment_upload_button(log_file)
    obj.click_on_attchment_upload_button(log_file)
    filePath = os.getcwd() + "\V3_Denton\\test_suites\Attachments_And_Notes.ts\Add_and_Delete_Attachments.tc\ImageForAttachment.PNG"
    autoit.win_wait_active("Open", 5)
    autoit.send(filePath)
    autoit.send("{ENTER}")
    obj.wait_loader_spiner_to_be_absent(log_file)

    # Verify that image was attached successfully
    if obj.locate_attachment_row(log_file, '1'):
        Collect_log.simple_log(log_file, "Info:\t\tImage  was attached successfully.")
    else:
        Collect_log.simple_log(log_file, "Error:\t\tSelected image doesn't attached.")

    # Upload a document
    obj.wait_loader_spiner_to_be_absent(log_file)
    obj.locate_attchment_upload_button(log_file)
    obj.click_on_attchment_upload_button(log_file)
    filePath = os.getcwd() + "\V3_Denton\\test_suites\Attachments_And_Notes.ts\Add_and_Delete_Attachments.tc\DocumentForAttachment.docx"
    autoit.win_wait_active("Open", 10)
    autoit.send(filePath)
    autoit.send("{ENTER}")
    obj.wait_loader_spiner_to_be_absent(log_file)

    # Verify that document was attached successfully
    if obj.locate_attachment_row(log_file, '2'):
        Collect_log.simple_log(log_file, "Info:\t\tDocument was attached successfully.")
    else:
        Collect_log.simple_log(log_file, "Error:\t\tSelected document doesn't attached.")

    # Upload a pdf file
    obj.wait_loader_spiner_to_be_absent(log_file)
    obj.locate_attchment_upload_button(log_file)
    obj.click_on_attchment_upload_button(log_file)
    filePath = os.getcwd() + "\V3_Denton\\test_suites\Attachments_And_Notes.ts\Add_and_Delete_Attachments.tc\PdfForAttachment.pdf"
    autoit.win_wait_active("Open", 10)
    autoit.send(filePath)
    autoit.send("{ENTER}")
    obj.wait_loader_spiner_to_be_absent(log_file)

    # Verify that PDF was attached successfully
    if obj.locate_attachment_row(log_file, '3'):
        Collect_log.simple_log(log_file, "Info:\t\tPDF file was attached successfully.")
    else:
        Collect_log.simple_log(log_file, "Error:\t\tSelected PDF file doesn't attached.")

    # Upload an excel sheet
    obj.wait_loader_spiner_to_be_absent(log_file)
    obj.locate_attchment_upload_button(log_file)
    obj.click_on_attchment_upload_button(log_file)
    filePath = os.getcwd() + "\V3_Denton\\test_suites\Attachments_And_Notes.ts\Add_and_Delete_Attachments.tc\ExcelForAttachment.xlsx"
    autoit.win_wait_active("Open", 15)
    autoit.send(filePath)
    autoit.send("{ENTER}")
    obj.wait_loader_spiner_to_be_absent(log_file)

    # Verify that Excel sheet was attached successfully
    if obj.locate_attachment_row(log_file, '4'):
        Collect_log.simple_log(log_file, "Info:\t\tExcel sheet was attached successfully.")
    else:
        Collect_log.simple_log(log_file, "Error:\t\tSelected Excel sheet doesn't attached.")

    # Click on Add to Order button
    obj.Click_AddToOrder(log_file)

    # Click on edit order icon and go to attachments tab
    obj.click_on_edit_icon_in_row(log_file)
    obj.open_attachment_tab(log_file)

    # read log_file
    file = open(log_file, 'r')
    string = file.read()
    file.close()

    if re.search("Error:", string, re.I):  # if found any error in run_out.log file:
        Collect_log.run_part_status(log_file, "FAIL", test_case)
        Collect_log.test_case_exit(log_file, startTime, obj.driver)
    else:
        Collect_log.run_part_status(log_file, "PASS\n")
except Exception as e:

    Collect_log.run_part_status(log_file, "FAIL", test_case)
    Collect_log.test_case_exit(log_file, startTime, obj.driver)


Collect_log.simple_log(log_file,"------------------ Validation Part ---------------")

# Make sure that all attachments are available
for i in [1,2,3,4]:
    if obj.locate_attachment_row(log_file, str(i)):
        Collect_log.simple_log(log_file, "Info:\t\tAttachment in row %s exists."%str(i))
    else:
        Collect_log.simple_log(log_file, "Error:\t\tAttachment in row %s doesn't exist."%str(i))

# remove All attachments
obj.remove_all_attachments(log_file)

# Get attachments count
if obj.return_attachments_count():
    Collect_log.simple_log(log_file, "Error:\t\tAttachment count is:" + str(obj.return_attachments_count()))
else:
    Collect_log.simple_log(log_file, "Info:\t\tThere is no attachments.")

#open log_file for read
file = open(log_file, 'r')
string = file.read()
file.close()

#search for error on log_file
if re.search("Error:", string,re.I):    # if found any error in run_out.log file:
    Collect_log.validate_part_status(log_file, "FAIL", test_case)
    Collect_log.test_case_exit(log_file, startTime, obj.driver)
else:
    Collect_log.validate_part_status(log_file, "PASS", test_case)
    Collect_log.simple_log(log_file, "\nTest case PASS")
    Collect_log.test_case_exit(log_file, startTime, obj.driver, True)

