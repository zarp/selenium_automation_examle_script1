import csv
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select

import time

def get_GC_list(filename_with_GCs):
    with open(filename_with_GCs, 'r') as f:
      reader = csv.reader(f)
      GC_list = list(reader)
      GC_list = [GC[0] for GC in GC_list]
    return GC_list

def generate_GC_urls(GC_list):
    prefix = "http://ccsweb/ccsV2/ccsRequestV2/ccsGCRequest.aspx?JobNumCode="
    return [prefix + GC for GC in GC_list]

def autodispo(GC_list):
    prefix = "http://internal_network/ccs/GCLotFcCodeDisposition.aspx?GCLotCode=GC"
    error_msg = "Error getting GC Lot Fc Code Information. Error - There is no row at position 0."
    for GC in GC_list:
        dash_num = 0 
        while True:
            url = prefix + GC + "-" + str(dash_num)
            driver.get(url)
            pagesource = driver.page_source
            if error_msg in pagesource:
                break
            else:
                print(url)
                
                dispos = driver.find_elements_by_name('ksw100$ContentPlaceHolder1$gridGCLotFcCodes$ctl03$grdGCFcCodeDisposition$ctl03$Add_ddlLossCodeType')
                if len(dispos)>0:                    
                    dispo_select = Select(dispos[0])
                    #dispo_select = Select(driver.find_element_by_name('ksw100$ContentPlaceHolder1$gridGCLotFcCodes$ctl03$grdGCFcCodeDisposition$ctl03$item_ddlLossCodeType'))
                    try:
                        dispo_select.select_by_value(str(6)) #'6' = pre-etest failure. 'lost part' would be '5'
                        status_select = Select(driver.find_element_by_name("ksw100$ContentPlaceHolder1$gridGCLotFcCodes$ctl03$grdGCFcCodeDisposition$ctl03$Add_ddlStatusID"))
                        status_select.select_by_value(str(100)) #100 = complete
                        qty_field = driver.find_element_by_name("ksw100$ContentPlaceHolder1$gridGCLotFcCodes$ctl03$grdGCFcCodeDisposition$ctl03$Add_txtLossTypeQty")
                        qty_field.send_keys(input()) # not fully automated to let engineer double check that dispositioning is safe in this case
                        save_button = driver.find_element_by_partial_link_text("Save Changes")
                        save_button.click() 
                    except:
                        pass
                             
            dash_num += 1
        
def update_forecast(urls, new_forecast_date):
    for url in urls:
        print('updating ', url)
        driver.get(url)
        assert "GC" in driver.title #GC Request Form
        lot_edit_buttons = driver.find_elements_by_link_text("Edit")
        number_of_lots = len(lot_edit_buttons)
        
        for cur_lot_number in range(number_of_lots):
            lot_edit_button = driver.find_elements_by_link_text("Edit")[cur_lot_number]
            print('updating lot forecast date')
            time.sleep(10) 
            lot_edit_button.click()

            ct_substring = str(cur_lot_number + 2)
            if len(ct_substring) == 1:
                ct_substring = "0" + ct_substring
                    
            elem = driver.find_elements_by_name("dgJobLots$ctl" + ct_substring + "$edit_ForecastBeginDate")[0] #example: name="dgJobLots$ctl02$edit_ForecastBeginDate" for lot 1
            elem.clear()
            elem.send_keys(new_forecast_date)


            elem = driver.find_elements_by_link_text("Update")[0]
            elem.click()

            pagesource = driver.page_source
            if ("TC and DC are required when checking submit to the characterization lab" in pagesource) or ("TC is required." in pagesource):
                elem = driver.find_elements_by_name("dgJobLots$ctl" + ct_substring + "$edit_TC")[0]
                elem.clear()
                elem.send_keys("n/a")
                elem = driver.find_elements_by_name("dgJobLots$ctl" + ct_substring + "$edit_DC")[0]
                elem.clear()
                elem.send_keys("n/a")
                elem = driver.find_elements_by_link_text("Update")[0]
                elem.click()

def cancel_GC_jobs(urls):
    for url in urls:
        print('cancelling ', url)
        driver.get(url)
        assert "GC" in driver.title
        job_status_select = Select(driver.find_element_by_id("ddlJobStatusID"))
        job_status_select.select_by_value(str(6)) # 6 == canceled job
        elem = driver.find_element_by_id("btnUpdate")
        elem.click()

if __name__ == "__main__":
    filename_with_GCs = "GCs_to_load.txt"
    command = input() #"update_forecast", "cancel_job", 'autodispo'
    if command == 'update_forecast':
        new_forecast_date = input() #yyyymmdd format


    GC_list = get_GC_list(filename_with_GCs)


    driver = webdriver.Ie()

    if command == "update_forecast":
        urls = generate_GC_urls(GC_list)
        update_forecast(urls, new_forecast_date)
    elif command == "cancel_job":
        urls = generate_GC_urls(GC_list)
        cancel_GC_jobs(urls)
    elif command == 'autodispo':
        autodispo(GC_list)
    else:
        raise ValueError("Unknown command")


    print("Done!")

