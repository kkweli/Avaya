import json
import os
import re
import time
from datetime import datetime

import pyautogui
import requests
from selenium.common.exceptions import NoSuchElementException, InvalidSessionIdException, TimeoutException, \
    NoAlertPresentException, NoSuchWindowException, UnexpectedAlertPresentException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait

import avayaFile
import browser


# noinspection PyTypeChecker,SpellCheckingInspection
class Avaya:
    # load credits
    dateCred = datetime.now()
    pgcreds = ["App: Avaya Recordings Bulk Download", "Version: Selenium ARBD-0.0.1", "Owner : Stanbic Bank LTD"]
    print('\n'.join(map(str, pgcreds)))
    print("Date : {0} {1}".format(dateCred.year, dateCred.strftime("%A")))
    # load config params
    loginC = avayaFile.Fileop.CjsonLoad('', "conf.json")
    # load select params
    selPar = avayaFile.Fileop.SjsonLoad('', 'selection.json')
    driver = browser.driverOpt('ie')
    # test url connection
    url1 = requests.head("http://10.235.6.180:8080/servlet/acr").status_code
    if url1 == 200:
        driver.get('http://10.235.6.180:8080/servlet/acr')
        print("\nConnected on address {0} : status code {1} \n".format('http://10.235.6.180:8080/servlet/acr', url1))
    else:
        url2 = requests.head("http://10.235.1.126:8080/servlet/acr").status_code
        if url2 == 200:
            driver.get('http://10.235.6.180:8080/servlet/acr')
            print("Connected on address {0} : status code {1} \n".format('http://10.235.6.180:8080/servlet/acr', url1))
        else:
            print(
                "Failed connection on main and standby avaya recorder with code : {0} and {1} respectively. \n".format(
                    url1, url2))

    @staticmethod
    def msgfinal(msg):
        return "window.alert('" + msg + "');"

    # Convert Function which takes in
    # 24hour time and convert it to
    # 12 hour format
    @staticmethod
    def convert12(timeStr):
        new_format = datetime.strptime(timeStr, "%H:%M:%S").strftime("%r")
        return new_format

    @staticmethod
    def writeToJson(tdata, jsFile):
        jfile = os.path.join(avayaFile.Fileop.dwnDir(''), jsFile)
        # print(''.join(tdata[-2]).split(" "))
        slData = ''.join(tdata[-2]).split(" ")
        lt = len(slData)
        # print(lt)
        d = Avaya.driver

        if len(tdata) != 0:
            if lt > 2:
                lenTime = slData[1] + " " + slData[2]
            else:
                lenTime = slData[1]
            lenDate = slData[0]
            print("Scrap time : ", lenTime)
            print("Scrap date : ", lenDate)
            with open(jfile, 'r+') as f:
                json_data = json.load(f)
                jtime = json_data['start_time']
                if len(lenTime) == 11 and len(json_data['start_time']) != 11:
                    jtime = Avaya.convert12(json_data['start_time'])

                print("Config time: ", jtime)
                print("Config date: ", json_data['start_date'])
                # if json_data['start_date'] <= lenDate and jtime != lenTime:
                if json_data['start_date'] == lenDate and jtime == lenTime:
                    try:
                        # delete downloaded files
                        avayaFile.Fileop.removefiles('', Avaya.loginC['download_dir'])
                        print("Exhausted date time series terminating program, ignoring download procedures.\n")
                        d.execute_script(
                            Avaya.msgfinal('Sorry the date time series has been exhausted for this query, exiting.'))
                        jlertsObj = d.switch_to.alert
                        time.sleep(5)
                        jlertsObj.accept()
                        d.close()
                    except NoAlertPresentException as nae:
                        print("Warning : {0}\n".format(nae))
                    except NoSuchWindowException as nse:
                        print("Warning : {0}\n".format(nse))
                    return False
                else:
                    json_data['start_date'] = "" + lenDate + ""
                    json_data['start_time'] = "" + lenTime + ""
                    f.seek(0)
                    f.write(json.dumps(json_data, indent=4, sort_keys=True))
                    f.truncate()
                    f.close()
                    return True
        else:
            try:
                # delete downloaded files
                avayaFile.Fileop.removefiles('', Avaya.loginC['download_dir'])
                print("Exhausted date time series terminating program, ignoring download procedures.\n")
                d.execute_script(
                    Avaya.msgfinal('Sorry the date and time series has been exhausted for this query, exiting.'))
                jlertsObj = d.switch_to.alert
                time.sleep(5)
                jlertsObj.accept()
                d.close()
            except NoAlertPresentException as nae:
                print("Warning : {0}\n".format(nae))
            except NoSuchWindowException as nse:
                print("Warning : {0}\n".format(nse))
            return False

    @staticmethod
    def writetoFile(wfileName, hook, oData):
        fdir = os.path.join(avayaFile.Fileop.dwnDir(''), "reports")
        # ensure directory exists
        avayaFile.Fileop.RecFileDir('', fdir)
        condir = avayaFile.Fileop.isDirectory('', fdir)
        wfName = os.path.join(fdir, wfileName)
        true_recs = []
        if condir:
            with open(wfName, hook, encoding='utf-8') as f:
                comp_str = ",, REFRESH,,"
                for rD in oData:
                    if comp_str in rD:
                        print("Sanity check completed.")
                    else:
                        f.write(str(rD))
                        true_recs.append(rD)
                        # print (rD)
            f.close()
        return true_recs

    @staticmethod
    def setFocus():
        dr = Avaya.driver
        # ndr.maximize_window()
        return dr.switch_to.window(dr.current_window_handle)

    @staticmethod
    def exceAvaya():
        try:
            # new_avaya = Avaya()
            eaDriver = Avaya.driver
            WebDriverWait(eaDriver, 30).until(
                EC.presence_of_element_located((By.XPATH,
                                                '//*[@id="workAreaWrapper"]/table[3]/tbody/tr/td/table[2]/tbody/tr[1]/td[2]'))
            )
            # set login
            print("Commencing Login, user intervention will be required.\n")
            eaDriver.execute_script("document.getElementById('username').value='" + Avaya.loginC['user'] + "'")
            eaDriver.find_element_by_id('j_password').click()
            eaDriver.find_element_by_id('j_password').clear()
            try:
                eaDriver.execute_script(
                    Avaya.msgfinal("Hello! Please enter password and click (OK) to login."))
                alertsObj1 = eaDriver.switch_to.alert
                time.sleep(2)
                alertsObj1.accept()
                # auto login enabled
                if len(Avaya.loginC['secret']) != 0:
                    time.sleep(2)
                    # set secret
                    eaDriver.execute_script(
                        "document.getElementById('j_password').value='" + Avaya.loginC['secret'] + "'")
                    time.sleep(1)
                    # click on ok button
                    eaDriver.find_element_by_css_selector('#loginToolbar_LOGINLabel > a').click()
            except NoAlertPresentException as nae:
                print("Warning : {0}\n".format(nae))
            except NoSuchWindowException as nse:
                print("Warning : {0}\n".format(nse))

            try:
                WebDriverWait(eaDriver, 50).until(EC.presence_of_element_located(
                    (By.CSS_SELECTOR, '#contentTitleDiv > table > tbody > tr > td.headerText1')))
                print("Great Logged in, selection criteria will be set next.\n")
                Avaya.setFocus()
                # navigate to call logs menu
                eaDriver.find_element_by_xpath(
                    '/html/body/table/tbody/tr/td/table[1]/tbody/tr/td[13]/table/tbody/tr/td').click()
                # input start date
                start_date = eaDriver.find_element_by_xpath(
                    '//*[@id="GeneralSetupGeneric5CollapsibleContainerr0Nodes"]/table/tbody/tr[3]/td[1]/input')
                eaDriver.execute_script("arguments[0].innerText = '" + Avaya.selPar['start_date'] + "'", start_date)
                # input end date
                end_date = eaDriver.find_element_by_xpath(
                    '//*[@id="GeneralSetupGeneric5CollapsibleContainerr0Nodes"]/table/tbody/tr[4]/td[1]/input')
                eaDriver.execute_script("arguments[0].innerText = '" + Avaya.selPar['end_date'] + "'", end_date)
                # input start time
                start_time = eaDriver.find_element_by_xpath(
                    '//*[@id="GeneralSetupGeneric5CollapsibleContainerr0Nodes"]/table/tbody/tr[3]/td[3]/input')
                eaDriver.execute_script("arguments[0].innerText = '" + Avaya.selPar['start_time'] + "'", start_time)
                # select parties incl
                optParties = Select(eaDriver.find_element_by_xpath(
                    '//*[@id="GeneralSetupGeneric5CollapsibleContainerr0Nodes"]/table/tbody/tr[10]/td[1]/select'))
                optParties.select_by_value('8')
                # input parties
                inc_parties = eaDriver.find_element_by_xpath(
                    '//*[@id="GeneralSetupGeneric5CollapsibleContainerr0Nodes"]/table/tbody/tr[10]/td[2]/input')
                eaDriver.execute_script("arguments[0].innerText = '" + Avaya.selPar['parties'] + "'", inc_parties)
                # click on search
                eaDriver.find_element_by_xpath(
                    '//*[@id="toolbarWrapper"]/table/tbody/tr/td[2]/table/tbody/tr/td').click()
                print("The result set will be expanded to show full list if paginated.\n")
            except NoSuchElementException as ee:
                print("Warning : {0}.\n".format(ee))
            except TimeoutException:
                try:
                    eaDriver.execute_script(
                        Avaya.msgfinal("Closing: Unfortunately, took too long to successfully login. Goodbye!."))
                    alertsObj1 = eaDriver.switch_to.alert
                    time.sleep(4)
                    alertsObj1.accept()
                    eaDriver.close()
                except NoAlertPresentException as nae:
                    print("Warning : {0}\n".format(nae))
                except NoSuchWindowException as nse:
                    print("Warning : {0}\n".format(nse))
            # collapse list
            try:
                try:
                    ele_col = eaDriver.find_element_by_link_text('Show All')
                    # ('//*[@id="GeneralSetupGeneric5CollapsibleContainerr0"]/table/tbody/tr/td[1]/a[6]')
                    ele_col.click()
                except NoSuchElementException as ee:
                    print("Collapse option unavailable due to limited result set. Msg: {0}\n".format(ee))

                print("Selecting full list for download.\n")
                # select all files for download
                eaDriver.find_element_by_xpath('//*[@id="crtableheading"]/td[10]/input').click()

                print("Collecting calls metadata from the result set.\n")

                # get count all search results
                alt_row = eaDriver.find_elements_by_class_name('formRowLight')
                first_row = eaDriver.find_elements_by_class_name('formRowLightAlternate')
                # collect elements into arrays
                el_atrs = []
                el_atrs2 = []
                time_vals = []

                # iterate through element id's
                for r_att1 in first_row:
                    # derive elements id's
                    r_att1 = r_att1.get_attribute('id')
                    # call parameters
                    call_timdat = '//*[@id="' + r_att1 + '"]/td[2]'
                    cell_row1a = eaDriver.find_element_by_xpath(call_timdat)
                    call_len = '//*[@id="' + r_att1 + '"]/td[3]'
                    cell_row1b = eaDriver.find_element_by_xpath(call_len)
                    call_agent = '//*[@id="' + r_att1 + '"]/td[4]'
                    cell_row1c = eaDriver.find_element_by_xpath(call_agent)
                    call_parties = '//*[@id="' + r_att1 + '"]/td[5]'
                    cell_row1d = eaDriver.find_element_by_xpath(call_parties)
                    # call_service = '//*[@id="'+r_att1+'"]/td[6]'
                    # cell_row1e = eaDriver.find_element_by_xpath(call_service)
                    # get cell texts values
                    cell_val_str = str(r_att1).translate({ord(i): None for i in
                                                          'row_'}) + "," + cell_row1a.text + "," + cell_row1b.text + "," + cell_row1c.text + "," + cell_row1d.text + "\n"
                    el_atrs.append(cell_val_str)
                    # get time series
                    # time_vals.append(cell_row1a.text)
                # file name parameters
                dataFile = datetime.now().strftime("%Y%m%d-%H%M%S") + "_" + re.sub("[(,)A-Za-z]", "",
                                                                                   Avaya.selPar[
                                                                                       'parties']) + ".txt"
                print("Writing to report log the first set call metadata.\n")
                # write data to file
                Avaya.writetoFile(dataFile, 'w+', el_atrs)

                for r_att2 in alt_row:
                    # derive elements id's
                    r_att2 = r_att2.get_attribute('id')
                    if r_att2 != "screentargetrow":
                        # call parameters
                        call_timdat_alt = '//*[@id="' + r_att2 + '"]/td[2]'
                        cell_row1a_alt = eaDriver.find_element_by_xpath(call_timdat_alt)
                        call_len_alt = '//*[@id="' + r_att2 + '"]/td[3]'
                        cell_row1b_alt = eaDriver.find_element_by_xpath(call_len_alt)
                        call_agent_alt = '//*[@id="' + r_att2 + '"]/td[4]'
                        cell_row1c_alt = eaDriver.find_element_by_xpath(call_agent_alt)
                        call_parties_alt = '//*[@id="' + r_att2 + '"]/td[5]'
                        cell_row1d_alt = eaDriver.find_element_by_xpath(call_parties_alt)
                        # call_service_alt = '//*[@id="'+r_att2+'"]/td[6]'
                        # cell_row1e_alt = eaDriver.find_element_by_xpath(call_service_alt)
                        # get cell texts values
                        cell_string = str(r_att2).translate({ord(i): None for i in
                                                             'row_'}) + "," + cell_row1a_alt.text + "," + cell_row1b_alt.text + "," + cell_row1c_alt.text + "," + cell_row1d_alt.text + "\n"
                        el_atrs2.append(cell_string)
                        # get time series
                        time_vals.append(cell_row1a_alt.text)
                # write data to file
                print("Writing to report log the second set call metadata.\n")
                Avaya.writetoFile(dataFile, 'a+', el_atrs2)
                print("Sanitizing metadata from the full list recordings.\n")
                tdseries = Avaya.writeToJson(time_vals, "selection.json")
                # print("time values : ", time_vals)
                if tdseries:
                    # click on save icon
                    try:
                        eaDriver.execute_script('doExportPop(); return false;')
                        time.sleep(10)
                    except UnexpectedAlertPresentException as uea:
                        print("Alert blocking : msg - {0}".format(uea))
                        eaDriver.close()
                    print("Commencing call recording files download.\n")
                    # folder exists
                    down_path = os.path.normpath(Avaya.loginC['download_dir'])
                    if not avayaFile.Fileop.isDirectory('', down_path):
                        avayaFile.Fileop.makDir('', down_path)

                    print("Checking local download folder ({0}) exists.\n".format(down_path))
                    # get folder count
                    item_saved = sum(f.endswith('.wav') for f in os.listdir(down_path))
                    print("download folder count before {0} file(s)\n".format(item_saved))

                    if item_saved >= 1:
                        # clear contents for run
                        avayaFile.Fileop.removefiles('', down_path)
                        print("Clearing local download folder ({0}).\n".format(down_path))

                    source_dr = os.path.join(avayaFile.Fileop.dwnDir(''), "reports")
                    source_fl = avayaFile.Fileop.newest('', source_dr)
                    total_files = len(open(source_fl).readlines())
                    total_files = total_files
                    while sum(f.endswith('.wav') for f in os.listdir(down_path)) <= 0:
                        # click on save icon
                        time.sleep(1)
                        eaDriver.switch_to.window(eaDriver.window_handles[-1])
                        pyautogui.press('enter')
                        print("Saving call recordings.\n")
                        break

                    while True:
                        tcount = sum(f.endswith('.wav') for f in os.listdir(down_path))
                        print('Download progress count : {0} of final count : {1}'.format(tcount, total_files))
                        if tcount >= 0:
                            # control the loop pulse
                            time.sleep(6)
                            # eaDriver.switch_to.window(eaDriver.window_handles[-1])
                            # pyautogui.press('enter')
                            print("Saving call recordings.\n")

                        if tcount == total_files:
                            # click to close download window
                            eaDriver.switch_to.window(eaDriver.window_handles[-1])
                            eaDriver.find_element_by_css_selector(
                                '#toolbarWrapper > table > tbody > tr > td.before > table > tbody > tr > td').click()
                            print("Closing download window.\n")
                            # eaDriver.execute_script('doScreenClose(); return false;')
                            time.sleep(1)
                            break

                    try:
                        if sum(f.endswith('.wav') for f in os.listdir(down_path)) == total_files:
                            print("Download now complete.\n")
                            # call main file operation method
                            avayaFile.Fileop.main('')
                            # focus on browser
                            eaDriver.switch_to.window(eaDriver.window_handles[-1])
                            eaDriver.execute_script(
                                Avaya.msgfinal("Great news download complete, Goodbye Closing Shortly."))
                            alertsObj = eaDriver.switch_to.alert
                            time.sleep(3)
                            alertsObj.accept()
                    except NoAlertPresentException as nae:
                        print("Warning : {0}\n".format(nae))
                    except NoSuchWindowException as nse:
                        print("Warning : {0}\n".format(nse))
                    finally:
                        # Close driver
                        eaDriver.quit()

            except NoSuchElementException as ee:
                print("Warning!! Msg: {0}\n".format(ee))
            except InvalidSessionIdException as se:
                print("Warning : {0} \n".format(se))

        except NoSuchElementException as ee:
            print("Avaya landing page exception error : {0}\n".format(ee))
            Avaya.driver.close()
        except TimeoutException as te:
            print("Avaya landing page timeout error : {0}\n".format(te))
            Avaya.driver.close()
        except InvalidSessionIdException as se:
            print("Warning : {0}\n".format(se))
            Avaya.driver.close()


def main():
    Avaya.exceAvaya()


if __name__ == "__main__":
    main()
