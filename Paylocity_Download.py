import datetime
import queue
import re
import time

from selenium.webdriver.support.ui import WebDriverWait,Select
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from msedge.selenium_tools import Edge, EdgeOptions
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from openpyxl import load_workbook
from time import sleep
import datetime
import os


class Paylocity:
        def __init__(self, gui_queue):
            self.start_date = self.end_date = self.username = self.password = self.company = self.coid = None
            self.setting_dict = None
            self.report_data = None
            self.gui_queue = gui_queue
            self.login_url = r'https://access.paylocity.com/'
            #self.logout_url = r'https://comptonbusinesssolutions.sharefile.com/Authentication/Login'

        def start_edge(self, download_pdf=True, download_prompt=False):
            self.downloadPath = os.path.join(os.getcwd(), 'Downloads','Paylocity',datetime.date.today().strftime("%m-%d-%Y"))
            if not os.path.isdir(self.downloadPath):
                os.makedirs(self.downloadPath)

            self.existing_files = os.listdir(self.downloadPath)
            self.existing_files = []

            edge_options = EdgeOptions()
            edge_options.use_chromium = True
            edge_options.add_experimental_option(
                "prefs", {
                    "behavior": "allow",
                    "download.prompt_for_download": download_prompt,
                    "plugins.always_open_pdf_externally": download_pdf,
                    "download.default_directory": self.downloadPath,
                    "safebrowsing.enabled": False,
                    "safebrowsing.disable_download_protection": True
                }
            )
            self.driver = Edge(
                executable_path=EdgeChromiumDriverManager(log_level=0).install(),
                options=edge_options,
            )
            self.driver.maximize_window()

        def load_login_page(self):
            self.driver.get(self.login_url)
            trial = 0
            while trial < 3:
                if self.driver.title == "Login | Paylocity":
                    return True
                else:
                    trial += 1
                    sleep(2)
            return False

        def login_pay(self):
            username = self.username
            password = self.password
            companyid = self.company
            coid = self.coid
            if not username or not password or not companyid:
                self.gui_queue.put({'status': f'Credentials not found in setting sheet to download reports.'}) \
                    if self.gui_queue else None
                return False
            try:
                companyXpath = '//*[@name="CompanyId"]'
                company = WebDriverWait(self.driver, 30).until(
                    EC.visibility_of_element_located((By.XPATH, companyXpath)))
                company.clear()
                company.send_keys(companyid)
                sleep(1)

                usernameXpath = '//*[@name="Username"]'
                user_name = WebDriverWait(self.driver, 30).until(
                    EC.visibility_of_element_located((By.XPATH, usernameXpath)))
                user_name.clear()
                user_name.send_keys(username)
                sleep(1)

                passwordXpath = '//*[@name="Password"]'
                password_ = WebDriverWait(self.driver, 30).until(
                    EC.visibility_of_element_located((By.XPATH, passwordXpath)))
                password_.clear()
                password_.send_keys(password)
                sleep(1)

                loginXpath = '//*[@type="submit" and text()="Login"]'
                login = WebDriverWait(self.driver, 30).until(EC.element_to_be_clickable((By.XPATH, loginXpath)))
                login.click()
                sleep(1)

                passcodeXpath = '//*[@type="radio" and @id="Device.OtpType1"]'
                passcode = WebDriverWait(self.driver, 60).until(EC.element_to_be_clickable((By.XPATH, passcodeXpath)))
                passcode.click()
                sleep(1)

                securebtnXpath = '//*[@type="submit" and text()="Send Code"]'
                securebtn = WebDriverWait(self.driver, 60).until(EC.element_to_be_clickable((By.XPATH, securebtnXpath)))
                securebtn.click()
                sleep(1)

                otpXpath = '//*[@id="OneTimePasscode"]'
                otpbtn = WebDriverWait(self.driver, 30).until(EC.element_to_be_clickable((By.XPATH, otpXpath)))
                if otpbtn:
                    self.gui_queue.put({'status': '\n\tAuthentication Code Required.'})
                    # print("Authentication code required")
                    sleep(3)
                    title = 'Paylocity | HR & Payroll'
                    WebDriverWait(self.driver, 300).until(EC.title_is(title))
                    sleep(2)
                else:
                    return False

                titleXpath = '//*[@class="c-header-company-id"]'
                title = 'G&L RESTAURANT LLC [34389]'
                pagetitle = self.driver.find_element(By.XPATH,titleXpath)
                if title != pagetitle:
                    companyXpath = '//*[contains(@name,"CompanyID") and @rptpromptname="CoIDFilter"]'
                    companyfilter = WebDriverWait(self.driver, 30).until(EC.element_to_be_clickable((By.XPATH, companyXpath)))
                    companyfilter.clear()
                    companyfilter.send_keys(coid)
                    searchXpath = '//*[contains(@class,"search_button") and text()=" Search"]'
                    searchbtn = WebDriverWait(self.driver, 30).until(EC.element_to_be_clickable((By.XPATH, searchXpath)))
                    searchbtn.click()
                    sleep(1)
                    companylinkXpath = '//*[@class="datarowlink" and text()="34389"]'
                    companylinkbtn = WebDriverWait(self.driver, 30).until(EC.element_to_be_clickable((By.XPATH, companylinkXpath)))
                    if companylinkbtn:
                        companylinkbtn.click()
                        sleep(1)
                    else:
                        self.gui_queue.put({'status': f'Company Id {coid} not found'}) if self.gui_queue else None
                        return False

                return True

            except Exception as e:
                print(str(e))
                return False

        def process_report(self):
            sleep(3)

            reportmenuXpath = '//*[contains(@data-automation-id,"Reports-&-Analytics") and text()="Reports & Analytics"]'
            reportmenu = WebDriverWait(self.driver, 30).until(EC.element_to_be_clickable((By.XPATH, reportmenuXpath)))
            reportmenu.click()

            reportingXpath = '//*[contains(@data-automation-id,"Reporting") and text()="Reporting"]'
            reporting = WebDriverWait(self.driver, 30).until(EC.element_to_be_clickable((By.XPATH, reportingXpath)))
            reporting.click()

            searchXpath = '//*[contains(@class,"search-box") and contains(@placeholder,"Search by Name, Description, Label")]'

            for file in self.report_data:
                searchbox = WebDriverWait(self.driver, 30).until(EC.element_to_be_clickable((By.XPATH, searchXpath)))
                searchbox.clear()
                searchbox.send_keys(file[0])
                sleep(0.5)
                filelinkXpath = f'//*[@class="report-link" and @title="{file[0]}"]'
                filelink = WebDriverWait(self.driver, 30).until(EC.element_to_be_clickable((By.XPATH, filelinkXpath)))
                if filelink:
                    filelink.click()
                    daterangeXpath = '//*[@id="ctl00_WorkSpaceContent_reportFilterCntrl_stdDateParms_rdoOverrideDates"]'
                    daterange = WebDriverWait(self.driver, 30).until(EC.element_to_be_clickable((By.XPATH, daterangeXpath)))
                    if daterange:
                        reporttypeXpath = '//*[@id="ctl00_WorkSpaceContent_reportFilterCntrl_ddOutputType"]'
                        reporttype = Select(self.driver.find_element(By.XPATH,reporttypeXpath))
                        reporttype.select_by_visible_text(file[1])
                        sleep(0.5)
                        daterange.click()
                        fromdateXpath = '//*[@id="ctl00_WorkSpaceContent_reportFilterCntrl_stdDateParms_ddStartDateRange"]'
                        fromdate = Select(self.driver.find_element(By.XPATH,fromdateXpath))

                        sleep(0.5)
                        dateselectXpath = '//*[@id="ctl00_WorkSpaceContent_reportFilterCntrl_stdDateParms_ddStartDateRange"]/option[1]'
                        dateselect = self.driver.find_element(By.XPATH,dateselectXpath)
                        fromdate.select_by_visible_text(dateselect.text)

                        sleep(0.5)
                        todateXpath = '//*[@id="ctl00_WorkSpaceContent_reportFilterCntrl_stdDateParms_ddEndDateRange"]'
                        todate = Select(self.driver.find_element(By.XPATH, todateXpath))

                        sleep(0.5)
                        dateselectXpath = '//*[@id="ctl00_WorkSpaceContent_reportFilterCntrl_stdDateParms_ddEndDateRange"]/option[1]'
                        dateselect = self.driver.find_element(By.XPATH, dateselectXpath)
                        todate.select_by_visible_text(dateselect.text)

                        sleep(0.5)
                    else:
                        self.gui_queue.put({'status': f'Date Range button not available for {file[0]}'}) if self.gui_queue else None
                        return False

                    runXpath = "//*[@id='UniveralReportingRunButton']//span"
                    runbtn = WebDriverWait(self.driver, 30).until(EC.element_to_be_clickable((By.XPATH, runXpath)))
                    if runbtn:
                        runbtn.click()
                    gotoXpath = '//*[contains(@id,"GoToReportList")]'
                    gotobtn = WebDriverWait(self.driver, 30).until(EC.element_to_be_clickable((By.XPATH, gotoXpath)))
                    if gotobtn:
                        gotobtn.click()
                    else:
                        self.gui_queue.put({'status': f'Unable to run report for {file[0]}'}) if self.gui_queue else None
                        return False
                else:
                    self.gui_queue.put({'status': f'Payroll report for {file[0]} not found'}) if self.gui_queue else None
                    return False
            sleep(2)
            reportpickupXpath = '//*[@data-automation-id="report-pickup-thumb"]'
            reportpickup = WebDriverWait(self.driver, 60).until(EC.element_to_be_clickable((By.XPATH, reportpickupXpath)))
            reportpickup.click()
            sleep(5)
            return True

        def download_report(self):
            sleep(5)
            refreshXpath = '//*[@id="refreshButton"]'
            refreshbtn = WebDriverWait(self.driver, 30).until(EC.element_to_be_clickable((By.XPATH, refreshXpath)))
            refreshbtn.click()
            sleep(3)

            report_data = [report_data[0] for report_data in self.report_data]
            for num in range(1,len(report_data)+1):
                WebDriverWait(self.driver, 30).until(EC.element_to_be_clickable((By.XPATH, refreshXpath)))
                sleep(1)
                reportnameXpath = f'//*[@id="report-pickup-display"]/app-report-pickup-scroller/div/div[2]/div[1]/div/table/tbody/tr[{num}]/td[2]//h4/*'
                reportname = self.driver.find_element(By.XPATH, reportnameXpath)
                # self.driver.execute_script('arguments[0].scrollIntoView();', reportname)
                name = str(reportname.text).split('[')[0].strip()
                sleep(0.5)

                reportlinkXpath = f'//*[@id="report-pickup-display"]/app-report-pickup-scroller/div/div[2]/div[1]/div/table/tbody/tr[{num}]/td[6]//a'
                reportlink = self.driver.find_element(By.XPATH,reportlinkXpath)

                reportdateXpath = f'//*[@id="report-pickup-display"]/app-report-pickup-scroller/div/div[2]/div[1]/div/table/tbody/tr[{num}]/td[4]'
                reportdate = self.driver.find_element(By.XPATH, reportdateXpath)
                date_ = datetime.datetime.strptime(reportdate.text, "%m/%d/%y %I:%M %p").strftime("%m/%d/%Y")

                reportgenXpath = f'//*[@id="report-pickup-display"]/app-report-pickup-scroller/div/div[2]/div[1]/div/table/tbody/tr[{num}]/td[3]'
                reportgen = self.driver.find_element(By.XPATH, reportgenXpath)
                sleep(0.5)

                if name in report_data and date_ == datetime.date.today().strftime("%m/%d/%Y") and reportgen.text == 'Satish Patel':
                    reportlink.click()
                    sleep(2)
            sleep(3)
            return True

        def logout(self):
            sleep(2)
            logoutimgXpath = '//*[@class="c-header-user-portrait"]'
            logoutimg = self.driver.find_element(By.XPATH,logoutimgXpath)
            logoutimg.click()

            logoutXpath = '//*[text()=" Logout"]'
            logout = self.driver.find_element(By.XPATH,logoutXpath)
            logout.click()
            loginXpath = '//*[text()="Login"]'
            WebDriverWait(self.driver, 30).until(EC.element_to_be_clickable((By.XPATH, loginXpath)))
            if self.driver.title == "Login | Paylocity":
                sleep(2)
                self.driver.quit()
                return True
            else:
                sleep(2)
                self.driver.quit()
                return False


class RunPay:
    def __init__(self):
        self.gui_queue = queue.Queue()

    def run(self):
        start_time = time.perf_counter()
        setting = 'PaylocitySettingSheet.xlsx'
        setting_wb = load_workbook(setting, data_only=True, read_only=True)
        setting_ws = setting_wb['Creds'].values
        setting_data = [list(row) for row in setting_ws if row]
        setting_dict = {}
        paylo = Paylocity(self.gui_queue)
        for row in setting_data:
            paylo.company = str(row[0]).strip()
            paylo.username = str(row[1]).strip()
            paylo.password = str(row[2]).strip()
            paylo.coid = re.findall(r'\d+', paylo.company)
            # paylo.start_date = start_date
            # paylo.end_date = end_date
            paylo.setting_dict = setting_dict

            setting_ws = setting_wb['Files'].values
            paylo.report_data = [list(row) for row in setting_ws if row]

            paylo.start_edge()
            login_page = paylo.load_login_page()
            if not login_page:
                self.gui_queue.put({'status': f'\nError : Unable to load login page.'}) if self.gui_queue else None
                return False

            login = paylo.login_pay()
            if not login:
                self.gui_queue.put({'status': f'\nError : Unable to Login.'}) if self.gui_queue else None
                return False

            report = paylo.process_report()
            if not report:
                self.gui_queue.put({'status': f'\nError : Unable to Process Files.'}) if self.gui_queue else None
                return False

            download = paylo.download_report()
            if not download:
                if not report:
                    self.gui_queue.put({'status': f'\nError : Unable to Download Files.'}) if self.gui_queue else None
                    return False

            logout = paylo.logout()
            if not logout:
                self.gui_queue.put({'status': f'\nError : Unable to Logout.'}) if self.gui_queue else None
                return False
        self.gui_queue.put({'status': f'\nFiles downloaded Successfully'}) if self.gui_queue else None

        end_time = time.perf_counter()
        time_taken = time.strftime("%H:%M:%S", time.gmtime(int(end_time - start_time)))
        self.gui_queue.put({'status': f'\nTime Taken : {time_taken}'}) if self.gui_queue else None

