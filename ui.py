import datetime, os, threading, queue
import PySimpleGUI as sg
from Paylocity_Download import RunPay
from openpyxl import load_workbook
from Paylocity_Process import runPay
sg.theme('DarkBlue3')
gui_queue = queue.Queue()


def download_pay(startdate,enddate):
    pay = RunPay()
    pay.gui_queue = gui_queue
    pay_status = pay.run(startdate,enddate)
    return pay_status

def process_pay(filepath,startdate,enddate):
    run = runPay()
    run.gui_queue = gui_queue
    run.run(filepath,startdate,enddate)
    return True

def run_gui(thread=None):
    end_date = (datetime.date.today() - datetime.timedelta(days=1)).strftime('%m/%d/%Y')
    start_date = datetime.datetime.strptime(end_date, '%m/%d/%Y').replace(day=1).strftime('%m/%d/%Y')
    layout = [
        [
            sg.Text('Paylocity',
                    size=(40, 1),
                    font=('Corbel', 18),
                    justification='center',
                    pad=((0, 0), (5, 10)))
        ],
        [
            sg.CalendarButton("Start Date", size=(12, 1), format='%m/%d/%Y', key='report_btn', enable_events=True),
            sg.Input(start_date, size=(15, 1), font=('Corbel', 11), key='startdate', disabled=True,
                     justification='center', enable_events=True, readonly=True),
        ],
        [
            sg.CalendarButton("End Date", size=(12, 1), format='%m/%d/%Y', key='report_btn', enable_events=True),
            sg.Input(start_date, size=(15, 1), font=('Corbel', 11), key='enddate', disabled=True,
                     justification='center', enable_events=True, readonly=True),
        ],
        [
            sg.Text("Choose a file: "), sg.Input(),
            sg.FolderBrowse(initial_folder=os.getcwd(), key="filepath"),
        ],
        [
            sg.OK('Report Download', key='report_download', size=(15, 1), font=('Corbel', 12), pad=((5, 5), (10, 0))),
            sg.OK('Prepare Report', key='prepare_report', size=(15, 1), font=('Corbel', 12), pad=((5, 5), (10, 0))),
            sg.Exit('Exit', key='exit', size=(15, 1), font=('Corbel', 12), pad=((5, 5), (10, 0))),
        ],
        [
            sg.Text("Status :", size=(15, 1), justification='left', font=('Corbel', 11)),
        ],
        [
            sg.Multiline(size=(60, 7), font='courier 10', background_color='white', text_color='black', key='status',
                         autoscroll=True, enable_events=True, change_submits=False)
        ],
    ]

    window = sg.Window('PayloCity',
                       element_justification='left',
                       text_justification='left',
                       auto_size_text=True).Layout(layout)

    while True:
        event, values = window.Read(timeout=1000)
        window.refresh()

        if event in ('Exit', None) or event == sg.WIN_CLOSED:
            window.close()
            break

        elif event == 'report_download':
            startdate = values['startdate']
            enddate = values['enddate']
            weekcount = enddate - startdate
            window['status'].print('Paylocity Report Download Processing...\n')
            window['report_download'].Update(disabled=True)
            window['prepare_report'].Update(disabled=True)
            thread = threading.Thread(target=download_pay, args=(startdate,enddate))
            thread.start()

        elif event == 'prepare_report':
            filepath = values['filepath']
            startdate = datetime.datetime.strptime(values['startdate'],"%m/%d/%Y")
            enddate = datetime.datetime.strptime(values['enddate'],"%m/%d/%Y")
            if startdate > enddate:
                window['status'].print('Start date cannot be greater then end date\n')
                continue

            filelist = None
            if filepath:
                filelist = os.listdir(filepath)
                filelist = [str(file).split('.')[0] for file in filelist]
                setting = 'PaylocitySettingSheet.xlsx'
                setting_wb = load_workbook(setting, data_only=True, read_only=True)
                setting_ws = setting_wb['Files'].values
                report_data = [list(row) for row in setting_ws if row]
                count = 0
                for report in report_data:
                    if report[0] in filelist:
                        count += 1
                if count != len(report_data):
                    window['status'].print('Some files are missing in the folder selected')
                    continue
                window['status'].print('Paylocity Processing...\n')
                window['report_download'].Update(disabled=True)
                window['prepare_report'].Update(disabled=True)
                thread = threading.Thread(target=process_pay, args=(filepath,startdate,enddate))
                thread.start()
            elif not filelist:
                window['status'].print('No path selected')
            else:
                window['status'].print('Please select folder properly')

        elif event == "exit" or event == sg.WIN_CLOSED:
            window.close()
            break

        if thread:
            if not thread.is_alive():
                window['report_download'].Update(disabled=False)
                window['prepare_report'].Update(disabled=False)
                window.refresh()

        try:
            message = gui_queue.get_nowait()
        except:
            message = None
        if message:
            for key, value in message.items():
                if key == 'status':
                    window['status'].print(value)
                    window.refresh()
                if key == 'Success':
                    sg.Popup(value, title='Status')
            window.refresh()


if __name__ == '__main__':
    # main function
    run_gui()
