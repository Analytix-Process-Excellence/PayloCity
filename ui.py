import datetime

import PySimpleGUI as sg
import threading
import queue
# from Paylocity import RunPaylocity
from Paylocity_Download import RunPay


sg.theme('DarkBlue3')
gui_queue = queue.Queue()


def download_pay():
    pay = RunPay()
    pay.gui_queue = gui_queue
    pay_status = pay.run()
    return pay_status


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
            sg.CalendarButton('Start Date', size=(12, 1), format='%m/%d/%Y', key='start_date_btn', enable_events=True),
            sg.Input(start_date, size=(12, 1), font=('Corbel', 11), key='start_date', disabled=True,
                     justification='center', enable_events=True, readonly=True),
        ],
        [
            sg.CalendarButton('End Date', size=(12, 1), format='%m/%d/%Y', key='end_date_btn', enable_events=True),
            sg.Input(end_date, size=(12, 1), font=('Corbel', 11), key='end_date', disabled=True,
                     justification='center', enable_events=True, readonly=True),
        ],
        # [
        #     sg.Text('Teams OTP : ', size=(12, 1), auto_size_text=False, justification='left'),
        #     sg.InputText(None, size=(15, 1), key='teams_otp', enable_events=True),
        #     # sg.OK('Submit OTP', key='submit_otp', size=(11, 1), font=('Corbel', 10), disabled=True),
        # ],
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
            window['status'].print('Paylocity Report Download Processing...\n')
            window['report_download'].Update(disabled=True)
            window['prepare_report'].Update(disabled=True)
            thread = threading.Thread(target=download_pay)
            thread.start()

        elif event == 'prepare_report':
            window['status'].print('Paylocity Processing...\n')
            window['report_download'].Update(disabled=True)
            window['prepare_report'].Update(disabled=True)
            # thread = threading.Thread(target=pate_je)
            # thread.start()

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
