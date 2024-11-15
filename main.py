import sys
import os
from pathlib import Path
from rpa_bot.bot import Bot
from process_sap import SapProcess
from notifications import Notifications
from rpa_bot.log import lte, log

class ConfigModel:
    sap_client = '010'
    sap_language = 'EN'
    pce_ace = None
    bot_version = '3.0.0'
    abspath = os.path.dirname(os.path.abspath(__file__))
    main_path = Path(abspath)

class Task(Bot): # Process
    __version__ = '3.0.0.0'


    def __init__(self, sysargs):
        self.config = ConfigModel  # Use if you want
        super().__init__(sysargs)

    def run_bot(self):
        from excel import ExcelProcess
        import xlwings as xw
        import os
        import datetime
        import shutil
        import pandas as pd

        # sap system, vault dict, credentials.item_name, client  (ew language)
        with SapProcess(self.config.sap_system, 
                     self.vault_dictionary, 
                     self.credentials.SAP_, # SAP user
                     self.config.pce_ace, 
                     client=self.config.sap_client,
                     main_path=self.config.main_path) as sap:
            try:
                self.replace_body = None
                excel = ExcelProcess(self.main_path)
                notifications = Notifications()
                # parameters
                exception_list = []
                company_code = self.config.company_code1
                company_code2 = self.config.company_code2
                # need to be setup first in SAP if needed
                # layout_me2n = self.config.layout_me2n
                variant_cockpit = self.config.variant_cockpit
                layout_cockpit1 = self.config.layout_cockpit1
                # layout_cockpit2 = self.config.layout_cockpit2

                file_name = 'Export.xlsx'
                action = 'r'
                posting_date = datetime.datetime.today()
                day = posting_date.day
                if day <= 2:
                    posting_date = excel.last_business_day()
                manual_trigger = False
                if len(sys.argv) > 3:
                    email_subject = sys.argv[3]
                    path_attachments = sys.argv[5]
                    manual_trigger = True
                    check_delay_start = lambda x: '_delayStart' in x
                    action = check_delay_start(email_subject) # email_subject[-1]

                if manual_trigger:
                    # list of documents provided by client
                    df_vendors = sap.get_vendors(manual_trigger)
                    sap.setup_cockpit(variant_cockpit, layout_cockpit1, df_vendors['Vendor Id'])
                    log('Cockpit with paramenters run')
                    df_processable_docs = sap.get_data_for_process(manual_trigger, file_name, self.temp_path, path_attachments)
                else:
                    df_vendors_by_comp = sap.get_vendors(manual_trigger)
                    df_processable_docs = pd.DataFrame()
                    for company_code, df_vendors in df_vendors_by_comp:
                        sap.setup_cockpit(variant_cockpit, layout_cockpit1, company_code, df_vendors['Vendor Id'])
                        log('Cockpit with paramenters run')
                        df_processable_docs = sap.get_data_for_process(manual_trigger, file_name, self.temp_path, None, df_processable_docs, company_code)

                file_name = sap.prepare_process_list(self.temp_path, df_processable_docs)
                log(f"List of documents to process {file_name} saved")

                report_path = os.path.join(self.temp_path, file_name)
                self.mail_attachments = [report_path]
                company_code = df_vendors_by_comp[0:][0][0]
                company_code2 = df_vendors_by_comp[1:][0][0]
                sap.setup_cockpit(variant_cockpit, layout_cockpit1, company_code, pd.DataFrame(), company_code2)

                wb = xw.Book(os.path.join(self.temp_path, file_name))
                log(f"Excel: {file_name} opened")
                app = wb.api.Application
                ws_data = wb.sheets["MM_FR_readyToProcess"]
                ws_data.range(1, 33).value = 'Invoice document number'
                ws_data.range(1, 34).value = 'Processing status'
                last_row = ws_data.api.Cells.Find(What="*",
                            After=ws_data.api.Cells(1, 1),
                            LookAt=xw.constants.LookAt.xlPart,
                            LookIn=xw.constants.FindLookIn.xlFormulas,
                            SearchOrder=xw.constants.SearchOrder.xlByRows,
                            SearchDirection=xw.constants.SearchDirection.xlPrevious,
                            MatchCase=False).Row
                log(f"{file_name} read")
                log("Document processing started...") 

                for index in range(2, last_row + 1): 
                    docNumber = int(ws_data.range(index, 5).value) 
                    company_code = ws_data.range(index, 7).value
                    status = ws_data.range(index, 34).value

                    if status is None: 
                        info = sap.process_item(docNumber, company_code) 
                        if str(info).isnumeric():
                            self.counters.inc_success()
                            ws_data.range(index, 33).value = info
                            ws_data.range(index, 34).value = f"Document posted with number {info}"
                            wb.save()
                        elif str(info).find('The control could not be found by id.') != -1 or str(info).find('Error in function openInvoice') != -1 or str(info).find('The object invoked has disconnected') != -1:
                            self.counters.inc_error()
                            info1 = f"Document {docNumber} cannot be processed. The control could not be found by id."
                            info2 = str(info).split('.')[0] + ". The control could not be found by id."
                            exception_list.append([docNumber, info2])
                            ws_data.range(index, 34).value = info1
                            wb.save()
                            if index < last_row:
                                sap.close_session()
                                sap.connect()
                                sap.setup_cockpit(variant_cockpit, layout_cockpit1, company_code, pd.DataFrame(), company_code2)
                        elif str(info).find('Document has not been found') != -1:
                            self.counters.inc_success()
                            ws_data.range(index, 34).value = f"{info}"
                            wb.save()
                        else:
                            self.counters.inc_success()
                            ws_data.range(index, 34).value = f"{info}"
                            wb.save()
                            sap.back_to_cockpit()
                log("Documents processing finished")
                wb.save()
                log(f"Report {file_name} saved")
                app.Quit()
                sap.close_session()
                log("SAP user logged out")
                # if there is no error and we reached to this point than it is success
                if self.bot_status == self.bot_status.unhandled:
                    self.bot_status = self.bot_status.success
            except Exception as e:
                if 'The object invoked has disconnected from its clients.' in str(e):
                    sap.kill_sap()
                if sap.sap_exists():       
                    sap.close_session()
                if os.path.exists(report_path):
                    wb.save()
                    app.Quit()
            # send exceptions 
            if exception_list:
                mail_subject = "RPA Bot: MMInvoiceFrance Exceptions"
                mail_body = notifications.exceptions_body(exception_list)
                self.mail_report(replace_body=mail_body, replace_subject=mail_subject)
                log("Exceptions list sent to RPA Team")
            # send report to receiver
            self.replace_subject = "RPA Bot: MMInvoiceFrance Finished"
            self.replace_body = notifications.report_body(self)
            self.mail_to = self.config.report_receiver_to
            # archive report
            archive_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'archive')
            if not os.path.exists(archive_path):
                os.makedirs(archive_path)
            shutil.copy2(report_path, os.path.join(archive_path, file_name))
            log(f"Report {file_name} archived")
            return True
        

if __name__ == '__main__':     
    # Child class context manager
    with Task(sysargs=sys.argv) as bot:
        bot.run_bot()
