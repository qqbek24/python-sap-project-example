from rpa_sap.sap import SAP
from rpa_bot.log import lte, log
import pandas as pd
import os
import glob
from datetime import datetime, timedelta
import time
from excel import ExcelProcess


class ToleranceRange:
    lowTolerance = 'lowTolerance'
    highTolerance = 'highTolerance'
    extremeTolerance = 'extremeTolerance'


class POType:
    EPO = 'EPO'
    Transport = 'Transport'
    interco = 'interco'
    Standard = 'Standard'


class SapProcess(SAP):


    """extension of SAP class"""
    def __init__(self, sap_system, vault_dict, credentials, config, client, main_path):
        super().__init__(sap_system, vault_dict, credentials, client=client)
        self.sap_system = sap_system
        self.excel = ExcelProcess(main_path)

    def process_item(self, doc_number, company_code):
        """
        This function is used to process an item in cockpit.

        Parameters:
        doc_number (str): The document number.
        company_code (str): The company code.

        Returns:
        Bool: True if the item was successfully processed, False otherwise.
        """
        try:
            info = ''
            go_to_final_steps = False
            check_po_line_entry_data = {
                'how_many_lines': 0, 
                'value': 0.00, 
                'different_pos': False, 
                'multiple_pos_exist': False, 
                'rule_5': False, 
                'searched_amount': 0.0, 
                'crop_result': False, 
                'multiple_only': False, 
                'sum_po_lines': 0.0, 
                'saldo': 0.0, 
                'tax_code_missing': False, 
                'missing_tax_codes': ""
            }

            # Step 
            if not self.find_invoice(doc_number, company_code):
                return f"{doc_number}. Document has not been found"

            # get workflow status, description, doc type (MM/FI), FuF, company code
            metaData = self.get_meta_data(doc_number)

            # Step 
            if not self.open_invoice(doc_number):
                return f"Document {doc_number} cannot be processed. Error during opening the document"

            # check workflow status and description
            # if self.check_meta_data(doc_number, metaData):

            # get PO number from workflow description
            info = self.get_wf_status_po_nr(doc_number, metaData)
            if 'cannot be processed' in info:
                return info
            given_po = info

            # check saldo based on workflow status
            info = self.wf_status_invoice_price_difference(doc_number, metaData)
            if info != False and 'final steps' in info:
                go_to_final_steps = True

            # Step
            document_source = self.check_document_source(doc_number)
            
            if go_to_final_steps == False and info is False and document_source != 'PDFCollector':
                # Document ws status saldo is not 0
                if 'cannot be processed' in document_source:
                    return info
                # do if source = Ariba
                elif document_source == 'Ariba':
                    process_data = self.process_ariba(doc_number, given_po)
                    if 'ERROR' in str(process_data).upper() or 'CANNOT BE PROCESSED' in str(process_data).upper():
                        return process_data
                    if 'final steps' in process_data:
                        go_to_final_steps = True
            
            if go_to_final_steps == False:
                # do if source = PDF_collector
                process_data = self.process_pdf_collector(doc_number, given_po)
                if 'ERROR' in str(process_data).upper() or 'CANNOT BE PROCESSED' in str(process_data).upper():
                    return process_data
                # Step
                if not self.take_over_document(doc_number):
                    return f"Document {doc_number} cannot be processed. Error during taking over the document"
                # Step
                info = self.check_doc_type(doc_number)
                if info is not True:
                    return info
                # Step
                info = self.check_process_data(doc_number, company_code, process_data[1], check_po_line_entry_data)
                # if 'ERROR' in str(info).upper() or 'CANNOT BE PROCESSED' in str(info).upper():
                if info is not True:
                    if 'ERROR' in str(info).upper() or 'CANNOT BE PROCESSED' in str(info).upper():
                        return info
                    if 'final steps' in info:
                        go_to_final_steps = True

                # Step
                if go_to_final_steps == False:
                    info = self.process_po_types(doc_number, company_code, process_data[1], check_po_line_entry_data)
                    if info is not True:
                        return info

            # Step
            check_vmd = self.check_vmd(doc_number)
            if not isinstance(check_vmd, tuple):
                return check_vmd

            # Step
            info = self.check_permitted_payee(doc_number)
            if info is True:
                return f'Vendor is excluded from posting. Document {doc_number} cannot be processed.'
            elif info != False: # and "is marked for deletion" in info:
                return f"Document {doc_number} cannot be processed. {info}"
            
            # Step
            info = self.check_fields(doc_number)
            if info is not True:
                return info                    

            #TODO: Step - (check)
            info = self.check_po(doc_number)
            if len(info) != 0:
                return f"PO and invoice have different: {info}"

            # Step
            info = self.check_dates(doc_number, company_code)
            if info is not True:
                return f"Error during setting posting date: {info}"

            # Step
            info = self.check_bank_ids(doc_number, check_vmd[1], process_data[1])
            if info is not True:
                return info

            # Step
            info = self.check_saldo(doc_number)
            if info is not True:
                return info

            # Step
            info = self.check_tax_code(doc_number, check_po_line_entry_data)
            if info is not True:
                return info[1]

            #TODO: Step - (check) check_before_book
            info = self.check_before_book(doc_number)
            if info is not True:
                return info

            #TODO: Step - (check) perform_booking_action
            info = self.perform_booking_action(doc_number)
            if info != '':
                postingNumber = self.get_posting_number(doc_number)
                return postingNumber
            else:
                return info              

        except Exception as e:
            return str(e)
    
    def get_vendors(self, manual_trigger):
        try:
            if manual_trigger:
                df_vendors = pd.read_excel('vendor matrix for RPA.xlsx')
            else:
                df = pd.read_excel('vendor matrix for RPA.xlsx')
                df_vendors = [
                    ('3B5', df[df['Company code'] == '3B5']),
                    ('V436', df[df['Company code'] == 'V436'])
                ]

            return df_vendors
        
        except Exception as e:
            log(f"Error in function 'get_vendors': {str(e)}")
            return False

    def kill_sap(self):
        from win32com.client import GetObject
        import os
        try:
            WMI = GetObject('winmgmts:')
            for p in WMI.ExecQuery('select * from Win32_Process where Name LIKE "%saplogon%"'):
                os.system("taskkill /F /T /pid " + str(p.ProcessId))
                log("SAP was killed")

        except Exception as e:
            log(e, lte.error)
            raise Exception("Error occured in kill_sap")
        
    def sap_exists(self):
        from win32com.client import GetObject
        try:
            WMI = GetObject('winmgmts:')
            for p in WMI.ExecQuery('select * from Win32_Process where Name LIKE "%saplogon%"'):
                return True
            
            return False
        
        except Exception as e:
            log(e, lte.error)
            raise Exception("Error occured in sap_exists")
    
    def logout(self):
        try:
            self.gui_connection.CloseConnection()
            self.gui_connection = None
            self.gui_session = None

            return True
        
        except Exception as e:
            log(str(e))
            return False
        
    def new_session(self):
        try:
            # Get the current connection from the session
            connection = self.gui_session.Parent
            sessions_nr = connection.Children.count
            # Check how many sessions are already open in this connection
            if sessions_nr < 6:
                # Open a new session if the maximum session limit is not reached
                self.gui_session.createSession()
                time.sleep(1)
                new_gui_session = connection.Children(sessions_nr)
            else:
                log("Max SAP sessions limit reached.")
                return False

            return new_gui_session
        
        except Exception as e:
            log(str(e))
            return False
        
    def close_additional_session(self, session_to_close):
        try:
            session_nr = str(session_to_close.info.SessionNumber - 1)
            session_id = str(f"/app/con[0]/ses[{session_nr}]")
            self.gui_connection.CloseSession(session_id)
            session_to_close = None

            return True
        
        except Exception as e:
            log(str(e))
            return False
    
    def back_to_cockpit(self):
        try:
            self.gui_session.findById("wnd[0]").sendVKey(12)
            try:
                self.gui_session.findById("wnd[1]/usr/btnBUTTON_1").press()
            except:
                pass
        except Exception as e:
            info = self.gui_session.findById("wnd[0]/sbar").Text
            log(e, lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            raise Exception(f"Error occured in function back_to_cockpit. SAP info: {info}")
    
    def setup_cockpit(self, variant_cockpit, layout=None, company_code=None, df_vendors=pd.DataFrame, company_code2=None):
        """
        The setup_cockpit method is designed to set up the cockpit for the SAP process.

        with LAYOUT - This method performs the following steps:
        1. Select a proper variant
        2. Enter layout name
        3. Run transaction.

        Note: This method does not return any value.
        """
        try:
            self.open_transaction("/n/cockpit/1")
            self.gui_session.findById("wnd[0]/tbar[1]/btn[17]").press()
            self.gui_session.findById("wnd[1]/usr/txtV-LOW").text = variant_cockpit
            self.gui_session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
            self.gui_session.findById("wnd[1]/tbar[0]/btn[8]").press()
            
            # Set the layout value.
            if layout is not None and layout != '':
                self.gui_session.findById("wnd[0]/usr/ctxtP_VAR_H").text = layout

            # Open the selection screen for company codes.
            self.gui_session.findById("wnd[0]/usr/btn%_SEL_BUKR_%_APP_%-VALU_PUSH").press()
            # Set values for the selection fields.
            self.gui_session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = company_code # "3b5"
            if company_code2 is not None and company_code2 != '':
                self.gui_session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = company_code2 # "V436"
            # Press the 'Enter' key to confirm the selection.
            self.gui_session.findById("wnd[1]").sendVKey(8)        

            if self.gui_session.info.systemname == "PCE":
                self.gui_session.findById("wnd[0]/usr/ctxtP_VAR_H").text = "/MONIKA"
                self.gui_session.findById("wnd[0]/usr/ctxtP_VAR_P").text = ""
                self.gui_session.findById("wnd[0]/usr/ctxtP_VAR_A").text = "/VISHAL S/4"
            
            # import_whitelist
            if not df_vendors.empty:
                df_vendors.to_clipboard(index=False, header=False)
                self.gui_session.findById("wnd[0]/usr/btn%_SEL_VEND_%_APP_%-VALU_PUSH").press()
                self.gui_session.findById("wnd[1]/tbar[0]/btn[16]").press()
                self.gui_session.findById("wnd[1]/tbar[0]/btn[24]").press()
                self.gui_session.findById("wnd[1]/tbar[0]/btn[8]").press()

            # # Press the 'Enter' key to confirm and execute.
            self.gui_session.findById("wnd[0]").sendVKey(8)

            # Set the first visible column.
            # self.gui_session.findById("wnd[0]/usr/cntlTOP_CONTAINER/shellcont/shell").firstVisibleColumn = "FI_MM_FLG"

            return True

        except Exception as e:
            log(e, lte.error)

    def choose_layout(self, layout_name):
        """
        This function is used to select a given layout in the system.

        Parameters:
        None

        Returns:
        Bool: True if the layout was successfully selected, False otherwise.
        """
        try:
            self.gui_session.findById("wnd[0]").sendVKey(33)
            dane_SAP = self.gui_session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell")
            a = 0
            # Loop until the variant is found
            while True:
                dane_SAP.firstVisibleRow = a
                variant = dane_SAP.getCellValue(a, "VARIANT")
                if variant == layout_name or variant == "/" + layout_name:
                    dane_SAP.currentCellRow = a
                    dane_SAP.selectedRows = a
                    dane_SAP.doubleClickCurrentCell()
                    return True
                a += 1
        except Exception as e:
            log(f"Error in function 'choose_layout': {str(e)}")
            return False
        
    def get_data_for_process(self, manual_trigger, file_name, temp_path, path_attachments=None, df_processable_docs=pd.DataFrame, company_code=None):
        try:
            if manual_trigger:
                # list of documents provided by client

                self.import_docs_from_file(path_attachments)
                log(f"Documents from {path_attachments} imported")
                self.generate_export_file(temp_path, file_name)
                log(f"Export file {file_name} generated")
                df = pd.read_excel(os.path.join(temp_path, file_name))

                return df
            
            else:                    
                self.exclude_credit_notes_prepare_kpi()
                log("credit notes excluded")
                file_name_comp_code = f'Export{company_code}.xlsx'
                self.generate_export_file(temp_path, file_name_comp_code)
                log(f"Export file {file_name_comp_code} generated")
                df = pd.read_excel(os.path.join(temp_path, file_name_comp_code))
                df = df.where(df.notna(), '')
                if not df.empty:
                    df_processable_docs = pd.concat([df_processable_docs, df]) if not df_processable_docs.empty else df

                return df_processable_docs
        
        except Exception as e:
            log(f"Error in function 'get_data_for_process': {str(e)}")
            return False

    def exclude_credit_notes_prepare_kpi(self):
        """
        This function is used to exclude credit notes, search only for invoices and setup workflow status to single values in cockpit.

        Returns:
        bool: True if the operation was successful, False otherwise.
        """
        try:
            self.gui_session.findById("wnd[0]/shellcont").dockerPixelSize = 80
            self.gui_session.findById("wnd[0]/usr/cntlTOP_CONTAINER/shellcont/shell").pressToolbarContextButton("&MB_VARIANT")
            self.gui_session.findById("wnd[0]/usr/cntlTOP_CONTAINER/shellcont/shell").selectContextMenuItem("&COL0")
            self.gui_session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_CONFIGURATION:SAPLSALV_CUL_COLUMN_SELECTION:0620/cntlCONTAINER1_LAYO/shellcont/shell").pressToolbarButton("&FIND")
            self.gui_session.findById("wnd[2]/usr/chkGS_SEARCH-EXACT_WORD").Selected = True
            self.gui_session.findById("wnd[2]/usr/txtGS_SEARCH-VALUE").text = "Invoice"
            self.gui_session.findById("wnd[2]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
            self.gui_session.findById("wnd[2]/tbar[0]/btn[0]").press()
            search_result = self.gui_session.findById("wnd[2]/usr/txtGS_SEARCH-SEARCH_INFO").text
            if 'No hits' in search_result:
                self.gui_session.findById("wnd[2]").Close()
                self.gui_session.findById("wnd[1]/tbar[0]/btn[0]").press()
            else:
                self.gui_session.findById("wnd[2]").Close()
                self.gui_session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_CONFIGURATION:SAPLSALV_CUL_COLUMN_SELECTION:0620/btnAPP_WL_SING").press()
                self.gui_session.findById("wnd[1]/tbar[0]/btn[0]").press()

            self.gui_session.findById("wnd[0]/usr/cntlTOP_CONTAINER/shellcont/shell").setCurrentCell(-1, "INVOICE_IND")
            self.gui_session.findById("wnd[0]/usr/cntlTOP_CONTAINER/shellcont/shell").selectColumn ("INVOICE_IND")
            self.gui_session.findById("wnd[0]/usr/cntlTOP_CONTAINER/shellcont/shell").pressToolbarButton ("&MB_FILTER")
            self.gui_session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = "X"
            self.gui_session.findById("wnd[1]/tbar[0]/btn[0]").press()

            # set workflow status to single values
            self.gui_session.findById("wnd[0]/usr/cntlTOP_CONTAINER/shellcont/shell").setCurrentCell(-1, "WC_ICON")
            self.gui_session.findById("wnd[0]/usr/cntlTOP_CONTAINER/shellcont/shell").selectColumn("WC_ICON")
            self.gui_session.findById("wnd[0]/usr/cntlTOP_CONTAINER/shellcont/shell").pressToolbarButton("&MB_FILTER")
            self.gui_session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN001_%_APP_%-VALU_PUSH").press()
            self.gui_session.findById("wnd[2]").sendVKey(2)
            self.gui_session.findById("wnd[3]/usr/cntlOPTION_CONTAINER/shellcont/shell").currentCellColumn = "TEXT"
            self.gui_session.findById("wnd[3]/usr/cntlOPTION_CONTAINER/shellcont/shell").doubleClickCurrentCell()
            self.gui_session.findById("wnd[2]/tbar[0]/btn[8]").press()
            self.gui_session.findById("wnd[1]/tbar[0]/btn[0]").press()

            return True
        
        except Exception as e:
            log(e, lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            raise Exception(f"Error occured in function exclude_credit_notes_prepare_kpi")
            

    def import_docs_from_file(self, katalog):
        """
        This function is used to import documents from a file and put them into document number selection in cockpit.

        Parameters:
        katalog (str): The path to the folder containing the file.

        Returns:
        bool: True if the operation was successful, False otherwise.
        """
        try:
            # Get the list of files in the folder
            files = glob.glob(os.path.join(katalog, "*.xls*"))
            file_count = len(files)

            if file_count == 0:
                info = "There was no file attached to the request. Please add an Excel file containing list of document numbers to process."
                return False, info
            else:
                for file in files:
                    # Open the Excel file
                    try:
                        df = pd.read_excel(file, header=None, sheet_name=0)
                    except Exception as e:
                        info = f"Could not open file {file}. Error: {str(e)}"
                        return False, info
                    
                    if df.empty or df.iloc[0, 0] == "":
                        info = f"File {file} contains no document numbers on the first sheet in column A."
                        os.remove(file)
                        continue

                    r_all = df.iloc[:, 0]  # Assuming we're only interested in the first column

                    # Interacting with SAP GUI
                    shell = self.gui_session.findById("wnd[0]/usr/cntlTOP_CONTAINER/shellcont/shell")
                    shell.setCurrentCell(-1, "DOCNO")
                    shell.selectColumn("DOCNO")
                    shell.pressToolbarButton("&MB_FILTER")

                    box = self.find_box_dyn("Document Number")

                    self.gui_session.findById(f"wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN00{box}_%_APP_%-VALU_PUSH").press()
                    if file == files[0]:
                        self.gui_session.findById("wnd[2]/tbar[0]/btn[16]").press()

                    # Copying the data to SAP GUI
                    r_all.to_clipboard(index=False, header=False)
                    self.gui_session.findById("wnd[2]/tbar[0]/btn[24]").press()
                    self.gui_session.findById("wnd[2]/tbar[0]/btn[8]").press()
                    self.gui_session.findById("wnd[1]/tbar[0]/btn[0]").press()

            return True

        except Exception as e:
            log(f"Error in procedure 'import_docs_from_file': {str(e)}")
            return False
    
    def find_box_dyn(self, column_name):
        """
        This function is used to find the dynamic box in the system.

        Parameters:
        column_name (str): The name of the column to find

        Returns:
        box_dyn (object): The dynamic box found in the system.
        """
        try:
            a = 1
            while True:
                element_id = f"wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/txt%_%%DYN00{a}_%_APP_%-TEXT"
                try:
                    element_text = self.gui_session.findById(element_id).text
                    if str(element_text).upper() == str(column_name).upper():
                        return a
                except Exception:
                    break
                a += 1
            return None
        except Exception as e:
            log(f"Error in find_box_dyn: {str(e)}")
            return False

    def generate_export_file(self, katalog, file_name):
        """
        This function is used to generate an export file from SAP transaction.

        Parameters:
        katalog (str): The path to the folder where the file will be saved.
        file_name (str): The name of the file to be generated.

        Returns:
        Bool: True if the operation was successful, False otherwise.
        """
        try:
            # Trigger export in SAP GUI
            shell = self.gui_session.findById("wnd[0]/usr/cntlTOP_CONTAINER/shellcont/shell")
            shell.pressToolbarContextButton("&MB_EXPORT")
            shell.selectContextMenuItem("&XXL")
            self.gui_session.findById("wnd[1]/tbar[0]/btn[0]").press()
            self.gui_session.findById("wnd[1]/usr/ctxtDY_PATH").text = str(katalog)
            self.gui_session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = file_name
            self.gui_session.findById("wnd[1]/tbar[0]/btn[11]").press()

            # Wait for the export file to appear
            stop_time = datetime.now() + timedelta(seconds=30)
            export_file = os.path.join(katalog, file_name)

            while datetime.now() < stop_time:
                if os.path.isfile(export_file):
                    self.excel.close_excel()
                    break
                time.sleep(1)
            return True

        except Exception as e:
            log(f"Error in procedure 'generate_export_file': {str(e)}")
            return False
    
    def prepare_process_list(self, katalog, df=pd.DataFrame):
        """
        This function is used to save a list of documents to be processed in an Excel file.

        Parameters:
        katalog (str): The path to the folder where the file will be saved.
        df (DataFrame): The data to be saved in the Excel file.

        Returns:
        file_name (str): The name of the file saved.
        """
        try:
            file_name = "FR_readyToProcess.xlsx"
            report_path = os.path.join(katalog, file_name)
            df.to_excel(report_path, sheet_name='MM_FR_readyToProcess', index=False)
            return file_name
        except Exception as e:
            log(f"Error in function 'prepare_process_list': {str(e)}")
            return False
    
    def get_procesable_documents(self, katalog, df_me2n, dates, variant_cockpit, layout_cockpit):
        """
        This function is used to get a list of documents that can be processed in the system.

        Parameters:
        katalog (str): The path to the folder where the file will be saved.
        df_me2n (DataFrame): The data to be processed.
        dates (list): A list of dates to be used for filtering.
        variant_cockpit (str): The variant to be used in the cockpit.
        layout_cockpit (str): The layout to be used in the cockpit.

        Returns:
        procesable_documents (list): A list of documents that can be processed.
        """
        try:
            self.open_cockpit()
            self.setup_cockpit(variant_cockpit, layout_cockpit)
            df_me2n['Pur. Doc.'].to_clipboard(index=False, header=False)
            self.gui_session.findById("wnd[0]/usr/cntlTOP_CONTAINER/shellcont/shell").clearSelection()
            self.gui_session.findById("wnd[0]/usr/cntlTOP_CONTAINER/shellcont/shell").selectColumn("PO_NUMBER")
            self.gui_session.findById("wnd[0]/usr/cntlTOP_CONTAINER/shellcont/shell").pressToolbarButton("&MB_FILTER")
            # set new filter - Purchasing Document
            self.gui_session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN001_%_APP_%-VALU_PUSH").press()
            self.gui_session.findById("wnd[2]/tbar[0]/btn[24]").press()
            self.gui_session.findById("wnd[2]/tbar[0]/btn[0]").press()
            self.gui_session.findById("wnd[2]/tbar[0]/btn[8]").press()
            self.gui_session.findById("wnd[1]/tbar[0]/btn[0]").press()
            self.gui_session.findById("wnd[0]/usr/cntlTOP_CONTAINER/shellcont/shell").clearSelection()
            file_name = "Export2.xlsx"
            self.generate_export_file(katalog, file_name)
            df = pd.read_excel(os.path.join(katalog, file_name))
            df = self.filter_dates(katalog, file_name, dates)
            file_name = self.prepare_process_list(katalog, df)
            return file_name
        except Exception as e:
            log(f"Error in function 'get_procesable_documents': {str(e)}")
            return False

    def filter_dates(self, katalog, file_name, dates):
        """
        This function is used to filter dates in a given Excel file.

        Parameters:
        katalog (str): The directory where the file is located.
        file_name (str): The name of the Excel file to be filtered.
        dates (list): A list of dates to be used for filtering.

        Returns:
        df (DataFrame): A pandas DataFrame after filtering the dates.
        """
        try:
            # Filter the data
            df = pd.read_excel(os.path.join(katalog, file_name))
            if len(dates) == 2:
                df = df[(df['Creation date'].dt.strftime('%Y-%m-%d') != dates[1].strftime('%Y-%m-%d')) & (df['Creation date'].dt.strftime('%Y-%m-%d') != dates[0].strftime('%Y-%m-%d'))]
            else:
                df = df[(df['Creation date'].dt.strftime('%Y-%m-%d') == dates[0].strftime('%Y-%m-%d'))]
            return df
        except Exception as e:
            log(f"Error in procedure 'filter_dates': {str(e)}")
            return False
        
    def find_invoice(self, doc_number, company_code):
        """
        This function is used to find an invoice in the system.

        Parameters:
        doc_number (str): The document number.

        Returns:
        Bool: True if the invoice was successfully opened, False otherwise.
        """
        try:
            self.gui_session = self.gui_connection.children.ElementAt(0)
            self.gui_session.findById("wnd[0]").maximize()
            # filter by document number
            self.gui_session.findById("wnd[0]/usr/cntlTOP_CONTAINER/shellcont/shell").setCurrentCell(-1, "DOCNO")
            self.gui_session.findById("wnd[0]/usr/cntlTOP_CONTAINER/shellcont/shell").selectColumn("DOCNO")
            self.gui_session.findById("wnd[0]/usr/cntlTOP_CONTAINER/shellcont/shell").pressToolbarButton("&MB_FILTER")
            #box = self.find_box_dyn("Document Number")
            self.gui_session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN001_%_APP_%-VALU_PUSH").press()
            self.gui_session.findById("wnd[2]").sendVKey(16)
            self.gui_session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = doc_number # '01010101'
            time.sleep(1)
            self.gui_session.findById("wnd[2]/tbar[0]/btn[8]").press()
            self.gui_session.findById("wnd[1]").sendVKey(0)
            # filter by company code
            self.gui_session.findById("wnd[0]/usr/cntlTOP_CONTAINER/shellcont/shell").setCurrentCell(-1, "COMP_CODE")
            self.gui_session.findById("wnd[0]/usr/cntlTOP_CONTAINER/shellcont/shell").selectColumn("COMP_CODE")
            self.gui_session.findById("wnd[0]/usr/cntlTOP_CONTAINER/shellcont/shell").pressToolbarButton("&MB_FILTER")
            self.gui_session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN002_%_APP_%-VALU_PUSH").press()
            self.gui_session.findById("wnd[2]").sendVKey(16)
            self.gui_session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = company_code
            time.sleep(1)
            self.gui_session.findById("wnd[2]/tbar[0]/btn[8]").press()
            self.gui_session.findById("wnd[1]").sendVKey(0)
            
            try:
                self.gui_session.findById("wnd[0]/usr/cntlTOP_CONTAINER/shellcont/shell").selectedRows = "0"
            except:
                return False
            return True
        
        except Exception as e:
            log(f"Error in function find_invoice. Doc number {doc_number}: {e}", lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            return str(f"Error in function find_invoice. Doc number {doc_number}: {e}")

    def get_meta_data(self, doc_number):
        try:
            line = self.line_numb()
            self.gui_session = self.gui_connection.children.ElementAt(0)
            self.gui_session.findById("wnd[0]").maximize()
            line = self.line_numb()
            
            metaData_fields = [
                ('WFDescription', "WC_NAME"),
                ('WFstatus', "WC_ICON"),
                ('DocType_MM_FI', "FI_MM_FLG"),
                ('followUpFlag', "FOLLOW_UP_ICON"),
                ('Company code', "COMP_CODE")
            ]

            updated_metaData_fields = {}
            field_path = 'wnd[0]/usr/cntlTOP_CONTAINER/shellcont/shell'

            for field_name, field_Id in metaData_fields:
                value = self.gui_session.findById(field_path).getCellValue(0, field_Id)
                updated_metaData_fields[field_name] = value

            return updated_metaData_fields
        
        except Exception as e:
            log(f"Error in function get_meta_data. Doc number {doc_number}: {e} Line: {line}", lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            return str(f"Error in function get_meta_data. Doc number {doc_number}: {e}")

    def check_meta_data(self, doc_number, metaData):
        """
        This function is used to check the metadata of a document.

        Parameters:
        doc_number (str): The document number.
        metaData (dict): The metadata of the document.

        Returns:
        Bool: True if the metadata is correct, False otherwise.
        """
        try:
            wf_status = metaData['WFstatus']
            wf_description = metaData['WFDescription']
            wf_doc_type_mm_fi = metaData['DocType_MM_FI']
            # follow_up_flag = metaData['followUpFlag']
            # company_code = metaData['Company code']

            if wf_status != '' and wf_description != '' or wf_doc_type_mm_fi == 'FI':
                return True
            else:
                return False
        
        except Exception as e:
            log(f"Error in function check_meta_data. Doc number {doc_number}: {e}", lte.error)
            return str(f"Error in function check_meta_data. Doc number {doc_number}: {e}")
        
    def open_invoice(self, doc_number):
        """
        This function is used to open an invoice in the system.

        Parameters:
        doc_number (str): The document number.

        Returns:
        Bool: True if the invoice was successfully opened, False otherwise.
        """
        try:
            try:
                self.gui_session.findById("wnd[0]/usr/cntlTOP_CONTAINER/shellcont/shell").selectedRows = "0"
            except:
                return False
            self.gui_session.findById("wnd[0]/tbar[1]/btn[8]").press()
            return True
        
        except Exception as e:
            log(e, lte.error)
            log(f"Error in function open_invoice. Doc number {doc_number}: {e}", lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            return str(f"Error in function open_invoice. Doc number {doc_number}: {e}")
    
    def get_wf_status_po_nr(self, doc_number, metaData):
        """
        This function is used to get the workflow status of a purchase order.

        Parameters:
        doc_number (str): The document number.
        metaData (dict): The metadata of the document.

        Returns:
        str: The workflow status of the purchase order or PO number.
        """
        try:
            given_po = ''
            wf_status = metaData['WFstatus']
            wf_description = metaData['WFDescription']
            wf_doc_type_mm_fi = metaData['DocType_MM_FI']
            if 'accepted' in str(wf_status).lower() and 'provide gr related to invoice' in str(wf_description).lower() or 'provide correct po number' in str(wf_description).lower():
                info = self.download_po_from_note(doc_number)
                if info == False:
                    info = f"Document {doc_number} cannot be processed, due to unexpected exception in processing 'Answered Workflows'(step 4c and 4f in 'Sending WC PDD')"
                    return info
                elif 'cannot be processed' in info:
                    return info

                given_po = info
                if given_po != '' and str(wf_doc_type_mm_fi).upper() == 'FI':
                    self.gui_session.findById("wnd[0]/tbar[1]/btn[17]").press()
                    self.gui_session.findById("wnd[1]/usr/btnBUTTON_1").press()
                    self.gui_session.findById("wnd[1]/usr/ctxtAUFM-EBELN").text = given_po
                    self.gui_session.findById("wnd[1]/tbar[0]/btn[11]").press()
                
                elif given_po == '':
                    screen_id = self.find_screen_id()
                    self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB1").Select()
                    given_po = self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB1/ssubSUB:/COCKPIT/SAPLDISPLAY46:0401/ctxt/COCKPIT/SHDR_DISP-PO_NUMBER").text = ""
                    if given_po == '':
                        info = f"Document {doc_number} cannot be processed, PO number was not found either on the tab 'Notes' or in the PO filled on tab 'General'"
                        return info
                    
            return given_po

        except Exception as e:
            log(e, lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            return str(f"Error in function wf_status_PO_nr. Doc number {doc_number}: {e}")

    def download_po_from_note(self, doc_number):
        """
        This function is used to ...

        Parameters:
        doc_number (str): The document number.

        Returns:
        download_po_from_note: True if the input is a message, False otherwise.
        """
        try:
            given_po = ''
            screen_id = self.find_screen_id()
            self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB6").Select()
            self.gui_session.findById("wnd[0]").sendVKey(31)
            i = 0
            while True:
                if self.is_message_wf(i):
                    text_note = self.gui_session.findById(f"wnd[1]/usr/tbl/COCKPIT/SAPLTEXTGO_TC_TXT_HEADER/txt/COCKPIT/STXTHDR_DISP-TEXT_DESC[1,{i}]").text
                    given_po = self.find_po_in_note(text_note, doc_number)
                    if given_po == "ZRM00":
                        self.gui_session.findById("wnd[1]").Close()
                        info = f"Document {doc_number} cannot be processed. ZRM purchase order type is not supported by this robot"
                        return info
                    elif given_po != '' and given_po != False:
                        self.gui_session.findById("wnd[1]").Close()
                        break # return given_po
                else:
                    break
                i += 1
            if self.gui_session.findById("wnd[1]", False):
                self.gui_session.findById("wnd[1]").Close()

            return given_po
        
        except Exception as e:
            log(e, lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            return str(f"Error in function download_po_from_note. Doc number {doc_number}: {e}")

    def is_message_wf(self, i):
        """
        This function is used to check if a given input is a message.

        Parameters:
        i (int): row number.

        Returns:
        is_message_wf (bool): True if the input is a message, False otherwise.
        """
        try:
            try:
                text_note = self.gui_session.findById(f"wnd[1]/usr/tbl/COCKPIT/SAPLTEXTGO_TC_TXT_HEADER/txt/COCKPIT/STXTHDR_DISP-TEXT_DESC[1,{i}]").text
                return True
            except:
                return False
            
        except Exception as e:
            info = self.gui_session.findById("wnd[0]/sbar").Text
            log(e, lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            return str(f"Error occured in function is_message_wf. SAP info: {info}")
    
    def find_po_in_note(self, text_note, doc_number):
        """
        This function is used to check if given text_note include PO number.

        Parameters:
        text_note (str): text note.

        Returns:
        find_po_in_note: adjusted_PO number or "ZRM00" if it's included in given text_note, False otherwise.
        """
        try:
            for a in range(len(text_note) - 1):
                substring_2 = text_note[a:a+2]
                substring_5 = text_note[a:a+5]
                if substring_2 in {"43", "45", "47", "40"}:
                    adjusted_PO = self.check_if_is_numeric(text_note, a)
                    if adjusted_PO.isdigit() and len(adjusted_PO) == 10:
                        return adjusted_PO

                if substring_5 == "ZRM00":
                    # info = 'ZRM purchase order type is not supported by this robot'
                    return "ZRM00"

            return False
        
        except Exception as e:
            log(e, lte.error)
            return str(f"Error occured in function is_message_wf. Doc number {doc_number}")
    
    def check_if_is_numeric(self, text_note, a):
        """
        This function is used to check if given text_note include PO number.

        Parameters:
        text_note (str): text note.

        Returns:
        check_if_is_numeric: adjusted_PO number, False otherwise.
        """
        try:
            temp = text_note[a:a+10]
            if temp.isdigit():
                return temp

            last_digit_idx = max(i for i in range(len(temp)) if temp[i].isdigit())
            adjusted_PO = temp[:2]

            for c in range(2, last_digit_idx + 1):
                # Replace "-" with "0"s to maintain a valid PO length
                if temp[c] == "-":
                    adjusted_PO += "0" * (10 - last_digit_idx)
                else:
                    adjusted_PO += temp[c]
            
            return adjusted_PO if adjusted_PO.isdigit() else False

        except Exception as e:
            info = self.gui_session.findById("wnd[0]/sbar").Text
            raise Exception(f"Error occured in function is_message_wf. SAP info: {info}")
    
    def wf_status_invoice_price_difference(self, doc_number, metaData):
        """
        This function is used to get the workflow status of a transaction with invoice price difference.

        Parameters:
        doc_number (str): The document number.
        metaData (dict): The metadata of the document.

        Returns:
        str: The workflow status of the transaction with invoice price difference.
        """
        try:
            wf_status = metaData['WFstatus']
            screen_id = self.find_screen_id()
            if 'accepted' in str(wf_status).lower() and 'transp.inv. price diff' in str(wf_status).lower():
                saldo = self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_SALDO:/COCKPIT/SAPLDISPLAY46:0440/txt/COCKPIT/SDUMMY-SALDO").text
                if self.convert_to_number(saldo) == 0.00: 
                    return "go to final steps"

            return False
        
        except Exception as e:
            log(e, lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            return str(f"Error in function wf_status_invoice_price_difference. Doc number {doc_number}: {e}")

    def check_document_source(self, doc_number):
        try:
            screen_id = self.find_screen_id()
            self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB5").Select()
            barcode = self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB5/ssubSUB:/COCKPIT/SAPLDISPLAY46:0436/ssubSUB_OTHERS:/COCKPIT/SAPLDISPLAY46:0700/txt/COCKPIT/SDYN_SUBSCR_0700-VALUE3").text
            userID = self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB5/ssubSUB:/COCKPIT/SAPLDISPLAY46:0436/ssubSUB_OTHERS:/COCKPIT/SAPLDISPLAY46:0700/txt/COCKPIT/SDYN_SUBSCR_0700-VALUE4").text

            documentSource = "Ariba" if not barcode and not userID else "PDFCollector" if barcode and userID else False

            self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB1").select()
            if documentSource == False:
                info = f"Document {doc_number} cannot be processed. Document could not be assigned neither to Ariba nor to PDF Collector category"
            else:
                info = documentSource           

            return info

        except Exception as e:
            log(e, lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            return str(f"Error in function check_document_source. Doc number {doc_number}: {e}")

    def process_ariba(self, doc_number, given_po):
        try:
            type = 1
            info = self.process_standard_po(doc_number, type, given_po)

            if 'ERROR' in str(info).upper() or 'CANNOT BE PROCESSED' in str(info).upper():
                return info
            elif info == False:
                return f"Document {doc_number}, cannot be processed. Document could not be assigned neither to Ariba nor to PDF Collector category"

            if info[1].get('two_way_match') == True:
                if not self.take_over_document(doc_number):
                    return f"Document {doc_number} cannot be processed. Error during taking over the document"
                screen_id = self.find_screen_id()
                saldo = self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_SALDO:/COCKPIT/SAPLDISPLAY46:0440/txt/COCKPIT/SDUMMY-SALDO").text
                self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0381/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB2").Select()
                net = self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB2/ssubSUB:/COCKPIT/SAPLDISPLAY46:0402/txt/COCKPIT/SHDR_DISP-NET_AMOUNT").text

                if self.convert_to_number(saldo) != 0.00:
                    total_ordered = info[1].get('po_totals')[0].get('val_ord')
                    
                    check_info = self.check_tolerance(doc_number, saldo, total_ordered, ToleranceRange.lowTolerance)
                    if check_info != True:
                        return check_info
                    
                    check_info = self.add_balance_to_first_line(doc_number, saldo)
                    if check_info != True:
                        return check_info
                    info = "go to final steps"
                return info # go to final step

            return info
            
        except Exception as e:
            log(e, lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            return str(f"Error in function process_ariba. Doc number {doc_number}: {e}")

    def process_pdf_collector(self, doc_number, given_po):
        try:
            if not self.take_over_document(doc_number):
                return f"Document {doc_number} cannot be processed. Error during taking over the document"

            type = 2
            info = self.process_standard_po(doc_number, type, given_po)

            if 'ERROR' in str(info).upper() or 'CANNOT BE PROCESSED' in str(info).upper():
                return info
            elif info == False:
                return f"Document {doc_number}, cannot be processed. Document could not be assigned neither to PDF Collector nor to Ariba category"

            return info
            
        except Exception as e:
            log(f"Error in function process_pdf_collector. Doc number {doc_number}: {e}", lte.error)
            return str(f"Error in function process_pdf_collector. Doc number {doc_number}: {e}")
        
    def process_standard_po(self, doc_number, type, given_po):
        try:
            tax_code = ''
            two_way_match = None
            gr_based = None
            po_line_details = []
            po_totals = []

            screen_id = self.find_screen_id(type=6, tab=1)
            po_number = self.check_given_po(doc_number, given_po)
            self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB1/ssubSUB:/COCKPIT/SAPLDISPLAY46:0401/ctxt/COCKPIT/SHDR_DISP-PO_NUMBER").SetFocus()
            self.gui_session.findById("wnd[0]").sendVKey(2)

            screen_id = self.find_screen_id(type=1)          
            po_line_details = self.get_po_line_details(doc_number, screen_id)
            
            # Ariba
            if type in {1, 3}:    # GR_based - TRUE, 
                gr_based = self.get_gr_based(doc_number)
                if not isinstance(gr_based, bool):
                    return gr_based
                two_way_match = False if gr_based else True

            # PDF Collector
            if type in {2, 3}:    # GR_based - FALSE, get_totals - TRUE
                tax_code = self.get_tax_code(doc_number)
                if 'cannot be processed' in tax_code:
                    return tax_code
                
            screen_id = self.find_screen_id(type=2, tab=9)
            po_totals = self.get_po_totals(doc_number)
            if po_number[:2] != '40':
                po_creator = self.get_po_creator(doc_number, po_number)
            else:
                po_creator = 'Missing'
            self.gui_session.findById("wnd[0]").sendVKey(3)

            vendor = self.get_vendor(doc_number)
            if vendor == '':
                return f"Document {doc_number}, cannot be processed. No Vendor account was assigned"
            netto = self.get_netto(doc_number)
            screen_id = self.find_screen_id()
            saldo = self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_SALDO:/COCKPIT/SAPLDISPLAY46:0440/txt/COCKPIT/SDUMMY-SALDO").text

            if type in {1, 2, 3}:  # GR_based - TRUE or get_totals - TRUE
                process_standard_result = {
                    'two_way_match': two_way_match,
                    'gr_based': gr_based,
                    'tax_code': tax_code,
                    'po_number': po_number,
                    'vendor': vendor,
                    'po_creator': po_creator,
                    'netto': netto,
                    'saldo': saldo,
                    'po_fully_booked': po_totals[0],
                    'po_totals': po_totals[1],
                    'po_line_details_many_lines': po_line_details[0],
                    'po_line_details': po_line_details[1]
                }
                return True, process_standard_result

            return False

        except Exception as e:
            log(e, lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            return str(f"Error in function process_standard_po. Doc number {doc_number}: {e}")

    def check_given_po(self, doc_number, given_po):
        """
        This function is used to enter given PO to system if founded.

        Parameters:
        None

        Returns:
        bool: True if given po entered, False otherwise.
        """
        try:
            screen_id = self.find_screen_id()
            if given_po != "":
                self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB1/ssubSUB:/COCKPIT/SAPLDISPLAY46:0401/ctxt/COCKPIT/SHDR_DISP-PO_NUMBER").text = given_po
            po_number = self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB1/ssubSUB:/COCKPIT/SAPLDISPLAY46:0401/ctxt/COCKPIT/SHDR_DISP-PO_NUMBER").text

            return po_number

        except Exception as e:
            log(f"Error in function check_given_po. Doc number {doc_number}: {e}", lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            raise Exception(f"Error in function check_given_po. Doc number {doc_number}: {e}")
        
    def get_gr_based(self, doc_number):
        try: 
            gr_based = ''
            screen_id = self.find_screen_id(type=4)
            if screen_id is None:
                middle_path = "subSUB3:SAPLMEVIEWS:1100/subSUB1:SAPLMEVIEWS:4002/btnDYN_4000-BUTTON"
                screen_id_button = self.find_screen_id(type=12, middle_path_id=middle_path)
                self.gui_session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{screen_id_button}/subSUB3:SAPLMEVIEWS:1100/subSUB1:SAPLMEVIEWS:4002/btnDYN_4000-BUTTON").press()
            
            screen_id = self.find_screen_id(type=4)
            if screen_id is None:
                screen_id = self.find_screen_id(type=13, tab=8)
                screen_id = self.find_screen_id(type=4)
            path = f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{screen_id}/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/chkMEPO1317-WEBRE"
            self.gui_session.findById(path).SetFocus()
            gr_based = self.gui_session.findById(path).Selected

            return gr_based

        except:
            self.gui_session.findById("wnd[0]").sendVKey(3)
            return f"Document {doc_number}, cannot be processed. [GR Based IV] field is not available in PO on tab Invoice"
        
    def get_tax_code(self, doc_number):
        try:
            tax_code = ''
            screen_id = self.find_screen_id(type=5)  
            if screen_id is None:
                middle_path = "subSUB3:SAPLMEVIEWS:1100/subSUB1:SAPLMEVIEWS:4002/btnDYN_4000-BUTTON"
                screen_id_button = self.find_screen_id(type=12, middle_path_id=middle_path)
                self.gui_session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{screen_id_button}/subSUB3:SAPLMEVIEWS:1100/subSUB1:SAPLMEVIEWS:4002/btnDYN_4000-BUTTON").press()
            
            screen_id = self.find_screen_id(type=5)  
            if screen_id is None:
                screen_id = self.find_screen_id(type=13, tab=8)
                screen_id = self.find_screen_id(type=5)  
            path = f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{screen_id}/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ"
            tax_code = self.gui_session.findById(path).text

            return tax_code

        except:
            return f"Document {doc_number}, cannot be processed. [Tax Code] field is not available in PO on tab Invoice. Tax Code cannot be downloaded"
    
    def get_vendor(self, doc_number):
        try:
            vendor = ''
            screen_id = self.find_screen_id(type=6, tab=2)
            path = f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB2/ssubSUB:/COCKPIT/SAPLDISPLAY46:0402/ctxt/COCKPIT/SHDR_DISP-VENDOR_NO"
            if self.gui_session.findById(path, False) is not None:
                vendor = self.gui_session.findById(path).text

            return vendor

        except Exception as e:
            log(e, lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            return str(f"Error in function get_vendor. Doc number {doc_number}: {e}")
        
    def get_netto(self, doc_number): 
        try:
            netto = ''
            screen_id = self.find_screen_id(type=6, tab=2)
            path = f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB2/ssubSUB:/COCKPIT/SAPLDISPLAY46:0402/txt/COCKPIT/SHDR_DISP-NET_AMOUNT"
            if self.gui_session.findById(path, False) is not None:
                netto = self.gui_session.findById(path).text

            return netto

        except Exception as e:
            log(e, lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            return str(f"Error in function get_netto. Doc number {doc_number}: {e}")

    def get_po_line_details(self, doc_number, screen_id):
        """
        This function is used to get PO line values.

        Parameters:
        doc_number (str): The document number.
        screen_id (str): The screen id.

        Returns:
        list: nested list with PO Lines values.
        """
        try:

            po_line_details = []
            b, c = 0, 0
            path = f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{screen_id}/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-EBELP[1,0]"
            if self.gui_session.findById(path, False) is None:
                middle_path = "subSUB2:SAPLMEVIEWS:1100/subSUB1:SAPLMEVIEWS:4001/btnDYN_4000-BUTTON"
                screen_id_button = self.find_screen_id(type=12, middle_path_id=middle_path)
                self.gui_session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{screen_id_button}/subSUB2:SAPLMEVIEWS:1100/subSUB1:SAPLMEVIEWS:4001/btnDYN_4000-BUTTON").press()

            screen_id = self.find_screen_id(type=1)
            how_many_rows = self.gui_session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{screen_id}/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/").RowCount
            how_many_rows = how_many_rows - 1
            visible_row_count = self.gui_session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{screen_id}/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/").VisibleRowCount
            scroll_row = 1 # how_many_rows % visible_row_count
            while True:
                po_line_item = self.gui_session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{screen_id}/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-EBELP[1,{c}]").text

                po_line_qty = self.gui_session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{screen_id}/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,{c}]").text
                po_line_net_price = self.gui_session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{screen_id}/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-NETPR[10,{c}]").text
                po_line_order_unit = self.gui_session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{screen_id}/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-MEINS[7,{c}]").text
                po_line_price_unit = self.gui_session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{screen_id}/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-PEINH[12,{c}]").text

                po_line_details.append({'item': po_line_item, 'qty': po_line_qty, 'net_price': po_line_net_price, 'order_unit': po_line_order_unit, 'price_unit': po_line_price_unit})
                b += 1
                
                po_line_item = self.gui_session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{screen_id}/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-EBELP[1,{c + 1}]").text
                if po_line_item == '':
                    break
                c += 1

                if c == scroll_row:
                    # Adjust scrollbar position
                    scroll_path = f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{screen_id}/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211"
                    self.gui_session.findById(scroll_path).verticalScrollbar.position = b
                    c = 0
            
            more_than_one_line = False
            if len(po_line_details) > 1:
                more_than_one_line = True

            return more_than_one_line, po_line_details

        except Exception as e:
            log(e, lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            return str(f"Error in function get_po_line_details. Doc number {doc_number}: {e}")

    def get_po_totals(self, doc_number):
        """
        This function is used to get PO totals (values and quantity).

        Parameters:
        doc_number (str): The document number.
        screen_id (str): The screen id.

        Returns:
        list: nested list with PO totals.
        """
        try:
            po_totals = []
            booked = False
            mepo_field_names = ['VALUE0', 'QUANTITY0']
            a = 0

            # 4 (quantity) and 5 (value) 
            for txtMEPO_nr in range(5, 3, -1):
                po_values = []

                # for iterating over different UI elements for each section (CUM_1, CUM_2)
                for i in range(1, 3):
                    field_nr = 1
                    txtMEPO_path = f"123{txtMEPO_nr}/txtMEPO123{txtMEPO_nr}-{mepo_field_names[a]}{field_nr}"
                    screen_id = self.find_screen_id(type=2, tab=9)
                    path = f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{screen_id}/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT9/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1232/ssubHEADER_CUM_{i}:SAPLMEGUI:{txtMEPO_path}"
                    if self.gui_session.findById(path, False) is not None:
                        break

                for field_nr in range(1, 5):
                    txtMEPO_path = f"123{txtMEPO_nr}/txtMEPO123{txtMEPO_nr}-{mepo_field_names[a]}{field_nr}"
                    screen_id = self.find_screen_id(type=2, tab=9)
                    field_path = f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{screen_id}/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT9/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1232/ssubHEADER_CUM_{i}:SAPLMEGUI:{txtMEPO_path}"
                
                    if self.gui_session.findById(field_path, False) is not None:
                        field_value = str(self.gui_session.findById(field_path).text).strip()
                        po_values.append(field_value)
                    else:
                        po_values.append("0")
                        # if not found or error occurs
                        if txtMEPO_nr == 4 and i == 2:
                            booked = False
                    
                if txtMEPO_nr == 4:
                    qty_ord = self.convert_to_number(po_values[0])
                    qty_del = self.convert_to_number(po_values[1])
                    qty_to_del = self.convert_to_number(po_values[2])
                    qty_inv = self.convert_to_number(po_values[3])

                    if qty_ord > 0 and qty_del <= qty_inv and qty_to_del == 0:
                        booked = True

                    po_totals.append({'qty_ord': qty_ord, 'qty_del': qty_del, 'qty_to_del': qty_to_del, 'qty_inv': qty_inv})
                else:
                    val_ord = self.convert_to_number(po_values[0])
                    val_del = self.convert_to_number(po_values[1])
                    val_to_del = self.convert_to_number(po_values[2])
                    val_inv = self.convert_to_number(po_values[3])
                    po_totals.append({'val_ord': val_ord, 'val_del': val_del, 'val_to_del': val_to_del, 'val_inv': val_inv})

                a += 1
            
            return booked, po_totals

        except Exception as e:
            log(e, lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            return str(f"Error in function get_po_totals. Doc number {doc_number}: {e}")

    def convert_to_number(self, value):
        try:
            if isinstance(value, str):
                value = str(value).strip()
            # Try to convert to integer
            return int(value)
        except ValueError:
            try:
                if value.count('.') > 0 and value.count(',') == 1:
                    value = value.replace('.', '')
                value = value.replace(',', '.')
                # If conversion to int fails, try to convert to float
                return float(value)
            except ValueError:
                return value

    def get_po_creator(self, doc_number, given_po):
        try:
            if given_po[:2] == "40":
                return "Missing"
            
            # find and select PO tab 14
            screen_id = self.find_screen_id(type=2, tab=14)
            # find screen id of sap element if exists
            screen_id = self.find_screen_id(type=3)
            try:
                POcreator = self.gui_session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{screen_id}/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT14/ssubTABSTRIPCONTROL2SUB:SAPLMEDCMV:0100/cntlDCMGRIDCONTROL1/shellcont/shell").getCellValue(0, "ERNAM")
            except: 
                return "Missing"
            
            return POcreator

        except Exception as e:
            log(f"Error in function get_po_creator. Doc number {doc_number}: {e}", lte.error)
            return str(f"Error in function get_po_creator. Doc number {doc_number}: {e}")

    def check_tolerance(self, doc_number, saldo, total_ordered, tolerance_type):
        try:
            procent = 0
            kwota = 0
            if saldo.endswith('-'):
                saldo = '-' + saldo[:-1]
            saldo = self.convert_to_number(str(saldo).strip())
            total_ordered = self.convert_to_number(str(total_ordered).strip())
            
            # Set tolerance values based on type
            if tolerance_type == ToleranceRange.lowTolerance:
                procent = 0.1
                kwota = 25
            elif tolerance_type == ToleranceRange.highTolerance:
                procent = 0.2
                kwota = 100
            elif tolerance_type == ToleranceRange.extremeTolerance:
                procent = 10000
                kwota = 20000000

            # Check tolerance conditions
            try:
                if abs(saldo / total_ordered) <= procent and abs(saldo) <= kwota:
                    return True
                else:
                    info = (f"Document {doc_number} cannot be processed. The tolerance of '{procent * 100}'% / {kwota} EUR was exceeded")
                    return info
            except:
                info = (f"Document {doc_number} cannot be processed. Check tolerance FAILED: {saldo} / {total_ordered} [saldo] / [total_ordered]")
                return info

        except Exception as e:
            log(f"Error in function check_tolerance. Doc number {doc_number}: {e}", lte.error)
            return str(f"Error in function check_tolerance. Doc number {doc_number}: {e}")
        
    def add_balance_to_first_line(self, doc_number, saldo):
        try:
            """
            This function adds or subtracts a saldo to the first line in a SAP session.

            Parameters:
            saldo (float): The saldo to add/subtract
            doc_number (str): SAP document number

            Returns:
            bool: True if operation was successful, False otherwise
            """
            kwota = 0 
            saldo2 = self.convert_to_number(saldo)

            if saldo2 > 0:
                info = self.add_balance(doc_number, saldo)
            else:
                if kwota < abs(saldo2):
                    info = f"Document {doc_number} cannot be processed. saldo of {saldo} cannot be substracted from the first line {kwota}"
                    return info
                else:
                    info = self.add_balance(doc_number, saldo)

            self.gui_session.findById("wnd[0]").sendVKey(0)

            return info

        except Exception as e:
            log(f"Document {doc_number} cannot be processed. 'add_balance_to_first_line', {e}", lte.error)
            return str(f"Document {doc_number} cannot be processed. 'add_balance_to_first_line', {e}")

    def add_balance(self, doc_number, saldo):
        try:
            try:
                saldo = self.convert_to_number(str(saldo).strip())
                kwotaLinii = self.convert_to_number(self.gui_session.findById("wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0381/subSUB_ITEM:/COCKPIT/SAPLDISPLAY46:0410/tbl/COCKPIT/SAPLDISPLAY46G_TC_ITEM_DET/txt/COCKPIT/SITEM_DISP-ITEM_AMOUNT[4,0]").text)
                self.gui_session.findById("wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0381/subSUB_ITEM:/COCKPIT/SAPLDISPLAY46:0410/tbl/COCKPIT/SAPLDISPLAY46G_TC_ITEM_DET/txt/COCKPIT/SITEM_DISP-ITEM_AMOUNT[4,0]").text = kwotaLinii + saldo
                return True
            except:
                self.gui_session.findById("wnd[0]/tbar[1]/btn[25]").press()
                self.gui_session.findById("wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0381/subSUB_ITEM:/COCKPIT/SAPLDISPLAY46:0410/tbl/COCKPIT/SAPLDISPLAY46G_TC_ITEM_DET/txt/COCKPIT/SITEM_DISP-ITEM_AMOUNT[4,0]").text = kwotaLinii + saldo
                return True
            
        except Exception as e:
            log(f"Document {doc_number} cannot be processed. 'add_balance', {e}", lte.error)
            return str(f"Document {doc_number} cannot be processed. 'add_balance', {e}")

    def enter_data(self, doc_number, process_standard_result): 
        try:
            po_number = process_standard_result.get('po_number')
            total_delivered_po = process_standard_result.get('po_totals')[0].get('val_del')
            still_to_deliver = process_standard_result.get('po_totals')[0].get('val_to_del')
            total_ordered_po = process_standard_result.get('po_totals')[0].get('val_ord')
            two_way_match = process_standard_result.get('two_way_match')
            netto = process_standard_result.get('netto')
            tax_code = process_standard_result.get('tax_code')
            po_line_details = process_standard_result.get('po_line_details')

            how_many_lines = self.count_accounting_lines()

            if process_standard_result.get('po_line_details_many_lines') == False:
                if total_delivered_po == 0 or (total_delivered_po > 0 and still_to_deliver == 0) or (total_delivered_po > 0 and still_to_deliver > 0 and two_way_match == True):
                    if netto == total_ordered_po:
                        screen_id = self.find_screen_id(type=8)
                        self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/{mid_path}ctxt/COCKPIT/SITEM_DISP-PO_NUMBER[2,0]").text = po_number
                        self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/{mid_path}ctxt/COCKPIT/SITEM_DISP-PO_ITEM[3,0]").text = item.get('item')
                        item_amount = round(
                            self.convert_to_number(item.get('net_price')) * self.convert_to_number(item.get('qty')) / self.convert_to_number(item.get('price_unit')), 2
                        )
                        self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/{mid_path}txt/COCKPIT/SITEM_DISP-ITEM_AMOUNT[4,0]").text = item_amount
                        self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/{mid_path}txt/COCKPIT/SITEM_DISP-QUANTITY[6,0]").text = item.get('qty')
                        self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/{mid_path}ctxt/COCKPIT/SITEM_DISP-TAX_CODE[10,0]").text = tax_code
                        self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/{mid_path}ctxt/COCKPIT/SITEM_DISP-PO_UNIT[7,0]").text = item.get('order_unit')
                    else:
                        check_tolerance_result = self.check_tolerance(doc_number, netto, total_ordered_po, ToleranceRange.highTolerance)
                        if check_tolerance_result == True:
                            screen_id = self.find_screen_id(type=8)
                            self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/{mid_path}ctxt/COCKPIT/SITEM_DISP-PO_NUMBER[2,0]").text = po_number
                            self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/{mid_path}ctxt/COCKPIT/SITEM_DISP-PO_ITEM[3,0]").text = item.get('item')
                            self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/{mid_path}txt/COCKPIT/SITEM_DISP-ITEM_AMOUNT[4,0]").text = netto
                            self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/{mid_path}txt/COCKPIT/SITEM_DISP-QUANTITY[6,0]").text = item.get('qty')
                            self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/{mid_path}ctxt/COCKPIT/SITEM_DISP-TAX_CODE[10,0]").text = tax_code
                            self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/{mid_path}ctxt/COCKPIT/SITEM_DISP-PO_UNIT[7,0]").text = item.get('order_unit')
                        else:
                            return False
                else:
                    return f"Document {doc_number} cannot be processed. PO is a three way match type therefore cannot be booked when still to deliver amount is greater than zero."
            else:
                mid_path = "subSUB_ITEM:/COCKPIT/SAPLDISPLAY46:0410/tbl/COCKPIT/SAPLDISPLAY46G_TC_ITEM_DET/"
                total_items = len(po_line_details)

                if how_many_lines < total_items:
                    screen_id = self.find_screen_id(type=8)
                    items_to_remove = {item['item'] for item in self.get_accounting_lines()}
                    # Filter main_list to exclude any items that are in items_to_remove
                    filtered_main_list = [item for item in po_line_details if item['item'] not in items_to_remove]
                    # add missing item rows
                    for item in filtered_main_list: # po_line_details[:-how_many_lines]:
                        self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_ITEM:/COCKPIT/SAPLDISPLAY46:0410/btnG_TC_ITEM_DET_INSERT").press()
                        self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/{mid_path}ctxt/COCKPIT/SITEM_DISP-PO_ITEM[3,0]").text = item.get('item')
                        self.gui_session.findById("wnd[0]").sendVKey(0)
                    self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_ITEM:/COCKPIT/SAPLDISPLAY46:0410/tbl/COCKPIT/SAPLDISPLAY46G_TC_ITEM_DET").columns.elementAt(3).selected = True
                    self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_ITEM:/COCKPIT/SAPLDISPLAY46:0410/btn%#AUTOTEXT001").press()
                    self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_ITEM:/COCKPIT/SAPLDISPLAY46:0410/tbl/COCKPIT/SAPLDISPLAY46G_TC_ITEM_DET").deselectAllColumns()
                    i = 1
                    for item in po_line_details:
                        self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_ITEM:/COCKPIT/SAPLDISPLAY46:0410/tbl/COCKPIT/SAPLDISPLAY46G_TC_ITEM_DET/txt/COCKPIT/SITEM_DISP-INVOICE_ITEM[1,{i - 1}]").text = i
                        i += 1

                # for index, item in enumerate(reversed(po_line_details)):
                for index, item in enumerate(po_line_details):
                    screen_id = self.find_screen_id(type=8)
                    self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/{mid_path}ctxt/COCKPIT/SITEM_DISP-PO_NUMBER[2,{index}]").text = po_number
                    self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/{mid_path}ctxt/COCKPIT/SITEM_DISP-PO_ITEM[3,{index}]").text = item.get('item')
                    item_amount = round(
                        self.convert_to_number(item.get('net_price')) * self.convert_to_number(item.get('qty')) / self.convert_to_number(item.get('price_unit')), 2
                    )
                    self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/{mid_path}txt/COCKPIT/SITEM_DISP-ITEM_AMOUNT[4,{index}]").text = item_amount
                    self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/{mid_path}txt/COCKPIT/SITEM_DISP-QUANTITY[6,{index}]").text = item.get('qty')
                    self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/{mid_path}ctxt/COCKPIT/SITEM_DISP-TAX_CODE[10,{index}]").text = tax_code
                    self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/{mid_path}ctxt/COCKPIT/SITEM_DISP-PO_UNIT[7,{index}]").text = item.get('order_unit')
                    self.gui_session.findById("wnd[0]").sendVKey(0)
            
            return True

        except Exception as e:
            info = self.gui_session.findById("wnd[0]/sbar").Text
            log(e, lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            raise Exception(f"Error occured in function enter_data. SAP info: {info}")

    def count_accounting_lines(self):
        try:
            screen_id = self.find_screen_id(type=8)
            path_part = f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_ITEM:/COCKPIT/SAPLDISPLAY46:0410/tbl/COCKPIT/SAPLDISPLAY46G_TC_ITEM_DET/ctxt/COCKPIT/SITEM_DISP-PO_NUMBER[2"
            a = 0
            
            while True:
                po_line_item = self.gui_session.findById(f"{path_part},{a}]").text
                if po_line_item == '' or po_line_item == '__________':
                    break
                a += 1

            return a

        except:
            a -= 1

    def get_accounting_lines(self):
        try:
            screen_id = self.find_screen_id(type=8)
            path_part = f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_ITEM:/COCKPIT/SAPLDISPLAY46:0410/tbl/COCKPIT/SAPLDISPLAY46G_TC_ITEM_DET/ctxt/COCKPIT/SITEM_DISP-PO_ITEM[3"
            existing_lines = []
            a = 0
            
            while True:
                po_line_item = self.gui_session.findById(f"{path_part},{a}]").text
                if po_line_item == '' or po_line_item == '_____':
                    break
                existing_lines.append({'item': po_line_item})
                a += 1

            return existing_lines

        except:
            existing_lines
    
    def take_over_document(self, doc_number):
        """
        This function is used to take over a document from cockpit.

        Parameters:
        doc_number (str): The document number.

        Returns:
        Bool: True if the document takeover was successful, error message otherwise.
        """
        try:
            if self.is_document_editable() is True:
                return True
            else:
                self.gui_session.findById("wnd[0]/tbar[1]/btn[25]").press()
                info = str(self.gui_session.findById("wnd[0]/sbar").text)
                if info.find("currently locked by user") != -1:
                    return info
                elif str(self.gui_session.findById("wnd[0]/sbar").messagetype) == "E":
                    info = str(self.gui_session.findById("wnd[0]/sbar").text)
                    return info
                elif info.find("currently locked by user") == -1:
                    if self.confirm_extra_window():
                        info = "[SAP comment] Documents with this status cannot be changed"
                        return info
                    else:
                        if self.gui_session.ActiveWindow.name == "wnd[1]":
                            self.gui_session.findById("wnd[1]/usr/btnBUTTON_1").press()
                        return True
                else:
                    return info
        except Exception as e:
            log(f"Error in function take_over_document. Doc number {doc_number}: {e}", lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            return str(f"Error in function take_over_document. Doc number {doc_number}: {e}")

    def is_document_editable(self):
        """
        This function is used to check if a document in the system is editable.

        Parameters:
        None

        Returns:
        editable (bool): True if the document is editable, False otherwise.
        """
        try:
            screen_id = self.find_screen_id()
            self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB1").Select()
            editable = self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB1/ssubSUB:/COCKPIT/SAPLDISPLAY46:0401/ctxt/COCKPIT/SHDR_DISP-SGTXT").Changeable

            return editable
        
        except Exception as e:
            info = self.gui_session.findById("wnd[0]/sbar").Text
            log(e, lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            raise Exception(f"Error occured in function is_document_editable. SAP info: {info}")
    
    def confirm_extra_window(self):
        """
        This function is used to close an extra window in the system.

        Parameters:
        None

        Returns:
        confirmed (bool): True if an extra window is present, error message otherwise.
        """
        try:
            if self.gui_session.ActiveWindow.name == "wnd[1]" and self.gui_session.ActiveWindow.text == "Information" and (str(self.gui_session.findById("wnd[1]/usr/txtMESSTXT1").text).find("Documents with this status cannot be changed") != -1):
                self.gui_session.findById("wnd[1]/tbar[0]/btn[0]").press()
                return True
            
            if self.gui_session.ActiveWindow.name == "wnd[1]" and self.gui_session.ActiveWindow.text == "Information" and (str(self.gui_session.findById("wnd[1]/usr/txtMESSTXT1").text).find("is marked for deletion") != -1):
                info = self.gui_session.findById("wnd[1]/usr/txtMESSTXT1").text
                self.gui_session.findById("wnd[1]/tbar[0]/btn[0]").press()
                self.gui_session.findById("wnd[0]/tbar[0]/btn[3]").press()
                return info
            
        except Exception as e:
            info = self.gui_session.findById("wnd[0]/sbar").Text
            log(e, lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            raise Exception(f"Error occured in function confirm_extra_window. SAP info: {info}")
        
    def check_doc_type(self, doc_number):
        """
        This function is used to check the document type in the system.

        Parameters:
        doc_number (str): The document number.

        Returns:
        bool: True if the document is an invoice or if subsequent debit change it to invoice, else info message
        """
        try:
            self.gui_session = self.gui_connection.children.ElementAt(0)
            self.gui_session.findById("wnd[0]").maximize()
            screen_id = self.find_screen_id()
            self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB1").Select()
            screen_id = self.find_screen_id()
            documentType = self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB1/ssubSUB:/COCKPIT/SAPLDISPLAY46:0401/cmb/COCKPIT/SHDR_DISP-TRANSACTION").text
            if str(documentType).strip() != 'Invoice':
                if str(documentType).strip() == 'Subsequent Debit':
                    self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB1/ssubSUB:/COCKPIT/SAPLDISPLAY46:0401/cmb/COCKPIT/SHDR_DISP-TRANSACTION").key = "1"
                    return True
                else:
                    info = f"Document {doc_number} cannot be processed. Document is not an invoice, it is {str(documentType).strip()}."
                    return info
                
            return True
            
        except Exception as e:
            log(f"Error in function check_doc_type. Doc number {doc_number}: {e}", lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            return str(f"Error in function check_doc_type. Doc number {doc_number}: {e}")
        
    def check_process_data(self, doc_number, comp_code, process_data, entry_data):
        try:
            info = ''

            po_number = process_data['po_number']
            po_full_booked = process_data['po_fully_booked']
            vendor = process_data['vendor']

            # step 1
            if po_number == '' or str(po_number).upper() == 'FI' or po_full_booked == True:
                multiple_po_exists = False
                entry_data['different_pos'] = True
                entry_data['multiple_only'] = True
                po_lines_result = self.check_po_lines(entry_data)[1]
                multiple_po_exists = po_lines_result['multiple_pos_exist']

                if multiple_po_exists == True:
                    info = f"Document {doc_number} cannot be processed. There are multiple PO numbers and this case was not included in PDD"
                    return info

                if comp_code == 'V436':
                    info = self.check_fi_vendors_v436(doc_number, vendor)
                    if "Cir.Code wasn not found" not in info:
                        return info
                
                info = self.check_vendor_critical(doc_number, vendor, comp_code, po_number, po_full_booked)
                return info
            
            # step 2
            po_type_check_result = self.check_po_type(doc_number)
            if not isinstance(po_type_check_result, tuple):
                if po_type_check_result != False and 'cannot be processed. SAP info' in po_type_check_result:
                    return po_type_check_result

                if po_type_check_result == False:
                    if comp_code == 'V436':
                        info = self.check_fi_vendors_v436(doc_number, vendor)
                        if "Cir.Code wasn not found" not in info:
                            return info
                    
                    info = self.check_vendor_critical(doc_number, vendor, comp_code, po_number, po_full_booked)
                    return info

            else:
                info = self.check_proposal(doc_number)
                if info is not True and 'NO ITEM PROPOSAL COULD BE GENERATED' in str(info).upper():
                    po_type_text = po_type_check_result[1]
                    if po_type_text != POType.Standard:
                        info = f"Document {doc_number} cannot be processed. Purchase Order {po_number} has no Goods Receipt"
                        return info

                    screen_id = self.find_screen_id(type=6, tab=1)
                    self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB1/ssubSUB:/COCKPIT/SAPLDISPLAY46:0401/ctxt/COCKPIT/SHDR_DISP-PO_NUMBER").SetFocus()
                    self.gui_session.findById("wnd[0]").sendVKey(2)
                    gr_based = self.get_gr_based(doc_number)
                    if not isinstance(gr_based, bool):
                        return gr_based
                    two_way_match = False if gr_based else True
                    self.gui_session.findById("wnd[0]").sendVKey(3)

                    if two_way_match == True:
                        saldo = process_data['saldo']
                        netto = process_data['netto']
                        more_than_one_line_po = process_data['po_line_details_many_lines']

                        if saldo != 0 and more_than_one_line_po == True:
                            # TODO: check if this is correct
                            info = self.enter_data(doc_number, process_data)
                            self.gui_session.findById("wnd[0]").sendVKey(0)
                            screen_id = self.find_screen_id()
                            saldo = self.convert_to_number(self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_SALDO:/COCKPIT/SAPLDISPLAY46:0440/txt/COCKPIT/SDUMMY-SALDO").text)
                            if saldo == 0:
                                info = "go to final steps"
                                return info
                            else:
                                info = f"Document {doc_number} cannot be processed. saldo is not zero and no matching line could be found. - OUT of scope"
                                return info
                        elif saldo != 0 and more_than_one_line_po == False:
                            total_ordered = process_data.get('po_totals')[0].get('val_ord')
                            check_info = self.check_tolerance(doc_number, saldo, total_ordered, ToleranceRange.highTolerance)
                            if check_info != True:
                                return check_info
                            
                            check_info = self.add_balance_to_first_line(doc_number, saldo)
                            if check_info != True:
                                return check_info
                            info = "go to final steps"
                            return info
                    else:
                        info = f"Document {doc_number} cannot be processed. Document is an Ariba 3 way match invoice without items generated by PO proposal"
                        return info

                elif info is not True:
                    return info

            return True
            
        except Exception as e:
            log(e, lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            raise Exception(f"Error occured in function check_process_data")
                
    def check_po_lines(self, entry_data):
        try:
            a = 0
            b = 0
            previous_nr_inv = ''
            number_inv = ''
            current_po = ''
            previous_po = ''
            list_lenght = 7
            hits = 0
            hits_crop = 0
            ktory = ktory3 = 0

            screen_id = self.find_screen_id()
            # Set scrollbar position to start
            self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_ITEM:/COCKPIT/SAPLDISPLAY46:0410/tbl/COCKPIT/SAPLDISPLAY46G_TC_ITEM_DET").verticalScrollbar.position = 0

            while True:
                if a > 0 and a % list_lenght == 0:
                    # Scroll down after processing a batch
                    self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_ITEM:/COCKPIT/SAPLDISPLAY46:0410/tbl/COCKPIT/SAPLDISPLAY46G_TC_ITEM_DET").verticalScrollbar.position = a
                    b = 0

                    number_inv = self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_ITEM:/COCKPIT/SAPLDISPLAY46:0410/tbl/COCKPIT/SAPLDISPLAY46G_TC_ITEM_DET/txt/COCKPIT/SITEM_DISP-INVOICE_ITEM[1,{b}]").text
                    po_number = self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_ITEM:/COCKPIT/SAPLDISPLAY46:0410/tbl/COCKPIT/SAPLDISPLAY46G_TC_ITEM_DET/ctxt/COCKPIT/SITEM_DISP-PO_NUMBER[2,{b}]").text
                    
                    if '____' in po_number:
                        break
                    
                    # If the same PO line as before, accumulate amount [sum_po_lines]
                    if previous_nr_inv == number_inv:
                        value = self.convert_to_number(self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_ITEM:/COCKPIT/SAPLDISPLAY46:0410/tbl/COCKPIT/SAPLDISPLAY46G_TC_ITEM_DET/txt/COCKPIT/SITEM_DISP-ITEM_AMOUNT[4,0]").text)
                        entry_data['sum_po_lines'] += value

                        if entry_data['rule_5'] and hits == 1:
                            self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_ITEM:/COCKPIT/SAPLDISPLAY46:0410/tbl/COCKPIT/SAPLDISPLAY46G_TC_ITEM_DET").verticalScrollbar.position = ktory - (ktory % list_lenght)
                            self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_ITEM:/COCKPIT/SAPLDISPLAY46:0410/btnG_TC_ITEM_DET_MARK_ALL").press()
                            self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_ITEM:/COCKPIT/SAPLDISPLAY46:0410/tbl/COCKPIT/SAPLDISPLAY46G_TC_ITEM_DET").getAbsoluteRow(ktory - 1).Selected = False
                            self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_ITEM:/COCKPIT/SAPLDISPLAY46:0410/btnG_TC_ITEM_DET_DELETE").press()
                            entry_data['crop_result'] = True
                        return True, entry_data  
                
                # Handle multiple POs [different_pos]
                if entry_data['different_pos']:
                    current_po = self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_ITEM:/COCKPIT/SAPLDISPLAY46:0410/tbl/COCKPIT/SAPLDISPLAY46G_TC_ITEM_DET/ctxt/COCKPIT/SITEM_DISP-PO_NUMBER[2,{b}]").text
                    if previous_po and previous_po != current_po and '____' not in current_po:
                        entry_data['multiple_pos_exist'] = True
                        if entry_data['multiple_only']:
                            return False, entry_data  # Exit if only multiple PO detection was needed
                
                # Check tax codes if required [tax_code_missing]
                if entry_data['tax_code_missing']:
                    taxCode = self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_ITEM:/COCKPIT/SAPLDISPLAY46:0410/tbl/COCKPIT/SAPLDISPLAY46G_TC_ITEM_DET/ctxt/COCKPIT/SITEM_DISP-TAX_CODE[10,{b}]").text
                    if not taxCode:
                        entry_data['missing_tax_codes'] += f"Tax code missing on line {a + 1}.\n"

                # Store previous values for comparison in next loop iteration [different_pos]
                previous_nr_inv = number_inv

                if entry_data['different_pos']:
                    previous_po = current_po
                
                # Check for end condition (empty or invalid text)
                text = self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_ITEM:/COCKPIT/SAPLDISPLAY46:0410/tbl/COCKPIT/SAPLDISPLAY46G_TC_ITEM_DET/txt/COCKPIT/SITEM_DISP-ITEM_AMOUNT[4,{b}]").text
                if text.startswith('__') or text == "":
                    break
                else:
                    entry_data['how_many_lines'] += 1

                # Process Rule 5 logic [rule_5]
                if entry_data['rule_5']:
                    if b >= 1:
                        value = self.convert_to_number(self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_ITEM:/COCKPIT/SAPLDISPLAY46:0410/tbl/COCKPIT/SAPLDISPLAY46G_TC_ITEM_DET/txt/COCKPIT/SITEM_DISP-ITEM_AMOUNT[4,{b}]").text)
                    else:
                        value = self.convert_to_number(self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_ITEM:/COCKPIT/SAPLDISPLAY46:0410/tbl/COCKPIT/SAPLDISPLAY46G_TC_ITEM_DET/txt/COCKPIT/SITEM_DISP-ITEM_AMOUNT[4,0]").text)
                    entry_data['value'] = value
                    if value == entry_data['searched_amount']:
                        hits += 1
                        ktory = a + 1
                    if value == entry_data['saldo'] * -1 and value != 0:
                        hits_crop += 1
                        ktory3 = a + 1

                a += 1
                b += 1

            # Check tax codes if required [tax_code_missing]
            if entry_data['tax_code_missing']:
                return True, entry_data
            
            if b >= 1:
                entry_data['value'] = self.convert_to_number(self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_ITEM:/COCKPIT/SAPLDISPLAY46:0410/tbl/COCKPIT/SAPLDISPLAY46G_TC_ITEM_DET/txt/COCKPIT/SITEM_DISP-ITEM_AMOUNT[4,{b - 1}]").text)
            else:
                entry_data['value'] = self.convert_to_number(self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_ITEM:/COCKPIT/SAPLDISPLAY46:0410/tbl/COCKPIT/SAPLDISPLAY46G_TC_ITEM_DET/txt/COCKPIT/SITEM_DISP-ITEM_AMOUNT[4,0]").text)
            
            # Final Rule 5 processing if needed [rule_5]
            if entry_data['rule_5']:
                if hits == 1:
                    self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/tbl/COCKPIT/SITEM_DISP").verticalScrollbar.position = ktory - (ktory % list_lenght)
                    self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SITEM_DISP-MARK_ALL").press()
                    self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SITEM_DISP").getAbsoluteRow(ktory - 1).Selected = False
                    self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SITEM_DISP-DELETE").press()
                    entry_data['crop_result'] = True
                elif hits_crop == 1:
                    self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SITEM_DISP").verticalScrollbar.position = ktory3 - (ktory3 % list_lenght)
                    self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SITEM_DISP").getAbsoluteRow(ktory3 - 1).Selected = True
                    self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SITEM_DISP-DELETE").press()
                    entry_data['crop_result'] = True

            return True, entry_data

        except Exception as e:
            log(e, lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            raise Exception(f"Error occured in function check_po_lines")

    def check_vendor_critical(self, doc_number, vendor, comp_code, po_number, po_full_booked):
        try:
            
            info = self.is_vendor_critical(vendor, comp_code, doc_number)

            if info == True:
                if po_full_booked == True:
                    info = f"Document {doc_number} cannot be processed. PO {po_number} is fully booked"
                else:
                    info = f"Document {doc_number} cannot be processed. Check vendor critical - PO number is missing"
                return info
                
            if info == False:
                if po_full_booked == True:
                    info = f"Document {doc_number} cannot be processed. Total invoiced amount equals total ordered amount for PO {po_number}. Document was rejected with status code '06B'. Rejection #ID1"
                else:
                    info = f"Document {doc_number} cannot be processed. Document was rejected with status code '06C'. Rejection #ID2"
                return info

            return info
        
        except Exception as e:
            log(e, lte.error)
            raise Exception(f"Error occured in function check_vendor_critical")

    def is_vendor_critical(self, vendor, comp_code, doc_number):
        '''This function is used to check if a vendor is critical.'''
        try:
            
            if comp_code == '3B5':
                df_critical_vendors3B5 = pd.read_excel('critical suppliers.xlsx', sheet_name='AM France 3b5', usecols="A", skiprows=3, engine="openpyxl")
                df_critical_vendor3B5 = df_critical_vendors3B5.loc[df_critical_vendors3B5['Vendor  ACE'] == vendor, 'Vendor  ACE']
                if not df_critical_vendor3B5.empty:
                    return True
            
            if comp_code == 'V436':
                df = pd.read_excel('critical suppliers.xlsx', sheet_name='AMMED v436', usecols="B,F,G", skiprows=3, engine="openpyxl")
                df = df.loc[df['Vendor number ACE'] == vendor]
                df_critical_vendorV436 = df.iloc[:, 0]
                # df_amei = df.iloc[:, 1]
                # df_enames = df.iloc[:, 2]
                if not df_critical_vendorV436.empty:
                    return True 

            return False

        except Exception as e:
            log(e, lte.error)
            raise Exception(f"Error occured in function is_vendor_critical")

    def check_fi_vendors_v436(self, doc_number, vendor):
        try:
            df_ammed_fi = pd.read_excel('AMMED FI Fournisseurs.xlsx', sheet_name='UPDATE FOS FI', usecols="A:B", engine="openpyxl")
            df_fi_vendorsV436 = df_ammed_fi[['Vendor', 'Cir.code']]
            cir_code = df_fi_vendorsV436.loc[df_fi_vendorsV436['Vendor'] == vendor, 'Cir.code']

            if not cir_code.empty:
                info = self.transfer_mm_to_fi(cir_code)
                if info != True:                    
                    return info
                else:
                    info = f"Document {doc_number} cannot be processed. MM invoice was transferred to FI"
                    return info
                
            info = f"Document {doc_number} cannot be processed. Cir.Code wasn not found in AMMED FI Fournisseurs V2.xlsx"
            return info

        except Exception as e:
            log(e, lte.error)
            raise Exception(f"Error occured in function check_fi_vendors_v436")

    def transfer_mm_to_fi(self, doc_number, cir_code):
        try:
            screen_id = self.find_screen_id(type=6, tab=1)
            self.gui_session.findById.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB1/ssubSUB:/COCKPIT/SAPLDISPLAY46:0401/ctxt/COCKPIT/SHDR_DISP-PO_NUMBER").text = ""
            self.gui_session.findById.findById("wnd[0]/tbar[1]/btn[17]").press()
            self.gui_session.findById.findById("wnd[1]/usr/btnBUTTON_1").press()
            self.gui_session.findById.findById("wnd[1]/tbar[0]/btn[0]").press()
            screen_id = self.find_screen_id(type=6, tab=5)
            self.gui_session.findById.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB5/ssubSUB:/COCKPIT/SAPLDISPLAY46:0436/ssubSUB_OTHERS:/COCKPIT/SAPLDISPLAY46:0700/ctxt/COCKPIT/SDYN_SUBSCR_0700-VALUE16").text = cir_code
            self.gui_session.findById.findById("wnd[0]").sendVKey(0)
            return True
        
        except Exception as e:
            log(e, lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            raise Exception(f"Error occured in function transfer_mm_to_fi")

    def check_po_type(self, doc_number):
        try:
            acceptable = [
                'Purchase outside EPO', 
                'Transport PO', 
                'Intercompany PO', 
                'Standard PO', 
                'Ivalua order', 
                'Emergency PO'
            ]
            po_types = [
                'EPO',
                'Transport',
                'interco',
                'Standard'
            ]

            screen_id = self.find_screen_id(type=6, tab=1)
            sap_status_bar_msg = self.gui_session.findById("wnd[0]/sbar").text
            if "PAYMENT CONDITIONS ARE BEING CHANGED" in str(sap_status_bar_msg).upper():
                info = f"'Purchase order' change sub. debit - {sap_status_bar_msg}" 
                return f"Document {doc_number} cannot be processed. SAP info: {info}"
            
            self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB1/ssubSUB:/COCKPIT/SAPLDISPLAY46:0401/ctxt/COCKPIT/SHDR_DISP-PO_NUMBER").SetFocus()
            self.gui_session.findById("wnd[0]").sendVKey(2)
            
            if self.gui_session.findById("wnd[0]/sbar").messagetype == "E":
                info = self.gui_session.findById("wnd[0]/sbar").text
                self.gui_session.findById("wnd[0]").sendVKey(3)
                return f"Document {doc_number} cannot be processed. SAP info: {info}"

            sap_status_bar_msg = self.gui_session.findById("wnd[0]/sbar").text
            if "does not exist" in sap_status_bar_msg:
                info = f"Purchase order {sap_status_bar_msg}" 
                return f"Document {doc_number} cannot be processed. SAP info: {info}"

            screen_id = self.find_screen_id(type=7)
            po_type_text = str(self.gui_session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{screen_id}/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/cmbMEPO_TOPLINE-BSART").text).strip()

            if po_type_text in acceptable:
                self.gui_session.findById("wnd[0]").sendVKey(3)
                if acceptable.index(po_type_text) >= 3:
                    po_type_text = po_types[3]
                else:
                    po_type_text = po_types[acceptable.index(po_type_text)]
                self.find_screen_id(type=6, tab=2)
                return True, po_type_text
            
            self.gui_session.findById("wnd[0]").sendVKey(3)
            self.find_screen_id(type=6, tab=2)

            return False
        
        except Exception as e:
            log(e, lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            raise Exception(f"Error occured in function check_po_type")

    def check_proposal(self, doc_number):
        """
        This function is used to check a proposal in the system.

        Parameters:
        doc_number (str): The document number.

        Returns:
        bool: True if the proposal is valid, info message otherwise.
        """
        try:
            line = self.line_numb()
            self.gui_session = self.gui_connection.children.ElementAt(0)
            self.gui_session.findById("wnd[0]").maximize()
            screen_id = self.find_screen_id()
            line = self.line_numb()
            self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB1").Select()
            screen_id = self.find_screen_id()
            line = self.line_numb()
            try:
                # button delete proposal
                self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_ITEM:/COCKPIT/SAPLDISPLAY46:0410/btnG_TC_ITEM_DET_PROPOSAL").press()
            except:
                screen_id = self.find_screen_id()
                self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/btnTOG_ITEM").press()
                # button delete proposal
                screen_id = self.find_screen_id()
                self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_ITEM:/COCKPIT/SAPLDISPLAY46:0410/btnG_TC_ITEM_DET_PROPOSAL").press()
            line = self.line_numb()
            try:
                self.gui_session.findById("wnd[1]/usr/btnBUTTON_1").press()
                i = 0
                while True:
                    if self.is_message(i):
                        txt = self.gui_session.findById("wnd[1]/usr/subSUBSCREEN:SAPLSBAL_DISPLAY:0101/cntlSAPLSBAL_DISPLAY_CONTAINER/shellcont/shell").getCellValue(i, "T_MSG")
                        if 'NO ITEM PROPOSAL COULD BE GENERATED' in str(txt).upper():
                            self.gui_session.findById("wnd[1]").sendVKey(0)
                            info = f"Document {doc_number} cannot be processed. {txt}"
                            return info
                    else:
                        break
                    i += 1

                self.gui_session.findById("wnd[1]").sendVKey(0)
            except:
                pass
            line = self.line_numb()
            saldo = self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_SALDO:/COCKPIT/SAPLDISPLAY46:0440/txt/COCKPIT/SDUMMY-SALDO").text
            if self.convert_to_number(str(saldo).strip()) == 0.00: 
                return True
            else:
                if self.find_matching_line(screen_id) is not True:
                    info = f"Document {doc_number} cannot be processed. saldo is not zero and no matching line could be found. - OUT of scope"
                    return info
                else:
                    return True
                
        except Exception as e:
            log(f"Error in function check_proposal. Doc number {doc_number}: {e} Line: {line}", lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            return str(f"Error in function check_proposal. Doc number {doc_number}: {e}")
    
    def find_matching_line(self, screen_id):
        """
        This function is used to find a line with amount on invoice.

        Parameters:
        screen_id (int): The screen ID.

        Returns:
        bool: True if the line was found, False otherwise.
        """
        try:
            self.gui_session = self.gui_connection.children.ElementAt(0)
            self.gui_session.findById("wnd[0]").maximize()
            self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB2").Select()
            net = self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB2/ssubSUB:/COCKPIT/SAPLDISPLAY46:0402/txt/COCKPIT/SHDR_DISP-NET_AMOUNT").text
            self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB1").Select()
            i = 0
            while self.check_line_exists(i, screen_id) is True:
                saldo = self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_SALDO:/COCKPIT/SAPLDISPLAY46:0440/txt/COCKPIT/SDUMMY-SALDO").text
                line = self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_ITEM:/COCKPIT/SAPLDISPLAY46:0410/tbl/COCKPIT/SAPLDISPLAY46G_TC_ITEM_DET/txt/COCKPIT/SITEM_DISP-ITEM_AMOUNT[4,{i}]").text
                if self.convert_to_number(str(saldo).strip()) == 0.00: 
                    return True
                else:
                    if str(line).strip() == str(net).strip():
                        j = i + 1
                        if self.check_line_exists(j, screen_id) is True:
                            self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_ITEM:/COCKPIT/SAPLDISPLAY46:0410/tbl/COCKPIT/SAPLDISPLAY46G_TC_ITEM_DET").getAbsoluteRow(j).Selected = True
                            self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_ITEM:/COCKPIT/SAPLDISPLAY46:0410/btnG_TC_ITEM_DET_DELETE").press()
                        else:
                            return False
                    else:
                        self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_ITEM:/COCKPIT/SAPLDISPLAY46:0410/tbl/COCKPIT/SAPLDISPLAY46G_TC_ITEM_DET").getAbsoluteRow(i).Selected = True
                        self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_ITEM:/COCKPIT/SAPLDISPLAY46:0410/btnG_TC_ITEM_DET_DELETE").press()
        
        except Exception as e:
            info = self.gui_session.findById("wnd[0]/sbar").Text
            log(e, lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            raise Exception(f"Error occured in function find_matching_line. SAP info: {info}")

    def check_line_exists(self, i, screen_id):
        """
        This function is used to check if a specific line exists in document.

        Parameters:
        i (int): The line number.
        screen_id (int): The screen ID.

        Returns:
        Bool: True if the line exists, False otherwise.
        """
        try:
            self.gui_session = self.gui_connection.children.ElementAt(0)
            self.gui_session.findById("wnd[0]").maximize()
            t = self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_ITEM:/COCKPIT/SAPLDISPLAY46:0410/tbl/COCKPIT/SAPLDISPLAY46G_TC_ITEM_DET/txt/COCKPIT/SITEM_DISP-ITEM_AMOUNT[4,{i}]").text
            if t != '' and t != '________________':
                return True
            else:
                return False
            
        except Exception as e:
            info = self.gui_session.findById("wnd[0]/sbar").Text
            log(e, lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            raise Exception(f"Error occured in function check_line_exists. SAP info: {info}")
        
    def process_po_types(self, doc_number, company_code, process_data, entry_data):
        try:
            screen_id = self.find_screen_id()
            saldo = self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_SALDO:/COCKPIT/SAPLDISPLAY46:0440/txt/COCKPIT/SDUMMY-SALDO").text
            if self.convert_to_number(str(saldo).strip()) != 0.00:
                vendor = process_data['vendor']
                po_number = process_data['po_number']
                po_lines_result = self.check_po_lines(entry_data)[1]
                is_transport_vendor = self.check_transport_vendor(doc_number, company_code, vendor)
                po_type_check_result = self.check_po_type(doc_number)

                if po_type_check_result in {POType.Transport, POType.EPO}:
                    total_delivered = process_data.get('po_totals')[0].get('val_del')
                    total_to_deliver = process_data.get('po_totals')[0].get('val_to_del')
                    total_invoiced = process_data.get('po_totals')[0].get('val_inv')
                    if total_delivered == 0 and total_to_deliver > 0:
                        info = f"Document {doc_number} cannot be processed. Purchase Order {po_number} has no Goods Receipt"
                        return info
                    
                    elif total_delivered > 0 and total_to_deliver == total_invoiced and total_to_deliver > 0:
                        info = f"Document {doc_number} cannot be processed. Purchase Order {po_number} has no Goods Receipt"
                        return info
                    
                if po_type_check_result == POType.interco:
                    info = f"Document {doc_number} cannot be processed. PO type - interco"
                    return info

                if po_type_check_result == POType.Transport:
                    entry_data['rule_5'] = True
                    entry_data['saldo'] = self.convert_to_number(saldo)
                    entry_data['searched_amount'] = self.convert_to_number(process_data['netto'])
                    po_lines_result = self.check_po_lines(entry_data)[1]
                    saldo = self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_SALDO:/COCKPIT/SAPLDISPLAY46:0440/txt/COCKPIT/SDUMMY-SALDO").text

                    if saldo != 0:
                        info = self.po_type_transport_process(doc_number, process_data, po_lines_result, po_number, saldo)
                        if info != True:
                            return info

                if po_type_check_result == POType.EPO:
                    entry_data['saldo'] = self.convert_to_number(saldo)
                    po_lines_result = self.check_po_lines(entry_data)[1]
                    saldo = self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_SALDO:/COCKPIT/SAPLDISPLAY46:0440/txt/COCKPIT/SDUMMY-SALDO").text
                    
                    info = self.po_type_epo_process(doc_number, po_lines_result, po_number, saldo)
                    if info != True:
                        return info

                if po_type_check_result == POType.Standard:
                    type = 3
                    saldo = self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_SALDO:/COCKPIT/SAPLDISPLAY46:0440/txt/COCKPIT/SDUMMY-SALDO").text
                    process_data_po = self.process_standard_po(doc_number, type, po_number)

                    if 'ERROR' in str(process_data_po).upper() or 'CANNOT BE PROCESSED' in str(process_data_po).upper() or 'DOCUMENT CAN NOT BE OPENED' in str(process_data_po).upper():
                        return process_data_po
                    elif process_data_po == False:
                        return f"Document {doc_number}, cannot be processed. Document could not be assigned neither to Ariba nor to PDF Collector category"

                    info = self.po_type_standard_process(doc_number, process_data_po, saldo)
                    if info != True:
                        return info

            return True
            
        except Exception as e:
            log(f"Error in function process_po_types. Doc number {doc_number}: {e} ", lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            return str(f"Error occured in function process_po_types. Doc number {doc_number}: {e}")
    
    def po_type_standard_process(self, doc_number, process_data_po, saldo):
        try:
            still_to_deliver = process_data_po[1].get('po_totals')[0].get('val_to_del')
            if still_to_deliver == 0:
                info_tolerance = self.check_tolerance(doc_number, saldo, process_data_po[1].get('po_totals')[0].get('val_ord'), ToleranceRange.lowTolerance)
                if info_tolerance != True:
                    return info_tolerance
                check_info = self.add_balance_to_first_line(doc_number, saldo)
                if check_info != True:
                    return check_info
            else:
                two_way_match = process_data_po[1].get('two_way_match')
                if two_way_match == True:
                    info_tolerance = self.check_tolerance(doc_number, saldo, process_data_po[1].get('po_totals')[0].get('val_ord'), ToleranceRange.lowTolerance)
                    if info_tolerance != True:
                        return info_tolerance
                    check_info = self.add_balance_to_first_line(doc_number, saldo)
                    if check_info != True:
                        return check_info
                else:
                    screen_id = self.find_screen_id()
                    find_line_info = self.find_matching_line_po_standard(screen_id)
                    check_result = find_line_info[0]
                    multiple_matches = find_line_info[1]
                    if check_result == True and multiple_matches == True:
                        info = f"Document {doc_number} cannot be processed. Document is a 3 way match invoice.Rule 5 failed. Multiple matching lines were found in the proposal"
                        return info
                    if check_result == False:
                        info = f"Document {doc_number} cannot be processed. Document is a 3 way match invoice.Rule 5 failed. No matching line was found in the proposal"
                        return info

            return True

        except Exception as e:
            log(e, lte.error)
            return str(f"Error occured in function po_type_standard_process. Doc number {doc_number}: {e}")

    def po_type_epo_process(self, doc_number, po_lines_result, po_number, saldo):
        try:
            # Case 1: Single PO line
            if po_lines_result['how_many_lines'] == 1:
                return f"Document {doc_number} cannot be processed. po type = EPO, one line (Rule 4)."

            # Case 2: Multiple PO lines and saldo > 0 (GR amount < invoiced amount)
            if saldo > 0:
                return f"Document {doc_number} cannot be processed. Matching PO line couldn't be found (Rule 4)."

            # Case 3: Find matching line out of available PO lines
            while True:
                screen_id = self.find_screen_id()
                saldo = self.convert_to_number(self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_SALDO:/COCKPIT/SAPLDISPLAY46:0440/txt/COCKPIT/SDUMMY-SALDO").text)

                if saldo < 0:
                    if abs(saldo) >= po_lines_result['value']:
                        self.remove_last_line(po_lines_result['how_many_lines'])
                        saldo = self.convert_to_number(self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_SALDO:/COCKPIT/SAPLDISPLAY46:0440/txt/COCKPIT/SDUMMY-SALDO").text)
                        screen_id = self.find_screen_id()
                        self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_ITEM:/COCKPIT/SAPLDISPLAY46:0410/tbl/COCKPIT/SAPLDISPLAY46G_TC_ITEM_DET").verticalScrollbar.position = 0
                        
                        # Count PO lines again
                        po_lines_result['how_many_lines'] = 0
                        check_lines_result = self.check_po_lines(po_lines_result)
                        how_many_lines = check_lines_result[1]['how_many_lines']
                        if how_many_lines == 0:
                            return f"Document {doc_number} cannot be processed. No matching line in PO {po_number} (Rule 4)."
                    elif saldo == 0:
                        break  # Exit the loop if the saldo is zero
                    elif abs(saldo) < po_lines_result['value']:
                        return f"Document {doc_number} cannot be processed due to Rule 4 failure (no matching line in PO type 47 with multiple lines) (Rule 4)."

                else:
                    return f"Document {doc_number} cannot be processed due to Rule 4 failure (no matching line in PO type 47 with multiple lines) (Rule 4)."

            return True
        
        except Exception as e:
            log(e, lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            return str(f"Error occured in function po_type_epo_process. Doc number {doc_number}: {e}")

    def po_type_transport_process(self, doc_number, process_data, po_lines_result, po_number, saldo):
        try:
            if po_lines_result['how_many_lines'] == 1 and saldo < 0:
                self.add_balance_to_first_line(doc_number, saldo)

            elif po_lines_result['how_many_lines'] == 1 and saldo > 0:
                total_ordered = process_data.get('po_totals')[0].get('val_ord')
                check_info = self.check_tolerance(doc_number, saldo, total_ordered, ToleranceRange.lowTolerance)
                if check_info == True:
                    self.add_balance_to_first_line(doc_number, saldo)
                else:
                    return check_info
                
            if po_lines_result['how_many_lines'] > 1 and saldo > 0:
                screen_id = self.find_screen_id(type=8)
                self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_ITEM:/COCKPIT/SAPLDISPLAY46:0410/tbl/COCKPIT/SAPLDISPLAY46G_TC_ITEM_DET").verticalScrollbar.position = 0
                item_amount = self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_ITEM:/COCKPIT/SAPLDISPLAY46:0410/tbl/COCKPIT/SAPLDISPLAY46G_TC_ITEM_DET/txt/COCKPIT/SITEM_DISP-ITEM_AMOUNT[4,0]").text
                value = round(self.convert_to_number(item_amount) + saldo, 2)
                self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_ITEM:/COCKPIT/SAPLDISPLAY46:0410/tbl/COCKPIT/SAPLDISPLAY46G_TC_ITEM_DET/txt/COCKPIT/SITEM_DISP-ITEM_AMOUNT[4,0]").text = value
                self.gui_session.findById("wnd[0]").sendVKey(0)

            elif po_lines_result['how_many_lines'] > 1 and saldo > 0:
                find_line_time_limit = datetime.now() + timedelta(minutes=1)

                while True:
                    if abs(saldo) >= po_lines_result['value']:
                        if po_lines_result['how_many_lines'] == 0:
                            return f"Document {doc_number} cannot be processed. Robot did not find any matching line in PO {po_number} (RULE 3b)"

                        elif datetime.now() > find_line_time_limit:
                            return f"Document {doc_number} cannot be processed. Time limit for finding any matching line has passed. Error in PO {po_number} (RULE 3b)"

                        # Remove the last line
                        self.remove_last_line(po_lines_result['how_many_lines'], doc_number)

                        # Update saldo and check lines
                        saldo = self.convert_to_number(self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_SALDO:/COCKPIT/SAPLDISPLAY46:0440/txt/COCKPIT/SDUMMY-SALDO").text)
                        self.gui_session.findById("wnd[0]").sendVKey(0)
                        
                        if not self.count_po_lines_fast(po_lines_result['how_many_lines'], po_lines_result['value'], doc_number):
                            return f"Document {doc_number} cannot be processed. Robot did not find any matching line in PO {po_number}"

                    elif saldo == 0:
                        break

                    elif abs(saldo) < po_lines_result['value']:
                        total_ordered = process_data.get('po_totals')[0].get('val_ord')
                        check_info = self.check_tolerance(doc_number, saldo, total_ordered, ToleranceRange.extremeTolerance)
                        if check_info == True:
                            if not self.change_last_amount(po_lines_result['how_many_lines'], doc_number):
                                return f"Document {doc_number} cannot be processed. Cannot change last amount in po type transport process. Rule 3b."
                            else:
                                break
                        else:
                            return check_info

            else:
                info = f"Document {doc_number} is not transport invoice. Vendor is not transport type"
                return info
            
            return True
            
        except Exception as e:
            log(e, lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            return str(f"Error occured in function po_type_transport_process. Doc number {doc_number}: {e}")

    def count_po_lines_fast(self, how_many_lines, value, doc_number):
        try:
            screen_id = self.find_screen_id(type=8)
            how_many_lines = how_many_lines - 1
            if (how_many_lines - 1) < 0:
                return False
            self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_ITEM:/COCKPIT/SAPLDISPLAY46:0410/tbl/COCKPIT/SAPLDISPLAY46G_TC_ITEM_DET").verticalScrollbar.position = how_many_lines
            value = self.convert_to_number(self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_ITEM:/COCKPIT/SAPLDISPLAY46:0410/tbl/COCKPIT/SAPLDISPLAY46G_TC_ITEM_DET/txt/COCKPIT/SITEM_DISP-ITEM_AMOUNT[4,0]").text)
            
            return True

        except Exception as e:
            log(e, lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            return str(f"Error occured in function count_po_lines_fast. Doc number {doc_number}: {e}")
        
    def change_last_amount(self, how_many_lines, doc_number):
        try:
            screen_id = self.find_screen_id(type=8)
            saldo = abs(self.convert_to_number(self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_SALDO:/COCKPIT/SAPLDISPLAY46:0440/txt/COCKPIT/SDUMMY-SALDO").text))
            line_amount = self.convert_to_number(self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_ITEM:/COCKPIT/SAPLDISPLAY46:0410/tbl/COCKPIT/SAPLDISPLAY46G_TC_ITEM_DET/txt/COCKPIT/SITEM_DISP-ITEM_AMOUNT[4,0]").text)
            if line_amount < saldo:
                return False
            else:
                self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_ITEM:/COCKPIT/SAPLDISPLAY46:0410/tbl/COCKPIT/SAPLDISPLAY46G_TC_ITEM_DET/txt/COCKPIT/SITEM_DISP-ITEM_AMOUNT[4,0]").text = round(line_amount - saldo, 2)
                self.gui_session.findById("wnd[0]").sendVKey(0)

            return True

        except Exception as e:
            log(e, lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            return str(f"Error occured in function change_last_amount. Doc number {doc_number}: {e}")
        
    def remove_last_line(self, how_many_lines, doc_number):
        try:
            screen_id = self.find_screen_id(type=8)
            self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_ITEM:/COCKPIT/SAPLDISPLAY46:0410/tbl/COCKPIT/SAPLDISPLAY46G_TC_ITEM_DET").getAbsoluteRow(how_many_lines - 1).Selected = True
            self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_ITEM:/COCKPIT/SAPLDISPLAY46:0410/btnG_TC_ITEM_DET_DELETE").press()

            return True
        
        except Exception as e:
            log(e, lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            raise Exception(f"Error occured in function remove_last_line. Doc number {doc_number}")

    def check_transport_vendor(self, doc_number, company_code, vendor):
        try:
            df_vendorsMatrix = pd.read_excel('vendor matrix for RPA.xlsx')
            df = df_vendorsMatrix[(df_vendorsMatrix['Company code'] == company_code) & (df_vendorsMatrix['Type'] == 'transport')]
            df_check = df.loc[df['Vendor Id'] == vendor]

            if not df_check.empty:
                return True
            
            return False

        except Exception as e:
            raise Exception(f"Error occured in function check_transport_vendor. Doc number {doc_number}")

    def find_matching_line_po_standard(self, screen_id):
        """
        This function is used to find a line with amount on invoice.

        Parameters:
        screen_id (int): The screen ID.

        Returns:
        bool: True if the line was found, False otherwise.
        """
        try:
            self.gui_session = self.gui_connection.children.ElementAt(0)
            self.gui_session.findById("wnd[0]").maximize()
            self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB2").Select()
            net = self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB2/ssubSUB:/COCKPIT/SAPLDISPLAY46:0402/txt/COCKPIT/SHDR_DISP-NET_AMOUNT").text
            self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB1").Select()
            multiple_matches = False
            i = 0
            while self.check_line_exists(i, screen_id) is True:
                line = self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_ITEM:/COCKPIT/SAPLDISPLAY46:0410/tbl/COCKPIT/SAPLDISPLAY46G_TC_ITEM_DET/txt/COCKPIT/SITEM_DISP-ITEM_AMOUNT[4,{i}]").text
                if self.convert_to_number(str(line).strip()) == 0.00: 
                    break

                if str(line).strip() == str(net).strip():
                    i += 1
                    return True, multiple_matches
                else:
                    self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_ITEM:/COCKPIT/SAPLDISPLAY46:0410/tbl/COCKPIT/SAPLDISPLAY46G_TC_ITEM_DET").getAbsoluteRow(i).Selected = True
                    self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_ITEM:/COCKPIT/SAPLDISPLAY46:0410/btnG_TC_ITEM_DET_DELETE").press()

            if i > 1:
                multiple_matches = True

            return False, multiple_matches

        except Exception as e:
            info = self.gui_session.findById("wnd[0]/sbar").Text
            log(e, lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            raise Exception(f"Error occured in function find_matching_line_po_standard. SAP info: {info}")
        
    def check_vmd(self, doc_number):
        try:
            vendor_details = self.get_vendor_details(doc_number)
            if vendor_details == False:
                return f"Document {doc_number} cannot be processed. Vendor details could not be downloaded (check vmd)"
            indexing_details = self.get_indexing_details(doc_number, vendor_details.get('interco_vendor'))
            if not isinstance(indexing_details, dict):
                return f"Document {doc_number} cannot be processed. Indexing details could not be downloaded (check vmd)"
            compare_vmd = self.compare_vmd(doc_number, vendor_details, indexing_details)
            if compare_vmd != True:
                return compare_vmd
            po_details = self.get_po_details(doc_number)

            vend_vat_numbers = vendor_details.get('vat_numbers')
            vend_bank_ids = vendor_details.get('bank_ids')
            po_vat_numbers = po_details.get('vendor_vat_numbers')
            po_bank_ids = po_details.get('vendor_bank_ids')

            if len(vend_vat_numbers) > 0 and len(po_vat_numbers) > 0:
                if vend_vat_numbers[0] != po_vat_numbers[0]:
                    info = f"Document {doc_number} cannot be processed. VAT number of vendor and VAT number of invoicing party in PO are different"
                    return info
            else:
                info = f"Document {doc_number} cannot be processed. VAT number of vendor or VAT number of invoicing party couldn't have been downloaded from VMD"
                return info
            
            if len(vend_bank_ids) > 0 and len(po_bank_ids) > 0:
                if vend_bank_ids[0][1] != po_bank_ids[0][1]:  # Assuming each entry has a [bank_id, bank_account]
                    info = str(f"Document {doc_number} cannot be processed. Bank account number of vendor and invoicing party in PO are different.")
                    return info
            else:
                info = f"Document {doc_number} cannot be processed. Bank IDs couldn't have been downloaded from VMD."
                return info

            return True, vendor_details, indexing_details, po_details

        except Exception as e:
            log(f"Error in function check_vmd. Doc number {doc_number}: {e}", lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            return str(f"Error in function check_vmd. Doc number {doc_number}: {e}")

    def get_vendor_details(self, doc_number, special_run=False, new_session=None):
        try:
            if new_session == None:
                new_session = self.gui_session
                session_nr = 0
                screen_id = self.find_screen_id(type=6, tab=2, session_nr=session_nr)
            else:
                session_nr = 1

            interco_vendor = False
            bank_ids = []

            if special_run == False:
                new_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB2/ssubSUB:/COCKPIT/SAPLDISPLAY46:0402/ctxt/COCKPIT/SHDR_DISP-VENDOR_NO").SetFocus()
                try:
                    new_session.findById("wnd[0]").sendVKey(2)
                except:
                    return False

            screen_id = self.find_screen_id(type=11, tab=None, session_nr=session_nr)
            middle_path = f"subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:20{screen_id}/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP"
            # Navigate to VAT numbers tab and call VAT download function
            new_session.findById(f"wnd[0]/usr/{middle_path}/tabpSCREEN_1100_TAB_03").Select()
            vat_numbers = self.download_vat_numbers(screen_id)
            if vat_numbers == False:
                new_session.findById("wnd[0]").sendVKey(3)
                return False

            new_session.findById(f"wnd[0]/usr/{middle_path}/tabpSCREEN_1100_TAB_04").Select()
            new_session.findById(f"wnd[0]/usr/{middle_path}/tabpSCREEN_1100_TAB_04/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7032/subA02P07:SAPLFS_BP_BDT_FS_ATTRIBUTES:1470/ctxtGS_BP001-VBUND").setFocus()
            info = new_session.findById(f"wnd[0]/usr/{middle_path}/tabpSCREEN_1100_TAB_04/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7032/subA02P07:SAPLFS_BP_BDT_FS_ATTRIBUTES:1470/ctxtGS_BP001-VBUND").text
            interco_vendor = bool(info)
            
            # Navigate to the bank IDs tab
            ile_bankow = 0
            b = 0
            scroll_row = 1
            new_session.findById(f"wnd[0]/usr/{middle_path}/tabpSCREEN_1100_TAB_05").Select()
            
            try:
                while True:                   
                    info = new_session.findById(f"wnd[0]/usr/{middle_path}/tabpSCREEN_1100_TAB_05/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7034/subA02P01:SAPLBUD0:1500/tblSAPLBUD0TCTRL_BUT0BK/txtGT_BUT0BK-BKVID[0,{b}]").text
                    if info:
                        bank_id = [
                            info,
                            new_session.findById(f"wnd[0]/usr/{middle_path}/tabpSCREEN_1100_TAB_05/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7034/subA02P01:SAPLBUD0:1500/tblSAPLBUD0TCTRL_BUT0BK/txtGT_BUT0BK-IBAN[6,{b}]").text,
                        ]
                        bank_ids.append(bank_id)
                    ile_bankow += 1
                    info = new_session.findById(f"wnd[0]/usr/{middle_path}/tabpSCREEN_1100_TAB_05/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7034/subA02P01:SAPLBUD0:1500/tblSAPLBUD0TCTRL_BUT0BK/txtGT_BUT0BK-BKVID[0,{b + 1}]").text
                    if info == '':
                        break
                    b += 1

                    # Scroll
                    if b == scroll_row:
                        try:
                            new_session.findById(f"wnd[0]/usr/{middle_path}/tabpSCREEN_1100_TAB_05/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7034/subA02P01:SAPLBUD0:1500/tblSAPLBUD0TCTRL_BUT0BK").verticalScrollbar.position = ile_bankow
                            b = 0
                        except:
                            pass
            except:
                pass
            
            result_dict = {
                'vat_numbers': vat_numbers,
                'bank_ids': bank_ids,
                'ile_bankow': ile_bankow,
                'interco_vendor': interco_vendor
            }
            # Go back to the previous screen
            new_session.findById("wnd[0]").sendVKey(3)

            return result_dict

        except Exception as e:
            info = new_session.findById("wnd[0]/sbar").Text
            log(e, lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            raise Exception(f"Error occured in function get_vendor_details. Doc number {doc_number}. SAP info: {info}")

    def download_vat_numbers(self, screen_id):
        try:
            vat_numbers = []
            a, b = 0, 0
            scroll_row = 1
            middle_path = f"subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:20{screen_id}/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_03/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7033/subA07P01:SAPLBUPA_BUTX_DIALOG:0100"

            while True:
                vat_number_text = self.gui_session.findById(f"wnd[0]/usr/{middle_path}/tblSAPLBUPA_BUTX_DIALOGTCTRL_BPTAX/txtDFKKBPTAXNUM-TAXNUMXL[2,{a}]").text
                vat_numbers.append(vat_number_text)
                b += 1

                text = self.gui_session.findById(f"wnd[0]/usr/{middle_path}/tblSAPLBUPA_BUTX_DIALOGTCTRL_BPTAX/txtDFKKBPTAXNUM-TAXNUMXL[2,{a + 1}]").text
                if text == '':
                    break
                a += 1

                # Scroll 
                if a == scroll_row:
                    self.gui_session.findById(f"wnd[0]/usr/{middle_path}/tblSAPLBUPA_BUTX_DIALOGTCTRL_BPTAX").verticalScrollbar.position = b
                    a = 0 

            return vat_numbers 

        except:
            return False
    
    def get_indexing_details(self, doc_number, interco_vendor):
        try:

            # Select the VAT tab and fetch vat_index
            screen_id = self.find_screen_id(type=6, tab=5)
            part_path1 = "COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR"
            part_path2 = "ssubSUB:/COCKPIT/SAPLDISPLAY46:0436/ssubSUB_OTHERS:/COCKPIT/SAPLDISPLAY46:0700/txt/COCKPIT"
            vat_index = self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/{part_path1}/tabpTAB5/{part_path2}/SDYN_SUBSCR_0700-VALUE6").text

            # Select the Bank tab to fetch bank_index
            screen_id = self.find_screen_id(type=6, tab=4)
            a = 0
            part_path2 = "ssubSUB:/COCKPIT/SAPLDISPLAY46:0404/subSUB_BANK:/COCKPIT/SAPLDISPLAY46:0407/tbl/COCKPIT/SAPLDISPLAY46G_TC_BANK_DET/txt/COCKPIT"

            while True:
                text = self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/{part_path1}/tabpTAB4/{part_path2}/SBANK_ACCT-BANKA[5,{a}]").text
                if text == "":
                    break

                if "__________________________" in text:
                    return False

                # Set bank_index based on vendor type
                if interco_vendor:
                    # For Intercompany vendors, check for treasury-related text
                    if "ARCELOR MITTAL TREASURY" in text or "ARCELORMITTAL TREASURY" in text:
                        bank_index = self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/{part_path1}/tabpTAB4/{part_path2}/SBANK_ACCT-IBAN[7,{a}]").text
                        break
                else:
                    # For regular vendors, check for "No check" entry
                    if "No check (see additional bank data check)" in text:
                        bank_index = self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/{part_path1}/tabpTAB4/{part_path2}/SBANK_ACCT-IBAN[7,{a}]").text
                        break

                a += 1

            indexing_details = {
                "vat_index": vat_index, 
                "bank_index": bank_index
            }
            
            return indexing_details

        except Exception as e:
            log(f"Doc number {doc_number}. Error in get_indexing_details: {e}", lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            return False
    
    def compare_vmd(self, doc_number, vendor_details, indexing_details):
        """
        Compares VAT and bank details between vendor master data and indexing level data.
        
        Parameters:
        - vat_numbers (list of str): VAT numbers for the processed vendor.
        - bank_ids (list of tuples): Bank details as [(bank_code, IBAN), ...].
        - ile_bankow (int): Number of bank accounts available in vendor data.
        - bank_index (str): The indexing level bank account to check against.
        - vat_index (str): The indexing level VAT number to check against.

        Returns:
        - bool: True if data matches, info otherwise.
        """
        try:
            vat_numbers = vendor_details.get('vat_numbers')
            bank_ids = vendor_details.get('bank_ids')
            ile_bankow = vendor_details.get('ile_bankow')
            vat_index = indexing_details.get('vat_index')
            bank_index = indexing_details.get('bank_index')

            # Compare VAT numbers
            for i, vat_number in enumerate(vat_numbers):
                if vat_number == vat_index:
                    break
                elif i == len(vat_numbers) - 1:
                    info = f"Document {doc_number} cannot be processed. VAT Numbers of Processed Vendor and Indexing Level Vendor do not match! (check vmd)"
                    return info

            # Compare bank IDs if available
            if ile_bankow > 0:
                for bank_code, iban in bank_ids:
                    if bank_index in iban or iban in bank_index or bank_index == iban:
                        return True

                # Check if there's only one bank account and no bank_index provided
                if ile_bankow == 1 and not bank_index:
                    return True
                
                info = f"Document {doc_number} cannot be processed. Bank account in VMD and indexing level bank account do not match! (check vmd)"
                return info
            else:
                info = f"Document {doc_number} cannot be processed. No bank account available in vendor master data! (check vmd)"
                return info

        except Exception as e:
            # TODO: check exception in 'compare_vmd' function
            log(f". Doc number {doc_number}. Error in compare_vmd: {e}", lte.error)
            return False
        
    def get_po_details(self, doc_number):
        """
        Retrieves PO details such as Invoicing Party, Company Code, VAT Numbers, Bank IDs, and PO Currency.
        
        Parameters:
        - session (object): SAP session object.
        - invoicing_party (str): Output variable to hold the invoicing party details.
        - comp_code_po (str): Output variable to hold the company code.
        - vat_numbers2 (list): List to store VAT numbers for the PO.
        - bank_ids2 (list): List to store Bank IDs for the PO.
        - ile_bankow2 (int): Variable to store the number of bank accounts.
        - po_currency (str): Output variable to hold the PO currency.
        
        Returns:
        - bool: True if details are successfully retrieved, False otherwise.
        """
        try:
            # Select initial PO screen and set focus on PO number
            screen_id = self.find_screen_id(type=6, tab=1)            
            self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB1/ssubSUB:/COCKPIT/SAPLDISPLAY46:0401/ctxt/COCKPIT/SHDR_DISP-PO_NUMBER").SetFocus()
            self.gui_session.findById("wnd[0]").sendVKey(2)

            # Access PO details by sending keys and selecting specific tabs
            self.gui_session.findById("wnd[0]").sendVKey(26)  # Simulates pressing key for transaction

            # Retrieve company code and invoicing party details
            screen_id = self.find_screen_id(type=2, tab=8)
            comp_code_po = self.gui_session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{screen_id}/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT8/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-BUKRS").text
            if comp_code_po == '':
                info = f"Document {doc_number} cannot be processed. There is no 'company code' on tab 'Org. Data' in this PO."
                return info
            
            screen_id = self.find_screen_id(type=2, tab=6)
            screen_id = self.find_screen_id(type=9)
            invoicing_party = self.gui_session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{screen_id}/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT6/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1224/subPARTNERS:SAPLEKPA:0111/tblSAPLEKPATC_0111/txtTPART-VTEXT[1,0]").text
            if invoicing_party == '':
                info = f"Document {doc_number} cannot be processed. There is no 'invoicing party' on tab 'Partners' in this PO."
                return info

            # Locate and set Invoicing Party if present in details
            row_index = 0
            while True:  
                text = self.gui_session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{screen_id}/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT6/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1224/subPARTNERS:SAPLEKPA:0111/tblSAPLEKPATC_0111/txtTPART-VTEXT[1,{row_index}]").text
                if "Invoicing Party" in text:
                    invoicing_party = self.gui_session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{screen_id}/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT6/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1224/subPARTNERS:SAPLEKPA:0111/tblSAPLEKPATC_0111/ctxtWRF02K-GPARN[2,{row_index}]").text
                    break
                if "_______" in text:
                    info = f"Document {doc_number} cannot be processed. There is no 'invoicing party' on tab 'Partners' in this PO."
                    return info
                row_index += 1

            # Retrieve PO Currency
            screen_id = self.find_screen_id(type=2, tab=1)
            screen_id = self.find_screen_id(type=10)
            po_currency = self.gui_session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{screen_id}/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT1/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1226/ctxtMEPO1226-WAERS").text
            po_currency = str(po_currency).upper()
            if po_currency == '':
                info = f"Document {doc_number} cannot be processed. There is no 'po currency' on tab 'Delivery/Invoice' in this PO."
                return info
            self.gui_session.findById("wnd[0]").sendVKey(3)

            # Retrieve Vendor PO details (Assuming get_vendor_po_details is defined)
            result = self.get_vendor_po_details(doc_number, invoicing_party, comp_code_po) # , vendor_vat_numbers, vendor_bank_ids, vendor_ile_bankow)
            self.gui_session = self.gui_connection.children.ElementAt(0)
            if "cannot be processed" in result:
                return result
            result_dict = {
                'invoicing_party': invoicing_party,
                'comp_code_po': comp_code_po,
                'vendor_vat_numbers': result.get('vat_numbers'),
                'vendor_bank_ids': result.get('bank_ids'),
                'vendor_ile_bankow': result.get('ile_bankow'),
                'po_currency': po_currency
            }

            if result:
                return result_dict
            else:
                info = f"Document {doc_number} cannot be processed. Can't retrieve Vendor PO details."
                return info

        except Exception as e:
            info = self.gui_session.findById("wnd[0]/sbar").Text
            log(e, lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            return str(f"Error occured in function get_po_details. Doc number {doc_number}. SAP info: {info}")
        
    def get_vendor_po_details(self, doc_number, invoicing_party, comp_code_po):
        try:
            new_gui_session = self.new_session()
            time.sleep(1)
            if new_gui_session == False:
                return f"Document {doc_number} cannot be processed. Vendor details could not be downloaded. New SAP session couldn't be started"
            new_gui_session.findById("wnd[0]/tbar[0]/okcd").text = "/nxk03"  # Example: Open the transaction SE38
            new_gui_session.findById("wnd[0]").sendVKey(0)
            new_gui_session.findById("wnd[0]/usr/ctxtRF02K-LIFNR").text = invoicing_party
            new_gui_session.findById("wnd[0]/usr/ctxtRF02K-BUKRS").text = comp_code_po
            new_gui_session.findById("wnd[0]").sendVKey(8)
            new_gui_session.findById("wnd[0]").sendVKey(7)
            new_gui_session.findById("wnd[0]").sendVKey(0)
            
            if str(new_gui_session.findById("wnd[0]/sbar").messagetype) == "E":
                sap_info = str(f"SAP info: {new_gui_session.findById("wnd[0]/sbar").text}")
                info = self.close_additional_session(new_gui_session)
                new_gui_session = None
                return sap_info

            vendor_data = self.get_vendor_details(doc_number, True, new_gui_session)
            if vendor_data == False:
                info = self.close_additional_session(new_gui_session)
                new_gui_session = None
                return f"Document {doc_number} cannot be processed. Vendor details could not be downloaded (check vmd)"

            info = self.close_additional_session(new_gui_session)
            new_gui_session = None

            return vendor_data

        except Exception as e:
            info = self.gui_session.findById("wnd[0]/sbar").Text
            log(e, lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            return str(f"Error occured in function get_vendor_po_details. Doc number {doc_number}. SAP info: {info}")

    def check_permitted_payee(self, doc_number):
        """
        This function is used to check if a payee is permitted.

        Parameters:
        doc_number (str): The document number.

        Returns:
        permitted (bool): True if the payee is permitted, False otherwise.
        """
        try:
            screen_id = self.find_screen_id()
            line = self.line_numb()
            self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB4").select()
            line = self.line_numb()
            self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB4/ssubSUB:/COCKPIT/SAPLDISPLAY46:0404/subSUB_VEND:/COCKPIT/SAPLDISPLAY46:0435/btnMAST").press()
            line = self.line_numb()
            # check if vendor is marked for deletion
            info = self.confirm_extra_window()
            if info != None and "is marked for deletion" in info:
                return info
            try:
                self.gui_session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_07").select()
                line = self.line_numb()
                self.gui_session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_07/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7120/subA10P01:SAPLCVI_FS_UI_VENDOR_ENH:0045/btnPUSH_CVIV_PAYEE").press()
            except:
                line = self.line_numb()
                self.gui_session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_08").select()
                line = self.line_numb()
                self.gui_session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_08/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7120/subA10P01:SAPLCVI_FS_UI_VENDOR_ENH:0045/btnPUSH_CVIV_PAYEE").press()
            try:
                line = self.line_numb()
                self.gui_session.findById("wnd[1]/tbar[0]/btn[12]").press()
                line = self.line_numb()
                self.gui_session.findById("wnd[0]/tbar[0]/btn[3]").press()
                # permitted payee
                return True
            except:
                line = self.line_numb()
                self.gui_session.findById("wnd[0]/tbar[0]/btn[3]").press()
                # not permitted payee
                return False
        except Exception as e:
            log(f"Error in function check_permitted_payee.Doc number: {doc_number} line: {line}. {e}", lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            return str(f"Error in function check_permitted_payee. {e}")
        
    def check_fields(self, doc_number):
        """
        This function is used to check if necessary fields are not empty.

        Parameters:
        doc_number (str): The document number.

        Returns:
        valid (bool): True if all fields are valid, list otherwise.
        """
        try:
            line = self.line_numb()
            self.gui_session = self.gui_connection.children.ElementAt(0)
            self.gui_session.findById("wnd[0]").maximize()
            fields = []
            line = self.line_numb()
            screen_id = self.find_screen_id()
            self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB2").Select()
            line = self.line_numb()
            
            # Define a list of fields with their corresponding paths and names
            field_checks = [
                ('Reference', f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB2/ssubSUB:/COCKPIT/SAPLDISPLAY46:0402/txt/COCKPIT/SHDR_DISP-REF_DOC_NO"),
                ('Document date', f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB2/ssubSUB:/COCKPIT/SAPLDISPLAY46:0402/ctxt/COCKPIT/SHDR_DISP-DOC_DATE"),
                ('Gross amount', f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB2/ssubSUB:/COCKPIT/SAPLDISPLAY46:0402/txt/COCKPIT/SHDR_DISP-GROSS_AMOUNT"),
                ('Net amount', f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB2/ssubSUB:/COCKPIT/SAPLDISPLAY46:0402/txt/COCKPIT/SHDR_DISP-NET_AMOUNT"),
                ('Vendor', f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB2/ssubSUB:/COCKPIT/SAPLDISPLAY46:0402/ctxt/COCKPIT/SHDR_DISP-VENDOR_NO"),
                ('Currency', f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB2/ssubSUB:/COCKPIT/SAPLDISPLAY46:0402/ctxt/COCKPIT/SHDR_DISP-CURRENCY"),
                ('Company code', f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB2/ssubSUB:/COCKPIT/SAPLDISPLAY46:0402/ctxt/COCKPIT/SHDR_DISP-COMP_CODE")
            ]
            
            # Check if any of the fields are empty
            for field_name, field_path in field_checks:
                value = self.gui_session.findById(field_path).text
                if value == '':
                    fields.append(field_name)
                if field_name == 'Gross amount':
                    gross_amount = value
                if field_name == 'Net amount':
                    net_amount = value
            
            if len(fields) != 0:
                field_label = "Field" if len(fields) == 1 else "Fields"
                info = f"Document {doc_number} cannot be processed. {field_label}: {fields} {'is' if len(fields) == 1 else 'are'} empty."
                return info
            
            if gross_amount < net_amount:
                info = f"Document {doc_number} cannot be processed. Value in field [Net amount] is greater than in [Gross amount]"
                return info
            
            return True
        
        except Exception as e:
            log(f"Error in function check_fields. Doc number {doc_number}: {e} Line: {line}", lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            return str(f"Error in function check_fields. Doc number {doc_number}: {e}")
        
    def check_po(self, doc_number):
        """
        This function is used to check the status of a Purchase Order (PO) in the system.

        Parameters:
        doc_number (str): The document number.

        Returns:
        status (str): The status of the PO.
        """
        try:
            info_list = []
            screen_id = self.find_screen_id()
            self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB2").select()
            currency_inv = self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB2/ssubSUB:/COCKPIT/SAPLDISPLAY46:0402/ctxt/COCKPIT/SHDR_DISP-CURRENCY").text
            vendor_inv = self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB2/ssubSUB:/COCKPIT/SAPLDISPLAY46:0402/ctxt/COCKPIT/SHDR_DISP-VENDOR_NO").text
            company_code_inv = self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB2/ssubSUB:/COCKPIT/SAPLDISPLAY46:0402/ctxt/COCKPIT/SHDR_DISP-COMP_CODE").text
            self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB1").select()
            self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB1/ssubSUB:/COCKPIT/SAPLDISPLAY46:0401/ctxt/COCKPIT/SHDR_DISP-PO_NUMBER").setFocus()
            self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB1/ssubSUB:/COCKPIT/SAPLDISPLAY46:0401/ctxt/COCKPIT/SHDR_DISP-PO_NUMBER").caretPosition = 4
            self.gui_session.findById("wnd[0]").sendVKey(2)
            # expand header
            self.gui_session.findById("wnd[0]").sendVKey(26)
            # select PO tab
            type = 1
            self.select_po_tab(type)
            # check vendor number
            screen_id = self.find_screen_id_po(type)
            vendor_po = self.gui_session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{screen_id}/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/ctxtMEPO_TOPLINE-SUPERFIELD").text
            p = str(vendor_po).find(" ")
            vendor_po = vendor_po[:p]
            if vendor_inv != vendor_po:
                info_list.append("vendor number")
            # check currency
            type = 2
            screen_id = self.find_screen_id_po(type)
            currency_po = self.gui_session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{screen_id}/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT1/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1226/ctxtMEPO1226-WAERS").text
            if currency_inv != currency_po:
                info_list.append("currency")
            # check company code
            type = 3
            screen_id = self.find_screen_id_po(type)
            self.gui_session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{screen_id}/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT8").select()
            company_code_po = self.gui_session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{screen_id}/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT8/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-BUKRS").text
            if company_code_inv != company_code_po:
                info_list.append("company code")
            # get invoice party
            b = 0
            for b in range(31):
                if self.switch_po_tab6(b):
                    break
            if not self.get_invoicing_party_from_po(doc_number):
                info_list.append("There is no invoicing party on tab 'Partners' in this PO.")
            self.gui_session.findById("wnd[0]").sendVKey(3)

            return info_list
        
        except Exception as e:
            log(f"Error in function check_po. Doc number {doc_number}: {e}", lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            return str(f"Error in function check_po. Doc number {doc_number}: {e}")
    
    def select_po_tab(self, type):
        try:
            if type == 1:
                screen_id = self.find_screen_id_po(type)
                self.gui_session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{screen_id}/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT6").Select()
            elif type == 2:
                screen_id = self.find_screen_id_po(type)
                if not self.get_tab_8_or_9(screen_id):
                    return False
            elif type == 3:
                self.gui_session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{screen_id}/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT1").Select()
            return True
        except Exception as e:
            log(f"Error in function select_po_tab: {e}", lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            return str("Error in function select_po_tab")
    
    def get_tab_8_or_9(self, screen_id):
        """
        This function is used to switch to the 8th or 9th tab of the Purchase Order (PO) in cockpit.

        Parameters:
        screen_id (str): The screen ID.

        Returns:
        success (bool): True if the switch was successful, False otherwise.
        """
        try:
            try:
                self.gui_session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{screen_id}/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT8").Select()
                return True
            except:
                self.gui_session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{screen_id}/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT9").Select()
                return True
        except Exception as e:
            log(f"Error in function get_tab_8_or_9: {e}", lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            return str("Error in function get_tab_8_or_9")
        
    def switch_po_tab6(self, b):
        """
        This function is used to switch to the 6th tab of the Purchase Order (PO) in cockpit.

        Parameters:
        b (int): The value for iteration

        Returns:
        success (bool): True if the switch was successful, False otherwise.
        """
        try:
            while True:
                try:
                    if b < 10:
                        self.gui_session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:000{b}/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT6").Select()
                        return True
                    else:
                        self.gui_session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{b}/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT6").Select()
                        return True
                except:
                    return False
        except Exception as e:
            log(f"Error in function check_po: {e}", lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            return str(f"Error in function switch_po_tab6: {e}")
    
    def find_current_saplmegui(self, b):
        """
        This function is used to find the current SAPLMEGUI (SAP Material Management GUI).

        Parameters:
        b (int): The value for iteration.

        Returns:
        bool: True if the current SAPLMEGUI is found, False otherwise.
        """
        try:
            # Determine the path based on the value of 'b'
            while True:
                try:
                    if b < 10:
                        txt = self.gui_session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:000{b}/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT6/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1224/subPARTNERS:SAPLEKPA:0111/tblSAPLEKPATC_0111/txtTPART-VTEXT[1,2]").text
                        return True
                    else:
                        txt = self.gui_session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{b}/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT6/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1224/subPARTNERS:SAPLEKPA:0111/tblSAPLEKPATC_0111/txtTPART-VTEXT[1,2]").text
                        return True
                except:
                    return False
        except Exception as e:
            log("Error in function find_current_saplmegui.", lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            return str(f"Error in function find_current_saplmegui: {e}")
        
    def get_invoicing_party_from_po(self, doc_number):
        """
        This function is used to get the invoicing party from a Purchase Order (PO).

        Parameters:
        doc_number (str): The document number of the PO.

        Returns:
        bool: True if the invoicing party is found, False otherwise.
        """
        try:
            b = 0
            # Find the current SAPLMEGUI session
            for b in range(31):
                if self.find_current_saplmegui(b):
                    break
                if b == 30:
                    special_msg = "find_current_saplmegui --> ID was not found."
                    raise Exception(special_msg)

            # Loop through to find the invoicing party
            for a in range(36):
                if b < 10:
                    txt = self.gui_session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:000{b}/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/"
                                        f"subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT6/"
                                        f"ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1224/subPARTNERS:SAPLEKPA:0111/"
                                        f"tblSAPLEKPATC_0111/txtTPART-VTEXT[1,{a}]").text
                else:
                    txt = self.gui_session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{b}/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/"
                                        f"subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT6/"
                                        f"ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1224/subPARTNERS:SAPLEKPA:0111/"
                                        f"tblSAPLEKPATC_0111/txtTPART-VTEXT[1,{a}]").text

                # Check if the string contains "_______"
                if "_______" in txt:
                    break

                if txt == "Invoicing Party":
                    if b < 10:
                        self.gui_session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:000{b}/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/"
                                                    f"subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT6/"
                                                    f"ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1224/subPARTNERS:SAPLEKPA:0111/"
                                                    f"tblSAPLEKPATC_0111/ctxtWRF02K-GPARN[2,{a}]").text
                        return True
                    else:
                        self.gui_session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{b}/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/"
                                                    f"subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT6/"
                                                    f"ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1224/subPARTNERS:SAPLEKPA:0111/"
                                                    f"tblSAPLEKPATC_0111/ctxtWRF02K-GPARN[2,{a}]").text
                        return True
            return False
        except Exception as e:
            log(f"Error in function check_po. Doc number {doc_number}: {e}", lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            return str(f"Error in function check_po. Doc number {doc_number}: {e}")
    
    def find_screen_id_po(self, type):
        """
        This function is used to find the screen ID of a Purchase Order (PO) in the system.

        Parameters:
        Type (int): The type of the PO screen.

        Returns:
        screen_id (str): The screen ID of the PO.
        """
        try:
            i = 0
            while True:
                if self.is_screen_po(i, type):
                    break 
                i += 1
            if type != 1:
                i = 0
                while True:
                    if self.is_details_po(i, type):
                        screen_id = i
                        return screen_id
                    i += 1
            else:
                screen_id = i
                return screen_id
        except Exception as e:
            info = self.gui_session.findById("wnd[0]/sbar").Text
            log(e, lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            raise Exception(f"Error occured in function find_screen_id_po. SAP info: {info}")
    
    def is_details_po(self, i, type):
        """
        This function is used to check if the current screen is a Purchase Order (PO) details screen.

        Parameters:
        Type (int): The type of the PO screen.
        i (int): The screen number.

        Returns:
        Bool value: True if the current screen is a PO details screen, False otherwise.
        """
        try:
            if type == 2:
                try:
                    self.gui_session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT1/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1226/ctxtMEPO1226-WAERS").text
                    return i
                except:
                    return False
            if type == 3:
                try:
                    self.gui_session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT8/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-BUKRS").text
                    return i
                except:
                    return False
        except Exception as e:
            info = self.gui_session.findById("wnd[0]/sbar").Text
            log(e, lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            raise Exception(f"Error occured in function is_details_po. SAP info: {info}")

    def is_screen_po(self, i, type):
        """
        This function is used to check if the current screen is a Purchase Order (PO) screen.

        Parameters:
        Type (int): The type of the PO screen.
        i (int): The screen number.

        Returns:
        bool value: True if the current screen is a PO screen, False otherwise.
        """
        try:
            if type == 1:
                try:
                    tekst = self.gui_session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/ctxtMEPO_TOPLINE-SUPERFIELD").text
                except:
                    return False
            if type == 2:
                try:
                    tekst = self.gui_session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT1").select()
                except:
                    return False
            if type == 3:
                try:
                    tekst = self.gui_session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT8").select()
                except:
                    return False
            return True
        except Exception as e:
            info = self.gui_session.findById("wnd[0]/sbar").Text
            log(e, lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            raise Exception(f"Error occured in function is_screen_po. SAP info: {info}")
        
    def check_dates(self, doc_number, comp_code):
        """
        This function is used to check the validity of dates in a given dataset.

        Parameters:
        doc_number (str): The document number.
        comp_code (str): The company code.

        Returns:
        valid: True if all dates are valid, string otherwise.
        """
        from datetime import datetime
        import datetime as dt
        try:
            screen_id = self.find_screen_id()
            screen_path = f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB2"
            self.gui_session.findById(screen_path).Select()
            dd = self.gui_session.findById(f"{screen_path}/ssubSUB:/COCKPIT/SAPLDISPLAY46:0402/ctxt/COCKPIT/SHDR_DISP-DOC_DATE").text
            doc_date = datetime.strptime(dd, '%d.%m.%Y')
            posting_date = self.get_date(comp_code)
            if posting_date is None:
                today = dt.date.today()
                posting_date = dt.datetime.strftime(today, '%d.%m.%Y')
            else:
                posting_date = dt.datetime.strftime(posting_date, '%d.%m.%Y')
            self.gui_session.findById(f"{screen_path}/ssubSUB:/COCKPIT/SAPLDISPLAY46:0402/ctxt/COCKPIT/SHDR_DISP-PSTNG_DATE").text = str(posting_date)
            self.gui_session.findById("wnd[0]").sendVKey(0)
            while self.gui_session.findById("wnd[0]/sbar").messagetype == "W":
                self.gui_session.findById("wnd[0]").sendVKey(0)
            pd = self.gui_session.findById(f"{screen_path}/ssubSUB:/COCKPIT/SAPLDISPLAY46:0402/ctxt/COCKPIT/SHDR_DISP-PSTNG_DATE").text
            posting_date = datetime.strptime(pd, '%d.%m.%Y')
            
            if posting_date < doc_date:
                return f"Document date is greater than posting date."
            
            return True
        
        except Exception as e:
            log(e, lte.error)
            log(f"Error in function check_dates. Doc number {doc_number}: {e}", lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            return str(f"Error in function check_dates. Doc number {doc_number}: {e}")
    
    def get_date(self, company_code):
        """
        This function is used to get posting date from calendar.

        Parameters:
        company_code (str): The company code.

        Returns:
        posting_date (str): The posting date otherwise None.
        """
        from datetime import datetime 
        try:
            current_date = datetime.now().strftime('%Y-%m-%d')
            dfs = self.excel.get_calendar()
            df_3B5 = dfs[0]
            df_V436 = dfs[1]
            if company_code == '3B5':
                if not df_3B5[df_3B5['Date']==current_date]['Date to be taken for posting'].empty:
                    row_3B5 = df_3B5[df_3B5['Date']==current_date]['Date to be taken for posting'].index[0]
                    value_3B5 = df_3B5.at[row_3B5, 'Date to be taken for posting']
                    return value_3B5
                else:
                    return None
            elif company_code == 'V436':
                if not df_V436[df_V436['Date']==current_date]['Date to be taken for posting'].empty:
                    row_V436 = df_V436[df_V436['Date']==current_date]['Date to be taken for posting'].index[0]
                    value_V436 = df_V436.at[row_V436, 'Date to be taken for posting']
                    return value_V436
                else:
                    return None
        except Exception as e:
            log(f"Error in function get_date. {e}", lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            return str(f"Error in function get_date. {e}")
    
    def check_bank_ids(self, doc_number, check_vmd, process_data):
        try:
            bank_ids = check_vmd.get('bank_ids')
            interco_vendor = check_vmd.get('interco_vendor')
            ile_bankow = check_vmd.get('ile_bankow')
            vendor = process_data.get('vendor')
            
            screen_id = self.find_screen_id(type=6, tab=2)
            bank_id = self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB2/ssubSUB:/COCKPIT/SAPLDISPLAY46:0402/ctxt/COCKPIT/SHDR_DISP-BVTYP").text
            
            # Check if bankID is empty
            if bank_id == "":
                if ile_bankow == 1:  # If only one bank is available, select it
                    self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB2/ssubSUB:/COCKPIT/SAPLDISPLAY46:0402/ctxt/COCKPIT/SHDR_DISP-BVTYP").text = bank_ids[0][0]

                elif interco_vendor:  # If multiple banks are available and vendor is intercompany, look for "CC" bank
                    for bank in bank_ids:
                        if bank[0] == "CC":
                            self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB2/ssubSUB:/COCKPIT/SAPLDISPLAY46:0402/ctxt/COCKPIT/SHDR_DISP-BVTYP").text = bank[0]
                            break
                    else:  # If "CC" bank is not found
                        error_message = f"Vendor {vendor} is an intercompany partner, but there is no 'CC' bank account in master data."
                        return error_message
                else:
                    error_message = "Bank account was not selected."
                    return error_message
            
            return True

        except Exception as e:
            log(f"Error occured in function check_bank_ids. Document {doc_number}. {e, lte.error}")
            return False
        
    def check_saldo(self, doc_number):
        try:
            screen_id = self.find_screen_id()
            saldo = self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_SALDO:/COCKPIT/SAPLDISPLAY46:0440/txt/COCKPIT/SDUMMY-SALDO").text
            if saldo == "":
                return False

            saldo = self.convert_to_number(saldo)
            if saldo != 0.00:
                return False
            
            return True

        except Exception as e:
            log(f"Error occured in function check_saldo. Document {doc_number}. {e, lte.error}")
            return False
        
    def check_tax_code(self, doc_number, entry_data):
        """
        This function is used to check if missing tax codes.

        Parameters:
        doc_number (str): The document number.
        entry_data (dict): The entry data.

        Returns:
        Bool or string: True if the tax code is correct, string otherwise.
        """
        try:
            entry_data['tax_code_missing'] = True
            po_lines_result = self.check_po_lines(entry_data)[1]
            if po_lines_result.get('missing_tax_codes') == '':
                return True
            else:
                return False, po_lines_result.get('missing_tax_codes')

        except Exception as e:
            log(f"Error in function check_tax_code. Doc number {doc_number}: {e}", lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            return str(f"Error in function check_tax_code. Doc number {doc_number}: {e}")
        
    def check_before_book(self, doc_number):
        """
        This function is used to perform checks before booking.

        Parameters:
        None

        Returns:
        valid (bool): True if the pre-booking checks are successful, False otherwise.
        """
        try:
            self.confirm_warning()
            self.gui_session.findById("wnd[0]/tbar[1]/btn[23]").press()
            self.confirm_warning()
            i = 0
            while True:
                if self.is_message(i):
                    type = self.gui_session.findById("wnd[1]/usr/subSUBSCREEN:SAPLSBAL_DISPLAY:0101/cntlSAPLSBAL_DISPLAY_CONTAINER/shellcont/shell").getCellValue(i, "%_ICON")
                    if str(type).find("Error") != -1:
                        txt = self.gui_session.findById("wnd[1]/usr/subSUBSCREEN:SAPLSBAL_DISPLAY:0101/cntlSAPLSBAL_DISPLAY_CONTAINER/shellcont/shell").getCellValue(i, "T_MSG")

                        if str(txt).find("Transport inv - WC price/quantity difference needed") != -1:
                            self.gui_session.findById("wnd[0]").sendVKey(3)
                            if self.gui_session.findById("wnd[1]", False) is not None:
                                self.gui_session.findById("wnd[1]/usr/btnBUTTON_2").press()
                            info = f"Document {doc_number} cannot be processed. Workflow 'Transport Invoice Difference'. WebCycle sending is disabled"

                            return info
                            
                        if str(txt).find("has been set as not relevant for tax") != -1:
                            # adjust Text Field = TRUE
                            self.fill_in_text_field(doc_number, txt)

                        info = f"Document {doc_number} cannot be processed. {txt}"
                        if self.gui_session.findById("wnd[1]", False) is not None:
                            self.gui_session.findById("wnd[1]/tbar[0]/btn[0]").press()

                        return info
                else:
                    break
                i += 1
            self.gui_session.findById("wnd[1]/tbar[0]/btn[0]").press()

            return True
        
        except Exception as e:
            log(f"Error in function check_before_book. Doc number {doc_number}: {e}", lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            return str(f"Error in function check_before_book. Doc number {doc_number}: {e}")
    
    def fill_in_text_field(self, doc_number, text):
        try:
            screen_id = self.find_screen_id()
            self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB1").Select()
            self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{screen_id}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB1/ssubSUB:/COCKPIT/SAPLDISPLAY46:0401/ctxt/COCKPIT/SHDR_DISP-SGTXT").text = text
            self.gui_session.findById("wnd[0]").sendVKey(3)
            if self.gui_session.findById("wnd[1]", False) is not None:
                self.gui_session.findById("wnd[1]/usr/btnBUTTON_1").press()

        except Exception as e:
            log(f"Error in function fill_in_text_field. Doc number {doc_number}: {e}", lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            return str(f"Error in function fill_in_text_field. Doc number {doc_number}: {e}")
        
    def confirm_warning(self):
        """
        This function is used to confirm a warning in the system.

        Parameters:
        None

        Returns:
        confirmed (bool): True if the warning is confirmed, False otherwise.
        """
        try:
            while self.gui_session.findById("wnd[0]/sbar").messagetype == "W":
                if self.gui_session.findById("wnd[0]/sbar").messagetype == "W":
                    self.gui_session.findById("wnd[0]").sendVKey(0)

        except Exception as e:
            info = self.gui_session.findById("wnd[0]/sbar").Text
            log(e, lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            raise Exception(f"Error occured in function confirm_warning. SAP info: {info}")
    
    def is_message(self, i):
        """
        This function is used to check if a given input is a message.

        Parameters:
        i (int): row number.

        Returns:
        is_message (bool): True if the input is a message, False otherwise.
        """
        try:
            try:
                self.gui_session.findById("wnd[1]/usr/subSUBSCREEN:SAPLSBAL_DISPLAY:0101/cntlSAPLSBAL_DISPLAY_CONTAINER/shellcont/shell").getCellValue(i, "T_MSG")
                return True
            except:
                return False
            
        except Exception as e:
            info = self.gui_session.findById("wnd[0]/sbar").Text
            log(e, lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            raise Exception(f"Error occured in function is_message. SAP info: {info}")

    def perform_booking_action(self, doc_number):
        """
        This function is used to perform a booking action in the system.

        Parameters:
        doc_number (str): The document number.

        Returns:
        result (str): posting number if the booking action was successful, string otherwise.
        """
        try:
            # click check
            self.gui_session.findById("wnd[0]/tbar[1]/btn[23]").press()
            i = 0
            while True:
                if self.is_message(i):
                    type = self.gui_session.findById("wnd[1]/usr/subSUBSCREEN:SAPLSBAL_DISPLAY:0101/cntlSAPLSBAL_DISPLAY_CONTAINER/shellcont/shell").getCellValue(i, "%_ICON")
                    if str(type).find("Error") != -1:
                        txt = self.gui_session.findById("wnd[1]/usr/subSUBSCREEN:SAPLSBAL_DISPLAY:0101/cntlSAPLSBAL_DISPLAY_CONTAINER/shellcont/shell").getCellValue(i, "T_MSG")
                        info = f"Document {doc_number} cannot be processed. {txt}"
                        return info
                else:
                    break
                i += 1
            self.gui_session.findById("wnd[1]/tbar[0]/btn[0]").press()
            # click post
            self.gui_session.findById("wnd[0]/tbar[1]/btn[16]").press()
            self.gui_session.findById("wnd[1]/tbar[0]/btn[0]").press()
            i = 0
            info = ''
            while True:
                if self.is_message(i):
                    type = self.gui_session.findById("wnd[1]/usr/subSUBSCREEN:SAPLSBAL_DISPLAY:0101/cntlSAPLSBAL_DISPLAY_CONTAINER/shellcont/shell").getCellValue(i, "%_ICON")
                    if str(type).find("Error") == -1:
                        info = self.gui_session.findById("wnd[1]/usr/subSUBSCREEN:SAPLSBAL_DISPLAY:0101/cntlSAPLSBAL_DISPLAY_CONTAINER/shellcont/shell").getCellValue(i, "T_MSG")
                else:
                    break
                i += 1
            self.gui_session.findById("wnd[1]/tbar[0]/btn[0]").press()
            self.gui_session.findById("wnd[0]").sendVKey(3)

            return info
        
        except Exception as e:
            log(f"Error in function perform_booking_action. Doc number {doc_number}: {e}", lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            return str(f"Error in function perform_booking_action. Doc number {doc_number}: {e}")

    def get_posting_number(self, doc_number):
        try:
            posting_number = self.gui_session.findById("wnd[0]/usr/cntlTOP_CONTAINER/shellcont/shell").getCellValue(0, "SAP_DOC_NO")

            return posting_number
        
        except Exception as e:
            log(f"Error in function get_posting_number. Doc number {doc_number}: {e}", lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            return str(f"Error in function get_posting_number. Doc number {doc_number}: {e}")

    def find_screen_id(self, type=None, tab=None, session_nr=0, middle_path_id=None):
        """
        This function is used to find the ID of the current screen in the system.

        Parameters:
        optional: type (int): The type of id path
        optional: tab (int): The tab number

        Returns:
        screen_id (str): The ID of the current screen and tab selected if given.
        """
        try:
            if type in [1, 2, 3, 4, 5, 7, 9, 10, 11, 12, 13]:
                start, end = 0, 50
            else:
                start, end = 370, 399

            for i in range(start, end):
                if self.is_screen(i, type=type, tab=tab, session_nr=session_nr, middle_path_id=middle_path_id) is True:
                    screen_id = i
                    if screen_id < 10:
                        screen_id = f"0{screen_id}"
                    return screen_id
                
        except Exception as e:
            info = self.gui_session.findById("wnd[0]/sbar").Text
            log(e, lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            raise Exception(f"Error occured in function find_screen_id. SAP info: {info}")
    
    def is_screen(self, i, type=None, tab=None, session_nr=0, middle_path_id=None):
        """
        This function is used to check if the current view is a screen in the system.

        Parameters:
        i (int): The screen number.

        Returns:
        Bool: True if the current view is a screen, False otherwise.
        """
        try:
            if i < 10 and type == 11:
                i = f"0{i}"
            self.gui_session = self.gui_connection.children.ElementAt(session_nr)
            paths = {
                1: f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-EBELP[1,0]",
                2: f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT{tab}",
                3: f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT14/ssubTABSTRIPCONTROL2SUB:SAPLMEDCMV:0100/cntlDCMGRIDCONTROL1/shellcont/shell",
                4: f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/chkMEPO1317-WEBRE",
                5: f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ",
                6: f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{i}/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:0405/tabsG_STRIP_HDR/tabpTAB{tab}",
                7: f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/cmbMEPO_TOPLINE-BSART",
                8: f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{i}/subSUB_ITEM:/COCKPIT/SAPLDISPLAY46:0410/tbl/COCKPIT/SAPLDISPLAY46G_TC_ITEM_DET/ctxt/COCKPIT/SITEM_DISP-PO_NUMBER[2,0]",
                9: f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT6/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1224/subPARTNERS:SAPLEKPA:0111/tblSAPLEKPATC_0111/txtTPART-VTEXT[1,0]",
                10: f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT1/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1226/ctxtMEPO1226-WAERS",
                11: f"wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:20{i}/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_03",
                12: f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/{middle_path_id}",
                13: f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT{tab}",
            }

            if type is None:
                if self.gui_session.findById(f"wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:0{i}/subSUB_SALDO:/COCKPIT/SAPLDISPLAY46:0440/txt/COCKPIT/SDUMMY-SALDO", False) is None:
                    return False
            else:
                # Get the path based on the type
                path = paths.get(type)

                # Check if the path (id) exists
                if self.gui_session.findById(path, False) is None:
                    return False
                
                if type in [2, 6, 13]:
                    self.gui_session.findById(path).select()
            
            return True
        
        except Exception as e:
            info = self.gui_session.findById("wnd[0]/sbar").Text
            log(e, lte.error)
            if 'The object invoked has disconnected from its clients.' in str(e):
                self.kill_sap()
            raise Exception(f"Error occured in function is_screen. SAP info: {info}")

    def line_numb(self):
        import inspect
        '''Returns the current line number in our program'''
        return inspect.currentframe().f_back.f_lineno
    