import pandas as pd
import win32com.client
import sys
import json
import time
import datetime
import pythoncom
import os, shutil, stat
import logging

sys.path.append('./excel_refresh_V2/')

import config

config.log


class ExcellApp():

    def __init__(self) -> None:
        
        jobj = open(config.DIRECTORIES_PATH, encoding='utf8' )
        self.directories = json.load(jobj)
        day = datetime.datetime.now()
        self.day = datetime.datetime.now().strftime("%A")
        self.day_num = int(day.strftime("%d"))

    def enable_excel(self) -> None:
        self.xlapp = win32com.client.DispatchEx("Excel.Application")
        logging.info('\033[34m Excel App enable \033[0m')


    def close_opener(self): #uždaro excel app
        self.xlapp.Quit()

    def woorkbook_opener(self, directory:str):
        #try:
        self.wb = self.xlapp.Workbooks.Open(directory) #Atidaro failą
        self.xlapp.Visible = True #True - refreshina atsidarius, False - šešėlyje
        self.xlapp.DisplayAlerts = False # True - Reikia tvirtinti Run, False - Nereikia tvirtinti

        #self.xlapp.AskToUpdateLinks = False
        #self.xlapp.UpdateLinks = 0  # xlUpdateLinksNever
        logging.info('\033[34m workbook open \033[0m')

        #except pythoncom.com_error:
            #print("Excel app is not disable")
        #except AttributeError:
         #   print("Excel app is not disable")
          #  del self.xlapp
           # self.xlapp = win32com.client.DispatchEx("Excel.Application")
            #self.wb = self.xlapp.Workbooks.Open(directory)
    
    def refresh_all(self):
        print('start refresh')
        print(".......refresh......>>>")
        self.wb.RefreshAll()
        #self.xlapp.CalculateUntilAsyncQueriesDone()
        time.sleep(60)

    def refreshes_reports(self):
        for report in self.directories[self.day]['Report'].items():
            if 'sharepoint' in report[1]:
                self.woorkbook_opener(directory=report[1])
            else:
                file_name = f'{report[1]}'
                os.chmod(file_name, stat.S_IWRITE) #Nuima read-only
                self.woorkbook_opener(directory=report[1]) #atidaro failą
                print('read-only off')
            self.refresh_all() #refreshina ataskaita
            print(f'{report[0]} refresh complete')


            t = time.localtime()
            current_time = time.strftime("%Y-%m-%d %H:%M:%S", t)

            self.close_file(True)

            if 'sharepoint' in report[1]:
                pass
            else:
                os.chmod(file_name, stat.FILE_ATTRIBUTE_READONLY)
                print('Read-only on')
        
    def close_file(self, save): #uždaro failą
            self.wb.Close(True)
    
    def remove_excel(self):
        del self.xlapp

    def refresh_report(self, report_url):
        if 'sharepoint' in report_url:
            self.woorkbook_opener(directory=report_url)
        else:
            file_name = f'{report_url}'
            os.chmod(file_name, stat.S_IWRITE) #Nuima read-only
            self.woorkbook_opener(directory=report_url) #atidaro failą
            print('read-only off')
        self.refresh_all() #refreshina ataskaita
        print(f'refresh complete')
        t = time.localtime()
        current_time = time.strftime("%Y-%m-%d %H:%M:%S", t)
        self.close_file(True)
        if 'sharepoint' in report_url:
            pass
        else:
            os.chmod(file_name, stat.FILE_ATTRIBUTE_READONLY)
            print('Read-only on')

    def start_refresh():
        win = ExcellApp()
        try:
            win.enable_excel()
            win.refreshes_reports()
        except Exception as e:
            print(f"Error: {e}")
        finally:
            try:
                win.close_opener()
                win.remove_excel()
            except:
                pass

if "__main__" == __name__:
    win = ExcellApp()
    try:
        win.enable_excel()
        win.refreshes_reports()
    except Exception as e:
        print(f"Error: {e}")
    finally:
        try:
            win.close_opener()
            win.remove_excel()
        except:
            pass
