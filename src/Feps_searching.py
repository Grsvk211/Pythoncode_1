import datetime
import sys
import ExcelInterface as EI
import os
import re
import InputConfigParser as ICF
import time
import logging
date_time = datetime.datetime.now()


def main(Feps):
    present_Feps = []
    Not_present_Feps= []
    TestPlan = EI.findInputFiles()[1]
    print("PT---->", TestPlan)
    test_Book = EI.openExcel(ICF.getInputFolder() + "\\" + TestPlan)
    test_Book.activate()
    Feps_sheet = test_Book.sheets['FEPS History']
    sheet_value = Feps_sheet.used_range.value
    Fepss = Feps.split(',')
    try:
        for Feps in Fepss:
            numeric_part = ''.join(filter(str.isdigit, Feps))
            logging.info("numeric_part--------->", numeric_part)
            Fun_name4 = EI.searchDataInColCache(sheet_value, 3, numeric_part.strip())
            logging.info("Fun_name4-------------->", Fun_name4)
            if Fun_name4['count'] > 0:
                row, col = Fun_name4['cellPositions'][0]
                time.sleep(2)
                logging.info(row, col)
                present_Feps.append(Feps)
                Impacted_sheets = EI.getDataFromCell(Feps_sheet, (row, col - 1))
                Impacted_sheets_reqs = EI.getDataFromCell(Feps_sheet, (row, col - 2))
                print("Feps--------->", Feps)
                print("Impacted_sheets ----------->", Impacted_sheets)
                print("Impacted_sheets_reqs ----------->", Impacted_sheets_reqs)
            elif Fun_name4['count'] == 0:
                Not_present_Feps.append(Feps)
        print("Feps Present in testplan Feps History sheet---------->", present_Feps)
        print("Feps not treated in testplan Feps History sheet---------->", Not_present_Feps)

        present_Feps_Summary = []
        Not_present_Feps_Summary = []
        if Not_present_Feps:
            Sommaire_sheet = test_Book.sheets['Sommaire']
            sheet_value = Sommaire_sheet.used_range.value
            for Feps in Not_present_Feps:
                logging.info("Feps--------->", Feps)
                data1 = EI.searchDataInColCache(sheet_value, 4, Feps)
                logging.info("data1------------->", data1)
                try:
                    if data1['count'] > 0:
                        roww, coll = data1['cellPositions'][0]
                        logging.info("DATA----------->", data1['cellValue'])
                        # Split each string in DATA based on '\n' and create a list of lists
                        output = [[line] for item in data1['cellValue'] for line in item.split('\n')]
                        logging.info(output)
                        filtered_list = [sublist for sublist in output if Feps in sublist[0]]
                        logging.info(filtered_list)
                        if filtered_list:
                            present_Feps_Summary.append(filtered_list)
                    elif data1['count'] == 0:
                        Not_present_Feps_Summary.append(Feps)
                except Exception as ex:
                    exc_type, exc_obj, exc_tb = sys.exc_info()
                    exp_fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
                    print(f"\nSomething went wrong {ex} line no. {exc_tb.tb_lineno} file name: {exp_fname}")
        consolidated_list = []
        for sublist in present_Feps_Summary:
            for inner_list in sublist:
                consolidated_list.extend(inner_list)
        print("Feps Present in the Summary tab of Testplan---------->", consolidated_list)
        print("Feps Not Present in the Summary tab and Feps History Tab of Testplan---------->", Not_present_Feps_Summary)
        test_Book.close()
    except Exception as ex:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        exp_fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        print(f"\nSomething went wrong {ex} line no. {exc_tb.tb_lineno} file name: {exp_fname}")


if __name__ == '__main__':
    start = time.time()
    print("Tool start time---------->", start)
    ICF.loadConfig()
    Feps = input("Enter the Feps to check in the Testplan :  ")
    print("Feps----------->", Feps)
    main(Feps)
    end1 = time.time()
    print("\nexecution time " + str(end1 - start))