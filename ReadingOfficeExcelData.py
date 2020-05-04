import pathlib
from com.modals.sheetmodals import UseCaseModel
from com.utils.ExcelLoader import LoadExcelUtil
from com.sql.SQLdatabase import MySQLUtils
import os, fnmatch
class ExcelReaderEntry:
    def __init__(self, mainExcel):
        self.mainExcel = mainExcel
    def get_use_case_details(self, sheet_ref):
        wb = LoadExcelUtil(self.mainExcel).get_workbook_instance()
        sheet_names = wb.sheet_names()
        ctr = 0
        for sheet_name in sheet_names:
            if sheet_name == sheet_ref:
                return wb.sheets()[ctr]
            ctr = ctr + 1
    def grid_landing_sheet(self, filename):
        wb = LoadExcelUtil(self.mainExcel).get_workbook_instance()
        sheet_landing = wb.sheets()[1]
        rows = []
        for row_number in range(sheet_landing.nrows):
            no = sheet_landing.cell(row_number, 1).value  # [row_number][1]
            oms_use_case = sheet_landing.cell(row_number, 2).value
            total_use_case_count = sheet_landing.cell(row_number, 3).value
            if no != "" and type(total_use_case_count) != str:
                if int(total_use_case_count) >= 0:
                    rows.append([no, oms_use_case, total_use_case_count])
        return rows
    def get_use_case_mode(self, row_number, sheet_reference):
        use_case_model = UseCaseModel(
            TestCaseId=sheet_reference.cell(row_number, 0).value,
            Module=sheet_reference.cell(row_number, 1).value,
            CarId=sheet_reference.cell(row_number, 2).value,
            RecordingId=int(sheet_reference.cell(row_number, 3).value),
            Steeringwheel=sheet_reference.cell(row_number, 4).value,
            TripParts=sheet_reference.cell(row_number, 5).value,
            Interior=sheet_reference.cell(row_number, 6).value,
            TestPerformedby=sheet_reference.cell(row_number, 7).value,
            OccupiedSeats=sheet_reference.cell(row_number, 8).value,
            TestcaseDescription=sheet_reference.cell(row_number, 9).value,
            ObjectId=sheet_reference.cell(row_number, 10).value,
            Evaluation=sheet_reference.cell(row_number, 11).value,
            DATfile=sheet_reference.cell(row_number, 12).value,
            FalseTriggers=sheet_reference.cell(row_number, 13).value,
            Result=sheet_reference.cell(row_number, 14).value)
        return use_case_model
    @staticmethod
    def get_all_files(fpath):
        listtpathfiles=[]
        for path in os.listdir(fpath):
            full_path = os.path.join(fpath, path)
            if os.path.isfile(full_path):
                #print( full_path)
                listtpathfiles.append(full_path)
        return  listtpathfiles
        # files = []
        # for filepath in pathlib.Path(path).glob("**/*"):
        #     files.append(filepath.absolute())
        # return files
if __name__ == '__main__':
    basepath =r"C:\Users\TIRUPATHI1614131\PycharmProjects\ReadExcelProject\PySQL"
    utils = MySQLUtils()
    print('Version Info ',utils.get_version_info()[0])
    wb = LoadExcelUtil(basepath).get_workbook_instance()
    # wb = LoadExcelUtil(basepath+'\\OMS_UseCase_GoldenVideosRegression_Testing_Report_Min_R9.3.4.0-M_15Nov_session2.xlsx').get_workbook_instance()
    tester_info = wb.sheets()[0]
    if utils.get_version_info()[0] ==tester_info.cell(4,2).value:
        print("Release version is found in the DataBase")
    else:
        files = ExcelReaderEntry.get_all_files(basepath)
        #completepaths = [f for f in files]# [basepath + "/" + f.name for f in files]
        print(files)
        for p in files:
            instance = ExcelReaderEntry(p)
            object = []
            excel_sheet =instance. grid_landing_sheet(p)
            namesOfWorkBook = [workbook for workbook in excel_sheet if workbook[2] > 0]
            names = [book[0] + "_" + book[1] for book in namesOfWorkBook]
            for name in names:
                # name:'OMS_UC6_Search_Light' , 'OMS_UC9_E_Call_NoOfPassengers'
                sheet = instance.get_use_case_details(name)
                if sheet is not None:
                    for row in range(2, sheet.nrows):
                        object.append(instance.get_use_case_mode(row, sheet))

            for o in object:
                row=o.__str__()
                print(row)
