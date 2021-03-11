import openpyxl
import re

class hr_calc():
    def __int__(self):
        self.count_days = list()

    def f1(self):
        wb = openpyxl.load_workbook("1.xlsx", data_only=True)
        ws = wb[wb.sheetnames[0]]

        for row in ws.iter_rows(min_row=3, ):
            yield [cell.value for cell in row if cell.value]

    def test(self):
        for day in range(31):
            self.count_days.append(day)
        print(self.count_days)
        for item in self.f1():
            for k in item:
                if type(k) == int:
                    continue
                try:
                    if re.search("до\s\d\d\.\d\d\.\d\d", k):
                        print(k)
                    elif re.search("з\s\d\d\.\d\d\.\d\d", k):
                        print(k)
                    # print(re.search("\sз\s\d\d\.\d\d\.\d\d", k).group(0))
                    # print(k)
                    # print(len(k), "==", k.index(";")+1)
                    # if len(k) == k.index(";")+1:
                    #     continue


                    # else:
                    #     print(re.findall("з", k))
                    #     print(k)
                    #     if re.search(".до.", k):
                    #         print("do")
                    #     # print(re.search("до", k))
                    #     # print(False)
                    #     # elif re.search("\sз\s", k):
                    #     #     print("З")
                    #     else:
                    #         print(k)
                except:
                    continue


hr_calc().test()

