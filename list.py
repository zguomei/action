import csv
import xlwt
import operator

#Author Yinglai Wang
#Date   11/08/2018


def csv_to_xlsx(csvfile):
    with open(csvfile, encoding='utf-8') as f:
        read_csv = csv.reader(f)
        workbook = xlwt.Workbook()
        sheet = workbook.add_sheet('库存状态')
        #l1 is used to store raw data
        l1 = []
        for line in read_csv:
            l1.append(line)
        first_line = str(l1[0]).replace("['","")
        first_line = first_line.replace("']","").split('，')
        #l2 is used to store rest of the content
        l2 = []
        for index in range(len(l1)):
            if len(l1[index]) > 0 and index > 0:
                l2.append(l1[index])
        l2.sort(key=operator.itemgetter(0))

        style1 = xlwt.XFStyle()
        pattern = xlwt.Pattern()
        pattern.pattern = xlwt.Pattern.SOLID_PATTERN
        pattern.pattern_fore_colour = xlwt.Style.colour_map['blue']
        font = xlwt.Font()
        font.bold = 'on'
        font.height = 260
        style1.pattern = pattern
        style1.font = font

        style2 = xlwt.XFStyle()
        pattern2 = xlwt.Pattern()
        pattern2.pattern = xlwt.Pattern.SOLID_PATTERN
        pattern2.pattern_fore_colour = xlwt.Style.colour_map['yellow']
        style2.pattern = pattern2

        style3 = xlwt.XFStyle()
        pattern3 = xlwt.Pattern()
        pattern3.pattern = xlwt.Pattern.SOLID_PATTERN
        pattern3.pattern_fore_colour = xlwt.Style.colour_map['red']
        style3.pattern = pattern3

        #write the first line to sheet
        i = 0
        for x in first_line:
            sheet.write(0, i, x, style1)
            i = i + 1
        #write the sorted content
        i = 1
        for line in l2:
            j = 0
            word_list = line[0].split('，')

            for var in word_list:
                if word_list[2] == "无货" and j == 1:
                    sheet.write(i, j, var, style2)
                    j = j + 1
                    continue
                if len(word_list) < 4 and j==1:
                    sheet.write(i, j, var, style3)
                    j = j + 1
                    continue
                sheet.write(i, j, var)
                j = j + 1
            i = i + 1


        workbook.save('workbook.xls')


if __name__ == '__main__':
    csv_to_xlsx("库存状态.csv")