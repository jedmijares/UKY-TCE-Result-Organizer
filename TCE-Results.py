import textract
# import wget
import urllib.request
import os
import xlsxwriter

# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('TCE-Results.xlsx')
worksheet = workbook.add_worksheet()

worksheet.add_table('A1:F2', {'columns': [{'header': 'Course Name'},
                                          {'header': 'First Name'},
                                          {'header': 'Last Name'},
                                          {'header': 'Course Rating'},
                                          {'header': 'Instructor Rating'},
                                          {'header': 'Average Hours Studied'},
                                          ]})

currentLine = 2

urls = [
    # 'https://www.uky.edu/eval/sites/www.uky.edu.eval/files/TCE/Fall%202015%20-%202016%20WEB.pdf',
    # 'https://www.uky.edu/eval/sites/www.uky.edu.eval/files/TCE/Winter%202015%202016%20Web.pdf',
    # 'https://www.uky.edu/eval/sites/www.uky.edu.eval/files/TCE/Spring%202015-2016%20WEB.pdf',
    # 'https://www.uky.edu/eval/sites/www.uky.edu.eval/files/TCE/Fall%202016-2017%20Web.pdf',
    # 'https://www.uky.edu/eval/sites/www.uky.edu.eval/files/TCE/Winter%202016-2017%20WEB.pdf',
    # 'https://www.uky.edu/eval/sites/www.uky.edu.eval/files/TCE/Spring%202016-2017.pdf',
    # 'https://www.uky.edu/eval/sites/www.uky.edu.eval/files/TCE/Summer%20I%202016-2017%20WEB.pdf',
    # 'https://www.uky.edu/eval/sites/www.uky.edu.eval/files/TCE/Summer%20II%202016-2017%20WEB.pdf',
    # 'https://www.uky.edu/eval/sites/www.uky.edu.eval/files/TCE/Fall%202017%202018%20WEB.pdf',
    # 'https://www.uky.edu/eval/sites/www.uky.edu.eval/files/TCE/Winter%202017-2018%20WEB1.pdf',
    # 'https://www.uky.edu/eval/sites/www.uky.edu.eval/files/TCE/Spring%202018%20Public%20Results.pdf',
    # 'https://www.uky.edu/eval/sites/www.uky.edu.eval/files/TCE/Summer%202018%20Public%20Results.pdf'
]

for filename in os.listdir('./PDFs'):
    if filename.endswith('.pdf'):
        # print(filename)
        text = textract.process('./PDFs/' + filename).decode("utf-8")

        pages = text.split(chr(12)) # this char splits pages
        # for page in pages:
        #     sections = page.split('\n\n')
        #     for section in sections:
        #         # print(section)
        #         print("----------------------------------")
        #     break

        pages.pop(0) # remove first page

        for page in pages:
            sections = page.split('\n\n')
            # print(len(sections))
            try:
                courseNames = sections[0].split('\n')
                # for name in courseNames:
                #     if(name.count('-') < 2):
                #         courseNames.remove(name)
                # courseCodes[]
                courseTitles = []
                for name in courseNames:
                    courseTitles.append(name.split(' ‐ ')[-1])
                    # print(name.split('‐')[-1])
                    # sys.exit()
                print(courseTitles)

                firstNames = sections[1].split('\n')
                lastNames = sections[2].split('\n')
                courseVal = sections[3].split('\n')
                # print(courseVal)
                # for val in courseVal:
                #     try:
                #         float(val)
                #         break
                #     except:
                #         courseVal.remove(val)
                #         print(val)
                instrVal = sections[4].split('\n')
                # for val in instrVal:
                #     try:
                #         float(val)
                #         break
                #     except:
                #         instrVal.remove(val)
                #         print(val)
                hoursStudied = sections[5].split('\n')
                # for val in hoursStudied:
                #     try:
                #         float(val)
                #         break
                #     except:
                #         hoursStudied.remove(val)
                #         print(val)
                if len(courseTitles) == len(firstNames) == len(lastNames) == len(courseVal) == len(instrVal) == len(hoursStudied):
                    for num in range(len(courseTitles)):
                        worksheet.write_row('A' + str(currentLine), [courseTitles[num], firstNames[num], lastNames[num], courseVal[num], instrVal[num], hoursStudied[num]])
                        currentLine += 1
                else:
                    # print(filename)
                    # print(courseNames)
                    # print(len(firstNames))
                    # print(len(lastNames))
                    # print(len(courseVal))
                    # print(len(instrVal))
                    # print(len(hoursStudied))
                    # print("------------------")
                    pass
            except:
                pass # last page

workbook.close()

# os.rmdir('./temp')