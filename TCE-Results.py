import textract
# import wget
import urllib.request
import os
import xlsxwriter
import re
import sys

# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('TCE-Results.xlsx')
worksheet = workbook.add_worksheet()

worksheet.add_table('A1:J2', {'columns': [{'header': 'Subject'},
                                          {'header': 'Course Code'},
                                          {'header': 'Course Title'},
                                          {'header': 'First Name'},
                                          {'header': 'Last Name'},
                                          {'header': 'Year'},
                                          {'header': 'Section'},
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

        pages = text.split(chr(12)) # this character splits pages
        # for page in pages:
        #     sections = page.split('\n\n')
        #     for section in sections:
        #         # print(section)
        #         print("----------------------------------")
        #     break

        pages.pop(0) # remove first page

        for pageNumber, page in enumerate(pages):
            sections = page.split('\n\n')
            if sections != ['']:
                courseNames = sections[0].split('\n')
                years = []
                classSections = []
                courseSubjects = []
                courseCodes = []
                courseTitles = []
                for name in courseNames:
                    if ' ‐ ' in name:
                        courseCodes.append(re.search(r'\d+', name).group())
                        courseTitles.append(name.split(' ‐ ')[-1])
                        courseSubjects.append(re.sub(r'(\d+)', ' ', name).split()[0]) # convert numbers to space, then take what's before the first space
                        years.append(name.split(' ‐ ')[0].split('‐')[-1])
                        if name.split(' ‐ ')[0][-4:] == "/010":
                            classSections.append('010')
                        elif name.split(' ‐ ')[0][-4:] == "/210":
                            classSections.append('210')
                        else:
                            classSections.append(name.split(' ‐ ')[0].split('‐')[-2])
                firstNames = sections[1].split('\n')
                lastNames = sections[2].split('\n')
                courseVal = sections[3].split('\n')
                instrVal = sections[4].split('\n')
                hoursStudied = sections[5].split('\n')
                if len(years) == len(classSections) == len(courseCodes) == len(courseSubjects) == len(courseTitles) == len(firstNames) == len(lastNames) == len(courseVal) == len(instrVal) == len(hoursStudied):
                    for num in range(len(courseTitles)):
                        worksheet.write_row('A' + str(currentLine), [courseSubjects[num], courseCodes[num], courseTitles[num], firstNames[num], lastNames[num], years[num], classSections[num], courseVal[num], instrVal[num], hoursStudied[num]])
                        currentLine += 1
                else:
                    print("Trouble parsing page " + str(pageNumber + 1) + " of " + filename)
                    # print(courseNames[0])
                    # sys.exit()
                    # quit()
                    # print(len(firstNames))
                    # print(len(lastNames))
                    # print(len(courseVal))
                    # print(len(years))
                    # print(len(classSections))
                    # print("------------------")
                    pass

workbook.close()

# os.rmdir('./temp')