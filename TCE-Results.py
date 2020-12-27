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

worksheet.add_table('A1:K2', {'columns': [{'header': 'Subject'},
                                          {'header': 'Course Code'},
                                          {'header': 'Course Title'},
                                          {'header': 'First Name'},
                                          {'header': 'Last Name'},
                                          {'header': 'Year'},
                                          {'header': 'Section'},
                                          {'header': 'Course Rating'},
                                          {'header': 'Instructor Rating'},
                                          {'header': 'Average Hours Studied'},
                                          {'header': 'Filename'},
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
                    if ' ‐ ' in name: # typical
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
                    elif "(" in name: # exception of form "CS 321/MA 321/001(INTRO NUMERICAL METHODS)"
                        classSections.append(name.split("(",1)[0][-3:])
                        years.append("N/A")
                        courseCodes.append(name.split()[1].split("/")[0])
                        courseTitles.append(name.split("(",1)[1][:-1])
                        courseSubjects.append(name.split()[0])
                    else: # exception
                        courseTitles.append(name)
                        years.append("N/A")
                        classSections.append("N/A")
                        courseCodes.append("N/A")
                        courseSubjects.append("N/A")
                firstNames = sections[1].split('\n')
                lastNames = sections[2].split('\n')
                courseVal = sections[3].split('\n')
                instrVal = sections[4].split('\n')
                hoursStudied = sections[5].split('\n')
                # prevent data from becoming misaligned if data is missing
                if len(years) == len(classSections) == len(courseCodes) == len(courseSubjects) == len(courseTitles) == len(firstNames) == len(lastNames) == len(courseVal) == len(instrVal) == len(hoursStudied):
                    for num in range(len(courseTitles)):
                        worksheet.write_row('A' + str(currentLine), [courseSubjects[num], courseCodes[num], courseTitles[num], firstNames[num], lastNames[num], years[num], classSections[num], courseVal[num], instrVal[num], hoursStudied[num], filename])
                        currentLine += 1
                else:
                    print("Trouble parsing page containing " + courseNames[0].split(' ‐ ')[0] + " of " + filename + ", skipping")
                    # print(courseNames[0])
                    # print(firstNames)
                    # print(lastNames)
                    # print(courseVal)
                    # print(instrVal)
                    # print("years " + str(len(years)))
                    # print("instrVal " + str(len(instrVal)))
                    # print("courseVal " + str(len(courseVal)))
                    # print("courseTitles " + str(len(courseTitles)))
                    # print("classSections " + str(len(classSections)))
                    # print("courseSubjects " + str(len(courseSubjects)))
                    print("------------------")
                    # sys.exit()
                    # pass

workbook.close()

# os.rmdir('./temp')