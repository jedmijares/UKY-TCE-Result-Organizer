import textract
text = (textract.process('Summer 2018 Public Results.pdf')).decode("utf-8")
# text = (textract.process('Fall 2015 - 2016 WEB.pdf')).decode("utf-8")

pages = text.split(chr(12)) # this char splits pages
# for page in pages:
#     sections = page.split('\n\n')
#     for section in sections:
#         # print(section)
#         print("----------------------------------")
#     break

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

for page in pages:
    sections = page.split('\n\n')
    try:
        courseNames = sections[0].split('\n')
        firstNames = sections[1].split('\n')
        lastNames = sections[2].split('\n')
        courseVal = sections[3].split('\n')
        instrVal = sections[4].split('\n')
        hoursStudied = sections[5].split('\n')

        for num in range(len(courseNames)):
            worksheet.write_row('A' + str(currentLine), [courseNames[num], firstNames[num], lastNames[num], courseVal[num], instrVal[num], hoursStudied[num]])
            currentLine += 1
    except:
        pass # last page

# # Some data we want to write to the worksheet.
# expenses = (
#     ['Rent', 1000],
#     ['Gas',   100],
#     ['Food',  300],
#     ['Gym',    50],
# )

# # Start from the first cell. Rows and columns are zero indexed.
# row = 0
# col = 0

# # Iterate over the data and write it out row by row.
# for item, cost in (expenses):
#     worksheet.write(row, col,     item)
#     worksheet.write(row, col + 1, cost)
#     row += 1

# # Write a total using a formula.
# worksheet.write(row, 0, 'Total')
# worksheet.write(row, 1, '=SUM(B1:B4)')

workbook.close()