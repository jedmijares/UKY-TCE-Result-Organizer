import textract
text = textract.process('Summer 2018 Public Results.pdf')
print(text.decode("utf-8"))