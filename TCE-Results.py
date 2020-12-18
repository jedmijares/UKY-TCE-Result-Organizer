import textract
text = (textract.process('Summer 2018 Public Results.pdf')).decode("utf-8")
# text = (textract.process('Fall 2015 - 2016 WEB.pdf')).decode("utf-8")

pages = text.split(chr(12)) # this char splits pages
for chunk in pages:
    print(chunk)
    print("000000000000000000000")