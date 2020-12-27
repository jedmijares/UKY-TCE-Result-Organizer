# UKY TCE Result Organizer

This script downloads the University of Kentucky's published course evaluation data (available to students [here](https://www.uky.edu/eval/results/tce-results-students)) and places all the data in an Excel table. The original data is presented in PDF files, so doing this allows students to sort the data by course subject, rating, or instructor.

You can check out the Excel output [here](https://github.com/jedmijares/UKY-TCE-Result-Organizer/releases/tag/v1.0).

## Issues

Missing data or irregular names can cause certain pages to not parse properly. In this case, that page will only save the course subject, code, and title, along with the filename, so the student can manually look up that course if needed.

The header data on the first page of each document also messes with parsing and is not consistently formatted between files, so I simply remove the first page of each PDF.

I wasn't able to get this working on Windows, but it ran on Ubuntu.
