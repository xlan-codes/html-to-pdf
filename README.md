This program has been developed using python 3.7.6.
Required modules:
1, matplotlib
2, pandas
3,yattag
4,pdfkit
When using pdfkit, you must set up the environment variable for that.( https://www.geeksforgeeks.org/python-convert-html-pdf/) 
After setting environment variable you have to restart your computer to fix this.

How to run these py files;
If you have already installed those modules, then 
1, First run drawcharts.py file(python drawcharts.py )
	This will produce images for the charts using matplotlib.(This will produce “images” folder in your working directory.)
	Caution: the “template_excel.xlsx” file must be in the same folder with drawcharts.py
2, if successful, run the py2html.py file.(This will produce the charts.html file)
3, then run the html2pdf.py file.(This will produce the “result.pdf” file in the working directory.)
	If you want re-run this file, at first please check whether “result.pdf” file exists or not.
	If it already exists, then please delete this file and re-run.




