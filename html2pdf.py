import pdfkit

options = {
    'page-size': 'Letter',
    'margin-top': '0.81in',
    'margin-right': '0.9in',
    'margin-bottom': '0.9in',
    'margin-left': '0.9in',
    'encoding': "UTF-8",
    'header-right':' ',
    'custom-header' : [
        ('Accept-Encoding', 'gzip')
    ],
    '--header-spacing':8,
    'outline-depth': 10,
    'footer-center':'Recursion, Co.2020,All rights reserved',
    'footer-font-size':9,
    '--header-html':'header.html'
}
pdfkit.from_file("charts.html","result.pdf",options=options,)