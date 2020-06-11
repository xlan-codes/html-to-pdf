from yattag import Doc,indent
import pandas as pd
import os


# Read table data from charts spreadsheet
df1 = pd.read_excel(open('template_excel.xlsx', 'rb'),
              sheet_name='charts',skiprows=1,nrows=10,usecols='T:W')
df2 = pd.read_excel(open('template_excel.xlsx', 'rb'),
              sheet_name='charts',skiprows=1,nrows=10,usecols='Y:AB')

doc,tag,text,line = Doc().ttl()

cwd = os.getcwd()



def write_head():
    with tag('head'):
        doc.asis('<meta charset="utf-8">')
        doc.asis('<meta name="viewport" content="width=device-width, initial-scale=1">')
        doc.asis(' <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.4.1/css/bootstrap.min.css">')
        with tag('style'):
            text('img { padding:4px;width: 100%;height: auto;}')
            text('.row { padding: 20px;}')
            text('h3 {color: #196494}')
            text('th {color:white; height:13px !important;font-size: 10px;}')
            text('td { height:13px;font-size: 10px;}')
            text('.rectangle {height: 6.42px;width: 100%; background-color: #196494;}')
            text('.column {float: left;width: 50%;padding: 5px;}')
            text('.row::after {content: "";clear: both;display: table;}')
            text('.center {display: block;margin-left: auto;margin-right: auto;width: 50%;}')
def write_table(data):
      with tag('div',klass="row"):
            k=1
            for df in data:
                
                with tag('div', klass="column"):
                    with tag('p',style="text-align:center;font-size:7.2pt;color:#7e7e7f;font-family:Century Gothic"):
                        with tag('b'):
                            text(f'Table Title {k}')
                    k+=1
                    with tag('font',size="6", face="Century Gothic",color="#7e7e7f"):
                        with tag('table',style="width:347px",align="center"):
                            with tag('thead',style="background-color:#007bff;color:white;height:13px"):
                                for col in df.columns:
                                    with tag('th',style="text-align:center"):
                                        text(str(col))
                            for i in range(0,len(df)):
                                if i%2==0:
                                    with tag('tr',style="background-color:white"):
                                        for value in df.loc[i]:
                                            with tag('td',style="text-align:center;height=13px"):
                                                text(str(value))
                                else:
                                    with tag('tr',style="background-color:#f2f2f2"):
                                        for value in df.loc[i]:
                                            with tag('td',style="text-align:center",height="13"):
                                                text(str(value))

def write_row(nCol,nImgNums,width,height):
    if nCol==1:
        with tag('div',klass='row'):
            src = "file:///"+cwd +f"/images/chart{nImgNums[0]}.png" 
            doc.stag('img',src=src,klass="mx-auto d-block img-fluid",style=f"width:{width}px;height:{height}px;border:1px solid black",align="middle")
    else:
        with tag('div',klass="row"):
                for num in nImgNums:   
                    with tag('div',klass=f"column"):
                        src = "file:///"+cwd +f"/images/chart{num}.png" 
                        doc.stag('img',src=src,klass="mx-auto d-block img-fluid",style=f"width:{width}px;height:{height}px;border:1px solid black")

def write_body():
    with tag('body'):
          with tag('div',klass="content"):
              with tag('div', klass='row'):
                  
                  doc.stag('img', src="file:///"+cwd+'/head.png',style="border: none; width:400px;height:150px;",klass="center")
                 
                  
              with tag('p',style="font-size:16pt;color:#0f4d74;font-family:Century Gothic"):
                  with tag('b'):
                        text('Section Name 1')
              with tag('div',klass="rectangle"):
                  text("")
             
              write_row(2,[1,2],347.7,212)
              write_table([df1,df2])
              with tag('p',style="font-size:16pt;color:#0f4d74;font-family:Century Gothic"):
                  with tag('b'):
                        text('Section Name 2')
              with tag('div',klass="rectangle"):
                  text("")
    
              write_row(1,[3],716.2,247.5)
              write_row(2,[4,5],347.7,192)
              write_row(2,[6,7],347.7,192)
              with tag('p',style="font-size:16pt;color:#0f4d74;font-family:Century Gothic"):
                  with tag('b'):
                        text('Section Name 3')
              with tag('div',klass="rectangle"):
                  text("")
             
              write_row(1,[8],715,189)
              write_row(2,[9,10],347.7,192)
              with tag('p',style="font-size:16pt;color:#0f4d74;font-family:Century Gothic"):
                  with tag('b'):
                        text('Section Name 4')
              with tag('div',klass="rectangle"):
                  text("")
             
              write_row(2,[11,12],347.7,170) 
              write_row(2,[13,14],347.7,170)  
              write_row(2,[15,16],347.7,196.5) 
              with tag('p',style="font-size:16pt;color:#0f4d74;font-family:Century Gothic"):
                    with tag('b'):
                            text('Section Name 5')
              doc.stag('div',klass="rectangle")
              write_row(2,[17,18],347.7,196.5) 
              with tag('p',style="font-size:16pt;color:#0f4d74;font-family:Century Gothic"):
                  with tag('b'):
                        text('Section Name 6')
              with tag('div',klass="rectangle"):
                  text("")
             
              write_row(2,[19,20],347.7,170) 
              write_row(2,[21,22],347.7,170)
              with tag('p',style="font-size:16pt;color:#0f4d74;font-family:Century Gothic"):
                  with tag('b'):
                        text('Section Name 7')
              with tag('div',klass="rectangle"):
                  text("")
             
              write_row(1,[23],715,178.8) 
              with tag('p',style="font-size:16pt;color:#0f4d74;font-family:Century Gothic"):
                  with tag('b'):
                        text('Section Name 8')
              with tag('div',klass="rectangle"):
                  text("")
             
              write_row(2,[24,25],347.7,170) 
doc.asis('<!DOCTYPE html>')
with tag('html'): 
    write_head()
    write_body()


html_content=indent(doc.getvalue())

with open('charts.html','w') as temp_file:
    temp_file.write(html_content)
