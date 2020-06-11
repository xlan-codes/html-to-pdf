import pandas as pd
import os
import numpy as np
import matplotlib as mpl
import matplotlib.pyplot as plt
from matplotlib.figure import figaspect
from matplotlib.figure import Figure
from pandas.plotting import register_matplotlib_converters
register_matplotlib_converters()


# Where to save the figures
PROJECT_ROOT_DIR = "."
IMAGES_PATH = os.path.join(PROJECT_ROOT_DIR, "images")
os.makedirs(IMAGES_PATH, exist_ok=True)

# Say, "the default sans-serif font is COMIC SANS"
mpl.rcParams['font.sans-serif'] = "Century Gothic"
# Then, "ALWAYS use sans-serif fonts"
mpl.rcParams['font.family'] = "sans-serif"

mpl.font_manager._rebuild()

#-------------------------------Chart 1 -----------------------------#
df = pd.read_excel(open('template_excel.xlsx', 'rb'),
              sheet_name='chart1data') 

fig = plt.figure(linewidth=1, edgecolor="#04253a")
ax1 = plt.subplot()

ax1.stackplot(df['date'],[df['Fed Holding'],df['Bank Holding']]
        ,labels=['Fed Holding','Bank Holding'],colors=['#8fd7f1','#2fa8e5'],alpha=0.8)
ax1.plot(df['date'],df['UPB'],color='#95aa00',linewidth='2.3',label='UPB')
#--range both axes 
ax1.set_ylim(0,2500)
ax1.set_xlim(0,len(plt.xticks()[0]))
#--set the xticks label rotate with step 3
for label in ax1.get_xticklabels() :
  label.set_rotation(90)
  label.set_ha('center')
ax1.set_xticks(np.arange(0,len(plt.xticks()[0]),3))
ax1.tick_params(axis=u'both',which=u'both',length=0,pad=5,labelsize=7)  
ax1.spines['top'].set_visible(False)
ax1.spines['right'].set_visible(False)
ax1.get_yaxis().set_major_formatter(
     mpl.ticker.FuncFormatter(lambda x, p: format(int(x), ',')))
ax1.set_axisbelow(True)
ax1.grid(axis='y',color="#f2f2f2",zorder=0)
handles, labels = ax1.get_legend_handles_labels()
ax1.legend(handles,labels,loc="upper center",
             ncol=4,bbox_to_anchor=(0.5,-0.4),frameon=False,prop={'size': 10},handlelength=4)
ax1.tick_params(axis='both',color='#7e7e7f')
for tick in ax1.get_xticklabels():
    tick.set_fontsize(10)
    
for tick in ax1.get_yticklabels():
     tick.set_fontsize(10)
[i.set_color("#7e7e7f") for i in ax1.get_xticklabels()]
[i.set_color("#7e7e7f") for i in ax1.get_yticklabels()]
plt.title('Chart Title 1', fontsize=14,pad=20,color='#7e7e7f',fontweight="bold")

plt.tight_layout()


fig.savefig( 'images/chart1.png', bbox_inches='tight') 

#-------------------------------------Chart 2-------------------------------------------#

df = pd.read_excel(open('template_excel.xlsx', 'rb'),
              sheet_name='chart2data') 

fig = plt.figure(linewidth=1, edgecolor="#04253a",figsize=(6.4,4.8))
ax1 = plt.subplot()

ax1.stackplot(df['date'],[df['Fed Holding'],df['Bank Holding'],df['Foreign Holding']]
        ,labels=['Fed Holding','Bank Holding','Foreign Holding'],colors=['#8fd7f1','#2fa8e5','#238acf'],alpha=0.8)
ax1.plot(df['date'],df['UPB'],color='#95aa00',linewidth='2.3',label='UPB')
#--range both axes 
ax1.set_ylim(0,8000)
ax1.set_xlim(-1,len(plt.xticks()[0]))
#--set the xticks label rotate with step 3
for label in ax1.get_xticklabels() :
  label.set_rotation(90)
  label.set_ha('center')
ax1.set_xticks(np.arange(0,len(plt.xticks()[0]),3))
ax1.tick_params(axis=u'both',which=u'both',length=0,pad=5,labelsize=7)  
ax1.spines['top'].set_visible(False)
ax1.spines['right'].set_visible(False)
ax1.get_yaxis().set_major_formatter(
     mpl.ticker.FuncFormatter(lambda x, p: format(int(x), ',')))
ax1.set_axisbelow(True)
ax1.grid(axis='y',color="#f2f2f2",zorder=0)
handles, labels = ax1.get_legend_handles_labels()
ax1.legend(handles,labels,loc="upper center",
             ncol=4,bbox_to_anchor=(0.5,-0.4),frameon=False,prop={'size': 7},handlelength=4)
ax1.tick_params(axis='both',color='#7e7e7f')
for tick in ax1.get_xticklabels():
    tick.set_fontsize(10)
    
for tick in ax1.get_yticklabels():
     tick.set_fontsize(10)
plt.title('Chart Title 2',fontsize=14,pad=20,color='#7e7e7f',fontweight="bold")
[i.set_color("#7e7e7f") for i in ax1.get_xticklabels()]
[i.set_color("#7e7e7f") for i in ax1.get_yticklabels()]
plt.tight_layout()

fig.savefig( 'images/chart2.png', bbox_inches='tight') 

#--------------------------------------Chart 3---------------------------------------------#
df = pd.read_excel(open('template_excel.xlsx', 'rb'),
              sheet_name='chart3data') 

fig = plt.figure(linewidth=1, edgecolor="#04253a",figsize=(15,4.8))

plt.plot(df['fctrdt'],df['All'],color='#95daf3',linestyle=':',linewidth='2.3',label="All")
plt.plot(df['fctrdt'],df['VA'],color='#96a816',label="VA",linewidth='2.3')
plt.plot(df['fctrdt'],df['FHA'],color='#238ace',label="FHA",linewidth='2.3')
[i.set_color("#7e7e7f") for i in plt.gca().get_xticklabels()]
#--range both axes 
plt.ylim(0,50)
plt.xlim(-1,len(plt.xticks()[0]))
#--set the xticks label rotate with step 3
plt.xticks(ticks=None,rotation=90)
plt.xticks(np.arange(0,len(plt.xticks()[0]),1))
plt.yticks(np.arange(0,55,5))
plt.grid(axis='y',color="#f2f2f2")
#--reorder the legends
handles, labels = plt.subplot().get_legend_handles_labels()
plt.legend(handles,labels,loc="upper center",
             ncol=3,bbox_to_anchor=(0.5,-0.4),frameon=False,prop={'size': 9},handlelength=5)
#--make the top and right border invisible
plt.subplot().spines['right'].set_visible(False)
plt.subplot().spines['top'].set_visible(False)
#--remove ticks from both axes
plt.tick_params(axis=u'both',which=u'both',length=0,pad=10)  
#--formate the number style of ticks
plt.subplot().get_yaxis().set_major_formatter(
    mpl.ticker.FuncFormatter(lambda x, p: format(int(x), ',')))
[i.set_color("#7e7e7f") for i in ax1.get_yticklabels()]
for tick in ax1.get_xticklabels():
    tick.set_fontsize(12)
    
for tick in ax1.get_yticklabels():
     tick.set_fontsize(12)
plt.title('Chart Title 3', fontsize=14,pad=20,color='#7e7e7f',fontweight="bold")
plt.tight_layout()
print(fig.get_size_inches())
fig.savefig( 'images/chart3.png', bbox_inches='tight') 

#----------------------------------Chart title 4---------------------------#
fig = plt.figure(linewidth=1, edgecolor="#04253a",figsize=(3.62205,2))
df = pd.read_excel(open('template_excel.xlsx', 'rb'),
              sheet_name='chart4data') 
df.sort_values('ratio_to_report',inplace=True)
ax1 = plt.subplot()
ax1.set_xlim(0,14.1)
ax1.barh(df['fctrdt'],df['ratio_to_report'],0.3)
ax1.set_xticks(np.arange(0,16,2))
ax1.grid(axis='x',color="#f2f2f2",zorder=0)
[i.set_color("red") for i in plt.gca().get_xticklabels()]
#--make the top and right border invisible
ax1.spines['right'].set_visible(False)
ax1.spines['top'].set_visible(False)
#--remove ticks from both axes
ax1.tick_params(axis=u'both',which=u'both',length=0,pad=10,labelsize=6) 

ax1.set_axisbelow(True)
[i.set_color("#7e7e7f") for i in ax1.get_xticklabels()]
[i.set_color("#7e7e7f") for i in ax1.get_yticklabels()]
plt.tight_layout()
for tick in ax1.get_xticklabels():
    tick.set_fontsize(6)
    
for tick in ax1.get_yticklabels():
     tick.set_fontsize(6)
plt.title('Chart Title 4', fontsize=7.2,pad=20,color='#7e7e7f',fontweight="bold")

fig.savefig( 'images/chart4.png', bbox_inches='tight') 
#-------------------------------------Chart title 5--------------------------#
  
fig = plt.figure(linewidth=1, edgecolor="#04253a",figsize=(3.62205,2))

df = pd.read_excel(open('template_excel.xlsx', 'rb'),
              sheet_name='chart5data') 
df.sort_values('ratio_to_report',inplace=True)
ax1=plt.subplot()
ax1.set_xlim(0,18.1)

ax1.barh(df['fctrdt'],df['ratio_to_report'],0.3,color="#96a816")
ax1.set_xticks(np.arange(0,20,2))
[i.set_color("#7e7e7f") for i in ax1.get_xticklabels()]
ax1.grid(axis='x',color="#f2f2f2",zorder=0)
#--make the top and right border invisible
ax1.spines['right'].set_visible(False)
ax1.spines['top'].set_visible(False)
#--remove ticks from both axes
ax1.tick_params(axis=u'both',which=u'both',length=0,pad=10,labelsize=6) 
ax1.set_axisbelow(True)
for tick in ax1.get_xticklabels():
    tick.set_fontsize(6)
    
for tick in ax1.get_yticklabels():
     tick.set_fontsize(6)
plt.title('Chart Title 5', fontsize=7.2,pad=20,color='#7e7e7f',fontweight="bold")
[i.set_color("red") for i in ax1.get_xticklabels()]
[i.set_color("#7e7e7f") for i in ax1.get_yticklabels()]
plt.tight_layout()

fig.savefig( 'images/chart5.png', bbox_inches='tight') 
#----------------------------------------Chart title 6------------------------------------#
fig = plt.figure(linewidth=1, edgecolor="#04253a",figsize=(3.62205,2))
ax1 = plt.subplot()
df = pd.read_excel(open('template_excel.xlsx', 'rb'),
              sheet_name='chart6data') 
df.sort_values('bal_to',inplace=True)
ax1.set_xlim(0,81)
ax1.set_axisbelow(True)
ax1.barh(df['name'],df['bal_to'],0.3,color="#238ace")
ax1.set_xticks(np.arange(0,90,10))
plt.grid(axis='x',color="#f2f2f2",zorder=0)
#--make the top and right border invisible
ax1.spines['right'].set_visible(False)
ax1.spines['top'].set_visible(False)
#--remove ticks from both axes
plt.tick_params(axis=u'both',which=u'both',length=0,pad=10,labelsize=6) 
for tick in ax1.get_xticklabels():
    tick.set_fontsize(6)
    
for tick in ax1.get_yticklabels():
     tick.set_fontsize(6)
[i.set_color("#7e7e7f") for i in ax1.get_xticklabels()]
[i.set_color("#7e7e7f") for i in ax1.get_yticklabels()]
plt.title('Chart Title 6', fontsize=7.2,pad=20,color='#7e7e7f',fontweight="bold")
plt.tight_layout()
fig.savefig( 'images/chart6.png', bbox_inches='tight') 


#------------------------------------Chart title 7-----------------------------------#
fig = plt.figure(linewidth=1, edgecolor="#04253a",figsize=(3.62205,2))

df = pd.read_excel(open('template_excel.xlsx', 'rb'),
              sheet_name='chart7data') 
df.sort_values('bal_to',inplace=True)
ax1 = plt.subplot()
ax1.set_xlim(0,25.1)

ax1.barh(df['name'],df['bal_to'],0.3,color="#96a816")
ax1.set_xticks(np.arange(0,30,5))
plt.grid(axis='x',color="#f2f2f2",zorder=0)
#--make the top and right border invisible
ax1.set_axisbelow(True)
ax1.spines['right'].set_visible(False)
ax1.spines['top'].set_visible(False)
#--remove ticks from both axes
plt.tick_params(axis=u'both',which=u'both',length=0,pad=10,labelsize=6) 
for tick in ax1.get_xticklabels():
    tick.set_fontsize(6)
    
for tick in ax1.get_yticklabels():
     tick.set_fontsize(6)
plt.title('Chart Title 7', fontsize=7.2,pad=20,color='#7e7e7f',fontweight="bold")
[i.set_color("#7e7e7f") for i in ax1.get_xticklabels()]
[i.set_color("#7e7e7f") for i in ax1.get_yticklabels()]
plt.tight_layout()
fig.savefig( 'images/chart7.png', bbox_inches='tight') 

#-------------------Chart 8 -----------------------#

df = pd.read_excel(open('template_excel.xlsx', 'rb'),
              sheet_name='chart8data') 
fig = plt.figure(linewidth=2, edgecolor="#04253a",frameon=True,figsize=(17,6.5))
ax1 = plt.subplot()

ax1.set_xlabel('Date')
ax1.set_ylabel('%yield')
l1 = ax1.plot(df['cpn_date'],df['G2SF CC % yield'],color='#2fa7e4',label="G2SF CC % yield",linewidth=2.3)
ax1.set_ylim(0,4.5)
ax1.tick_params(axis=u'both',which=u'both',length=0,pad=5,labelsize=7)

ax1.set_xticks(range(0,len(ax1.get_xticklabels()),7))
ax1.set_xticklabels(df['cpn_date'][::7])
ax1.spines['top'].set_visible(False)
ax2 = ax1.twinx()
ax2.set_xlabel('Date')
ax2.set_ylabel('spread bps')
l2 = ax2.plot(df['cpn_date'],df['G2SF CC % yield nominal spread bps vs UST 5/10 blend(50%/50%)'],
    color='#96a816',label="G2SF CC % yield nominal spread bps vs UST 5/10 blend(50%/50%)",linewidth=2)
ax2.set_ylim(0,90)
ax2.set_yticklabels(range(0,91,10))
ax2.set_yticks(range(0,91,10))
ax2.tick_params(axis=u'both',which=u'both',length=0,pad=5,labelsize=7)
ax2.set_xticks(range(0,len(ax2.get_xticklabels()),7))
ax2.set_xticklabels(df['cpn_date'][::7])
ax2.spines['top'].set_visible(False)


for label in ax1.get_xticklabels() + ax2.get_xticklabels():
  label.set_rotation(90)
  label.set_ha('center')

plt.setp(ax1,xlim=(-1,len(df['cpn_date'])))
plt.setp(ax2,xlim=(-1,len(df['cpn_date'])))



#--range both axes 
# plt.setp(ax1.get_xticklabels(), rotation=90, ha='right')
plt.grid(axis='y',color="#f2f2f2")

line1,label1 = ax1.get_legend_handles_labels()
line2,label2 = ax2.get_legend_handles_labels()

ax2.legend(line1+line2, label1+label2,loc="lower center", bbox_to_anchor=(0.5, -0.7)
                ,columnspacing=3, ncol=4,frameon=False,prop={'size': 9},handlelength=5)

#----space between ticks label----#
plt.gca().margins(x=0)
plt.gcf().canvas.draw()

t1 = plt.gca().get_xticklabels()
maxsize = max([t.get_window_extent(renderer=fig.canvas.get_renderer()).width for t in t1])
m = 0.2 # inch margin
s = maxsize/plt.gcf().dpi*len(df['cpn_date'])/3+2*m
margin = m/plt.gcf().get_size_inches()[0]

plt.gcf().subplots_adjust(left=margin, right=1.-margin)
plt.gcf().set_size_inches(s, plt.gcf().get_size_inches()[1])

for tick in ax1.get_xticklabels():
    tick.set_fontsize(12)
    
for tick in ax1.get_yticklabels():
     tick.set_fontsize(12)

for tick in ax2.get_yticklabels():
     tick.set_fontsize(12)
plt.title('Chart Title 8', fontsize=15,pad=20,color='#7e7e7f',fontweight="bold")
[i.set_color("#7e7e7f") for i in ax1.get_xticklabels()]
[i.set_color("#7e7e7f") for i in ax1.get_yticklabels()]
[i.set_color("#7e7e7f") for i in ax2.get_yticklabels()]
plt.tight_layout()
fig.savefig( 'images/chart8.png', bbox_inches='tight') 

#------------------------------------Chart 9--------------------------------#
fig = plt.figure(linewidth=1, edgecolor="#04253a",figsize=(8,6))
df = pd.read_excel(open('template_excel.xlsx', 'rb'),
              sheet_name='chart10data') 

ax1 = plt.subplot()
[i.set_color("#7e7e7f") for i in ax1.get_xticklabels()]
ax1.plot(df['fctrdt'],df[df.columns[1]],color='#a6e1f5',linestyle=':',label="3",linewidth='2.3')
ax1.plot(df['fctrdt'],df[df.columns[2]],color='#2fa7e4',label="3.5",linewidth='2.3')
ax1.plot(df['fctrdt'],df[df.columns[3]],color='#2389ce',label="4",linewidth='2.3')
ax1.plot(df['fctrdt'],df[df.columns[4]],color='#96a916',label="4.5",linewidth='2.3')


ax1.tick_params(axis=u'both',which=u'both',length=0,pad=5,labelsize=7)
ax1.set_xticks(df['fctrdt'][::2])
ax1.set_ylim(0,45)
ax1.set_yticklabels(range(0,46,5))

ax1.spines['top'].set_visible(False)
ax1.spines['right'].set_visible(False)
# ax1.set_xticklabels(df['fctrdt'])
for label in ax1.get_xticklabels() :
  label.set_rotation(90)

line1,label1 = ax1.get_legend_handles_labels()

ax1.legend(line1, label1,loc="lower center", bbox_to_anchor=(0.5, -0.7)
                ,columnspacing=2.5, ncol=4,frameon=False,prop={'size': 9},handlelength=5)

ax1.set_xlim(left=df['fctrdt'][0])

plt.grid(axis='y',color="#f2f2f2")
for tick in ax1.get_xticklabels():
    tick.set_fontsize(12)
    
for tick in ax1.get_yticklabels():
     tick.set_fontsize(12)
plt.title('Chart Title 9', fontsize=15,pad=20,color='#7e7e7f',fontweight="bold")
[i.set_color("#7e7e7f") for i in ax1.get_yticklabels()]
plt.tight_layout()
fig.savefig( 'images/chart9.png', bbox_inches='tight')

#---------------------------------------------chart 10 -----------#
df = pd.read_excel(open('template_excel.xlsx', 'rb'),
              sheet_name='chart10data') 

fig = plt.figure(linewidth=1, edgecolor="#04253a",figsize=(8,6))

df = pd.read_excel(open('template_excel.xlsx', 'rb'),
              sheet_name='chart10data') 

ax1 = plt.subplot()
ax1.plot(df['fctrdt'],df[df.columns[1]],color='#a6e1f5',linestyle=':',label="3",linewidth='2.3')
ax1.plot(df['fctrdt'],df[df.columns[2]],color='#2fa7e4',label="3.5",linewidth='2.3')
ax1.plot(df['fctrdt'],df[df.columns[3]],color='#2389ce',label="4",linewidth='2.3')
ax1.plot(df['fctrdt'],df[df.columns[4]],color='#96a916',label="4.5",linewidth='2.3')


ax1.tick_params(axis=u'both',which=u'both',length=0,pad=5,labelsize=7)
ax1.set_xticks(df['fctrdt'][::2])
ax1.set_ylim(0,70)


ax1.spines['top'].set_visible(False)
ax1.spines['right'].set_visible(False)
# ax1.set_xticklabels(df['fctrdt'])
for label in ax1.get_xticklabels() :
  label.set_rotation(90)

line1,label1 = ax1.get_legend_handles_labels()

ax1.legend(line1, label1,loc="lower center", bbox_to_anchor=(0.5, -0.7)
                ,columnspacing=2.5, ncol=4,frameon=False,prop={'size': 9},handlelength=5)
print(ax1.get_xticklabels())
# ax1.get_xticklabels[0]
ax1.set_xlim(left=df['fctrdt'][0])
[i.set_color("#7e7e7f") for i in ax1.get_xticklabels()]
plt.grid(axis='y',color="#f2f2f2")
for tick in ax1.get_xticklabels():
    tick.set_fontsize(12)
    
for tick in ax1.get_yticklabels():
     tick.set_fontsize(12)
plt.title('Chart Title 10', fontsize=15,pad=20,color='#7e7e7f',fontweight="bold")
[i.set_color("#7e7e7f") for i in ax1.get_yticklabels()]
plt.tight_layout()
fig.savefig( 'images/chart10.png', bbox_inches='tight') 

#----------------------------------------------------chart 11 --------------------#
df = pd.read_excel(open('template_excel.xlsx', 'rb'),
              sheet_name='chart11data') 

fig = plt.figure(linewidth=1, edgecolor="#04253a")
ax1 = plt.subplot()
ax1.plot(df['fctrdt'],df[df.columns[1]],color='#2389ce',linewidth=2.3,label="FHL")
ax1.plot(df['fctrdt'],df[df.columns[2]],color='#96a916',linewidth=2.3,label="FNM")
ax1.plot(df['fctrdt'],df[df.columns[3]],color='#a5a5a5',linewidth=2.3,label="GNM")
ax1.tick_params(axis=u'both',which=u'both',length=0,pad=5,labelsize=7)
ax1.set_xticks(df['fctrdt'])
ax1.set_ylim(0,25000)
ax1.spines['top'].set_visible(False)
ax1.spines['right'].set_visible(False)

for label in ax1.get_xticklabels() :
  label.set_rotation(45)
  label.set_ha('right')

line1,label1 = ax1.get_legend_handles_labels()
ax1.legend(line1, label1,loc="lower center", bbox_to_anchor=(0.5, -0.6)
                ,columnspacing=2.5, ncol=4,frameon=False,prop={'size': 9},handlelength=5)
plt.subplot().get_yaxis().set_major_formatter(
    mpl.ticker.FuncFormatter(lambda x, p: format(int(x), ',')))

plt.grid(axis='y',color="#f2f2f2")
for tick in ax1.get_xticklabels():
    tick.set_fontsize(12)
    
for tick in ax1.get_yticklabels():
     tick.set_fontsize(12)
plt.title('Chart Title 11', fontsize=15,pad=20,color='#7e7e7f',fontweight="bold")
[i.set_color("#7e7e7f") for i in ax1.get_yticklabels()]
[i.set_color("#7e7e7f") for i in ax1.get_xticklabels()]
plt.tight_layout()
fig.savefig( 'images/chart11.png', bbox_inches='tight') 


#--------------------chart 12------------------------#

fig = plt.figure(linewidth=1, edgecolor="#04253a")
df = pd.read_excel(open('template_excel.xlsx', 'rb'),
              sheet_name='chart12data') 

ax1 = plt.subplot()
ax1.plot(df['fctrdt'],df[df.columns[1]],color='#2389ce',linewidth=2.3,label="FHL")
ax1.plot(df['fctrdt'],df[df.columns[2]],color='#96a916',linewidth=2.3,label="FNM")
ax1.plot(df['fctrdt'],df[df.columns[3]],color='#a5a5a5',linewidth=2.3,label="GNM")
ax1.tick_params(axis=u'both',which=u'both',length=0,pad=5,labelsize=7)
ax1.set_xticks(df['fctrdt'])
ax1.set_ylim(0,80000)
ax1.spines['top'].set_visible(False)
ax1.spines['right'].set_visible(False)
# ax1.set_xticklabels(df['fctrdt'])
for label in ax1.get_xticklabels() :
  label.set_rotation(45)
  label.set_ha('right')

line1,label1 = ax1.get_legend_handles_labels()
ax1.legend(line1, label1,loc="lower center", bbox_to_anchor=(0.5, -0.5)
                ,columnspacing=2.5, ncol=4,frameon=False,prop={'size': 9},handlelength=5)
plt.subplot().get_yaxis().set_major_formatter(
    mpl.ticker.FuncFormatter(lambda x, p: format(int(x), ',')))

plt.grid(axis='y',color="#f2f2f2")
for tick in ax1.get_xticklabels():
    tick.set_fontsize(12)
    
for tick in ax1.get_yticklabels():
     tick.set_fontsize(12)
plt.title('Chart Title 12', fontsize=15,pad=20,color='#7e7e7f',fontweight="bold")
[i.set_color("#7e7e7f") for i in ax1.get_xticklabels()]
[i.set_color("#7e7e7f") for i in ax1.get_yticklabels()]
plt.tight_layout()
fig.savefig( 'images/chart12.png', bbox_inches='tight') 

#-----------------------------------------------------#

#--------------------------------------------Chart 13 -------------------------#
df = pd.read_excel(open('template_excel.xlsx', 'rb'),
              sheet_name='chart13data') 

fig = plt.figure(linewidth=1, edgecolor="#04253a")
ax1 = plt.subplot()
ax1.plot(df['fctrdt'],df[df.columns[1]],color='#2389ce',linewidth=2.3,label="FHA")
ax1.plot(df['fctrdt'],df[df.columns[2]],color='#96a916',linewidth=2.3,label="VA")

ax1.tick_params(axis=u'both',which=u'both',length=0,pad=5,labelsize=7)
ax1.set_xticks(df['fctrdt'])
ax1.set_ylim(0,10000)
ax1.spines['top'].set_visible(False)
ax1.spines['right'].set_visible(False)
# ax1.set_xticklabels(df['fctrdt'])
for label in ax1.get_xticklabels() :
  label.set_rotation(45)
  label.set_ha('right')

line1,label1 = ax1.get_legend_handles_labels()
ax1.legend(line1, label1,loc="lower center", bbox_to_anchor=(0.5, -0.5)
                ,columnspacing=2.5, ncol=4,frameon=False,prop={'size': 9},handlelength=5)
plt.subplot().get_yaxis().set_major_formatter(
    mpl.ticker.FuncFormatter(lambda x, p: format(int(x), ',')))

plt.grid(axis='y',color="#f2f2f2")
for tick in ax1.get_xticklabels():
    tick.set_fontsize(12)
    
for tick in ax1.get_yticklabels():
     tick.set_fontsize(12)
plt.title('Chart Title 13', fontsize=14,pad=20,color='#7e7e7f',fontweight="bold")
[i.set_color("#7e7e7f") for i in ax1.get_xticklabels()]
[i.set_color("#7e7e7f") for i in ax1.get_yticklabels()]
plt.tight_layout()
fig.savefig( 'images/chart13.png', bbox_inches='tight') 

#----------------------------------------------------#

#-------------------------------------------------------------chart 14 -------------------------------------------------#
df = pd.read_excel(open('template_excel.xlsx', 'rb'),
              sheet_name='chart14data') 

fig = plt.figure(linewidth=1, edgecolor="#04253a")
ax1 = plt.subplot()
ax1.plot(df['fctrdt'],df[df.columns[1]],color='#2389ce',linewidth=2.3,label="FHA Jumbo %")
ax1.plot(df['fctrdt'],df[df.columns[2]],color='#96a916',linewidth=2.3,label="VA Jumbo %")
ax1.plot(df['fctrdt'],df[df.columns[3]],color='#2fa7e4',linewidth=2.3,linestyle=':',label="Total Jumbo %")
ax1.tick_params(axis=u'both',which=u'both',length=0,pad=5,labelsize=7)
ax1.set_xticks(df['fctrdt'])
ax1.set_ylim(0,14)
ax1.spines['top'].set_visible(False)
ax1.spines['right'].set_visible(False)

for label in ax1.get_xticklabels() :
  label.set_rotation(45)
  label.set_ha('right')

line1,label1 = ax1.get_legend_handles_labels()
ax1.legend(line1, label1,loc="lower center", bbox_to_anchor=(0.5, -0.5)
                ,columnspacing=2.5, ncol=4,frameon=False,prop={'size': 9},handlelength=7)
plt.subplot().get_yaxis().set_major_formatter(
    mpl.ticker.FuncFormatter(lambda x, p: format(int(x), ',')))

plt.grid(axis='y',color="#f2f2f2")
for tick in ax1.get_xticklabels():
    tick.set_fontsize(12)
    
for tick in ax1.get_yticklabels():
     tick.set_fontsize(12)
plt.title('Chart Title 14', fontsize=14,pad=20,color='#7e7e7f',fontweight="bold")
[i.set_color("#7e7e7f") for i in ax1.get_xticklabels()]
[i.set_color("#7e7e7f") for i in ax1.get_yticklabels()]
plt.tight_layout()
fig.savefig( 'images/chart14.png', bbox_inches='tight') 
#----------------------------------------------------------------#

#--------------------------------Chart 15 -------------------------#
df = pd.read_excel(open('template_excel.xlsx', 'rb'),
              sheet_name='chart15data') 
fig = plt.figure(linewidth=1, edgecolor="#04253a")

ax1 = plt.subplot()
ax1.plot(df['fctrdt'],df[df.columns[1]],color='#104d73',linewidth=2.3,label="% LTV>90")
ax1.plot(df['fctrdt'],df[df.columns[2]],color='#2fa7e4',linewidth=2.3,label="% LTV>90")
ax1.plot(df['fctrdt'],df[df.columns[3]],color='#2389ce',linewidth=2.3,label="% LTV>90")
ax1.tick_params(axis=u'both',which=u'both',length=0,pad=5,labelsize=7)
ax1.set_xticks(df['fctrdt'][::2])
ax1.set_ylim(0,80)
ax1.spines['top'].set_visible(False)
ax1.spines['right'].set_visible(False)
ax1.set_xlim(left=df['fctrdt'][0])

for label in ax1.get_xticklabels() :
  label.set_rotation(90)
  label.set_ha('right')

line1,label1 = ax1.get_legend_handles_labels()
ax1.legend(line1, label1,loc="lower center", bbox_to_anchor=(0.5, -0.6)
                ,columnspacing=2.5, ncol=4,frameon=False,prop={'size': 9},handlelength=7)
plt.subplot().get_yaxis().set_major_formatter(
    mpl.ticker.FuncFormatter(lambda x, p: format(int(x), ',')))

plt.grid(axis='y',color="#f2f2f2")

for tick in ax1.get_xticklabels():
    tick.set_fontsize(10)
    
for tick in ax1.get_yticklabels():
     tick.set_fontsize(10)
plt.title('Chart Title 15', fontsize=13,pad=20,color='#7e7e7f',fontweight="bold")
[i.set_color("#7e7e7f") for i in ax1.get_xticklabels()]
[i.set_color("#7e7e7f") for i in ax1.get_yticklabels()]
plt.tight_layout()
fig.savefig( 'images/chart15.png', bbox_inches='tight') 
#----------------------------------------------------------------------------#

#-----------------------Chart 16-----------------------------------#


fig = plt.figure(linewidth=1, edgecolor="#04253a",figsize=(9,6.4))
df = pd.read_excel(open('template_excel.xlsx', 'rb'),
              sheet_name='chart16data') 

ax1 = plt.subplot()
ax1.plot(df['fctrdt'],df[df.columns[1]],color='#95dbf3',linewidth=2.3,label="GSE avg cs")
ax1.plot(df['fctrdt'],df[df.columns[2]],color='#ffc000',linewidth=2.3,label="GSE median cs")
ax1.plot(df['fctrdt'],df[df.columns[3]],color='#104d73',linewidth=2.3,label="FHA avg cs")
ax1.plot(df['fctrdt'],df[df.columns[4]],color='#2fa7e4',linewidth=2.3,label="FHA median cs")
ax1.plot(df['fctrdt'],df[df.columns[5]],color='#96a916',linewidth=2.3,label="VA avg cs")
ax1.plot(df['fctrdt'],df[df.columns[6]],color='#56bff0',linewidth=2.3,label="VA median cs")
ax1.tick_params(axis=u'both',which=u'both',length=0,pad=5,labelsize=7)
ax1.set_xticks(df['fctrdt'][::2])
ax1.set_ylim(600,800)
ax1.set_yticks(range(600,801,50))

ax1.spines['top'].set_visible(False)
ax1.spines['right'].set_visible(False)
ax1.set_xlim(left=df['fctrdt'][0])

for label in ax1.get_xticklabels() :
  label.set_rotation(90)
  label.set_ha('left')

line1,label1 = ax1.get_legend_handles_labels()
ax1.legend(line1, label1,loc="lower center", bbox_to_anchor=(0.5, -0.7)
                ,columnspacing=2.5, ncol=3,frameon=False,prop={'size': 9},handlelength=7)
plt.subplot().get_yaxis().set_major_formatter(
    mpl.ticker.FuncFormatter(lambda x, p: format(int(x), ',')))

plt.grid(axis='y',color="#f2f2f2")

for tick in ax1.get_xticklabels():
    tick.set_fontsize(12)
    
for tick in ax1.get_yticklabels():
     tick.set_fontsize(12)
plt.title('Chart Title 16', fontsize=15,pad=20,color='#7e7e7f',fontweight="bold")
[i.set_color("#7e7e7f") for i in ax1.get_xticklabels()]
[i.set_color("#7e7e7f") for i in ax1.get_yticklabels()]
plt.tight_layout()
fig.savefig( 'images/chart16.png', bbox_inches='tight') 
#----------------------------------------------------------------------------#

#------------------------Chart 17------------------------------------#
df = pd.read_excel(open('template_excel.xlsx', 'rb'),
              sheet_name='chart17data') 

fig = plt.figure(linewidth=1, edgecolor="#04253a")
ax1 = plt.subplot()
x=df['fctrdt']
y6=df[df.columns[6]]
y5=df[df.columns[5]]
y4=df[df.columns[4]]
y3=df[df.columns[3]]
y2=df[df.columns[2]]
y1=df[df.columns[1]]
labels=["<=2",">2 and <=2.5",">2.5 and <=3",">3 and <=3.5",">3.5 and <=4",">4 and <=4.5"]
y= np.vstack([y1,y2,y3,y4,y5,y6])
ax1.stackplot(x,y1,y2,y3,y4,y5,y6,labels=labels,
      colors=['#96dbf6','#95dbf3','#56bff0','#2fa8e4','#238ace','#104d73'])


ax1.tick_params(axis=u'both',which=u'both',length=0,pad=5,labelsize=7)
ax1.set_xticks(df['fctrdt'][::3])
ax1.set_ylim(0,10000)
ax1.set_axisbelow(True)

ax1.spines['top'].set_visible(False)
ax1.spines['right'].set_visible(False)
ax1.set_xlim(left=df['fctrdt'][0],right=df['fctrdt'].iloc[-1])

for label in ax1.get_xticklabels() :
  label.set_rotation(90)
  label.set_ha('left')

line1,label1 = ax1.get_legend_handles_labels()

ax1.legend(reversed(line1), reversed(label1),loc="lower center", bbox_to_anchor=(1.2, 0.4)
                ,columnspacing=0.5, ncol=1,frameon=False,prop={'size': 9},handlelength=1)
plt.subplot().get_yaxis().set_major_formatter(
    mpl.ticker.FuncFormatter(lambda x, p: format(int(x), ',')))

plt.grid(axis='y',color="#f2f2f2",b=None,zorder=0)
for tick in ax1.get_xticklabels():
    tick.set_fontsize(11)
    
for tick in ax1.get_yticklabels():
     tick.set_fontsize(11)
plt.title('Chart Title 17', fontsize=14,pad=20,color='#7e7e7f',fontweight="bold")
[i.set_color("#7e7e7f") for i in ax1.get_xticklabels()]
[i.set_color("#7e7e7f") for i in ax1.get_yticklabels()]
plt.tight_layout()
fig.savefig( 'images/chart17.png', bbox_inches='tight') 

 #----------------------------------------------------------------------------------------#

#---------------------------------chart 18----------------------------------------#
df = pd.read_excel(open('template_excel.xlsx', 'rb'),
              sheet_name='chart18data') 

fig = plt.figure(linewidth=1, edgecolor="#04253a")
ax1 = plt.subplot()
x=df['fctrdt']
y9=df[df.columns[9]]
y8=df[df.columns[8]]
y7=df[df.columns[7]]
y6=df[df.columns[6]]
y5=df[df.columns[5]]
y4=df[df.columns[4]]
y3=df[df.columns[3]]
y2=df[df.columns[2]]
y1=df[df.columns[1]]
labels=["LLB85","MLB110","MLB125","HLB150","HLB175","HLB200","HLB225","Large","Jumbo"]
y= np.vstack([y1,y2,y3,y4,y5,y6])
ax1.stackplot(x,y1,y2,y3,y4,y5,y6,y7,y8,y9, labels=labels,
      colors=['#dee1e7','#ccd2dc','#aeb7c6','#95dbf3','#56bff0','#2fa8e4','#238ace','#1a6798','#104d73'])
ax1.set_axisbelow(True)

ax1.tick_params(axis=u'both',which=u'both',length=0,pad=5,labelsize=7)
ax1.set_xticks(df['fctrdt'][::3])
ax1.set_ylim(0,70000)


ax1.spines['top'].set_visible(False)
ax1.spines['right'].set_visible(False)
ax1.set_xlim(left=df['fctrdt'][0],right=df['fctrdt'].iloc[-1])

for label in ax1.get_xticklabels() :
  label.set_rotation(90)
  label.set_ha('left')
line1,label1 = ax1.get_legend_handles_labels()

ax1.legend(reversed(line1), reversed(label1),loc="lower center", bbox_to_anchor=(1.1, 0.6)
                ,columnspacing=0.5, ncol=1,frameon=False,prop={'size': 9},handlelength=1)
plt.subplot().get_yaxis().set_major_formatter(
    mpl.ticker.FuncFormatter(lambda x, p: format(int(x), ',')))

plt.grid(axis='y',color="#f2f2f2",b=None,zorder=0)
for tick in ax1.get_xticklabels():
    tick.set_fontsize(12)
    
for tick in ax1.get_yticklabels():
     tick.set_fontsize(12)
plt.title('Chart Title 18', fontsize=14,pad=20,color='#7e7e7f',fontweight="bold")
[i.set_color("#7e7e7f") for i in ax1.get_xticklabels()]
[i.set_color("#7e7e7f") for i in ax1.get_yticklabels()]
plt.tight_layout()
fig.savefig( 'images/chart18.png', bbox_inches='tight') 

#-------------------------------------------------------------------------------------#

#----------------------------Chart19--------------------------------------#
df = pd.read_excel(open('template_excel.xlsx', 'rb'),
              sheet_name='chart19data') 

fig = plt.figure(linewidth=1, edgecolor="#04253a",figsize=(8,6))
ax1 = plt.subplot()
ax1.plot(df['fctrdt'],df[df.columns[1]],color='#95dbf3',linewidth=2.3,label="All")
ax1.plot(df['fctrdt'],df[df.columns[2]],color='#ffc000',linewidth=2.3,label="Bank")
ax1.plot(df['fctrdt'],df[df.columns[3]],color='#104d73',linewidth=2.3,label="Non Bank")

ax1.tick_params(axis=u'both',which=u'both',length=0,pad=5)
ax1.set_xticks(df['fctrdt'][::2])
ax1.set_ylim(0,5)


ax1.spines['top'].set_visible(False)
ax1.spines['right'].set_visible(False)
ax1.set_xlim(left=df['fctrdt'][0],right=df['fctrdt'].iloc[-1])

for label in ax1.get_xticklabels() :
  label.set_rotation(90)
  label.set_ha('left')

line1,label1 = ax1.get_legend_handles_labels()
ax1.legend(line1, label1,loc="lower center", bbox_to_anchor=(0.5, -0.6)
                ,columnspacing=2.5, ncol=3,frameon=False,prop={'size': 9},handlelength=5)
plt.subplot().get_yaxis().set_major_formatter(
    mpl.ticker.FuncFormatter(lambda x, p: format(int(x), ',')))

plt.grid(axis='y',color="#f2f2f2")
for tick in ax1.get_xticklabels():
    tick.set_fontsize(12)
    
for tick in ax1.get_yticklabels():
     tick.set_fontsize(12)
plt.title('Chart Title 19', fontsize=16,pad=20,color='#7e7e7f',fontweight="bold")
[i.set_color("#7e7e7f") for i in ax1.get_xticklabels()]
[i.set_color("#7e7e7f") for i in ax1.get_yticklabels()]
plt.tight_layout()
fig.savefig( 'images/chart19.png', bbox_inches='tight') 
#-------------------------------------------------------------------#

#-------------------------------Chart 20-----------------------------#
df = pd.read_excel(open('template_excel.xlsx', 'rb'),
              sheet_name='chart20data') 

fig = plt.figure(linewidth=1, edgecolor="#04253a",figsize=(8,6))
ax1 = plt.subplot()
ax1.plot(df['fctrdt'],df[df.columns[1]],color='#95dbf3',linewidth=2.3,label="All")
ax1.plot(df['fctrdt'],df[df.columns[2]],color='#ffc000',linewidth=2.3,label="Bank")
ax1.plot(df['fctrdt'],df[df.columns[3]],color='#104d73',linewidth=2.3,label="Non Bank")

ax1.tick_params(axis=u'both',which=u'both',length=0,pad=5)
ax1.set_xticks(df['fctrdt'][::2])
ax1.set_ylim(0,3.5)
ticks=[0,0.5,1,1.5,2,2.5,3,3.5]
tickslabels = map(str,ticks)
ax1.set_yticks(ticks)
ax1.set_yticklabels(tickslabels)
ax1.spines['top'].set_visible(False)
ax1.spines['right'].set_visible(False)
ax1.set_xlim(left=df['fctrdt'][0],right=df['fctrdt'].iloc[-1])

for label in ax1.get_xticklabels() :
  label.set_rotation(90)
  label.set_ha('left')

line1,label1 = ax1.get_legend_handles_labels()
ax1.legend(line1, label1,loc="lower center", bbox_to_anchor=(0.5, -0.6)
                ,columnspacing=2.5, ncol=3,frameon=False,prop={'size': 9},handlelength=4)

plt.grid(axis='y',color="#f2f2f2")
for tick in ax1.get_xticklabels():
    tick.set_fontsize(12)
    
for tick in ax1.get_yticklabels():
     tick.set_fontsize(12)
plt.title('Chart Title 20', fontsize=16,pad=20,color='#7e7e7f',fontweight="bold")
[i.set_color("#7e7e7f") for i in ax1.get_xticklabels()]
[i.set_color("#7e7e7f") for i in ax1.get_yticklabels()]
plt.tight_layout()
fig.savefig( 'images/chart20.png', bbox_inches='tight') 

#--------------------------------------------------------------#

#----------------------------Chart 21---------------------------------#

fig = plt.figure(linewidth=1, edgecolor="#04253a",figsize=(8,6))
df = pd.read_excel(open('template_excel.xlsx', 'rb'),
              sheet_name='chart21data') 

ax1 = plt.subplot()
x=df['fctrdt']

y4=df[df.columns[4]]
y3=df[df.columns[3]]
y2=df[df.columns[2]]
y1=df[df.columns[1]]
labels=["%Firsttime Buyer","%Other Purchase","%Streamline Refi","%Cashou Refi"]
y= np.vstack([y1,y2,y3,y4])
ax1.stackplot(x,y1,y2,y3,y4, labels=labels,
      colors=['#96dbf5','#55c1f2','#2fa8e6','#2389ce'])


ax1.tick_params(axis=u'both',which=u'both',length=0,pad=5,labelsize=7)
ax1.set_xticks(df['fctrdt'][::3])
ax1.set_ylim(0,100)

ax1.set_axisbelow(True)
ax1.spines['top'].set_visible(False)
ax1.spines['right'].set_visible(False)
ax1.set_xlim(left=df['fctrdt'][0],right=df['fctrdt'].iloc[-1])

for label in ax1.get_xticklabels() :
  label.set_rotation(90)
  label.set_ha('center')

line1,label1 = ax1.get_legend_handles_labels()

ax1.legend(line1, label1,loc="lower center", bbox_to_anchor=(0.5, -0.7)
                ,columnspacing=3.5, ncol=4,frameon=False,prop={'size': 9},handlelength=1)
plt.subplot().get_yaxis().set_major_formatter(
    mpl.ticker.FuncFormatter(lambda x, p: format(int(x), ',')))

plt.grid(axis='y',color="#f2f2f2",b=None,zorder=0)
for tick in ax1.get_xticklabels():
    tick.set_fontsize(12)
    
for tick in ax1.get_yticklabels():
     tick.set_fontsize(12)
plt.title('Chart Title 21', fontsize=16,pad=20,color='#7e7e7f',fontweight="bold")
[i.set_color("#7e7e7f") for i in ax1.get_xticklabels()]
[i.set_color("#7e7e7f") for i in ax1.get_yticklabels()]
plt.tight_layout()
fig.savefig( 'images/chart21.png', bbox_inches='tight') 
#------------------------------------------------------------------------------------#

#----------------------------Chart 22-----------------------------------------#
df = pd.read_excel(open('template_excel.xlsx', 'rb'),
              sheet_name='chart22data') 

fig = plt.figure(linewidth=1, edgecolor="#04253a",figsize=(8,6))
ax1 = plt.subplot()
x=df['fctrdt']

y4=df[df.columns[4]]
y3=df[df.columns[3]]
y2=df[df.columns[2]]
y1=df[df.columns[1]]
labels=["%Firsttime Buyer","%Other Purchase","%Streamline Refi","%Cashou Refi"]
y= np.vstack([y1,y2,y3,y4])
ax1.stackplot(x,y1,y2,y3,y4, labels=labels,
      colors=['#96dbf5','#55c1f2','#2fa8e6','#2389ce'])


ax1.tick_params(axis=u'both',which=u'both',length=0,pad=5,labelsize=7)
ax1.set_xticks(df['fctrdt'][::3])
ax1.set_ylim(0,100)
ax1.set_axisbelow(True)

ax1.spines['top'].set_visible(False)
ax1.spines['right'].set_visible(False)
ax1.set_xlim(left=df['fctrdt'][0],right=df['fctrdt'].iloc[-1])

for label in ax1.get_xticklabels() :
  label.set_rotation(90)
  label.set_ha('center')

line1,label1 = ax1.get_legend_handles_labels()

ax1.legend(line1, label1,loc="lower center", bbox_to_anchor=(0.5, -0.7)
                ,columnspacing=3.5, ncol=4,frameon=False,prop={'size': 9},handlelength=1)
plt.subplot().get_yaxis().set_major_formatter(
    mpl.ticker.FuncFormatter(lambda x, p: format(int(x), ',')))

plt.grid(axis='y',color="#f2f2f2",b=None,zorder=0)
for tick in ax1.get_xticklabels():
    tick.set_fontsize(12)
    
for tick in ax1.get_yticklabels():
     tick.set_fontsize(12)
plt.title('Chart Title 22', fontsize=16,pad=20,color='#7e7e7f',fontweight="bold")
[i.set_color("#7e7e7f") for i in ax1.get_xticklabels()]
[i.set_color("#7e7e7f") for i in ax1.get_yticklabels()]
plt.tight_layout()
fig.savefig( 'images/chart22.png', bbox_inches='tight') 


#-----------------------Chart 23--------------------------------#
df = pd.read_excel(open('template_excel.xlsx', 'rb'),
              sheet_name='chart23data') 
fig = plt.figure(linewidth=2, edgecolor="#04253a",frameon=True,figsize=(18,6))
ax1 = plt.subplot()
ax1.set_xlabel('Date')
ax1.set_ylabel('$m')
l1 = ax1.plot(df['fctrdt'],df[df.columns[1]],color='#95daf3',label="Issuance(left)",linewidth=2.3)
l2 = ax1.plot(df['fctrdt'],df[df.columns[2]],color='#2389ce',label="Mandatory Purchase(right)",linewidth=2.3)
ax1.set_ylim(0,1601)
ax1.set_yticklabels(range(0,1601,200))
ax1.tick_params(axis=u'both',which=u'both',length=0,pad=5,labelsize=7)
ax1.set_xticks(df['fctrdt'][::1])
ax1.spines['top'].set_visible(False)
plt.grid(axis='y',color="#f2f2f2",b=None)
ax2 = ax1.twinx()
ax2.set_xlabel('Date')
ax2.set_ylabel('$m')
l3 = ax2.plot(df['fctrdt'],df[df.columns[3]],
    color='#96a916',label="HECM UPB(right)",linewidth=2.3)
ax2.set_ylim(46000,58001)

ax2.set_yticklabels(range(46000,58001,2000))
ax2.tick_params(axis=u'both',which=u'both',length=0,pad=5,labelsize=7)
ax1.set_xticks(df['fctrdt'][::1])
ax2.spines['top'].set_visible(False)


for label in ax1.get_xticklabels() + ax2.get_xticklabels():
  label.set_rotation(90)
  label.set_ha('center')

plt.setp(ax1,xlim=(df['fctrdt'][0],df['fctrdt'].iloc[-1]))
plt.setp(ax2,xlim=(df['fctrdt'][0],df['fctrdt'].iloc[-1]))

line1,label1 = ax1.get_legend_handles_labels()
line2,label2 = ax2.get_legend_handles_labels()

ax2.legend(line1+line2, label1+label2,loc="lower center", bbox_to_anchor=(0.5, -0.7)
                ,columnspacing=2.5, ncol=3,frameon=False,prop={'size': 9},handlelength=7)

for tick in ax1.get_xticklabels():
    tick.set_fontsize(12)
    
for tick in ax1.get_yticklabels():
     tick.set_fontsize(12)
plt.title('Chart Title 23', fontsize=15,pad=20,color='#7e7e7f',fontweight="bold")
[i.set_color("#7e7e7f") for i in ax1.get_xticklabels()]
[i.set_color("#7e7e7f") for i in ax1.get_yticklabels()]
[i.set_color("#7e7e7f") for i in ax2.get_yticklabels()]
plt.tight_layout()
fig.savefig( 'images/chart23.png', bbox_inches='tight') 
#----------------------------------------------------------------------------#

#-------------------------------chart 24-------------------------------------#
df = pd.read_excel(open('template_excel.xlsx', 'rb'),
              sheet_name='chart24data') 

fig = plt.figure(linewidth=1, edgecolor="#04253a",figsize=(8,6))
ax1 = plt.subplot()
x=df['fctrdt']
y3=df[df.columns[3]]
y2=df[df.columns[2]]
y1=df[df.columns[1]]
labels=["GNMCL","GNMPL","GNMPN"]
ax1.stackplot(x,y1,y2,y3,labels=labels,
      colors=['#95daf3','#2fa7e4','#238acf'])
ax1.tick_params(axis=u'both',which=u'both',length=0,pad=5,labelsize=7)
ax1.set_xticks(df['fctrdt'][::2])
ax1.set_ylim(0,3500)
ax1.spines['top'].set_visible(False)
ax1.spines['right'].set_visible(False)
ax1.set_xlim(left=df['fctrdt'][0],right=df['fctrdt'].iloc[-1])
ax1.set_axisbelow(True)
for label in ax1.get_xticklabels() :
  label.set_rotation(90)
  label.set_ha('center')
line1,label1 = ax1.get_legend_handles_labels()

ax1.legend(line1,label1,loc="lower center", bbox_to_anchor=(0.5, -0.7)
                ,columnspacing=0.5, ncol=3,frameon=False,prop={'size': 9},handlelength=1)
ax1.get_yaxis().set_major_formatter(
    mpl.ticker.FuncFormatter(lambda x, p: format(int(x), ',')))

plt.grid(axis='y',color="#f5f5f5",b=None)
for tick in ax1.get_xticklabels():
    tick.set_fontsize(12)
    
for tick in ax1.get_yticklabels():
     tick.set_fontsize(12)
plt.title('Chart Title 24', fontsize=16,pad=20,color='#7e7e7f',fontweight="bold")
[i.set_color("#7e7e7f") for i in ax1.get_xticklabels()]
[i.set_color("#7e7e7f") for i in ax1.get_yticklabels()]
plt.tight_layout()
fig.savefig( 'images/chart24.png', bbox_inches='tight') 

#------------------------------Chart 25---------------------------#

df = pd.read_excel(open('template_excel.xlsx', 'rb'),
              sheet_name='chart25data') 

fig = plt.figure(linewidth=1, edgecolor="#04253a")
ax1 = plt.subplot()
x=df['fctrdt']
y3=df[df.columns[3]]
y2=df[df.columns[2]]
y1=df[df.columns[1]]
labels=["GNMCL","GNMPL","GNMPN"]
ax1.stackplot(x,y1,y2,y3,labels=labels,
      colors=['#95daf3','#2fa7e4','#238acf'])
ax1.tick_params(axis=u'both',which=u'both',length=0,pad=5,labelsize=7)
ax1.set_xticks(df['fctrdt'][::3])
ax1.set_ylim(0,140000)
ax1.spines['top'].set_visible(False)
ax1.spines['right'].set_visible(False)
ax1.set_xlim(left=df['fctrdt'][0],right=df['fctrdt'].iloc[-1])

for label in ax1.get_xticklabels() :
  label.set_rotation(45)
  label.set_ha('right')
line1,label1 = ax1.get_legend_handles_labels()
ax1.set_axisbelow(True)
ax1.legend(line1,label1,loc="lower center", bbox_to_anchor=(0.5, -0.5)
                ,columnspacing=0.5, ncol=3,frameon=False,prop={'size': 9},handlelength=1)
plt.subplot().get_yaxis().set_major_formatter(
    mpl.ticker.FuncFormatter(lambda x, p: format(int(x), ',')))

plt.grid(axis='y',color="#f5f5f5",b=None,zorder=0)
for tick in ax1.get_xticklabels():
    tick.set_fontsize(12)
    
for tick in ax1.get_yticklabels():
     tick.set_fontsize(12)
plt.title('Chart Title 25', fontsize=15,pad=20,color='#7e7e7f',fontweight="bold")
[i.set_color("#7e7e7f") for i in ax1.get_xticklabels()]
[i.set_color("#7e7e7f") for i in ax1.get_yticklabels()]
plt.tight_layout()
fig.savefig( 'images/chart25.png', bbox_inches='tight') 
