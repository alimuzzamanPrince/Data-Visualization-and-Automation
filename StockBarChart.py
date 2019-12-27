import pyodbc as db
import numpy as np
import matplotlib.pyplot as plt
import matplotlib
from matplotlib.ticker import MaxNLocator
from PIL import Image
import functions as fn
import linechart as ln

#####----webdriverpart

from selenium import webdriver
import os
from IPython.display import display, HTML
import smtplib, ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email import encoders
import time
from io import BytesIO
from Screenshot import Screenshot_Clipping
from PIL import ImageEnhance
import base64
from datetime import datetime

####-----Table

dirpath = os.path.dirname(os.path.realpath(__file__))

try:
    os.remove(dirpath+'/SKFStockReport.xlsx')
    print("File Removed!")
except WindowsError: 
    print("no file found")
    pass

plt.figure(0)
bars = plt.bar(fn.y_pos, fn.NIL, width=0.5)
fn.bardecoration(bars)
plt.title('Branch Status - Stock NIL', fontdict=fn.font1)
plt.savefig(dirpath+'/nilbar.png')

plt.figure(1)
bars = plt.bar(fn.y_pos, fn.SUS, width=0.5)
fn.bardecoration(bars)
plt.title('Branch Status - Stock SUS (1-15 Days)', fontdict=fn.font1)
plt.savefig(dirpath+'/susbar.png')

plt.figure(2)
bars = plt.bar(fn.y_pos, fn.US, width=0.5)
fn.bardecoration(bars)
plt.title('Branch Status - Stock US (16-35 Days)', fontdict=fn.font1)
plt.savefig(dirpath+'/usbar.png')

plt.figure(3)
bars = plt.bar(fn.y_pos, fn.NS, width=0.5)
fn.bardecoration(bars)
plt.title('Branch Status - Stock NS (36-45 Days)', fontdict=fn.font1)
plt.savefig(dirpath+'/nsbar.png')

plt.figure(4)
bars = plt.bar(fn.y_pos, fn.OS, width=0.5)
fn.bardecoration(bars)
plt.title('Branch Status - Stock OS (46-60 Days)', fontdict=fn.font1)
plt.savefig(dirpath+'/osbar.png')

plt.figure(5)
bars = plt.bar(fn.y_pos, fn.SOS, width=0.5)
for bar in bars:
    yval = bar.get_height()
    plt.text(bar.get_x(), yval + 20, yval, rotation=90, fontsize = 8)

for bar in bars[0:7]:
    bar.set_color('#be4748')

for bar in bars[7:15]:
    bar.set_color('#1a58c5')

for bar in bars[15:23]:
    bar.set_color('#c5871a')

for bar in bars[23:31]:
    bar.set_color('#58c51a')

plt.xticks(fn.y_pos, fn.branchlist, rotation=90)
plt.yticks(np.arange(0, 500, 30), fontsize=6)
plt.ylabel('No. of Item', fontdict=fn.font)
###plt.legend(handles=fn.legend_elements, loc='best')

plt.title('Branch Status - Stock SOS (More Than 60 Days)', fontdict=fn.font1)
plt.savefig(dirpath+'/sosbar.png')

image = Image.new('RGB', (1281, 1442))
im  = Image.open(dirpath+"/nilbar.png")
width,height=im.size
im1 = Image.open(dirpath+"/susbar.png")
im2 = Image.open(dirpath+"/usbar.png")
im3 = Image.open(dirpath+"/sosbar.png")
im4 = Image.open(dirpath+"/osbar.png")
im5 = Image.open(dirpath+"/nsbar.png")


####image = Image.new('RGB', (width1+width2+width3+width4, height+height1+1))
image.paste(im,(0,0))
image.paste(im5,(width+1,0))

image.paste(im2,(0,height+1))
image.paste(im1,(width+1,height+1))

image.paste(im4,(0,height*2+2))
image.paste(im3,(width+1,height*2+2))
#######image.show()

image.save(dirpath+"/Final.png")

##############-------Pie Chart
colors = ['#990000', '#ff0000', '#e6e600', '#008000', '#00ff16', 'darkorange']
xxx = ['NIL','SUS','US','NS','OS','SOS']

plt.figure(6)
wtf = plt.pie(fn.nuruta, colors=colors, autopct='%1.1f%%', pctdistance=1.2)
plt.title('Mr. Nurul', fontdict=fn.font1)
plt.legend(labels = xxx, title="Status", loc="center left",bbox_to_anchor=(1, 0, 0.5, 1))

fraction_text_list = wtf[2]
for text in fraction_text_list:
    text.set_rotation(70)
plt.savefig(dirpath+'/nurupie.png')

plt.figure(7)
wtf = plt.pie(fn.batikta, colors=colors, autopct='%1.1f%%', pctdistance=1.2)
plt.title('Mr. Atik', fontdict=fn.font1)
plt.legend(labels = xxx, title="Status", loc="center left",bbox_to_anchor=(1, 0, 0.5, 1))

fraction_text_list = wtf[2]
for text in fraction_text_list:
    text.set_rotation(70)
plt.savefig(dirpath+'/batikpie.png')

plt.figure(8)
wtf = plt.pie(fn.warta, colors=colors, autopct='%1.1f%%', pctdistance=1.2)
plt.title('Mr. Anwar', fontdict=fn.font1)
plt.legend(labels = xxx, title="Status", loc="center left",bbox_to_anchor=(1, 0, 0.5, 1))

fraction_text_list = wtf[2]
for text in fraction_text_list:
    text.set_rotation(70)
plt.savefig(dirpath+'/warpie.png')

plt.figure(9)
wtf = plt.pie(fn.kamta, colors=colors, autopct='%1.1f%%', pctdistance=1.2)
plt.title('Mr. Kamrul', fontdict=fn.font1)
plt.legend(labels = xxx, title="Status", loc="center left",bbox_to_anchor=(1, 0, 0.5, 1))

fraction_text_list = wtf[2]
for text in fraction_text_list:
    text.set_rotation(70)
plt.savefig(dirpath+'/kampie.png')


imageX = Image.new('RGB', (1281, 961))

imp1  = Image.open(dirpath+"/kampie.png")
widthx,heightx = imp1.size
imp2  = Image.open(dirpath+"/warpie.png")
imp3  = Image.open(dirpath+"/batikpie.png")
imp4  = Image.open(dirpath+"/nurupie.png")


imageX.paste(imp1,(0,0))
imageX.paste(imp2,(widthx+1,0))
imageX.paste(imp3,(0,heightx+1))
imageX.paste(imp4,(widthx+1,heightx+1))

imageX.save(dirpath+"/ndmpie.png")


i = 0
x = 10
y_pos = np.arange(len(xxx))
plt.figure(x)
mapdata = [fn.NIL[i],fn.SUS[i],fn.US[i],fn.NS[i],fn.OS[i],fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('CTG Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/ctgbar.png')
i=i+1
x=x+1

plt.figure(x)
mapdata = [fn.NIL[i],fn.SUS[i],fn.US[i],fn.NS[i],fn.OS[i],fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('CTN Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/ctnbar.png')
i=i+1
x=x+1

plt.figure(x)
mapdata = [fn.NIL[i],fn.SUS[i],fn.US[i],fn.NS[i],fn.OS[i],fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('JES Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/jesbar.png')
i=i+1
x=x+1

plt.figure(x)
mapdata = [fn.NIL[i],fn.SUS[i],fn.US[i],fn.NS[i],fn.OS[i],fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('MOT Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/motbar.png')
i=i+1
x=x+1

plt.figure(x)
mapdata = [fn.NIL[i],fn.SUS[i],fn.US[i],fn.NS[i],fn.OS[i],fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('MYM Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/mymbar.png')
i=i+1
x=x+1

plt.figure(x)
mapdata = [fn.NIL[i],fn.SUS[i],fn.US[i],fn.NS[i],fn.OS[i],fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('NAJ Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/najbar.png')
i=i+1
x=x+1

plt.figure(x)
mapdata = [fn.NIL[i],fn.SUS[i],fn.US[i],fn.NS[i],fn.OS[i],fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('NOK Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/nokbar.png')
i=i+1
x=x+1

plt.figure(x)
mapdata = [fn.NIL[i],fn.SUS[i],fn.US[i],fn.NS[i],fn.OS[i],fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('COM Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/combar.png')
i=i+1
x=x+1

plt.figure(x)
mapdata = [fn.NIL[i],fn.SUS[i],fn.US[i],fn.NS[i],fn.OS[i],fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('FEN Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/fenbar.png')
i=i+1
x=x+1

plt.figure(x)
mapdata = [fn.NIL[i],fn.SUS[i],fn.US[i],fn.NS[i],fn.OS[i],fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('FRD Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/frdbar.png')
i=i+1
x=x+1

plt.figure(x)
mapdata = [fn.NIL[i],fn.SUS[i],fn.US[i],fn.NS[i],fn.OS[i],fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('GZP Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/gzpbar.png')
i=i+1
x=x+1

plt.figure(x)
mapdata = [fn.NIL[i],fn.SUS[i],fn.US[i],fn.NS[i],fn.OS[i],fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('KHL Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/khlbar.png')
i=i+1
x=x+1

plt.figure(x)
mapdata = [fn.NIL[i],fn.SUS[i],fn.US[i],fn.NS[i],fn.OS[i],fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('KSG Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/ksgbar.png')
i=i+1
x=x+1

plt.figure(x)
mapdata = [fn.NIL[i],fn.SUS[i],fn.US[i],fn.NS[i],fn.OS[i],fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('MIR Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/mirbar.png')
i=i+1
x=x+1

plt.figure(x)
mapdata = [fn.NIL[i],fn.SUS[i],fn.US[i],fn.NS[i],fn.OS[i],fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('PAT Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/patbar.png')
i=i+1
x=x+1

plt.figure(x)
mapdata = [fn.NIL[i],fn.SUS[i],fn.US[i],fn.NS[i],fn.OS[i],fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('BOG Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/bogbar.png')
i=i+1
x=x+1

plt.figure(x)
mapdata = [fn.NIL[i],fn.SUS[i],fn.US[i],fn.NS[i],fn.OS[i],fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('COX Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/coxbar.png')
i=i+1
x=x+1

plt.figure(x)
mapdata = [fn.NIL[i],fn.SUS[i],fn.US[i],fn.NS[i],fn.OS[i],fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('HZJ Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/hzjbar.png')
i=i+1
x=x+1

plt.figure(x)
mapdata = [fn.NIL[i],fn.SUS[i],fn.US[i],fn.NS[i],fn.OS[i],fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('KUS Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/kusbar.png')
i=i+1
x=x+1

plt.figure(x)
mapdata = [fn.NIL[i],fn.SUS[i],fn.US[i],fn.NS[i],fn.OS[i],fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('MHK Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/mhkbar.png')
i=i+1
x=x+1

plt.figure(x)
mapdata = [fn.NIL[i],fn.SUS[i],fn.US[i],fn.NS[i],fn.OS[i],fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('MLV Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/mlvbar.png')
i=i+1
x=x+1

plt.figure(x)
mapdata = [fn.NIL[i],fn.SUS[i],fn.US[i],fn.NS[i],fn.OS[i],fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('PBN Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/pbnbar.png')
i=i+1
x=x+1

plt.figure(x)
mapdata = [fn.NIL[i],fn.SUS[i],fn.US[i],fn.NS[i],fn.OS[i],fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('SYL Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/sylbar.png')
i=i+1
x=x+1

plt.figure(x)
mapdata = [fn.NIL[i],fn.SUS[i],fn.US[i],fn.NS[i],fn.OS[i],fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('BSL Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/bslbar.png')
i=i+1
x=x+1

plt.figure(x)
mapdata = [fn.NIL[i],fn.SUS[i],fn.US[i],fn.NS[i],fn.OS[i],fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('DNJ Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/dnjbar.png')
i=i+1
x=x+1

plt.figure(x)
mapdata = [fn.NIL[i],fn.SUS[i],fn.US[i],fn.NS[i],fn.OS[i],fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('KRN Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/krnbar.png')
i=i+1
x=x+1

plt.figure(x)
mapdata = [fn.NIL[i],fn.SUS[i],fn.US[i],fn.NS[i],fn.OS[i],fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('RAJ Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/rajbar.png')
i=i+1
x=x+1

plt.figure(x)
mapdata = [fn.NIL[i],fn.SUS[i],fn.US[i],fn.NS[i],fn.OS[i],fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('RNG Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/rngbar.png')
i=i+1
x=x+1

plt.figure(x)
mapdata = [fn.NIL[i],fn.SUS[i],fn.US[i],fn.NS[i],fn.OS[i],fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('SAV Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/savbar.png')
i=i+1
x=x+1

plt.figure(x)
mapdata = [fn.NIL[i],fn.SUS[i],fn.US[i],fn.NS[i],fn.OS[i],fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('TGL Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/tglbar.png')
i=i+1
x=x+1

plt.figure(x)
mapdata = [fn.NIL[i],fn.SUS[i],fn.US[i],fn.NS[i],fn.OS[i],fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('VRB Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/vrbbar.png')
i=i+1
x=x+1

im1  = Image.open(dirpath+"/ctgbar.png")
width, height = im1.size
image = Image.new('RGB', (width*3, height*3))

im2  = Image.open(dirpath+"/ctnbar.png")
im3  = Image.open(dirpath+"/jesbar.png")
im4  = Image.open(dirpath+"/motbar.png")
im5  = Image.open(dirpath+"/mymbar.png")
im6  = Image.open(dirpath+"/najbar.png")
im7  = Image.open(dirpath+"/nokbar.png")
imfk  = Image.open(dirpath+"/fakebar.png")
imkam  = Image.open(dirpath+"/kam.png")


image.paste(imkam,(0,0))
image.paste(im1,(width,0))
image.paste(im2,(width*2,0))
image.paste(im3,(0,height))
image.paste(im4,(width,height))
image.paste(im5,(width*2,height))
image.paste(im6,(0,height*2))
image.paste(im7,(width,height*2))
image.paste(imfk,(width*2,height*2))

image.save(dirpath+"/ndm1.png")

image1 = Image.new('RGB', (width*3, height*3))

im8  = Image.open(dirpath+"/combar.png")
im9  = Image.open(dirpath+"/fenbar.png")
im10  = Image.open(dirpath+"/frdbar.png")
im11  = Image.open(dirpath+"/gzpbar.png")
im12  = Image.open(dirpath+"/khlbar.png")
im13  = Image.open(dirpath+"/ksgbar.png")
im14  = Image.open(dirpath+"/mirbar.png")
im15  = Image.open(dirpath+"/patbar.png")
imwar  = Image.open(dirpath+"/war.png")

image1.paste(imwar,(0,0))
image1.paste(im8,(width,0))
image1.paste(im9,(width*2,0))
image1.paste(im10,(0,height))
image1.paste(im11,(width,height))
image1.paste(im12,(width*2,height))
image1.paste(im13,(0,height*2))
image1.paste(im14,(width,height*2))
image1.paste(im15,(width*2,height*2))

image1.save(dirpath+"/ndm2.png")

image2 = Image.new('RGB', (width*3, height*3))

im16  = Image.open(dirpath+"/bogbar.png")
im17  = Image.open(dirpath+"/coxbar.png")
im18  = Image.open(dirpath+"/hzjbar.png")
im19  = Image.open(dirpath+"/kusbar.png")
im20  = Image.open(dirpath+"/mhkbar.png")
im21  = Image.open(dirpath+"/mlvbar.png")
im22  = Image.open(dirpath+"/pbnbar.png")
im23  = Image.open(dirpath+"/sylbar.png")
imtik  = Image.open(dirpath+"/batik.png")

image2.paste(imtik,(0,0))
image2.paste(im16,(width,0))
image2.paste(im17,(width*2,0))
image2.paste(im18,(0,height))
image2.paste(im19,(width,height))
image2.paste(im20,(width*2,height))
image2.paste(im21,(0,height*2))
image2.paste(im22,(width,height*2))
image2.paste(im23,(width*2,height*2))

image2.save(dirpath+"/ndm3.png")

image3 = Image.new('RGB', (width*3, height*3))

im24  = Image.open(dirpath+"/bslbar.png")
im25  = Image.open(dirpath+"/dnjbar.png")
im26  = Image.open(dirpath+"/krnbar.png")
im27  = Image.open(dirpath+"/rajbar.png")
im28  = Image.open(dirpath+"/rngbar.png")
im29  = Image.open(dirpath+"/savbar.png")
im30  = Image.open(dirpath+"/tglbar.png")
im31  = Image.open(dirpath+"/vrbbar.png")
imnuru  = Image.open(dirpath+"/nuru.png")

image3.paste(imnuru,(0,0))
image3.paste(im24,(width,0))
image3.paste(im25,(width*2,0))
image3.paste(im26,(0,height))
image3.paste(im27,(width,height))
image3.paste(im28,(width*2,height))
image3.paste(im29,(0,height*2))
image3.paste(im30,(width,height*2))
image3.paste(im31,(width*2,height*2))

image3.save(dirpath+"/ndm4.png")

#########Line chart started
plt.figure(x)
plt.plot(ln.days, ln.nilkam30, color='#be4748')
plt.plot(ln.days, ln.nilwar30, color='#1a58c5')
plt.plot(ln.days, ln.niltik30, color='#c5871a')
plt.plot(ln.days, ln.nilnuru30, color='#58c51a')

plt.xticks(range(1,31), ln.datenuru30, rotation=90)
plt.ylabel('No. of Item', fontdict=ln.font)
plt.xlabel('Days', fontdict=ln.font)
plt.title('NDM\'s Last 30 Days NIL Status', fontdict=ln.font1)
plt.legend(['Mr. Kamrul', 'Mr. Anwar', 'Mr. Atik', 'Mr. Nurul'], loc='best')
plt.savefig(dirpath+'/nilline.png')
x=x+1

plt.figure(x)
plt.plot(ln.days, ln.uskam30, color='#be4748')
plt.plot(ln.days, ln.uswar30, color='#1a58c5')
plt.plot(ln.days, ln.ustik30, color='#c5871a')
plt.plot(ln.days, ln.usnuru30, color='#58c51a')

plt.xticks(range(1,31), ln.datenuru30, rotation=90)
plt.ylabel('No. of Item', fontdict=ln.font)
plt.xlabel('Days', fontdict=ln.font)
plt.title('NDM\'s Last 30 Days US Status', fontdict=ln.font1)
plt.legend(['Mr. Kamrul', 'Mr. Anwar', 'Mr. Atik', 'Mr. Nurul'], loc='best')
plt.savefig(dirpath+'/usline.png')
x=x+1

plt.figure(x)
plt.plot(ln.days, ln.suskam30, color='#be4748')
plt.plot(ln.days, ln.suswar30, color='#1a58c5')
plt.plot(ln.days, ln.sustik30, color='#c5871a')
plt.plot(ln.days, ln.susnuru30, color='#58c51a')

plt.xticks(range(1,31), ln.datenuru30, rotation=90)
plt.ylabel('No. of Item', fontdict=ln.font)
plt.xlabel('Days', fontdict=ln.font)
plt.title('NDM\'s Last 30 Days SUS Statsus', fontdict=ln.font1)
plt.legend(['Mr. Kamrul', 'Mr. Anwar', 'Mr. Atik', 'Mr. Nurul'], loc='best')
plt.savefig(dirpath+'/susline.png')
x=x+1

plt.figure(x)
plt.plot(ln.days, ln.nskam30, color='#be4748')
plt.plot(ln.days, ln.nswar30, color='#1a58c5')
plt.plot(ln.days, ln.nstik30, color='#c5871a')
plt.plot(ln.days, ln.nsnuru30, color='#58c51a')

plt.xticks(range(1,31), ln.datenuru30, rotation=90)
plt.ylabel('No. of Item', fontdict=ln.font)
plt.xlabel('Days', fontdict=ln.font)
plt.title('NDM\'s Last 30 Days NS Status', fontdict=ln.font1)
plt.legend(['Mr. Kamrul', 'Mr. Anwar', 'Mr. Atik', 'Mr. Nurul'], loc='best')
plt.savefig(dirpath+'/nsline.png')
x=x+1

plt.figure(x)
plt.plot(ln.days, ln.oskam30, color='#be4748')
plt.plot(ln.days, ln.oswar30, color='#1a58c5')
plt.plot(ln.days, ln.ostik30, color='#c5871a')
plt.plot(ln.days, ln.osnuru30, color='#58c51a')

plt.xticks(range(1,31), ln.datenuru30, rotation=90)
plt.ylabel('No. of Item', fontdict=ln.font)
plt.xlabel('Days', fontdict=ln.font)
plt.title('NDM\'s Last 30 Days OS Status', fontdict=ln.font1)
plt.legend(['Mr. Kamrul', 'Mr. Anwar', 'Mr. Atik', 'Mr. Nurul'], loc='best')
plt.savefig(dirpath+'/osline.png')
x=x+1

plt.figure(x)
plt.plot(ln.days, ln.soskam30, color='#be4748')
plt.plot(ln.days, ln.soswar30, color='#1a58c5')
plt.plot(ln.days, ln.sostik30, color='#c5871a')
plt.plot(ln.days, ln.sosnuru30, color='#58c51a')

plt.xticks(range(1,31), ln.datenuru30, rotation=90)
plt.ylabel('No. of Item', fontdict=ln.font)
plt.xlabel('Days', fontdict=ln.font)
plt.title('NDM\'s Last 30 Days SOS Status', fontdict=ln.font1)
plt.legend(['Mr. Kamrul', 'Mr. Anwar', 'Mr. Atik', 'Mr. Nurul'], loc='best')
plt.savefig(dirpath+'/sosline.png')
x=x+1

imageline = Image.new('RGB', (1281, 1442))

imp1  = Image.open(dirpath+"/nilline.png")
widthx,heightx = imp1.size
imp2  = Image.open(dirpath+"/nsline.png")
imp3  = Image.open(dirpath+"/usline.png")
imp4  = Image.open(dirpath+"/susline.png")
imp5  = Image.open(dirpath+"/osline.png")
imp6  = Image.open(dirpath+"/sosline.png")


imageline.paste(imp1,(0,0))
imageline.paste(imp2,(widthx+1,0))
imageline.paste(imp3,(0,heightx+1))
imageline.paste(imp4,(widthx+1,heightx+1))
imageline.paste(imp5,(0,heightx*2+2))
imageline.paste(imp6,(widthx+1,heightx*2+2))

imageline.save(dirpath+"/ndmline.png")
########---------Wbdriver Part
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
prefs = {"profile.default_content_settings.popups": 0,"download.default_directory": dirpath,"directory_upgrade": True, "safebrowsing.enabled": True}
options.add_experimental_option("prefs", prefs)

driver=webdriver.Chrome('chromedriver.exe', options=options)
driver.set_page_load_timeout(1500)

ob=Screenshot_Clipping.Screenshot()
driver.get("http://Alim:alimerp@10.168.2.245:9081/Report/Pages/ReportViewer.aspx?%2Fsum")
time.sleep(150)
ob.full_Screenshot(driver, save_path=dirpath, image_name='dhd.png')

im = Image.open(dirpath+"/dhd.png")
croppedIm = im.crop((0, 43, 1010, 580))

enhancer = ImageEnhance.Sharpness(croppedIm)
enhanced_im = enhancer.enhance(0.8)

enhanced_im.save(dirpath+"/dhd_col.png", "png", quality=100, optimize=True, progressive=True)

driver.get('http://Alim:alimerp@10.168.2.245:9081/reports/')
print("Downloading Stock Report... ...")
driver.get("http://Alim:alimerp@10.168.2.245:9081/Report?%2FSKFStockReport&rs:Format=EXCELOPENXML")
print("Download Complete! Make Happy Face :D ")

time.sleep(17)
driver.close()
###########-------------

msgRoot = MIMEMultipart('related')

me  = 'erp-bi.service@transcombd.com'

#to = ['hislam@skf.transcombd.com','muhammad.mainuddin@tdcl.transcombd.com']
#cc = ['almamun@transcombd.com','tdclndm@tdcl.transcombd.com','tdclpharma@transcombd.com','monowar@tdcl.transcombd.com','mosaddek.hossain@skf.transcombd.com','monirul@skf.transcombd.com']
#bcc = ['biswascma@yahoo.com','tawhid@transcombd.com','yakub@transcombd.com','zubair@transcombd.com','m.alimuzzaman@transcombd.com']
#recipient=to+cc+bcc

recipient='m.alimuzzaman@transcombd.com'

date = datetime.today()
today = str(date.day)+'/'+str(date.month)+'/'+str(date.year)+' '+date.strftime("%I:%M %p")
today1 = str(date.day)+' '+str(date.strftime("%B"))+' '+str(date.year)+' at '+date.strftime("%I:%M %p")

subject="SK+F Formulation – Stock Status Report - "+today

email_server_host = 'mail.transcombd.com'
port = 25

msgRoot['to'] = recipient
msgRoot['from'] = me
msgRoot['subject'] = subject

## Encapsulate the plain and HTML versions of the message body in an
## 'alternative' part, so message agents can decide which they want to display.
msgAlternative = MIMEMultipart('alternative')
msgRoot.attach(msgAlternative)

msgText = MIMEText('This is the alternative plain text message.')
msgAlternative.attach(msgText)

## We reference the image in the IMG SRC attribute by the ID we give it below
msgText = MIMEText("""<b>Dear Sir,</b><br><br>Enclosed, Please find herewith the Stock Status report of """
                   + today1 + """ for TDCL All Branches including Central Warehouse.<br><br>

                   <img src="cid:image1"><br> <br>
                   <img src="cid:image2"><br> <br>
                   <img src="cid:image8"><br> <br>
                   <img src="cid:image3"><br> <br>
                   <img src="cid:image4"><br> <br>
                   <img src="cid:image5"><br> <br>
                   <img src="cid:image6"><br> <br>
                   <img src="cid:image7"><br> <br>

                   If there is any inconvenience, you are requested to communicate with the ERP BI Service:
                   <br><b>(Mobile: 01713-389972, 01713-380499)</b><br><br>
                   Regards<br><b>ERP BI Service</b><br>Information System Automation (ISA)<br><br>
                   <i><font color="blue">****This is a system generated stock report ****</i></font>""", 'html')


msgAlternative.attach(msgText)

# Assigning the 2nd image directory
fp = open(dirpath+'/Final.png', 'rb')
msgImage1 = MIMEImage(fp.read())
fp.close()

# Define the 2nd image's ID as referenced above
msgImage1.add_header('Content-ID', '<image1>')
msgRoot.attach(msgImage1)

# Assigning the 2nd image directory
fp = open(dirpath+'/ndmpie.png', 'rb')
msgImage2 = MIMEImage(fp.read())
fp.close()

# Define the 2nd image's ID as referenced above
msgImage2.add_header('Content-ID', '<image2>')
msgRoot.attach(msgImage2)

## Assigning the 2nd image directory
fp = open(dirpath+'/ndm1.png', 'rb')
msgImage3 = MIMEImage(fp.read())
fp.close()

# Define the 2nd image's ID as referenced above
msgImage3.add_header('Content-ID', '<image3>')
msgRoot.attach(msgImage3)

# Assigning the 2nd image directory
fp = open(dirpath+'/ndm2.png', 'rb')
msgImage4 = MIMEImage(fp.read())
fp.close()

# Define the 2nd image's ID as referenced above
msgImage4.add_header('Content-ID', '<image4>')
msgRoot.attach(msgImage4)

# Assigning the 2nd image directory
fp = open(dirpath+'/ndm3.png', 'rb')
msgImage5 = MIMEImage(fp.read())
fp.close()

# Define the 2nd image's ID as referenced above
msgImage5.add_header('Content-ID', '<image5>')
msgRoot.attach(msgImage5)

# Assigning the 2nd image directory
fp = open(dirpath+'/ndm4.png', 'rb')
msgImage6 = MIMEImage(fp.read())
fp.close()

# Define the 2nd image's ID as referenced above
msgImage6.add_header('Content-ID', '<image6>')
msgRoot.attach(msgImage6)

# Assigning the 2nd image directory
fp = open(dirpath+'/dhd_col.png', 'rb')
msgImage7 = MIMEImage(fp.read())
fp.close()

# Define the 2nd image's ID as referenced above
msgImage7.add_header('Content-ID', '<image7>')
msgRoot.attach(msgImage7)

# Assigning the 2nd image directory
fp = open(dirpath+'/ndmline.png', 'rb')
msgImage8 = MIMEImage(fp.read())
fp.close()

# Define the 2nd image's ID as referenced above
msgImage8.add_header('Content-ID', '<image8>')
msgRoot.attach(msgImage8)

#Excel attachment
part = MIMEBase('application', "octet-stream")
file_location = dirpath+'/SKFStockReport.xlsx'
###print(file_location)
import os
# Create the attachment file (only do it once)
filename = os.path.basename(file_location)
attachment = open(file_location, "rb")
part = MIMEBase('application', 'octet-stream')
part.set_payload(attachment.read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
msgRoot.attach(part)

#msgRoot['From'] = me
#msgRoot['To'] = ', '.join(to)
#msgRoot['Cc'] = ', '.join(cc)
#msgRoot['Bcc'] = ', '.join(bcc)
#msgRoot['Subject'] = subject

server = smtplib.SMTP(email_server_host, port)
server.ehlo()
print('Sending Mail')
server.sendmail(me, recipient, msgRoot.as_string())
print('Mail Sent')
server.close()