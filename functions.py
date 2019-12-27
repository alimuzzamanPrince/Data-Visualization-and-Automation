import matplotlib.pyplot as plt
from matplotlib.patches import Patch
import pyodbc as db
import numpy as np


connection = db.connect('DRIVER={SQL Server};'
                        'SERVER=10.168.2.241;'
                        'DATABASE=ARCHIVESKF;'
                        'UID=sa;PWD=erp')#####Add trusted_connection=yes and remove uid and pwd while connecting to local

cursor = connection.cursor()


query = """DECLARE @TotalItem VARCHAR(20)
SET @TotalItem = (SELECT count( distinct  ITEMNO) as TotalMovingItem FROM ICStockStatusCurrentLOT 
		where  left(ITEMNO,1)<>'9' and AUDTORG<>'SKFDAT')
SELECT NDMNAME, LEFT(AUDTORG,3) AS AUDTORG,
 Sum( CASE WHEN [Days]<=15 THEN 1 END) AS SUS
,Sum( CASE WHEN [Days]>15 AND [Days]<=35 THEN 1 END) AS US
,Sum( CASE WHEN [Days]>35 AND [Days]<=45 THEN 1 END) AS NS
,Sum( CASE WHEN [Days]>45 AND [Days]<=60 THEN 1 END) AS OS
,Sum( CASE WHEN [Days]>60 THEN 1 END) AS SOS
,@TotalItem-(Sum( CASE WHEN [Days]<=15 THEN 1 END)+Sum( CASE WHEN [Days]>15 AND [Days]<=35 THEN 1 END)+Sum( CASE WHEN [Days]>35 AND [Days]<=45 THEN 1 END)
+Sum( CASE WHEN [Days]>45 AND [Days]<=60 THEN 1 END)+Sum( CASE WHEN [Days]>60 THEN 1 END)) AS NIL
,Sum( CASE WHEN [Days]>=0 THEN 1 END) AS TotalItem
,@TotalItem AS AllItem

FROM

(
SELECT NDMNAME, Stock.ITEMNO, Stock.AUDTORG, Sum(QTYONHAND) as QTYONHAND , Sum(QTYSHIPPED)as QTYSHIPPED
,CAST(ISNULL(CASE WHEN SUM(QTYSHIPPED)>0 THEN (SUM(QTYONHAND)+ISNULL(SUM(SIT),0))/(SUM(QTYSHIPPED)/90) END,0) AS INT) AS [Days]FROM
-----
(SELECT ITEMNO, AUDTORG, SUM(QTYONHAND) AS QTYONHAND FROM ICStockStatusCurrentLOT
where LEN(LOCATION)>'3' AND left(ITEMNO,1)<>'9' AND AUDTORG <> 'SKFDAT' 
GROUP BY ITEMNO, AUDTORG) AS Stock
LEFT JOIN
(SELECT ITEM, AUDTORG, SUM(QTYSHIPPED) AS QTYSHIPPED FROM OESalesDetails
WHERE TRANSDATE BETWEEN CONVERT(varchar(8), GETDATE()-91,112) AND CONVERT(varchar(8), GETDATE()-1,112)
GROUP BY ITEM, AUDTORG) AS Sales
ON RTRIM(Stock.ITEMNO) = RTRIM(Sales.ITEM) AND RTRIM(Stock.AUDTORG) = RTRIM(Sales.AUDTORG)
LEFT JOIN
(SELECT ITEMNO, AUDTORG,SUM(QTY) AS SIT FROM GIT WHERE OPENINGDATE = convert(varchar, getdate(), 23) GROUP BY ITEMNO, AUDTORG) as GIT
ON Stock.ITEMNO = GIT.ITEMNO AND Stock.AUDTORG=GIT.AUDTORG
left join
(select distinct  BRANCHNAME,BRANCH,NDMNAME from NDM ) as NDM
ON (Stock.AUDTORG=NDM.BRANCH)
-------
Group BY NDMNAME,Stock.ITEMNO, Stock.AUDTORG) AS TX
GROUP BY NDMNAME,AUDTORG
ORDER BY NDMNAME, AUDTORG
"""

cursor.execute(query)
data = list(cursor.fetchall())

branchlist = []
SUS = []
US = []
NS = []
OS = []
SOS = []
NIL = []
BranchItem = []
TotalItem = []
rownumber = 0
for row in data:
        branchname=data[rownumber][1]
        branchlist.append(branchname)
        sus = data[rownumber][2]
        SUS.append(sus)
        us = data[rownumber][3]
        US.append(us)
        ns = data[rownumber][4]
        NS.append(ns)
        os = data[rownumber][5]
        OS.append(os)
        sos = data[rownumber][6]
        SOS.append(sos)
        nil = data[rownumber][7]
        NIL.append(nil)
        bi = data[rownumber][8]
        BranchItem.append(bi)
        ti = data[rownumber][9]
        TotalItem.append(ti)
        rownumber = rownumber+1


font = {'family': 'serif',
        'color':  '#3b5998',
        'weight': 400,
        'size': 15,
        }

font1 = {'family': 'serif',
        'color':  'darkred',
        'weight': 700,
        'size': 15,
        }


legend_elements = [Patch(facecolor='#be4748',label='Mr. Kamrul')
                   , Patch(facecolor='#1a58c5',label='Mr. Anwar')
                   , Patch(facecolor='#c5871a',label='Mr. Atik')
                   , Patch(facecolor='#58c51a',label='Mr. Nurul')]

y_pos = np.arange(len(branchlist))

def bardecoration(bars):
    for bar in bars:
        yval = bar.get_height()
        plt.text(bar.get_x(), yval + 3.5, yval, rotation=90, fontsize = 8)

    for bar in bars[0:7]:
        bar.set_color('#be4748')

    for bar in bars[7:15]:
        bar.set_color('#1a58c5')

    for bar in bars[15:23]:
        bar.set_color('#c5871a')

    for bar in bars[23:31]:
        bar.set_color('#58c51a')

    plt.xticks(y_pos, branchlist, rotation=90)
    plt.yticks(np.arange(0, 150, 10), fontsize=6)
    plt.ylabel('No. of Item', fontdict=font)
    plt.legend(handles=legend_elements, loc='best')


suskam = 0
uskam = 0
nskam = 0
soskam = 0
oskam = 0
nilkam = 0

for items in NIL[0:7]:
    nilkam += items

for items in SUS[0:7]:
    suskam += items

for items in US[0:7]:
    uskam += items

for items in SOS[0:7]:
    soskam += items

for items in OS[0:7]:
    oskam += items

for items in NS[0:7]:
    nskam += items

suswar = 0
uswar = 0
nswar = 0
soswar = 0
oswar = 0
nilwar = 0

for items in NIL[7:15]:
    nilwar += items

for items in SUS[7:15]:
    suswar += items

for items in US[7:15]:
    uswar += items

for items in SOS[7:15]:
    soswar += items

for items in OS[7:15]:
    oswar += items

for items in NS[7:15]:
    nswar += items

susbatik = 0
usbatik = 0
nsbatik = 0
sosbatik = 0
osbatik = 0
nilbatik = 0

for items in NIL[15:23]:
    nilbatik += items

for items in SUS[15:23]:
    susbatik += items

for items in US[15:23]:
    usbatik += items

for items in SOS[15:23]:
    sosbatik += items

for items in OS[15:23]:
    osbatik += items

for items in NS[15:23]:
    nsbatik += items

susnuru = 0
usnuru = 0
nsnuru = 0
sosnuru = 0
osnuru = 0
nilnuru = 0

for items in NIL[23:31]:
    nilnuru += items

for items in SUS[23:31]:
    susnuru += items

for items in US[23:31]:
    usnuru += items

for items in SOS[23:31]:
    sosnuru += items

for items in OS[23:31]:
    osnuru += items

for items in NS[23:31]:
    nsnuru += items

MovingItem = int(TotalItem[0])
#print(MovingItem)

nuruta = [nilnuru/MovingItem,susnuru/MovingItem,usnuru/MovingItem,nsnuru/MovingItem,osnuru/MovingItem,sosnuru/MovingItem]
batikta = [nilbatik/MovingItem,susbatik/MovingItem,usbatik/MovingItem,nsbatik/MovingItem,osbatik/MovingItem,sosbatik/MovingItem]
warta = [nilwar/MovingItem,suswar/MovingItem,uswar/MovingItem,nswar/MovingItem,oswar/MovingItem,soswar/MovingItem]
kamta = [nilkam/MovingItem,suskam/MovingItem,uskam/MovingItem,nskam/MovingItem,oskam/MovingItem,soskam/MovingItem]

xxx = ['NIL','SUS','US','NS','OS','SOS']
y_pos1 = np.arange(len(xxx))

def branchbars(bars):
    for bar in bars:
        yval = bar.get_height()
        plt.text(bar.get_x()+0.42, yval + 10.25, yval, horizontalalignment='center', verticalalignment='center', fontsize=14)

    plt.xticks(y_pos1, xxx, rotation=0)
    plt.yticks(np.arange(0, 320, 20), fontsize=10)
    plt.ylabel('No. of Item', fontdict=font)
    #plt.tight_layout()
    bars[0].set_color('#990000')
    bars[1].set_color('#ff0000')
    bars[2].set_color('#e6e600')
    bars[3].set_color('#008000')
    bars[4].set_color('#00ff16')
    bars[5].set_color('darkorange')