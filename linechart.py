import pyodbc as db
import numpy as np
#####-----Table



font = {'family': 'serif',
        'color':  '#3b5998',
        'weight': 400,
        'size': 15,
        }

font1 = {'family': 'serif',
        'color':  '#3b5998',
        'weight': 700,
        'size': 17,
        }

connection = db.connect('DRIVER={SQL Server};'
                        'SERVER=10.168.2.241;'
                        'DATABASE=ARCHIVESKF;'
                        'UID=sa;PWD=erp')#####Add trusted_connection=yes and remove uid and pwd while connecting to local

cursor = connection.cursor()

##########AtikLine-------------------------------------------------------------------------------------------------------------------
querytik = """SELECT [NDMNAME]
      ,[AUDTORG]
      ,[SUS]
      ,[US]
      ,[NS]
      ,[OS]
      ,[SOS]
      ,[NIL]
      ,[TotalItem]
      ,[AllItem]
      ,RIGHT([AUDTDATE],2)
  FROM [ARCHIVESKF].[dbo].[StockChart] WHERE NDMNAME='Md. Atikullaha'"""

days = np.arange(1,31,1)
xvalues = np.arange(1,31,2)
yvalues = np.arange(1,15,2)

cursor.execute(querytik)
datatik = list(cursor.fetchall())

branchlist = []
SUStik = []
UStik = []
NStik = []
OStik = []
SOStik = []
NILtik = []
BranchItem = []
TotalItem = []
AUDTDATE = []
rownumber = 0
for row in datatik:
        branchname=datatik[rownumber][1]
        branchlist.append(branchname)
        sustik = datatik[rownumber][2]
        SUStik.append(sustik)
        ustik = datatik[rownumber][3]
        UStik.append(ustik)
        nstik = datatik[rownumber][4]
        NStik.append(nstik)
        ostik = datatik[rownumber][5]
        OStik.append(ostik)
        sostik = datatik[rownumber][6]
        SOStik.append(sostik)
        niltik = datatik[rownumber][7]
        NILtik.append(niltik)
        bi = datatik[rownumber][8]
        BranchItem.append(bi)
        ti = datatik[rownumber][9]
        TotalItem.append(ti)
        audtdate = datatik[rownumber][10]
        AUDTDATE.append(audtdate)
        rownumber = rownumber+1

sustik30=[]
i = 0
j=0
while j<30:
    data = sum(SUStik[i:i+8])
    sustik30.append((data/(8*TotalItem[i]))*100)
    i=i+8
    j=j+1
sustik30.reverse()

ustik30=[]
i = 0
j=0
while j<30:
    data = sum(UStik[i:i+8])
    ustik30.append((data/(8*TotalItem[i]))*100)
    i=i+8
    j=j+1
ustik30.reverse()

nstik30=[]
i = 0
j=0
while j<30:
    data = sum(NStik[i:i+8])
    nstik30.append((data/(8*TotalItem[i]))*100)
    i=i+8
    j=j+1
nstik30.reverse()

ostik30=[]
i = 0
j=0
while j<30:
    data = sum(OStik[i:i+8])
    ostik30.append((data/(8*TotalItem[i]))*100)
    i=i+8
    j=j+1
ostik30.reverse()

sostik30=[]
i = 0
j=0
while j<30:
    data = sum(SOStik[i:i+8])
    sostik30.append((data/(8*TotalItem[i]))*100)
    i=i+8
    j=j+1
sostik30.reverse()

niltik30=[]
i = 0
j=0
while j<30:
    data = sum(NILtik[i:i+8])
    niltik30.append((data/(8*TotalItem[i]))*100)
    i=i+8
    j=j+1
niltik30.reverse()

datetik30=[]
i = 0
j=0
while j<30:
    data = AUDTDATE[i]
    datetik30.append(data)
    i=i+8
    j=j+1

datetik30.reverse()


#######-------------------------------------------------------------------------------
querywar = """SELECT [NDMNAME]
      ,[AUDTORG]
      ,[SUS]
      ,[US]
      ,[NS]
      ,[OS]
      ,[SOS]
      ,[NIL]
      ,[TotalItem]
      ,[AllItem]
      ,RIGHT([AUDTDATE],2)
  FROM [ARCHIVESKF].[dbo].[StockChart] WHERE NDMNAME='Md. Anwar Hossain'"""

cursor.execute(querywar)
datawar = list(cursor.fetchall())

branchlist = []
SUSwar = []
USwar = []
NSwar = []
OSwar = []
SOSwar = []
NILwar = []
BranchItem = []
TotalItem = []
AUDTDATE = []
rownumber = 0
for row in datawar:
        branchname=datawar[rownumber][1]
        branchlist.append(branchname)
        suswar = datawar[rownumber][2]
        SUSwar.append(suswar)
        uswar = datawar[rownumber][3]
        USwar.append(uswar)
        nswar = datawar[rownumber][4]
        NSwar.append(nswar)
        oswar = datawar[rownumber][5]
        OSwar.append(oswar)
        soswar = datawar[rownumber][6]
        SOSwar.append(soswar)
        nilwar = datawar[rownumber][7]
        NILwar.append(nilwar)
        bi = datawar[rownumber][8]
        BranchItem.append(bi)
        ti = datawar[rownumber][9]
        TotalItem.append(ti)
        audtdate = datawar[rownumber][10]
        AUDTDATE.append(audtdate)
        rownumber = rownumber+1

suswar30=[]
i = 0
j=0
while j<30:
    data = sum(SUSwar[i:i+8])
    suswar30.append((data/(8*TotalItem[i]))*100)
    i=i+8
    j=j+1
suswar30.reverse()

uswar30=[]
i = 0
j=0
while j<30:
    data = sum(USwar[i:i+8])
    uswar30.append((data/(8*TotalItem[i]))*100)
    i=i+8
    j=j+1
uswar30.reverse()

nswar30=[]
i = 0
j=0
while j<30:
    data = sum(NSwar[i:i+8])
    nswar30.append((data/(8*TotalItem[i]))*100)
    i=i+8
    j=j+1
nswar30.reverse()

oswar30=[]
i = 0
j=0
while j<30:
    data = sum(OSwar[i:i+8])
    oswar30.append((data/(8*TotalItem[i]))*100)
    i=i+8
    j=j+1
oswar30.reverse()

soswar30=[]
i = 0
j=0
while j<30:
    data = sum(SOSwar[i:i+8])
    soswar30.append((data/(8*TotalItem[i]))*100)
    i=i+8
    j=j+1
soswar30.reverse()

nilwar30=[]
i = 0
j=0
while j<30:
    data = sum(NILwar[i:i+8])
    nilwar30.append((data/(8*TotalItem[i]))*100)
    i=i+8
    j=j+1
nilwar30.reverse()

datewar30=[]
i = 0
j=0
while j<30:
    data = AUDTDATE[i]
    datewar30.append(data)
    i=i+8
    j=j+1

datewar30.reverse()

querykam = """SELECT [NDMNAME]
      ,[AUDTORG]
      ,[SUS]
      ,[US]
      ,[NS]
      ,[OS]
      ,[SOS]
      ,[NIL]
      ,[TotalItem]
      ,[AllItem]
      ,RIGHT([AUDTDATE],2)
  FROM [ARCHIVESKF].[dbo].[StockChart] WHERE NDMNAME='Islam Md. Kamrul Ahsan'"""

cursor.execute(querykam)
datakam = list(cursor.fetchall())

branchlist = []
SUSkam = []
USkam = []
NSkam = []
OSkam = []
SOSkam = []
NILkam = []
BranchItem = []
TotalItem = []
AUDTDATE = []
rownumber = 0
for row in datakam:
        branchname=datakam[rownumber][1]
        branchlist.append(branchname)
        suskam = datakam[rownumber][2]
        SUSkam.append(suskam)
        uskam = datakam[rownumber][3]
        USkam.append(uskam)
        nskam = datakam[rownumber][4]
        NSkam.append(nskam)
        oskam = datakam[rownumber][5]
        OSkam.append(oskam)
        soskam = datakam[rownumber][6]
        SOSkam.append(soskam)
        nilkam = datakam[rownumber][7]
        NILkam.append(nilkam)
        bi = datakam[rownumber][8]
        BranchItem.append(bi)
        ti = datakam[rownumber][9]
        TotalItem.append(ti)
        audtdate = datakam[rownumber][10]
        AUDTDATE.append(audtdate)
        rownumber = rownumber+1

suskam30=[]
i = 0
j=0
while j<30:
    data = sum(SUSkam[i:i+7])
    suskam30.append((data/(7*TotalItem[i]))*100)
    i=i+7
    j=j+1
suskam30.reverse()

uskam30=[]
i = 0
j=0
while j<30:
    data = sum(USkam[i:i+7])
    uskam30.append((data/(7*TotalItem[i]))*100)
    i=i+7
    j=j+1
uskam30.reverse()

nskam30=[]
i = 0
j=0
while j<30:
    data = sum(NSkam[i:i+7])
    nskam30.append((data/(7*TotalItem[i]))*100)
    i=i+7
    j=j+1
nskam30.reverse()

oskam30=[]
i = 0
j=0
while j<30:
    data = sum(OSkam[i:i+7])
    oskam30.append((data/(7*TotalItem[i]))*100)
    i=i+7
    j=j+1
oskam30.reverse()

soskam30=[]
i = 0
j=0
while j<30:
    data = sum(SOSkam[i:i+7])
    soskam30.append((data/(7*TotalItem[i]))*100)
    i=i+7
    j=j+1
soskam30.reverse()

nilkam30=[]
i = 0
j=0
while j<30:
    data = sum(NILkam[i:i+7])
    nilkam30.append((data/(7*TotalItem[i]))*100)
    i=i+7
    j=j+1
nilkam30.reverse()

datekam30=[]
i = 0
j=0
while j<30:
    data = AUDTDATE[i]
    datekam30.append(data)
    i=i+7
    j=j+1

datekam30.reverse()

querynuru = """SELECT [NDMNAME]
      ,[AUDTORG]
      ,[SUS]
      ,[US]
      ,[NS]
      ,[OS]
      ,[SOS]
      ,[NIL]
      ,[TotalItem]
      ,[AllItem]
      ,RIGHT([AUDTDATE],2)
  FROM [ARCHIVESKF].[dbo].[StockChart] WHERE NDMNAME='Md. Nurul Amin'"""

cursor.execute(querynuru)
datanuru = list(cursor.fetchall())

branchlist = []
SUSnuru = []
USnuru = []
NSnuru = []
OSnuru = []
SOSnuru = []
NILnuru = []
BranchItem = []
TotalItem = []
AUDTDATE = []
rownumber = 0
for row in datanuru:
        branchname=datanuru[rownumber][1]
        branchlist.append(branchname)
        susnuru = datanuru[rownumber][2]
        SUSnuru.append(susnuru)
        usnuru = datanuru[rownumber][3]
        USnuru.append(usnuru)
        nsnuru = datanuru[rownumber][4]
        NSnuru.append(nsnuru)
        osnuru = datanuru[rownumber][5]
        OSnuru.append(osnuru)
        sosnuru = datanuru[rownumber][6]
        SOSnuru.append(sosnuru)
        nilnuru = datanuru[rownumber][7]
        NILnuru.append(nilnuru)
        bi = datanuru[rownumber][8]
        BranchItem.append(bi)
        ti = datanuru[rownumber][9]
        TotalItem.append(ti)
        audtdate = datanuru[rownumber][10]
        AUDTDATE.append(audtdate)
        rownumber = rownumber+1

susnuru30=[]
i = 0
j=0
while j<30:
    data = sum(SUSnuru[i:i+8])
    susnuru30.append((data/(8*TotalItem[i]))*100)
    i=i+8
    j=j+1
susnuru30.reverse()

usnuru30=[]
i = 0
j=0
while j<30:
    data = sum(USnuru[i:i+8])
    usnuru30.append((data/(8*TotalItem[i]))*100)
    i=i+8
    j=j+1
usnuru30.reverse()

nsnuru30=[]
i = 0
j=0
while j<30:
    data = sum(NSnuru[i:i+8])
    nsnuru30.append((data/(8*TotalItem[i]))*100)
    i=i+8
    j=j+1
nsnuru30.reverse()

osnuru30=[]
i = 0
j=0
while j<30:
    data = sum(OSnuru[i:i+8])
    osnuru30.append((data/(8*TotalItem[i]))*100)
    i=i+8
    j=j+1
osnuru30.reverse()

sosnuru30=[]
i = 0
j=0
while j<30:
    data = sum(SOSnuru[i:i+8])
    sosnuru30.append((data/(8*TotalItem[i]))*100)
    i=i+8
    j=j+1
sosnuru30.reverse()

nilnuru30=[]
i = 0
j=0
while j<30:
    data = sum(NILnuru[i:i+8])
    nilnuru30.append((data/(8*TotalItem[i]))*100)
    i=i+8
    j=j+1
nilnuru30.reverse()

datenuru30=[]
i = 0
j=0
while j<30:
    data = AUDTDATE[i]
    datenuru30.append(data)
    i=i+8
    j=j+1

datenuru30.reverse()