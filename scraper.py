
import xlsxwriter
import win32com.client as win32
from pathlib import Path
from oppai import main
import os
def writer(worksheet,setlink,bursts,row,filepath,setid,result):
    #setlink
    worksheet.write_url(row,1,"https://osu.ppy.sh/s/%s"%(setid))
    #difficulty link to oppai here
    #stars/pp/pp95
    try:
        
        worksheet.write(row,2,result[0])
        worksheet.write(row,4,result[1])
        worksheet.write(row,5,result[2])
    except:
        pass
    #bursts
    worksheet.write(row,3,bursts)
#bpm to ms =60,000 / BPM = one beat in milliseconds
#320bpm=188ms upper bound--notes must be spaced more than this
upper=93
lower=188

workbook = xlsxwriter.Workbook('Alternating Map Collection.xlsx')
bold = workbook.add_format({'bold': 1})
NomodSheet = workbook.add_worksheet('Nomod')
NomodSheet.write('A1', 'Name', bold)
NomodSheet.write('B1', 'SetLink', bold)
NomodSheet.write('C1', 'Difficulty', bold)
NomodSheet.write('D1', 'Bursts', bold)
NomodSheet.write('E1', 'SS pp', bold)
NomodSheet.write('F1', '95% pp', bold)


AR10Sheet = workbook.add_worksheet('AR8.7+DT')
AR10Sheet.write('A1', 'Name', bold)
AR10Sheet.write('B1', 'SetLink', bold)
AR10Sheet.write('C1', 'Difficulty', bold)
AR10Sheet.write('D1', 'Bursts', bold)
AR10Sheet.write('E1', 'SS pp', bold)
AR10Sheet.write('F1', '95% pp', bold)

DTSheet = workbook.add_worksheet('AR8.7-DT')
DTSheet.write('A1', 'Name', bold)
DTSheet.write('B1', 'SetLink', bold)
DTSheet.write('C1', 'Difficulty', bold)
DTSheet.write('D1', 'Bursts', bold)
DTSheet.write('E1', 'SS pp', bold)
DTSheet.write('F1', '95% pp', bold)


percent_fmt = workbook.add_format({'num_format': '0.00%'})
NomodSheet.set_column('D:D',None, percent_fmt)
AR10Sheet.set_column('D:D',None, percent_fmt)
DTSheet.set_column('D:D',None, percent_fmt)

decimal_fmt=workbook.add_format({'num_format': '0.00'})
NomodSheet.set_column('C:C',None, decimal_fmt)
AR10Sheet.set_column('C:C',None, decimal_fmt)
DTSheet.set_column('C:C',None, decimal_fmt)
NomodSheet.set_column('E:E',None, decimal_fmt)
AR10Sheet.set_column('E:E',None, decimal_fmt)
DTSheet.set_column('E:E',None, decimal_fmt)
NomodSheet.set_column('F:F',None, decimal_fmt)
AR10Sheet.set_column('F:F',None, decimal_fmt)
DTSheet.set_column('F:F',None, decimal_fmt)

url_format = workbook.add_format({
    'font_color': 'blue',
    'underline':  1
})
row1=1
row2=1
row3=1
counter=0
for (subdir, dirs, files) in os.walk('.'):

    for file in files:
        filepath = subdir + os.sep + file
        if filepath.endswith('.osu'):
            #print(filepath)
            with open(filepath, 'r',encoding='utf-8',errors='ignore') as f:
                hit = False
                AR=0
                notes=0
                prev=0
                match=0    
                matchdt=0
                ub=0
                ubdt=0
                lb=0
                lbdt=0;
                osu=False
                beatMapId=''
                for line in f:
                    if 'Mode: 0' in line:
                        osu=True
                        #print(file)
                        setid=(os.path.split(os.path.abspath(filepath))[0].split('\\')[-1].split(' ')[0])
                        
                    if 'HitObject' in line:
                        hit = True
                    if 'ApproachRate' in line:
                        AR=line.split(':')[1]
                    if 'BeatmapID' in line:
                        beatMapId=line.split(':')[1]
                    if 'BeatmapSetID' in line:
                        setid=line.split(':')[1]
                    if hit == True and osu==True:

                        currentline = line.split(',')
                        if len(currentline)>1:
                            notes=notes+1
                            difference=int(currentline[2])-int(prev)
                            if difference>=upper and difference<=lower:
                                match=match+1
                            if difference*2/3>=upper and difference*2/3<=lower:
                                matchdt=matchdt+1
                            if difference<upper:
                                print(int(currentline[2]))
                                ub=ub+1
                            if difference*2/3<upper :
                                ubdt=ubdt+1
                            if difference>=lower:
                                lb=lb+1 
                            if difference*2/3>=lower:
                                lbdt=lbdt+1 
                            prev=currentline[2]
                
                print("hitobjects:",notes," match:",match,"percentage:",match/notes*100,"match+dt:",matchdt,"percentage:",matchdt/notes*100) #percentage of notes that are alt notes
                print("lb:",lb,"percentage:",lb/notes*100,"lb+dt:",lbdt,"percentage:",lbdt/notes*100) #percentage of notes that are too slow
                print("ub:",ub,"percentage:",ub/notes*100,"ub+dt:",ubdt,"percentage:",ubdt/notes*100) #percentage of notes that are too fast
                if osu==True and notes>=10:
                    pmatch=match/notes*100
                    pmatchdt=matchdt/notes*100
                    plb=lb/notes*100
                    plbdt=lbdt/notes*100
                    pub=ub/notes
                    pubdt=ubdt/notes
                    
                    print(file)
                    if pmatchdt>=25 and plbdt<=80 and pubdt<=0.05 and float(AR)<8.7:
                        result=main("+ DT",filepath)
                        if result[0]>3.5:
                            if beatMapId:
                                 DTSheet.write_url(row1,0,'https://osu.ppy.sh/b/%s'%(beatMapId),url_format,"%s +DT"%(file.split('.os')[0]))
                            else:
                                   DTSheet.write(row1,0,'%s +DT'%(file.split('.os')[0]))
                            writer(DTSheet,file.split('.os')[0],pubdt,row1,filepath,setid,result)
                            row1=row1+1
                            counter=counter+1
                    elif pmatch>=25 and plb<=80 and pub<=0.05:
                        result=main("",filepath)
                        if result[0]>3.5:
                            #name/set/diff/bursts
                            if beatMapId:
                                NomodSheet.write_url(row2,0,'https://osu.ppy.sh/b/%s'%(beatMapId),url_format,"%s"%(file.split('.os')[0]))
                            else:
                                NomodSheet.write(row2,0,'%s'%(file.split('.os')[0]))
                            writer(NomodSheet,file.split('.os')[0],pub,row2,filepath,setid,result)
                            row2=row2+1
                            counter=counter+1
                    elif pmatchdt>=25 and plbdt<=80 and pubdt<=0.05 and float(AR)>=8.7:
                        result=main("+ DT",filepath)
                        if result[0]>3.5:
                            if beatMapId:
                                AR10Sheet.write_url(row3,0,'https://osu.ppy.sh/b/%s'%(beatMapId),url_format,"%s +DT"%(file.split('.os')[0]))
                            else:
                               AR10Sheet.write(row3,0,'%s +DT'%(file.split('.os')[0]))
                            writer(AR10Sheet,file.split('.os')[0],pubdt,row3,filepath,setid,result)
                            row3=row3+1
                            counter=counter+1
print(counter)
print(row1)
print(row2)
print(row3)
workbook.close()

excel = win32.gencache.EnsureDispatch('Excel.Application')
dir = os.path.dirname(__file__)
loc=os.path.join(dir, 'Alternating Map Collection.xlsx')
wb = excel.Workbooks.Open(loc)
ws = wb.Worksheets("AR8.7+DT")
ws2 = wb.Worksheets("Nomod")
ws3 = wb.Worksheets("AR8.7-DT")
ws.Columns.AutoFit()
ws2.Columns.AutoFit()
ws3.Columns.AutoFit()
wb.Save()
excel.Application.Quit()
#name link written
#name/setlink/stars/bursts/ss/95
