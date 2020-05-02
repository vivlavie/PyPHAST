import os
import win32com.client

NumDirections = 6

def EventTree_OLF2007(pv,hole,weather,lEvent,print2pdf=False):
    import shutil
    #Failure of fire detector for immediate ignition
    P_FD_Fail = 0.0
    #Failure of gas detectin for delayed ignition
    P_GD_Fail = 0.0
    et_filename = '.\\et\\ET_'+pv+'_'+hole+'_'+weather+'.xlsx'
    if et_filename in os.listdir():
        iExl=load_workbook(filename=et_filename)
    else:
        shutil.copy('ET_template.xlsx',et_filename)
        iExl=load_workbook(filename=et_filename)    
    
    shET = iExl['et']
    basic_freq_set = False
    for e in lEvent:
        if ((pv in e.Key) and (hole in e.Hole) and (weather == e.Weather)):
            if e.Discharge.ReleaseRate <= 1:
                P_FD_Fail = 0.1
                P_GD_Fail = 0.1
            elif e.Discharge.ReleaseRate <= 10:
                P_FD_Fail = 0.05
                P_GD_Fail = 0.05
            elif e.Discharge.ReleaseRate > 10:
                P_FD_Fail = 0.005
                P_GD_Fail = 0.005
            else:
                print('something wrong in P_FD_Fail')
            epv,ehole,eweather = e.Key.split("\\")
            print("{:20s}{:10s} {:6s} {:8.2e} {:8.2e} {:8.2e} - Leak Freq {:8.2e} ".format(epv,ehole,eweather,e.Frequency,e.PESD,e.PBDV,e.Frequency/e.PESD/e.PBDV))
            if basic_freq_set == False:
                shET.cell(1,6).value = pv
                shET.cell(1,7).value = hole
                shET.cell(1,8).value = weather
                shET.cell(30,2).value = e.Frequency/e.PESD/e.PBDV
                shET.cell(31,2).value = e.Discharge.ReleaseRate
                shET.cell(20,6).value = e.PImdIgn
                shET.cell(26,10).value = e.PDelIgn
                shET.cell(15,7).value = 1-P_FD_Fail
                shET.cell(42,7).value = 1-P_GD_Fail
                # shET.cell(10,8).value = 1-0.01*numESDVs[pv]
                shET.cell(10,8).value = e.PESD #PESD: p of successful ESD
                # if "BN" in e.Hole:
                #     shET.cell(9,9).value = 1
                # else:
                #     shET.cell(9,9).value = 1-0.005
                shET.cell(9,9).value = e.PBDV #PBDV: p of successful BDV

                basic_freq_set = True
            else:
                if abs(shET.cell(30,2).value - e.Frequency/e.PESD/e.PBDV)/(e.Frequency/e.PESD/e.PBDV) > 0.01:
                    print(pv, epv, ehole,'Basic frequency wrong',shET.cell(30,2).value, e.Frequency/e.PESD/e.PBDV, e.PESD, e.PBDV)
                    # break
                # else:
                #     print(epv,hole,'Basic frequency match')
                    
                if shET.cell(20,6).value != e.PImdIgn:
                    print(epv, hole, 'ImdIgn Wrong')
                    # break
                # else:
                #     print(epv,hole,'ImdIgn match')
                    
            if e.Explosion == None:
                e.Explosion = Explosion(0.,0.,0.)
            #Eithter fire detection for jet fire or gas detection for explosion has succedded
            if ("EOBO" in ehole) or ("EOBN" in ehole):
                shET.cell(9,13).value = e.JetFire.Frequency*NumDirections*(1-P_FD_Fail) #Jet fire frequency = LeakFreq x PEO x PBO x PImgIgn / NumDir
                shET.cell(25,13).value = e.Explosion.Frequency*(1-P_GD_Fail) #VCE                
                # shET.cell(28,13).value = e.Frequency*(1-e.PImdIgn)*(e.PDelIgn)*(1-e.PExp_Ign)*(1-P_GD_Fail) #Flash fire
                # shET.cell(31,13).value = e.Frequency*(1-e.PImdIgn)*(1-e.PDelIgn)*(1-P_GD_Fail) #Gas disperison
                shET.cell(28,13).value = e.Frequency*(e.PDelIgn)*(1-e.PExp_Ign)*(1-P_GD_Fail) #Flash fire
                shET.cell(31,13).value = e.Frequency*(1-e.PDelIgn)*(1-P_GD_Fail) #Gas disperison
                # print(pv,epv,ehole,shET.cell(9,13).value, shET.cell(25,13).value, shET.cell(31,13).value)
            if "EOBX" in ehole:
                shET.cell(12,13).value = e.JetFire.Frequency*NumDirections*(1-P_FD_Fail)
                shET.cell(34,13).value = e.Explosion.Frequency*(1-P_GD_Fail)                
                # shET.cell(37,13).value = e.Frequency*(1-e.PImdIgn)*(e.PDelIgn)*(1-e.PExp_Ign)*(1-P_GD_Fail) #Flash fire
                # shET.cell(40,13).value = e.Frequency*(1-e.PImdIgn)*(1-e.PDelIgn)*(1-P_GD_Fail)
                shET.cell(37,13).value = e.Frequency*(e.PDelIgn)*(1-e.PExp_Ign)*(1-P_GD_Fail) #Flash fire
                shET.cell(40,13).value = e.Frequency*(1-e.PDelIgn)*(1-P_GD_Fail)
                # print(pv,epv,ehole,shET.cell(12,13).value, shET.cell(34,13).value, shET.cell(37,13).value)
            if ("EXBO" in ehole) or ("EXBN" in ehole):
                shET.cell(15,13).value = e.JetFire.Frequency*NumDirections*(1-P_FD_Fail)
                shET.cell(43,13).value = e.Explosion.Frequency*(1-P_GD_Fail)
                # shET.cell(46,13).value = e.Frequency*(1-e.PImdIgn)*(e.PDelIgn)*(1-e.PExp_Ign)*(1-P_GD_Fail) #Flash fire
                # shET.cell(49,13).value = e.Frequency*(1-e.PImdIgn)*(1-e.PDelIgn)*(1-P_GD_Fail)
                shET.cell(46,13).value = e.Frequency*(e.PDelIgn)*(1-e.PExp_Ign)*(1-P_GD_Fail) #Flash fire
                shET.cell(49,13).value = e.Frequency*(1-e.PDelIgn)*(1-P_GD_Fail)
                # print(pv,epv,ehole,shET.cell(15,13).value, shET.cell(43,13).value, shET.cell(49,13).value)
            if "EXBX" in ehole:
                #Two cases should be distinguished; either detectors are successful or no
                #Detection Success but EXBX
                shET.cell(18,13).value = e.JetFire.Frequency*NumDirections*(1-P_FD_Fail)
                shET.cell(52,13).value = e.Explosion.Frequency*(1-P_GD_Fail) #explosion
                # shET.cell(55,13).value = e.Frequency*(1-e.PImdIgn)*(e.PDelIgn)*(1-e.PExp_Ign)*(1-P_GD_Fail) #Flash fire
                # shET.cell(58,13).value = e.Frequency*(1-e.PImdIgn)*(1-e.PDelIgn)*(1-P_GD_Fail) #Dispersion only                
                shET.cell(55,13).value = e.Frequency*(e.PDelIgn)*(1-e.PExp_Ign)*(1-P_GD_Fail) #Flash fire
                shET.cell(58,13).value = e.Frequency*(1-e.PDelIgn)*(1-P_GD_Fail) #Dispersion only                

                #Detection failure & EXBX
                shET.cell(22,13).value = e.JetFire.Frequency*NumDirections/e.PESD/e.PBDV*(P_FD_Fail)                
                shET.cell(61,13).value = e.Explosion.Frequency/e.PESD/e.PBDV*(P_GD_Fail)
                # shET.cell(64,13).value = e.Frequency/e.PESD/e.PBDV*(1-e.PImdIgn)*(e.PDelIgn)*(1-e.PExp_Ign)*(P_GD_Fail) #Flash fire
                # shET.cell(67,13).value = e.Frequency/e.PESD/e.PBDV*(1-e.PImdIgn)*(1-e.PDelIgn)*(P_GD_Fail)
                shET.cell(64,13).value = e.Frequency/e.PESD/e.PBDV*(e.PDelIgn)*(1-e.PExp_Ign)*(P_GD_Fail) #Flash fire
                shET.cell(67,13).value = e.Frequency/e.PESD/e.PBDV*(1-e.PDelIgn)*(P_GD_Fail)
                
                # print(pv,epv,ehole,shET.cell(22,13).value, shET.cell(61,13).value, shET.cell(66,13).value)

    iExl.save(et_filename)
    # print(et_filename)

    if print2pdf == True:        
        o = win32com.client.Dispatch("Excel.Application")
        o.Visible = False
        wb_path = "C:\\Users\\seoshk\\LR\Energy - PRJ11100223773 - Documents\\6. Project Work place\\01 FRA\\PyExdCrv\\Rev.B\\"+ et_filename
        wb = o.Workbooks.Open(wb_path)

        ws_index_list = ['et'] #say you want to print these sheets
        path_to_pdf = wb_path[:-4]+".pdf"
        print_area = 'A1:M71'

        # for index in ws_index_list:
        #     #off-by-one so the user can start numbering the worksheets at 1
        #     ws = wb.Worksheets[index - 1]
        #     ws.PageSetup.Zoom = False
        #     ws.PageSetup.FitToPagesTall = 1
        #     ws.PageSetup.FitToPagesWide = 1
        #     ws.PageSetup.PrintArea = print_area

        wb.WorkSheets(ws_index_list).Select()
        wb.ActiveSheet.ExportAsFixedFormat(0, path_to_pdf)