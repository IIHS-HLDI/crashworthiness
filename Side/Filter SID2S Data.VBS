'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 04/07/2010 14:15:08
'-- Author: 
'-- Comment: Filter data
'-------------------------------------------------------------------------------

Dim group_count

group_count = Data.Root.ChannelGroups.Count

Call Data.Root.ChannelGroups.Add("Vehicle", group_count + 1).Activate 

IF B1 = TRUE THEN  'Filter driver data if checkbox was selected
  Call Filterdata(11)
END IF

IF B2 = TRUE THEN 'Filter rear passenger data if checkbox was selected
  Call Filterdata(14)
END IF
'-----

Sub Filterdata(Occ_Position)

    'HEAD 
    Call Data.Root.ChannelGroups.Add(Occ_Position & "_Head_Neck", group_count + 1).Activate 
    group_count = group_count + 1
    Call ChnCFCFiltCalc("","[1]/" & Occ_Position & "HEAD0000__ACX_","/" & Occ_Position & "HEADCG00__ACXA","CFC_1000",0,"EndPoints",10) 
    Call ChnCFCFiltCalc("","[1]/" & Occ_Position & "HEAD0000__ACY_","/" & Occ_Position & "HEADCG00__ACYA","CFC_1000",0,"EndPoints",10) 
    Call ChnCFCFiltCalc("","[1]/" & Occ_Position & "HEAD0000__ACZ_","/" & Occ_Position & "HEADCG00__ACZA","CFC_1000",0,"EndPoints",10) 

    Call ChnCFCFiltCalc("","[1]/" & Occ_Position & "HEADCG0000AVX_","/" & Occ_Position & "HEADCG0000AVXC","CFC_180",0,"EndPoints",10) 
    Call ChnCFCFiltCalc("","[1]/" & Occ_Position & "HEADCG0000AVY_","/" & Occ_Position & "HEADCG0000AVYC","CFC_180",0,"EndPoints",10) 
    Call ChnCFCFiltCalc("","[1]/" & Occ_Position & "HEADCG0000AVZ_","/" & Occ_Position & "HEADCG0000AVZC","CFC_180",0,"EndPoints",10) 

    'NECK 
    Call ChnCFCFiltCalc("","[1]/" & Occ_Position & "NECKUP00__FOX_","/" & Occ_Position & "NECKUP00__FOXA","CFC_1000",0,"EndPoints",10) 
    Call ChnCFCFiltCalc("","[1]/" & Occ_Position & "NECKUP00__FOY_","/" & Occ_Position & "NECKUP00__FOYA","CFC_1000",0,"EndPoints",10) 
    Call ChnCFCFiltCalc("","[1]/" & Occ_Position & "NECKUP00__FOZ_","/" & Occ_Position & "NECKUP00__FOZA","CFC_1000",0,"EndPoints",10)
    Call ChnCFCFiltCalc("","[1]/" & Occ_Position & "NECKUP00__MOX_","/" & Occ_Position & "NECKUP00__MOXB","CFC_600",0,"EndPoints",10) 
    Call ChnCFCFiltCalc("","[1]/" & Occ_Position & "NECKUP00__MOZ_","/" & Occ_Position & "NECKUP00__MOZB","CFC_600",0,"EndPoints",10) 

    Call Data.Root.ChannelGroups.Add(Occ_Position & "_Spine", group_count + 1).Activate 
    group_count = group_count + 1
    'SPINE accelerometers 
    Call ChnCFCFiltCalc("","[1]/" & Occ_Position & "SPIN0100__ACY_","/" & Occ_Position & "SPIN0100__ACYA","CFC_1000",0,"EndPoints",10) 
    Call ChnCFCFiltCalc("","[1]/" & Occ_Position & "SPIN0400__ACY_","/" & Occ_Position & "SPIN0400__ACYA","CFC_1000",0,"EndPoints",10) 
    Call ChnCFCFiltCalc("","[1]/" & Occ_Position & "SPIN1200__ACY_","/" & Occ_Position & "SPIN1200__ACYA","CFC_1000",0,"EndPoints",10) 

    Call Data.Root.ChannelGroups.Add(Occ_Position & "_Shoulder", group_count + 1).Activate 
    group_count = group_count + 1
    'Shoulder
    Call ChnCFCFiltCalc("","[1]/" & Occ_Position & "SHRILE00__DSY_","/" & Occ_Position & "SHRILE00__DSYB","CFC_600",0,"EndPoints",10)
    Call ChnCFCFiltCalc("","[1]/" & Occ_Position & "SHLDLE00__FOX_","/" & Occ_Position & "SHLDLE00__FOXA","CFC_1000",0,"EndPoints",10) 
    Call ChnCFCFiltCalc("","[1]/" & Occ_Position & "SHLDLE00__FOY_","/" & Occ_Position & "SHLDLE00__FOYA","CFC_1000",0,"EndPoints",10) 
    Call ChnCFCFiltCalc("","[1]/" & Occ_Position & "SHLDLE00__FOZ_","/" & Occ_Position & "SHLDLE00__FOZA","CFC_1000",0,"EndPoints",10) 

    Call Data.Root.ChannelGroups.Add(Occ_Position & "_Torso", group_count + 1).Activate 
    group_count = group_count + 1
    'Rib accelerometers
    Call ChnCFCFiltCalc("","[1]/" & Occ_Position & "TRRI01LE__ACY_","/" & Occ_Position & "TRRI01LE__ACYA","CFC_1000",0,"EndPoints",10) 
    Call ChnCFCFiltCalc("","[1]/" & Occ_Position & "TRRI02LE__ACY_","/" & Occ_Position & "TRRI02LE__ACYA","CFC_1000",0,"EndPoints",10) 
    Call ChnCFCFiltCalc("","[1]/" & Occ_Position & "TRRI03LE__ACY_","/" & Occ_Position & "TRRI03LE__ACYA","CFC_1000",0,"EndPoints",10) 
    Call ChnCFCFiltCalc("","[1]/" & Occ_Position & "ABRI01LE__ACY_","/" & Occ_Position & "ABRI01LE__ACYA","CFC_1000",0,"EndPoints",10) 
    Call ChnCFCFiltCalc("","[1]/" & Occ_Position & "ABRI02LE__ACY_","/" & Occ_Position & "ABRI02LE__ACYA","CFC_1000",0,"EndPoints",10) 
    'Rib potentiometers
    Call ChnCFCFiltCalc("","[1]/" & Occ_Position & "TRRI01LE__DSY_","/" & Occ_Position & "TRRI01LE__DSYB","CFC_600",0,"EndPoints",10) 
    Call ChnCFCFiltCalc("","[1]/" & Occ_Position & "TRRI02LE__DSY_","/" & Occ_Position & "TRRI02LE__DSYB","CFC_600",0,"EndPoints",10) 
    Call ChnCFCFiltCalc("","[1]/" & Occ_Position & "TRRI03LE__DSY_","/" & Occ_Position & "TRRI03LE__DSYB","CFC_600",0,"EndPoints",10) 
    Call ChnCFCFiltCalc("","[1]/" & Occ_Position & "ABRI01LE__DSY_","/" & Occ_Position & "ABRI01LE__DSYB","CFC_600",0,"EndPoints",10) 
    Call ChnCFCFiltCalc("","[1]/" & Occ_Position & "ABRI02LE__DSY_","/" & Occ_Position & "ABRI02LE__DSYB","CFC_600",0,"EndPoints",10) 


    Call Data.Root.ChannelGroups.Add(Occ_Position & "_Pelvis_Femur", group_count + 1).Activate 
    group_count = group_count + 1
    'Pelvis
    Call ChnCFCFiltCalc("","[1]/" & Occ_Position & "PELV0000__ACY_","/" & Occ_Position & "PELV0000__ACYA","CFC_1000",0,"EndPoints",10) 
    Call ChnCFCFiltCalc("","[1]/" & Occ_Position & "ILACLE00__FOY_","/" & Occ_Position & "ILACLE00__FOYB","CFC_600",0,"EndPoints",10)
    Call ChnCFCFiltCalc("","[1]/" & Occ_Position & "ACTBLE00__FOY_","/" & Occ_Position & "ACTBLE00__FOYB","CFC_600",0,"EndPoints",10)
    'Femur
    Call ChnCFCFiltCalc("","[1]/" & Occ_Position & "FEMRLELO__FOY_","/" & Occ_Position & "FEMRLELO__FOYB","CFC_600",0,"EndPoints",10)
    Call ChnCFCFiltCalc("","[1]/" & Occ_Position & "FEMRLELO__MOX_","/" & Occ_Position & "FEMRLELO__MOXB","CFC_600",0,"EndPoints",10)
    Call ChnCFCFiltCalc("","[1]/" & Occ_Position & "FEMRLELO__MOY_","/" & Occ_Position & "FEMRLELO__MOYB","CFC_600",0,"EndPoints",10)


    'VEHICLE 
    IF Occ_Position = 11 THEN
      Call Data.Root.ChannelGroups("Vehicle").Activate 
      Call ChnCFCFiltCalc("","[1]/" & Occ_Position & "VEHC0000__ACY_","/" & "10VEHCRIFR__ACYC","CFC_180",0,"EndPoints",10)
      Call ChnCFCFiltCalc("","[1]/" & Occ_Position & "VEHC0000__ACY_","/" & "10VEHCRIFR__ACYD","CFC_60",0,"EndPoints",10)
    END IF

    IF Occ_Position = 14 THEN
      Call Data.Root.ChannelGroups("Vehicle").Activate 
      Call ChnCFCFiltCalc("","[1]/" & Occ_Position & "VEHC0000__ACY_","/" & "10VEHCRIRE__ACYC","CFC_180",0,"EndPoints",10)
      Call ChnCFCFiltCalc("","[1]/" & Occ_Position & "VEHC0000__ACY_","/" & "10VEHCRIRE__ACYD","CFC_60",0,"EndPoints",10)
    END IF

End Sub
