'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 04/27/2010 09:36:16
'-- Author: 
'-- Comment: Rear Impact Filters
'-------------------------------------------------------------------------------

'Change units from G to gn
Data.Root.ChannelGroups(1).Channels("S0VEHC0100__ACX_").Properties("unit_string").Value = "gn"
Data.Root.ChannelGroups(1).Channels("S0VEHC0200__ACX_").Properties("unit_string").Value = "gn"
Data.Root.ChannelGroups(1).Channels("S0VEHC0300__ACX_").Properties("unit_string").Value = "gn"
Data.Root.ChannelGroups(1).Channels("S0VEHC0300__ACY_").Properties("unit_string").Value = "gn"
Data.Root.ChannelGroups(1).Channels("S0VEHC0300__ACZ_").Properties("unit_string").Value = "gn"
Data.Root.ChannelGroups(1).Channels("S1VEHC0400__ACX_").Properties("unit_string").Value = "gn"
Data.Root.ChannelGroups(1).Channels("S1HEAD0000__ACX_").Properties("unit_string").Value = "gn"
Data.Root.ChannelGroups(1).Channels("S1THSP01LE__ACX_").Properties("unit_string").Value = "gn"
Data.Root.ChannelGroups(1).Channels("S1THSP01RI__ACX_").Properties("unit_string").Value = "gn"

'Filter Head Data
Call ChnCFCFiltCalc("","[1]/S1HEAD0000__ACX_","[2]/S1HEAD0000__ACXA","CFC_1000",0,"EndPoints",10) '... XW,Y,E,CFCFiltType,Fir100RemoveBias,CFCPreEventType,CFCFreeValue 
Call ChnCFCFiltCalc("","[1]/S1HEAD0000__ACX_","[2]/S1HEAD0000__ACXC","CFC_180",0,"EndPoints",10) '... XW,Y,E,CFCFiltType,Fir100RemoveBias,CFCPreEventType,CFCFreeValue
Call ChnCFCFiltCalc("","[1]/S1HEAD0000__ACX_","[2]/S1HEAD0000__ACXD","CFC_60",0,"EndPoints",10) '... XW,Y,E,CFCFiltType,Fir100RemoveBias,CFCPreEventType,CFCFreeValue
Call ChnUnitConvert("[2]/S1HEAD0000__ACXD","[2]/S1HEAD0000__ACXDSI","m/s^2") '... Y,E,TargetUnit CONVERT HEAD ACCELERATION to SI units FOR NIC CALCULATION
Call ChnCFCFiltCalc("","[1]/S1HEAD0000__ACY_","[2]/S1HEAD0000__ACYA","CFC_1000",0,"EndPoints",10) '... XW,Y,E,CFCFiltType,Fir100RemoveBias,CFCPreEventType,CFCFreeValue 
Call ChnCFCFiltCalc("","[1]/S1HEAD0000__ACZ_","[2]/S1HEAD0000__ACZA","CFC_1000",0,"EndPoints",10) '... XW,Y,E,CFCFiltType,Fir100RemoveBias,CFCPreEventType,CFCFreeValue 
'Call ChnCFCFiltCalc("","[1]/S1HEADRE00__FOX_","[2]/S1HEADRE00__FOXA","CFC_1000",0,"EndPoints",10) '... XW,Y,E,CFCFiltType,Fir100RemoveBias,CFCPreEventType,CFCFreeValue 
'Call ChnCFCFiltCalc("","[1]/S1HEADRE00__FOY_","[2]/S1HEADRE00__FOYA","CFC_1000",0,"EndPoints",10) '... XW,Y,E,CFCFiltType,Fir100RemoveBias,CFCPreEventType,CFCFreeValue 
'Call ChnCFCFiltCalc("","[1]/S1HEADRE00__FOZ_","[2]/S1HEADRE00__FOZA","CFC_1000",0,"EndPoints",10) '... XW,Y,E,CFCFiltType,Fir100RemoveBias,CFCPreEventType,CFCFreeValue 

'Filter Upper Neck Data
Call ChnCFCFiltCalc("","[1]/S1NECKUP00__FOX_","[2]/S1NECKUP00__FOXA","CFC_1000",0,"EndPoints",10) '... XW,Y,E,CFCFiltType,Fir100RemoveBias,CFCPreEventType,CFCFreeValue 
Call ChnCFCFiltCalc("","[1]/S1NECKUP00__FOX_","[2]/S1NECKUP00__FOXB","CFC_600",0,"EndPoints",10) '... XW,Y,E,CFCFiltType,Fir100RemoveBias,CFCPreEventType,CFCFreeValue 
Call ChnCFCFiltCalc("","[1]/S1NECKUP00__MOY_","[2]/S1NECKUP00__MOYB","CFC_600",0,"EndPoints",10) '... XW,Y,E,CFCFiltType,Fir100RemoveBias,CFCPreEventType,CFCFreeValue 
Call ChnCFCFiltCalc("","[1]/S1NECKUP00__FOZ_","[2]/S1NECKUP00__FOZA","CFC_1000",0,"EndPoints",10) '... XW,Y,E,CFCFiltType,Fir100RemoveBias,CFCPreEventType,CFCFreeValue
Call ChnCFCFiltCalc("","[1]/S1NECKUP00__FOZ_","[2]/S1NECKUP00__FOZB","CFC_600",0,"EndPoints",10) '... XW,Y,E,CFCFiltType,Fir100RemoveBias,CFCPreEventType,CFCFreeValue

'Filter Lower Neck Data
Call ChnCFCFiltCalc("","[1]/S1NECKLO00__FOX_","[2]/S1NECKLO00__FOXA","CFC_1000",0,"EndPoints",10) '... XW,Y,E,CFCFiltType,Fir100RemoveBias,CFCPreEventType,CFCFreeValue 
Call ChnCFCFiltCalc("","[1]/S1NECKLO00__MOY_","[2]/S1NECKLO00__MOYB","CFC_600",0,"EndPoints",10) '... XW,Y,E,CFCFiltType,Fir100RemoveBias,CFCPreEventType,CFCFreeValue 
Call ChnCFCFiltCalc("","[1]/S1NECKLO00__FOZ_","[2]/S1NECKLO00__FOZA","CFC_1000",0,"EndPoints",10) '... XW,Y,E,CFCFiltType,Fir100RemoveBias,CFCPreEventType,CFCFreeValue 

'Filter Spine Data
Call ChnCFCFiltCalc("","[1]/S1THSP01LE__ACX_","[2]/S1THSP01LE__ACXD","CFC_60",0,"EndPoints",10) '... XW,Y,E,CFCFiltType,Fir100RemoveBias,CFCPreEventType,CFCFreeValue 
Call ChnCFCFiltCalc("","[1]/S1THSP01LE__ACZ_","[2]/S1THSP01LE__ACZD","CFC_60",0,"EndPoints",10) '... XW,Y,E,CFCFiltType,Fir100RemoveBias,CFCPreEventType,CFCFreeValue 
Call ChnCFCFiltCalc("","[1]/S1THSP01RI__ACX_","[2]/S1THSP01RI__ACXD","CFC_60",0,"EndPoints",10) '... XW,Y,E,CFCFiltType,Fir100RemoveBias,CFCPreEventType,CFCFreeValue 
Call ChnCFCFiltCalc("","[1]/S1LUSP0100__ACX_","[2]/S1LUSP0100__ACXD","CFC_60",0,"EndPoints",10) '... XW,Y,E,CFCFiltType,Fir100RemoveBias,CFCPreEventType,CFCFreeValue 
Call ChnCFCFiltCalc("","[1]/S1LUSP0100__ACZ_","[2]/S1LUSP0100__ACZD","CFC_60",0,"EndPoints",10) '... XW,Y,E,CFCFiltType,Fir100RemoveBias,CFCPreEventType,CFCFreeValue 

'Filter Pelvis Data
Call ChnCFCFiltCalc("","[1]/S1PELV0000__ACX_","[2]/S1PELV0000__ACXD","CFC_60",0,"EndPoints",10) '... XW,Y,E,CFCFiltType,Fir100RemoveBias,CFCPreEventType,CFCFreeValue 
Call ChnCFCFiltCalc("","[1]/S1PELV0000__ACY_","[2]/S1PELV0000__ACYD","CFC_60",0,"EndPoints",10) '... XW,Y,E,CFCFiltType,Fir100RemoveBias,CFCPreEventType,CFCFreeValue 
Call ChnCFCFiltCalc("","[1]/S1PELV0000__ACZ_","[2]/S1PELV0000__ACZD","CFC_60",0,"EndPoints",10) '... XW,Y,E,CFCFiltType,Fir100RemoveBias,CFCPreEventType,CFCFreeValue

'Filter Vehicle Accelerations using SAE60
Call ChnCFCFiltCalc("","[1]/S0VEHC0100__ACX_","[2]/S0VEHC0100__ACXD","CFC_60",0,"EndPoints",10) '... XW,Y,E,CFCFiltType,Fir100RemoveBias,CFCPreEventType,CFCFreeValue 
Call ChnCFCFiltCalc("","[1]/S0VEHC0200__ACX_","[2]/S0VEHC0200__ACXD","CFC_60",0,"EndPoints",10) '... XW,Y,E,CFCFiltType,Fir100RemoveBias,CFCPreEventType,CFCFreeValue 
Call ChnCFCFiltCalc("","[1]/S0VEHC0300__ACX_","[2]/S0VEHC0300__ACXD","CFC_60",0,"EndPoints",10) '... XW,Y,E,CFCFiltType,Fir100RemoveBias,CFCPreEventType,CFCFreeValue 
Call ChnCFCFiltCalc("","[1]/S0VEHC0300__ACY_","[2]/S0VEHC0300__ACYD","CFC_60",0,"EndPoints",10) '... XW,Y,E,CFCFiltType,Fir100RemoveBias,CFCPreEventType,CFCFreeValue 
Call ChnCFCFiltCalc("","[1]/S0VEHC0300__ACZ_","[2]/S0VEHC0300__ACZD","CFC_60",0,"EndPoints",10) '... XW,Y,E,CFCFiltType,Fir100RemoveBias,CFCPreEventType,CFCFreeValue 
Call ChnCFCFiltCalc("","[1]/S1VEHC0400__ACX_","[2]/S1VEHC0400__ACXD","CFC_60",0,"EndPoints",10) '... XW,Y,E,CFCFiltType,Fir100RemoveBias,CFCPreEventType,CFCFreeValue 

'Filter Vehicle x accelerations for delta V calculations using SAE SAE180
Call ChnCFCFiltCalc("","[1]/S0VEHC0100__ACX_","[2]/S0VEHC0100__ACXC","CFC_180",0,"EndPoints",10) '... XW,Y,E,CFCFiltType,Fir100RemoveBias,CFCPreEventType,CFCFreeValue 
Call ChnCFCFiltCalc("","[1]/S0VEHC0200__ACX_","[2]/S0VEHC0200__ACXC","CFC_180",0,"EndPoints",10) '... XW,Y,E,CFCFiltType,Fir100RemoveBias,CFCPreEventType,CFCFreeValue 
Call ChnCFCFiltCalc("","[1]/S0VEHC0300__ACX_","[2]/S0VEHC0300__ACXC","CFC_180",0,"EndPoints",10) '... XW,Y,E,CFCFiltType,Fir100RemoveBias,CFCPreEventType,CFCFreeValue 
Call ChnCFCFiltCalc("","[1]/S1VEHC0400__ACX_","[2]/S1VEHC0400__ACXC","CFC_180",0,"EndPoints",10) '... XW,Y,E,CFCFiltType,Fir100RemoveBias,CFCPreEventType,CFCFreeValue 

'Convert the vehicle x accelerations to SI units for the delta V calculations
Call ChnUnitConvert("[2]/S0VEHC0100__ACXC","[2]/S0VEHC0100__ACXCSI","m/s^2") '... Y,E,TargetUnit 
Call ChnUnitConvert("[2]/S0VEHC0200__ACXC","[2]/S0VEHC0200__ACXCSI","m/s^2") '... Y,E,TargetUnit 
Call ChnUnitConvert("[2]/S0VEHC0300__ACXC","[2]/S0VEHC0300__ACXCSI","m/s^2") '... Y,E,TargetUnit 
Call ChnUnitConvert("[2]/S1VEHC0400__ACXC","[2]/S1VEHC0400__ACXCSI","m/s^2") '... Y,E,TargetUnit