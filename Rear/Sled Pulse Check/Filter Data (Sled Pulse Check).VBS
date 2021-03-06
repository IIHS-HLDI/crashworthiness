'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 04/27/2010 09:36:16
'-- Author: 
'-- Comment: Pulse Check Filters
'-------------------------------------------------------------------------------


'Change units from G to gn
Dim Channel
For Each Channel In Data.Root.ChannelGroups(1).Channels
  If Data.Root.ChannelGroups(1).Channels(Channel.Name).Properties("unit_string").Value = "G" Then
    Data.Root.ChannelGroups(1).Channels(Channel.Name).Properties("unit_string").Value = "gn"
  End If
Next

'Filter Vehicle Accelerations using SAE60
Call ChnCFCFiltCalc("","[1]/S0VEHC0100__ACX_","[2]/S0VEHC0100__ACXD","CFC_60",0,"EndPoints",10) '... XW,Y,E,CFCFiltType,Fir100RemoveBias,CFCPreEventType,CFCFreeValue 
Call ChnCFCFiltCalc("","[1]/S0VEHC0200__ACX_","[2]/S0VEHC0200__ACXD","CFC_60",0,"EndPoints",10) '... XW,Y,E,CFCFiltType,Fir100RemoveBias,CFCPreEventType,CFCFreeValue 
Call ChnCFCFiltCalc("","[1]/S0VEHC0300__ACX_","[2]/S0VEHC0300__ACXD","CFC_60",0,"EndPoints",10) '... XW,Y,E,CFCFiltType,Fir100RemoveBias,CFCPreEventType,CFCFreeValue 
Call ChnCFCFiltCalc("","[1]/S0VEHC0300__ACY_","[2]/S0VEHC0300__ACYD","CFC_60",0,"EndPoints",10) '... XW,Y,E,CFCFiltType,Fir100RemoveBias,CFCPreEventType,CFCFreeValue 
Call ChnCFCFiltCalc("","[1]/S0VEHC0300__ACZ_","[2]/S0VEHC0300__ACZD","CFC_60",0,"EndPoints",10) '... XW,Y,E,CFCFiltType,Fir100RemoveBias,CFCPreEventType,CFCFreeValue 
'Call ChnCFCFiltCalc("","[1]/S1VEHC0400__ACX_","[2]/S1VEHC0400__ACXD","CFC_60",0,"EndPoints",10) '... XW,Y,E,CFCFiltType,Fir100RemoveBias,CFCPreEventType,CFCFreeValue 


'Filter extra channels for Sean and Add resultant

'Call ChnCFCFiltCalc("","[1]/S9SEBECY0000ACX_","[2]/OBD__ACXD","CFC_60",0,"EndPoints",10) '... XW,Y,E,CFCFiltType,Fir100RemoveBias,CFCPreEventType,CFCFreeValue 
'Call ChnCFCFiltCalc("","[1]/S9SEBECY0000ACZ_","[2]/OBD__ACZD","CFC_60",0,"EndPoints",10) '... XW,Y,E,CFCFiltType,Fir100RemoveBias,CFCPreEventType,CFCFreeValue 
'Call ChnCFCFiltCalc("","[1]/S9SEBECY0000ACY_","[2]/OBD__ACYD","CFC_60",0,"EndPoints",10) '... XW,Y,E,CFCFiltType,Fir100RemoveBias,CFCPreEventType,CFCFreeValue 
'Call ChnXYZAbsValue("[2]/OBD__ACXD","[2]/OBD__ACZD","[2]/OBD__ACYD","/OBD_ResultantXYZ")



'Filter Vehicle x accelerations for delta V calculations using SAE SAE180
Call ChnCFCFiltCalc("","[1]/S0VEHC0100__ACX_","[2]/S0VEHC0100__ACXC","CFC_180",0,"EndPoints",10) '... XW,Y,E,CFCFiltType,Fir100RemoveBias,CFCPreEventType,CFCFreeValue 
Call ChnCFCFiltCalc("","[1]/S0VEHC0200__ACX_","[2]/S0VEHC0200__ACXC","CFC_180",0,"EndPoints",10) '... XW,Y,E,CFCFiltType,Fir100RemoveBias,CFCPreEventType,CFCFreeValue 
Call ChnCFCFiltCalc("","[1]/S0VEHC0300__ACX_","[2]/S0VEHC0300__ACXC","CFC_180",0,"EndPoints",10) '... XW,Y,E,CFCFiltType,Fir100RemoveBias,CFCPreEventType,CFCFreeValue 


'Convert the vehicle x accelerations to SI units for the delta V calculations
Call ChnUnitConvert("[2]/S0VEHC0100__ACXC","[2]/S0VEHC0100__ACXCSI","m/s^2") '... Y,E,TargetUnit 
Call ChnUnitConvert("[2]/S0VEHC0200__ACXC","[2]/S0VEHC0200__ACXCSI","m/s^2") '... Y,E,TargetUnit 
Call ChnUnitConvert("[2]/S0VEHC0300__ACXC","[2]/S0VEHC0300__ACXCSI","m/s^2") '... Y,E,TargetUnit 
