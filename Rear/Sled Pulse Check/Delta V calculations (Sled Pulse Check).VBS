'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 05/07/2010 16:05:04
'-- Author:  
'-- Comment: Delta V calculations
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.

'Call script to determine the zero crossing points and save them as channel properties (zerocross1, zerocross2)
Scriptinclude(CurrentScriptPath & "Zero cross function (Sled Pulse Check).vbs")

Dim ACXC1_t1  'First zero cross point for S0VEHC0100__ACXC
Dim ACXC1_t2  'Second zero cross point for S0VEHC0100__ACXC
Dim ACXC2_t1  'First zero cross point for S0VEHC0200__ACXC
Dim ACXC2_t2  'Second zero cross point for S0VEHC0200__ACXC
Dim ACXC3_t1  'First zero cross point for S0VEHC0300__ACXC
Dim ACXC3_t2  'Second zero cross point for S0VEHC0300__ACXC


'Calculate the time values to use in the Channel Time Area Copy Function
ACXC1_t1 = (Data.Root.ChannelGroups(2).Channels("S0VEHC0100__ACXC").Properties("zerocross1").Value)
ACXC1_t2 = (Data.Root.ChannelGroups(2).Channels("S0VEHC0100__ACXC").Properties("zerocross2").Value)
ACXC2_t1 = (Data.Root.ChannelGroups(2).Channels("S0VEHC0200__ACXC").Properties("zerocross1").Value)
ACXC2_t2 = (Data.Root.ChannelGroups(2).Channels("S0VEHC0200__ACXC").Properties("zerocross2").Value)
ACXC3_t1 = (Data.Root.ChannelGroups(2).Channels("S0VEHC0300__ACXC").Properties("zerocross1").Value)
ACXC3_t2 = (Data.Root.ChannelGroups(2).Channels("S0VEHC0300__ACXC").Properties("zerocross2").Value)

'Run Channel Time Area Copy to extract deltaV
Call ChnTimeAreaCopy("","[2]/S0VEHC0100__ACXCSI","","/DeltaV_1",ACXC1_t1,ACXC1_t2) '... XW,Y,E,E,TimeAreaCopyBeg,TimeAreaCopyEnd 
Call ChnTimeAreaCopy("","[2]/S0VEHC0200__ACXCSI","","/DeltaV_2",ACXC2_t1,ACXC2_t2) '... XW,Y,E,E,TimeAreaCopyBeg,TimeAreaCopyEnd 
Call ChnTimeAreaCopy("","[2]/S0VEHC0300__ACXCSI","","/DeltaV_3",ACXC3_t1,ACXC3_t2) '... XW,Y,E,E,TimeAreaCopyBeg,TimeAreaCopyEnd 
 
'Calculate the DeltaV for each sled acceleration and save 
Call ChnIntegrate("[1]/Time axis","[3]/DeltaV_1","/DeltaV_1") '... XW,Y,E 
Call ChnIntegrate("[1]/Time axis","[3]/DeltaV_2","/DeltaV_2") '... XW,Y,E 
Call ChnIntegrate("[1]/Time axis","[3]/DeltaV_3","/DeltaV_3") '... XW,Y,E 
Call ChnToWfChn("[1]/Time axis","'[3]/DeltaV_1' - '[3]/DeltaV_3'",0) '... X,ChnNoStr,XChnDelete 

'Shift data so channels start at correct time value
Call ChnWFPropSet ("'[3]/DeltaV_1'","Time", "[1]/Time axis", ACXC1_t1, .0001)
Call ChnWFPropSet ("'[3]/DeltaV_2'","Time", "[1]/Time axis", ACXC2_t1, .0001)
Call ChnWFPropSet ("'[3]/DeltaV_3'","Time", "[1]/Time axis", ACXC3_t1, .0001)

'Convert DeltaV channels from m/s to km/h
Call ChnUnitConvert("[3]/DeltaV_1","/DeltaV_1_kph","km/h") '... Y,E,TargetUnit 
Call ChnUnitConvert("[3]/DeltaV_2","/DeltaV_2_kph","km/h") '... Y,E,TargetUnit 
Call ChnUnitConvert("[3]/DeltaV_3","/DeltaV_3_kph","km/h") '... Y,E,TargetUnit  









