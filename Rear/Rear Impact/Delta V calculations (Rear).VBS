'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 05/07/2010 16:05:04
'-- Author: 
'-- Comment: Delta V calculations
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.

'Call script to determine the zero crossing points and save them as channel properties (zerocross1, zerocross2)
Scriptinclude(CurrentScriptPath & "Zero cross function (Rear).vbs")

Dim ACXC1_t1  'First zero cross point for S0VEHC0100__ACXC
Dim ACXC1_t2  'Second zero cross point for S0VEHC0100__ACXC
Dim ACXC2_t1  'First zero cross point for S0VEHC0200__ACXC
Dim ACXC2_t2  'Second zero cross point for S0VEHC0200__ACXC
Dim ACXC3_t1  'First zero cross point for S0VEHC0300__ACXC
Dim ACXC3_t2  'Second zero cross point for S0VEHC0300__ACXC
Dim ACXC4_t1  'First zero cross point for S1VEHC0400__ACXC
Dim ACXC4_t2  'Second zero cross point for S1VEHC0400__ACXC


'Calculate the time values to use in the Channel Time Area Copy Function
ACXC1_t1 = (Data.Root.ChannelGroups(2).Channels("S0VEHC0100__ACXC").Properties("zerocross1").Value)
ACXC1_t2 = (Data.Root.ChannelGroups(2).Channels("S0VEHC0100__ACXC").Properties("zerocross2").Value)
ACXC2_t1 = (Data.Root.ChannelGroups(2).Channels("S0VEHC0200__ACXC").Properties("zerocross1").Value)
ACXC2_t2 = (Data.Root.ChannelGroups(2).Channels("S0VEHC0200__ACXC").Properties("zerocross2").Value)
ACXC3_t1 = (Data.Root.ChannelGroups(2).Channels("S0VEHC0300__ACXC").Properties("zerocross1").Value)
ACXC3_t2 = (Data.Root.ChannelGroups(2).Channels("S0VEHC0300__ACXC").Properties("zerocross2").Value)
ACXC4_t1 = (Data.Root.ChannelGroups(2).Channels("S1VEHC0400__ACXC").Properties("zerocross1").Value)
ACXC4_t2 = (Data.Root.ChannelGroups(2).Channels("S1VEHC0400__ACXC").Properties("zerocross2").Value)

'Run Channel Time Area Copy to extract deltaV
Call Data.Root.ChannelGroups.Add("DeltaV", 4).Activate()
Call ChnTimeAreaCopy("","[2]/S0VEHC0100__ACXCSI","","/DeltaV_1",ACXC1_t1,ACXC1_t2) '... XW,Y,E,E,TimeAreaCopyBeg,TimeAreaCopyEnd 
Call ChnTimeAreaCopy("","[2]/S0VEHC0200__ACXCSI","","/DeltaV_2",ACXC2_t1,ACXC2_t2) '... XW,Y,E,E,TimeAreaCopyBeg,TimeAreaCopyEnd 
Call ChnTimeAreaCopy("","[2]/S0VEHC0300__ACXCSI","","/DeltaV_3",ACXC3_t1,ACXC3_t2) '... XW,Y,E,E,TimeAreaCopyBeg,TimeAreaCopyEnd 
Call ChnTimeAreaCopy("","[2]/S1VEHC0400__ACXCSI","","/DeltaV_4",ACXC4_t1,ACXC4_t2) '... XW,Y,E,E,TimeAreaCopyBeg,TimeAreaCopyEnd 

'Calculate the DeltaV for each sled acceleration and save 
Call ChnIntegrate("[1]/Time axis","[4]/DeltaV_1","/DeltaV_1") '... XW,Y,E 
Call ChnIntegrate("[1]/Time axis","[4]/DeltaV_2","/DeltaV_2") '... XW,Y,E 
Call ChnIntegrate("[1]/Time axis","[4]/DeltaV_3","/DeltaV_3") '... XW,Y,E 
Call ChnIntegrate("[1]/Time axis","[4]/DeltaV_4","/DeltaV_4") '... XW,Y,E 
Call ChnToWfChn("[1]/Time axis","'[4]/DeltaV_1' - '[4]/DeltaV_4'",0) '... X,ChnNoStr,XChnDelete 

'Shift data so channels start at correct time value
Call ChnWFPropSet ("'[4]/DeltaV_1'","Time", "[1]/Time axis", ACXC1_t1, .0001)
Call ChnWFPropSet ("'[4]/DeltaV_2'","Time", "[1]/Time axis", ACXC2_t1, .0001)
Call ChnWFPropSet ("'[4]/DeltaV_3'","Time", "[1]/Time axis", ACXC3_t1, .0001)
Call ChnWFPropSet ("'[4]/DeltaV_4'","Time", "[1]/Time axis", ACXC4_t1, .0001)

'Convert DeltaV channels from m/s to km/h
Call ChnUnitConvert("[4]/DeltaV_1","/DeltaV_1_kph","km/h") '... Y,E,TargetUnit 
Call ChnUnitConvert("[4]/DeltaV_2","/DeltaV_2_kph","km/h") '... Y,E,TargetUnit 
Call ChnUnitConvert("[4]/DeltaV_3","/DeltaV_3_kph","km/h") '... Y,E,TargetUnit 
Call ChnUnitConvert("[4]/DeltaV_4","/DeltaV_4_kph","km/h") '... Y,E,TargetUnit 


'------------------------------------------------------------------------------------------------------------------

'Determine when the deltaV reaches 4 m/s
Dim ch_name


'Calculate deltaV crossing for DeltaV_1
ch_name = "DeltaV_1"  
Call Calculate_DeltaV_Crossing

'Calculate deltaV crossing for DeltaV_2
ch_name = "DeltaV_2"
Call Calculate_DeltaV_Crossing

'Calculate deltaV crossing for DeltaV_3
ch_name = "DeltaV_3"
Call Calculate_DeltaV_Crossing

'Calculate deltaV crossing for DeltaV_4
ch_name = "DeltaV_4"
Call Calculate_DeltaV_Crossing



'Rearrange the channels and add time channels
Call Data.Move(Data.Root.ChannelGroups(4).Channels("DeltaV_1_kph"),Data.Root.ChannelGroups(4).Channels,2)
Call Data.Move(Data.Root.ChannelGroups(4).Channels("DeltaV_2_kph"),Data.Root.ChannelGroups(4).Channels,4)
Call Data.Move(Data.Root.ChannelGroups(4).Channels("DeltaV_3_kph"),Data.Root.ChannelGroups(4).Channels,6)
Call ChnFromWfXGen("[4]/DeltaV_2","[4]/Time") '... Y,E 
Call ChnFromWfXGen("[4]/DeltaV_3","[4]/Time1") '... Y,E 
Call ChnFromWfXGen("[4]/DeltaV_4","[4]/Time2") '... Y,E 
Call Data.Move(Data.Root.ChannelGroups(4).Channels("Time"),Data.Root.ChannelGroups(4).Channels,3)
Call Data.Move(Data.Root.ChannelGroups(4).Channels("Time1"),Data.Root.ChannelGroups(4).Channels,6)
Call Data.Move(Data.Root.ChannelGroups(4).Channels("Time2"),Data.Root.ChannelGroups(4).Channels,9)
Data.Root.ChannelGroups(4).Channels("Time").Name = "Time_DeltaV2"
Data.Root.ChannelGroups(4).Channels("Time1").Name = "Time_DeltaV3"
Data.Root.ChannelGroups(4).Channels("Time2").Name = "Time_DeltaV4"

 

Sub Calculate_DeltaV_Crossing
Dim deltaV_value1
Dim deltaV_value2
Dim deltaV_cross_point
Dim deltaV_cross_time
'Find the first point after DeltaV crosses 4 m/s 
For i = 1 to (Data.Root.ChannelGroups(4).Channels(ch_name).Properties("length").Value)
  deltaV_value1 = Data.Root.ChannelGroups(4).Channels(ch_name).Values(i)
  'deltaV_value2 = Data.Root.ChannelGroups(4).Channels(ch_name).Values(i+1)

  IF deltaV_value1 > 4 THEN
    deltaV_cross_point = i - 1  'choose the point after crossing 4
    deltaV_cross_time = (deltaV_cross_point/10000) + Data.Root.ChannelGroups(4).Channels(ch_name).Properties("wf_start_offset").Value
    Call Data.Root.ChannelGroups(4).Channels(ch_name).Properties.Add("Max_Time", deltaV_cross_time, DataTypeFloat64)
    Exit For
  END IF

Next

'Find the time just before the deltaV reaches 4 m/s by splitting the deltaV into two channels at an amplitdue of 4 m/s
'Call ChnSplitAtValue(ch_name,4,1,"/SplittedUpperAmplitude","/SplittedLowerAmplitude") '... Y,SplitValue,FloorIsNovalue,E,E 
'Call ChnTMinTMaxCalc("","[4]/SplittedLowerAmplitude") '... XW,Y 

'Save the time as a channel property of the deltaV
'Call Data.Root.ChannelGroups(4).Channels(ch_name).Properties.Add("Max_Time", MaxTime, DataTypeFloat64)
'Dim test
'Test = Data.Root.ChannelGroups(4).Channels(ch_name).Properties("Max_Time").Value
'msgbox(Test)

'Delete the split channels
'Call Data.Root.ChannelGroups(4).Channels.Remove("SplittedUpperAmplitude")
'Call Data.Root.ChannelGroups(4).Channels.Remove("SplittedLowerAmplitude")

End Sub





