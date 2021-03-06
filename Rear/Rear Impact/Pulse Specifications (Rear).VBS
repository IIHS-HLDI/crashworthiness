'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 05/25/2010 13:42:52
'-- Author: 
'-- Comment: Load IIWPG Ideal Pulse and compare the test pulse
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.


'Load the IIWPG Ideal Pulse
'-------------------------------------------------------------------------------
Call DataFileLoad(CurrentScriptPath & "Target Pulse.tdms","TDMS","Load") '... DataFilename,FileImportFilter,ImportAction 


'Display the Sled Acceleration Pulse Specifications
Dim Peak_Acceleration
Dim Peak_DeltaV
Dim t0_Acceleration
Dim End_Time
Dim Peak_Acceleration_Criteria
Dim Peak_DeltaV_Criteria
Dim t0_Acceleration_Criteria
Dim End_Time_Criteria
Dim temp_value
Dim Average_Acceleration

'Determine peak acceleration, deltaV, acceleration at t0, and endtime from sled acceleration #3
Peak_Acceleration = Data.Root.ChannelGroups(2).Channels("S0VEHC0300__ACXD").Properties("maximum").Value
Peak_DeltaV = Data.Root.ChannelGroups(4).Channels("DeltaV_3_kph").Properties("maximum").Value
temp_value = (Data.Root.ChannelGroups(2).Channels("S0VEHC0300__ACXD").Properties("wf_start_offset").Value)*-10000+1
t0_acceleration = Data.Root.ChannelGroups(2).Channels("S0VEHC0300__ACXD").Values(temp_value)
End_Time = (Data.Root.ChannelGroups(2).Channels("S0VEHC0300__ACXD").Properties("zerocross2").Value)*1000

'Create properties for End_Time and t0_acceleration
Call Data.Root.ChannelGroups(2).Channels("S0VEHC0300__ACXD").Properties.Add("t0_acceleration", t0_acceleration, DataTypeFloat64)
Call Data.Root.ChannelGroups(2).Channels("S0VEHC0300__ACXD").Properties.Add("End_Time", End_Time, DataTypeFloat64)

'Determine if the sled characteristics are within specifications
If Peak_Acceleration >= 9.5 and Peak_Acceleration <= 10.5 then
  Peak_Acceleration_Criteria = "PASS"
  Else 
  Peak_Acceleration_Criteria = "FAIL"
End IF

If Peak_DeltaV >= 14.8 and Peak_DeltaV <= 16.2 then
  Peak_DeltaV_Criteria = "PASS"
  Else 
  Peak_DeltaV_Criteria = "FAIL"
End IF

If t0_Acceleration >= -.25 and t0_Acceleration <= 0.5 then
  t0_Acceleration_Criteria = "PASS"
  Else 
  t0_Acceleration_Criteria = "FAIL"
End IF

If End_Time >= 88 and End_Time <= 94 then
  End_Time_Criteria = "PASS"
  Else 
  End_Time_Criteria = "FAIL"
End IF

'Add Pulse Criteria to TDMS properties
Call Data.Root.Properties.Add("Peak Acceleration(g)", Peak_Acceleration, datatypefloat64)
Call Data.Root.Properties.Add("Peak Acceleration Criteria", Peak_Acceleration_Criteria, Datatypestring)
Call Data.Root.Properties.Add("DeltaV(kph)", Peak_DeltaV, DataTypefloat64)
Call Data.Root.Properties.Add("DeltaV Criteria", Peak_DeltaV_Criteria, DataTypestring)
Call Data.Root.Properties.Add("t0 Acceleration(g)", t0_Acceleration, DataTypefloat64)
Call Data.Root.Properties.Add("t0_Acceleration_Criteria", t0_Acceleration_Criteria, DataTypestring)
Call Data.Root.Properties.Add("End Time(ms)", End_Time, DataTypefloat64)
Call Data.Root.Properties.Add("End Time Criteria", End_Time_Criteria, DataTypestring)

'Delete the IIWPG Time channel
Call Data.Root.ChannelGroups(5).Channels.Remove("IIWPG Time (sec)")

'Find the average sled acceleration and save as a custom property
Call ChnTimeAreaCopy("","[2]/S0VEHC0300__ACXD","","[2]/S0VEHC0300__ACXD_trim",0,End_Time/1000) '... XW,Y,E,E,TimeAreaCopyBeg,TimeAreaCopyEnd

StatSel(1)       = "No"
StatSel(2)       = "No"
StatSel(3)       = "No"
StatSel(4)       = "No"
StatSel(5)       = "No"
StatSel(6)       = "Yes"
StatSel(7)       = "No"
StatSel(8)       = "No"
StatSel(9)       = "No"
StatSel(10)      = "No"
StatSel(11)      = "No"
StatSel(12)      = "No"
StatSel(13)      = "No"
StatSel(14)      = "No"
StatSel(15)      = "No"
StatSel(16)      = "No"
StatSel(17)      = "No"
StatSel(18)      = "No"
StatSel(19)      = "No"
StatSel(20)      = "No"
StatSel(21)      = "No"
StatSel(22)      = "No"
StatSel(23)      = "No"
StatClipCopy     = 0
StatClipValue    = 0
StatFormat       = ""
StatResChn       = 0
StatResChnNames  = 0
StatResChnNameFormat= "NameName"
Call StatBlockCalc("Channel","1-","[2]/S0VEHC0300__ACXD_trim") '... StatDirec,RowNoStr,ChnNoStr 
Average_Acceleration = Data.Root.ChannelGroups(2).Channels("S0VEHC0300__ACXD_trim").Properties("ResultStatArithMean").Value
Call Data.Root.ChannelGroups(2).Channels("S0VEHC0300__ACXD").Properties.Add("Average_Acceleration", Average_Acceleration, DataTypeFloat64)
Call Data.Root.ChannelGroups(2).Channels.Remove("S0VEHC0300__ACXD_trim")