'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 05/25/2010 13:42:52
'-- Author: 
'-- Comment: 
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.

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

'Determine peak acceleration, deltaV, acceleration at t0, and endtime from sled acceleration #3
Peak_Acceleration = Data.Root.ChannelGroups(2).Channels("S0VEHC0300__ACXD").Properties("maximum").Value
Peak_DeltaV = Data.Root.ChannelGroups(3).Channels("DeltaV_3_kph").Properties("maximum").Value
temp_value = (Data.Root.ChannelGroups(2).Channels("S0VEHC0300__ACXD").Properties("wf_start_offset").Value)*-10000+1
t0_acceleration = Data.Root.ChannelGroups(2).Channels("S0VEHC0300__ACXD").Values(temp_value)
End_Time = (Data.Root.ChannelGroups(2).Channels("S0VEHC0300__ACXD").Properties("zerocross2").Value)*1000

'Create properties for End_Time and t0_acceleration
Call Data.Root.ChannelGroups(2).Channels("S0VEHC0300__ACXD").Properties.Add("t0_acceleration", t0_acceleration, DataTypeFloat64)
Call Data.Root.ChannelGroups(2).Channels("S0VEHC0300__ACXD").Properties.Add("End_Time", End_Time, DataTypeFloat64)

'Determine if the sled characteristics are within specifications
If 9.5 <= Peak_Acceleration and Peak_Acceleration <= 10.5 then
  Peak_Acceleration_Criteria = "PASS"
  Else 
  Peak_Acceleration_Criteria = "FAIL"
End IF


If 14.8 <= Peak_DeltaV and Peak_DeltaV <= 16.2 then
  Peak_DeltaV_Criteria = "PASS"
  Else 
  Peak_DeltaV_Criteria = "FAIL"
End IF

If -.25 <= t0_Acceleration and t0_Acceleration <= 0.5 then
  t0_Acceleration_Criteria = "PASS"
  Else 
  t0_Acceleration_Criteria = "FAIL"
End IF

If 88 <= End_Time and End_Time <= 94 then
  End_Time_Criteria = "PASS"
  Else 
  End_Time_Criteria = "FAIL"
End IF

'Add Pulse Criteria to TDMS properties
Call Data.Root.Properties.Add("Peak_Acceleration_g_", Peak_Acceleration, datatypefloat64)
Call Data.Root.Properties.Add("Peak_Acceleration_Criteria", Peak_Acceleration_Criteria, Datatypestring)
Call Data.Root.Properties.Add("DeltaV_kph_", Peak_DeltaV, DataTypefloat64)
Call Data.Root.Properties.Add("DeltaV_Criteria", Peak_DeltaV_Criteria, DataTypestring)
Call Data.Root.Properties.Add("t0_Acceleration_g_", t0_Acceleration, DataTypefloat64)
Call Data.Root.Properties.Add("t0_Acceleration_Criteria", t0_Acceleration_Criteria, DataTypestring)
Call Data.Root.Properties.Add("End_Time_ms_", End_Time, DataTypefloat64)
Call Data.Root.Properties.Add("End_Time_Criteria", End_Time_Criteria, DataTypestring)



