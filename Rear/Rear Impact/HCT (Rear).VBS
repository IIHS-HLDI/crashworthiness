'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 05/20/2010 14:06:13
'-- Author: 
'-- Comment: Shift head contact time based on deltaV at 4 m/s
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.

Dim temp_value

R1 = (Data.Root.ChannelGroups(4).Channels("DeltaV_3").Properties("Max_Time").Value)

'Round the time to the nearest whole ms for the deltaV at 4 m/s
If (Round(R1,3) - Round(R1,4)) = 0 Then  'if the value occurs at a whole ms don't round

temp_value = R1 

Else 

temp_value = Round(R1 + 0.0005,3)  'round up to the nearest whole ms

End If    

'Create a copy of the contact switch channel and rename it HCT 
Call ChnClpCopy("[1]/S1HEADRE00__DSX_") '... ClpSource 
Call Data.Root.ChannelGroups(3).Activate()
Call ChnClpPaste(88)                    '... ClpTarget 
Data.Root.ChannelGroups(3).Channels("S1HEADRE00__DSX_").Name = "HCT"
Call Data.Root.ChannelGroups(4).Activate()

'Time shift the data
R1 = .027 - Data.Root.ChannelGroups(2).Channels("S0VEHC0300__ACXD").Properties("Max_Time_Preshift").Value
Call ChnWFPropSet ("[3]/HCT","Time", "[1]/Time axis", (-0.1+R1) + (0.070 - temp_value), .0001)

'-------------------------------------------------------------------------------------------------------------     
   
'Use the cross function to determine when head contact occurs     
Dim i 'channel point count
Dim channel_length
Dim temp_value1
Dim temp_value2
Dim zero_cross_point, zero_cross_point2
Dim zero_cross_time, zero_cross_time2


'Find the first point after HCT channel reaches -50
For i = 1 to (Data.Root.ChannelGroups(3).Channels("HCT").Properties("length").Value - 1)
  temp_value1 = Data.Root.ChannelGroups(3).Channels("HCT").Values(i)
  temp_value2 = Data.Root.ChannelGroups(3).Channels("HCT").Values(i+1)

  IF temp_value2 < -50 THEN
   zero_cross_point = i + 1 'choose the point after crossing -50
  Exit For
  END IF

Next

'Find point where HCT returns to zero
For i = zero_cross_point to (Data.Root.ChannelGroups(3).Channels("HCT").Properties("length").Value - 1)
  temp_value1 = Data.Root.ChannelGroups(3).Channels("HCT").Values(i)
  temp_value2 = Data.Root.ChannelGroups(3).Channels("HCT").Values(i+1)
 
  IF temp_value2 > 0 THEN
    zero_cross_point2 = i  'choose the point before the axis is crossed
    Exit For
  END IF

Next

'Calculate time from cross points
zero_cross_time = Data.Root.ChannelGroups(3).Channels("HCT").Properties("wf_start_offset").Value + ((zero_cross_point - 1) * .0001)
zero_cross_time2 = Data.Root.ChannelGroups(3).Channels("HCT").Properties("wf_start_offset").Value + ((zero_cross_point2 - 1) * .0001)

'Save HCT_begin and HCT_end as channel properties
Call Data.Root.ChannelGroups(3).Channels("HCT").Properties.Add("HCT_begin", zero_cross_time, DataTypeFloat64)
Call Data.Root.ChannelGroups(3).Channels("HCT").Properties.Add("HCT_end", zero_cross_time2, DataTypeFloat64)


'Calculate the Force Balance of the head
Dim offset

Call ChnLinScale("[2]/S1HEAD0000__ACXDSI","[3]/Head Force",4.5,0) '... Y,E,ChnScaleFactor,ChnScaleOffset 
Call ChnSub("[3]/Head Force","[2]/S1NECKUP00__FOXA","[3]/Force Balance") '... Y,Y1,E 
Call Data.Root.ChannelGroups(3).Channels.Remove("Head Force")
offset = Data.Root.ChannelGroups(3).Channels("HCT").Properties("wf_start_offset").Value
Data.Root.ChannelGroups(3).Channels("Force Balance").Properties("wf_start_offset").Value = offset
Data.Root.ChannelGroups(3).Channels("Force Balance").Properties("unit_string").Value = "N"


