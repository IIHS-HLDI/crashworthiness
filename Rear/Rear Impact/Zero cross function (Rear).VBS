'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 05/07/2010 11:08:39
'-- Author: 
'-- Comment: Find time when sled accelerations cross zero
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.

Dim output
Dim i 'channel point count
Dim ch_group, ch_name
Dim channel_name,channel_length, channel_offset
Dim channel_target, channel_target_point
dim target
Dim temp_value
Dim temp_value2
Dim zero_cross_point, zero_cross_point2
Dim zero_cross_time, zero_cross_time2

ch_group = 2

'Calculate zero crossing for sled acceleration 1
ch_name = "S0VEHC0100__ACXC"  
Call Calculate_Zero_Crossing

'Calculate zero crossing for sled acceleration 2
ch_name = "S0VEHC0200__ACXC"
Call Calculate_Zero_Crossing

'Calculate zero crossing for sled acceleration 3
ch_name = "S0VEHC0300__ACXC"
Call Calculate_Zero_Crossing

'Calculate zero crossing for sled acceleration 4
ch_name = "S1VEHC0400__ACXC"
Call Calculate_Zero_Crossing

'Calculate zero crossing for sled acceleration 1
ch_name = "S0VEHC0100__ACXD"  
Call Calculate_Zero_Crossing

'Calculate zero crossing for sled acceleration 2
ch_name = "S0VEHC0200__ACXD"
Call Calculate_Zero_Crossing

'Calculate zero crossing for sled acceleration 3
ch_name = "S0VEHC0300__ACXD"
Call Calculate_Zero_Crossing

'Calculate zero crossing for sled acceleration 4
ch_name = "S1VEHC0400__ACXD"
Call Calculate_Zero_Crossing



Sub Calculate_Zero_Crossing

channel_name = Data.Root.ChannelGroups(ch_group).Channels(ch_name).Name
channel_length = Data.Root.ChannelGroups(ch_group).Channels(ch_name).Properties("length").Value
channel_offset = Data.Root.ChannelGroups(ch_group).Channels(ch_name).Properties("wf_start_offset").Value


'Find point that corresponds to maximum sled acceleration
channel_target = Data.Root.ChannelGroups(ch_group).Channels(ch_name).Properties("maximum").Value
For i = 1 to (channel_length - 1)
  temp_value = Data.Root.ChannelGroups(ch_group).Channels(ch_name).Values(i)
  If temp_value = channel_target THEN
    channel_target_point = i
    Exit For
  End If
Next


'Find last zero cross point before maximum value is reached
For i = 1 to channel_target_point
  temp_value = Data.Root.ChannelGroups(ch_group).Channels(ch_name).Values(i)
  temp_value2 = Data.Root.ChannelGroups(ch_group).Channels(ch_name).Values(i+1)

  'If values cross from negative to positive
  IF temp_value < 0 THEN  
    IF temp_value2 >  0 THEN
      zero_cross_point = i + 1 'choose the point after the axis crossing
    END IF
  END IF

Next

'Find point where y axis is crossed again
For i = channel_target_point to (channel_length - channel_target_point - 1)
  temp_value = Data.Root.ChannelGroups(ch_group).Channels(ch_name).Values(i)
  temp_value2 = Data.Root.ChannelGroups(ch_group).Channels(ch_name).Values(i+1)

  'If values cross from positive to negative
  IF temp_value > 0 THEN  
    IF temp_value2 < 0 THEN
      zero_cross_point2 = i  'choose the point before the axis is crossed
      Exit For
    END IF
  END IF

Next

'Calculate Time of zero cross points
zero_cross_time = Data.Root.ChannelGroups(2).Channels(ch_name).Properties("wf_start_offset").Value + ((zero_cross_point - 1) * .0001)
zero_cross_time2 = Data.Root.ChannelGroups(2).Channels(ch_name).Properties("wf_start_offset").Value + ((zero_cross_point2 - 1) * .0001)

'Save zero cross time as channel properties
Call Data.Root.ChannelGroups(2).Channels(ch_name).Properties.Add("zerocross1", zero_cross_time, DataTypeFloat64)
Call Data.Root.ChannelGroups(2).Channels(ch_name).Properties.Add("zerocross2", zero_cross_time2, DataTypeFloat64)

End Sub
