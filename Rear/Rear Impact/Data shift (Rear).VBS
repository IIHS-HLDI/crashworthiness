'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 05/07/2010 15:52:53
'-- Author: 
'-- Comment: Shift data so peak sled acceleration occurs at 27 ms
'-------------------------------------------------------------------------------

'Get the maximum sled acceleration time
Call ChnTMinTMaxCalc("","[2]/S0VEHC0300__ACXD") '... XW,Y 
Call Data.Root.ChannelGroups(2).Channels("S0VEHC0300__ACXD").Properties.Add("Max_Time_Preshift", MaxTime, DataTypeFloat64)

'Shift the filtered data so the peak sled acceleration occurs at 27 ms
R1 = .027 - Data.Root.ChannelGroups(2).Channels("S0VEHC0300__ACXD").Properties("Max_Time_Preshift").Value
Call ChnWFPropSet ("'[2]/S1HEAD0000__ACXA' - '[2]/S1VEHC0400__ACXCSI'","Time", "[1]/Time axis", -0.1+R1, .0001)