'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 04/07/2010 13:41:50
'-- Author: 
'-- Comment: Convert to waveforms, remove offset, and create filtered and calculated groups
'-------------------------------------------------------------------------------

dim i

'Convert all the channels from numeric to waveform (except time axis)
For i = 2 to data.Root.ChannelGroups(1).Channels.Count
  Call ChnToWfChn("[1]/Time axis",i,0) '... X,ChnNoStr,XChnDelete 
Next

'Remove offset from the raw data
For i = 2 to data.Root.ChannelGroups(1).Channels.Count
  Call ChnOffset(i,i,100,"mean value offset") '... Y,E,ChnOffsetValue,ChnOffsetMode 
Next

'Create a Filtered data group
Call Data.Root.ChannelGroups.Add("Filtered Data", 2).Activate()

'Create a Calculated data group
Call Data.Root.ChannelGroups.Add("Calculated Channels", 3).Activate()


