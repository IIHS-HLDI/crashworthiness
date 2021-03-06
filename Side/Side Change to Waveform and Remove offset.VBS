'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 04/07/2010 13:41:50
'-- Author: 
'-- Comment: Convert to waveforms and group creation
'-------------------------------------------------------------------------------

dim i
Dim current_name

Call Data.Root.ChannelGroups(1).Channels("Time axis").Properties.Add("Test ID", T1, DataTypeString)
Call Data.Root.ChannelGroups(1).Channels("Time axis").Properties.Add("Test Type", T4, DataTypeString)  

'Change numeric channels to waveform
For Each Chn In Data.Root.ChannelGroups(1).Channels
  If Chn.Name <> "Time axis" Then
    Call ChnToWfChn("[1]/Time axis", Chn, 0, "WfXRelative") 'Changes channel to waveform
    Call ChnOffset(Chn,Chn,100,"mean value offset") 'Removes initial offset by the average of the first 100 points
    Call Chn.Properties.Add("Test ID", T1, DataTypeString)
    Call Chn.Properties.Add("Test Type", T4, DataTypeString)
    If Chn.Properties("unit_string").Value = "G" Then
        Chn.Properties("unit_string").Value = "gn"
    End If
    If Chn.Properties("unit_string").Value = "Deg/Sec" Then
        Chn.Properties("unit_string").Value = "°/s"
    End If
    Chn.Properties("wf_increment").Value = 0.0001 'Forces the x incriment to be .0001 s 
  End If
Next



