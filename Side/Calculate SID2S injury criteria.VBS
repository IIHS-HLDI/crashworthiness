'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 04/08/2010 09:22:22
'-- Author: 
'-- Comment: Injury Calculations
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.

If Data.Root.ChannelGroups.Exists("11_Head_Neck") Then 'Driver data is present
  Call Calculate_injuries(11)
End If

If Data.Root.ChannelGroups.Exists("14_Head_Neck") Then 'Rear passenger data is present
  Call Calculate_injuries(14)
End If

'Collapse groups in data portal
Call UIAutoRefreshSet(True) 
Call Portal.Structure.Collapse(Data.Root) 
Call UIAutoRefreshSet(False) 

Function MaxValue(Channel)
  If Channel.Properties("maximum").Value >= ABS(Channel.Properties("minimum").value) Then
    MaxValue = round(Channel.Properties("maximum").Value,1)
  Else
    MaxValue = round(Channel.Properties("minimum").Value,1)
  End If
End Function

Function AbsMaxTime(Channel)
  If Channel.Properties("maximum").Value >= ABS(Channel.Properties("minimum").value) Then
    Call ChnTMinTMaxCalc("",Channel)
    AbsMaxTime = round(maxtime*1000,1)
  Else
    Call ChnTMinTMaxCalc("",Channel)
    AbsMaxTime = round(mintime*1000,1)
  End If
End Function

Function BRICcalc(Xmax, Ymax, Zmax)
  BRICcalc = SQRT((Xmax/3795.84)^2+(Ymax/3234.34)^2+(Zmax/2456.27)^2)
End Function

Sub Calculate_injuries(Occ_Pos)

  'Head injury calcuations
  Call Data.Root.ChannelGroups(Occ_Pos & "_Head_Neck").Activate  

  'Peak Head Resultant Acceleration Time & HIC Values
  Call ChnXYZAbsValue("/" & Occ_Pos & "HEADCG00__ACXA","/" & Occ_Pos & "HEADCG00__ACYA","/" & Occ_Pos & "HEADCG00__ACZA","/" & Occ_Pos & "HEADCG00__ACR") 
  Call ChnHICCalc("","/" & Occ_Pos & "HEADCG00__ACR",1,1,0,0,10,0) 
  Call ChnTMinTMaxCalc("","/" & Occ_Pos & "HEADCG00__ACR") 
  
  Call Data.Move(Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "HEADCG00__ACR"),Data.Root.ActiveChannelGroup.Channels,4)
  
  Call Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "HEADCG00__ACR").Properties.Add("Max_time",round( maxtime*1000,1), DataTypeFloat64) 
  Call Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "HEADCG00__ACR").Properties.Add("HIC36", HICRes(1,2), DataTypeFloat64)
  Call Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "HEADCG00__ACR").Properties.Add("HIC36_T1",HICRes(2,2), DataTypeFloat64)
  Call Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "HEADCG00__ACR").Properties.Add("HIC36_T2", HICRes(3,2), DataTypeFloat64)
  Call Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "HEADCG00__ACR").Properties.Add("HIC15", HICRes(1,3), DataTypeFloat64)
  Call Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "HEADCG00__ACR").Properties.Add("HIC15_T1", HICRes(2,3), DataTypeFloat64)
  Call Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "HEADCG00__ACR").Properties.Add("HIC15_T2", HICRes(3,3), DataTypeFloat64)

  '3ms clip
  Call ChnXMSCalc("","/" & Occ_Pos & "HEADCG00__ACR",3,0,"in one peak",0)  
  Call Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "HEADCG00__ACR").Properties.Add("Clip_3_ms", XmsAcceleration, DataTypeFloat64)
  Call Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "HEADCG00__ACR").Properties.Add("Clip_3_ms_T1", XmsTimeBegin, DataTypeFloat64)
  Call Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "HEADCG00__ACR").Properties.Add("Clip_3_ms_T2",XmsTimeEnd, DataTypeFloat64)
  
  'Angular displacement                   
  Call ChnIntegrate("","/" & Occ_Pos & "HEADCG0000AVXC","/" & Occ_Pos & "HEADCG0000ANX") 
  Call ChnIntegrate("","/" & Occ_Pos & "HEADCG0000AVYC","/" & Occ_Pos & "HEADCG0000ANY") 
  Call ChnIntegrate("","/" & Occ_Pos & "HEADCG0000AVZC","/" & Occ_Pos & "HEADCG0000ANZ") 

  Call ChnOffset("/" & Occ_Pos & "HEADCG0000ANX","/" & Occ_Pos & "HEADCG0000ANX",-data.Root.ActiveChannelGroup.Channels(Occ_Pos & "HEADCG0000ANX").Values(1001),"free offset")
  Call ChnOffset("/" & Occ_Pos & "HEADCG0000ANY","/" & Occ_Pos & "HEADCG0000ANY",-data.Root.ActiveChannelGroup.Channels(Occ_Pos & "HEADCG0000ANY").Values(1001),"free offset")
  Call ChnOffset("/" & Occ_Pos & "HEADCG0000ANZ","/" & Occ_Pos & "HEADCG0000ANZ",-data.Root.ActiveChannelGroup.Channels(Occ_Pos & "HEADCG0000ANZ").Values(1001),"free offset")

  'BRIC
  Call Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "HEADCG0000AVXC").Properties.Add("AbsMax", MaxValue(Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "HEADCG0000AVXC")))
  Call Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "HEADCG0000AVYC").Properties.Add("AbsMax", MaxValue(Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "HEADCG0000AVYC")))
  Call Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "HEADCG0000AVZC").Properties.Add("AbsMax", MaxValue(Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "HEADCG0000AVZC")))
  
  Call Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "HEADCG0000AVXC").Properties.Add("AbsMax_time", AbsMaxTime(Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "HEADCG0000AVXC")))
  Call Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "HEADCG0000AVYC").Properties.Add("AbsMax_time", AbsMaxTime(Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "HEADCG0000AVYC")))
  Call Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "HEADCG0000AVZC").Properties.Add("AbsMax_time", AbsMaxTime(Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "HEADCG0000AVZC")))
  
  Call Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "HEADCG00__ACR").Properties.Add("BRIC", BRICcalc(Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "HEADCG0000AVXC").Properties("AbsMax").Value, Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "HEADCG0000AVYC").Properties("AbsMax").Value, Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "HEADCG0000AVZC").Properties("AbsMax").Value))
  
  'Neck 
  'A-P Shear force
  If abs(CMin(Occ_Pos & "NECKUP00__FOXA")) > CMax(Occ_Pos & "NECKUP00__FOXA") Then
    R2 = (CMin(Occ_Pos & "NECKUP00__FOXA"))
    Call Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "NECKUP00__FOXA").Properties.Add("Max_force", R2, DataTypeFloat64)
    Call ChnTMinTMaxCalc("","/" & Occ_Pos & "NECKUP00__FOXA") 
    Call Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "NECKUP00__FOXA").Properties.Add("Max_time", round( mintime*1000,1), DataTypeFloat64)
  Else
    R2 = (CMax(Occ_Pos & "NECKUP00__FOXA"))
    Call Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "NECKUP00__FOXA").Properties.Add("Max_force", R2, DataTypeFloat64)
    Call ChnTMinTMaxCalc("","/" & Occ_Pos & "NECKUP00__FOXA") 
    Call Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "NECKUP00__FOXA").Properties.Add("Max_time",round( maxtime*1000,1), DataTypeFloat64)
  End If
    
  'L-M Shear force
  If abs(CMin(Occ_Pos & "NECKUP00__FOYA")) > CMax(Occ_Pos & "NECKUP00__FOYA") Then
    R2 = (CMin(Occ_Pos & "NECKUP00__FOYA"))
    Call Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "NECKUP00__FOYA").Properties.Add("Max_force", R2, DataTypeFloat64)
    Call ChnTMinTMaxCalc("","/" & Occ_Pos & "NECKUP00__FOYA") 
    Call Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "NECKUP00__FOYA").Properties.Add("Max_time", round( mintime*1000,1), DataTypeFloat64)
  Else
    R2 = (CMax(Occ_Pos & "NECKUP00__FOYA"))
    Call Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "NECKUP00__FOYA").Properties.Add("Max_force", R2, DataTypeFloat64)
    Call ChnTMinTMaxCalc("","/" & Occ_Pos & "NECKUP00__FOYA")
    Call Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "NECKUP00__FOYA").Properties.Add("Max_time",round( maxtime*1000,1), DataTypeFloat64)
  End If
    
  'Compression time
  Call ChnTMinTMaxCalc("","/" & Occ_Pos & "NECKUP00__FOZA") 
  Call Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "NECKUP00__FOZA").Properties.Add("Min_time", round( mintime*1000,1), DataTypeFloat64)
  Call Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "NECKUP00__FOZA").Properties.Add("Max_time",round( maxtime*1000,1), DataTypeFloat64)
  
  'X Moment
  If abs(CMin(Occ_Pos & "NECKUP00__MOXB")) > CMax(Occ_Pos & "NECKUP00__MOXB") Then
    R2 = (CMin(Occ_Pos & "NECKUP00__MOXB"))
    Call Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "NECKUP00__MOXB").Properties.Add("Max_moment", R2, DataTypeFloat64)
    Call ChnTMinTMaxCalc("","/" & Occ_Pos & "NECKUP00__MOXB")  
    Call Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "NECKUP00__MOXB").Properties.Add("Max_time", round( mintime*1000,1), DataTypeFloat64)
  Else
    R2 = (CMax(Occ_Pos & "NECKUP00__MOXB"))
    Call Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "NECKUP00__MOXB").Properties.Add("Max_moment", R2, DataTypeFloat64)
    Call ChnTMinTMaxCalc("","/" & Occ_Pos & "NECKUP00__MOXB")  
    Call Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "NECKUP00__MOXB").Properties.Add("Max_time",round( maxtime*1000,1), DataTypeFloat64)
  End If
  
  'Z Moment
  If abs(CMin(Occ_Pos & "NECKUP00__MOZB")) > CMax(Occ_Pos & "NECKUP00__MOZB") Then
    R2 = (CMin(Occ_Pos & "NECKUP00__MOZB"))
    Call Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "NECKUP00__MOZB").Properties.Add("Max_moment", R2, DataTypeFloat64)
    Call ChnTMinTMaxCalc("","/" & Occ_Pos & "NECKUP00__MOZB") 
    Call Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "NECKUP00__MOZB").Properties.Add("Max_time", round( mintime*1000,1), DataTypeFloat64)
  Else
    R2 = (CMax(Occ_Pos & "NECKUP00__MOZB"))
    Call Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "NECKUP00__MOZB").Properties.Add("Max_moment", R2, DataTypeFloat64)
    Call ChnTMinTMaxCalc("","/" & Occ_Pos & "NECKUP00__MOZB") 
    Call Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "NECKUP00__MOZB").Properties.Add("Max_time",round( maxtime*1000,1), DataTypeFloat64)
  End If

  'Shoulder
  Call Data.Root.ChannelGroups(Occ_Pos & "_Shoulder").Activate  
  Call ChnTMinTMaxCalc("","/" & Occ_Pos & "SHRILE00__DSYB")
  Call Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "SHRILE00__DSYB").Properties.Add("Min_time", round( mintime*1000,1), DataTypeFloat64)
  Call ChnTMinTMaxCalc("","/" & Occ_Pos & "SHLDLE00__FOYA")
  Call Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "SHLDLE00__FOYA").Properties.Add("Max_time",round( maxtime*1000,1), DataTypeFloat64)
  
  'Find shoulder loading time
  Dim temp_value, i, load_time, channel_offset, channel_length, channel_target_point, DX
  
  load_time = 1

  channel_length = Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "SHRILE00__DSYB").Properties("length").Value
  channel_offset = Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "SHRILE00__DSYB").Properties("wf_start_offset").Value
  DX = Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "SHRILE00__DSYB").Properties("wf_increment").Value
  
  For i = 1 to (channel_length - 1)
    temp_value = Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "SHRILE00__DSYB").Values(i)
    If temp_value <= -0.5 THEN
      channel_target_point = i
      Exit For
    End If
  Next
  
  load_time = (channel_offset + (channel_target_point - 1)*DX)
  Call Data.Root.ActiveChannelGroup.Properties.Add("Shoulder_loading_time", 1000*load_time,DataTypeFloat64)  
    
  'Ribs
  Call Data.Root.ChannelGroups(Occ_Pos & "_Torso").Activate
  Dim Rib_name, Rib_count, Rib_defl_sum
  Rib_name = array("TRRI01LE__", "TRRI02LE__", "TRRI03LE__", "ABRI01LE__", "ABRI02LE__")
  Spine_accel_name = array("SPIN0400__ACYA", "SPIN0400__ACYA", "SPIN1200__ACYA", "SPIN1200__ACYA", "SPIN1200__ACYA")
  Rib_defl_sum = 0
  
  'Deflection
  For Rib_count = 0 to 4
    Call ChnTMinTMaxCalc("","/" & Occ_Pos & Rib_name(rib_count) & "DSYB") 
    Call Data.Root.ActiveChannelGroup.Channels(Occ_Pos & Rib_name(rib_count) & "DSYB").Properties.Add("Min_time", round( mintime*1000,1), DataTypeFloat64)
    Rib_defl_sum = Rib_defl_sum + Data.Root.ActiveChannelGroup.Channels(Occ_Pos & Rib_name(rib_count) & "DSYB").Properties("minimum").Value
  Next
  Call Data.Root.ActiveChannelGroup.Properties.Add("Avg_rib_defl",Round(Rib_defl_sum/5,0),DataTypeFloat64)
  
  'Deflection rate
  For Rib_count = 0 to 4
    Call ChnCFCFiltCalc("","[1]/" & Occ_Pos & Rib_name(rib_count) & "DSY_","/TempRib","CFC_180",0,"EndPoints",10) 
    Call ChnDifferentiate("","/TempRib","","/DifferentiatedY")
    Call Data.Root.ActiveChannelGroup.Channels.Remove("TempRib")
    Data.Root.ActiveChannelGroup.Channels("DifferentiatedY").Name =  Occ_Pos & Rib_name(rib_count) & "YVF"
    Call ChnUnitConvert("/" & Occ_Pos & Rib_name(rib_count) & "YVF","/" & Occ_Pos & Rib_name(rib_count) & "YVF","m/s") 
    Call ChnLinScale("/" & Occ_Pos & Rib_name(rib_count) & "YVF","/" & Occ_Pos & Rib_name(rib_count) & "YVF",-1,0) 
    Call ChnTMinTMaxCalc("","/" & Occ_Pos & Rib_name(rib_count) & "YVF") 
    Call Data.Root.ActiveChannelGroup.Channels(Occ_Pos & Rib_name(rib_count) & "YVF").Properties.Add("Max_time",round( maxtime*1000,1), DataTypeFloat64)
    Data.Root.ActiveChannelGroup.Channels(Occ_Pos & Rib_name(rib_count) & "YVF").Properties("wf_xname").Value = "Time axis"
  Next
  
  'VC
  For Rib_count = 0 to 4
    Call ChnVCCalc("","[1]/" & Occ_Pos & Rib_name(rib_count) & "DSY_","/" & Occ_Pos & Rib_name(rib_count) & "VC",138,1,"ECE_94")  
    Call ChnUnitConvert("/" & Occ_Pos & Rib_name(rib_count) & "VC","/" & Occ_Pos & Rib_name(rib_count) & "VC","m/s") 
    Call ChnTMinTMaxCalc("","/" & Occ_Pos & Rib_name(rib_count) & "VC") '... XW,Y 
    Call Data.Root.ActiveChannelGroup.Channels(Occ_Pos & Rib_name(rib_count) & "VC").Properties.Add("Max_time",round( maxtime*1000,1), DataTypeFloat64)
  Next
  
  'Find torso loading time
  
  load_time = 1
  For Rib_count = 0 to 4
    channel_length = Data.Root.ActiveChannelGroup.Channels(Occ_Pos & Rib_name(rib_count) & "DSYB").Properties("length").Value
    channel_offset = Data.Root.ActiveChannelGroup.Channels(Occ_Pos & Rib_name(rib_count) & "DSYB").Properties("wf_start_offset").Value
    DX = Data.Root.ActiveChannelGroup.Channels(Occ_Pos & Rib_name(rib_count) & "DSYB").Properties("wf_increment").Value
    
    For i = 1 to (channel_length - 1)
      temp_value = Data.Root.ActiveChannelGroup.Channels(Occ_Pos & Rib_name(rib_count) & "DSYB").Values(i)
      If temp_value <= -0.5 Then
        channel_target_point = i
        Exit For
      End If
    Next
    
    If (channel_offset + (channel_target_point - 1)*DX) < load_time Then
      load_time = (channel_offset + (channel_target_point - 1)*DX)
    End If
  Next

  Call Data.Root.ActiveChannelGroup.Properties.Add("Torso_loading_time", 1000*load_time,DataTypeFloat64)
        
  'Calculate Accel-based Deflection Rate
  'Subtract the rib accel from the spine box accel, convert units from gn to m/s^2, then integrate
  
  Dim spine_accel_index, Spine_accel_name
  spine_accel_index = Data.Root.ChannelGroups(Occ_Pos & "_Spine").Properties("index").Value
  'These are the spine accelerometers that are paired with the Rib_name array above.
  Spine_accel_name = array("SPIN0400__ACYA", "SPIN0400__ACYA", "SPIN1200__ACYA", "SPIN1200__ACYA", "SPIN1200__ACYA") 
  
  For Rib_count = 0 to 4
    Call ChnSub(Occ_Pos & Rib_name(rib_count) & "ACYA","[" & spine_accel_index & "]/" & Occ_Pos & Spine_accel_name(rib_count),"/Subtracted") 
    Call ChnUnitConvert("/Subtracted","/Subtracted","m/s^2") 
    Call ChnIntegrate("","/Subtracted","/Subtracted") 
    Data.Root.ActiveChannelGroup.Channels("Subtracted").Name = Occ_Pos & Rib_name(rib_count) & "Accel_Vel"
  Next
        
  'Pelvis  
  Call Data.Root.ChannelGroups(Occ_Pos & "_Pelvis_Femur").Activate

  Call ChnSplitAtValue("/" & Occ_Pos & "ILACLE00__FOYB",0,0,"/SplittedUpperAmplitude","/SplittedLowerAmplitude")
  Call ChnSplitAtValue("/" & Occ_Pos & "ACTBLE00__FOYB",0,0,"/SplittedUpperAmplitude1","/SplittedLowerAmplitude1") 
  Call Data.Root.ActiveChannelGroup.Channels.Remove("SplittedLowerAmplitude")
  Call Data.Root.ActiveChannelGroup.Channels.Remove("SplittedLowerAmplitude1")
  Call ChnAdd("/SplittedUpperAmplitude","/SplittedUpperAmplitude1","/" & Occ_Pos & "COMBINED_PELV_FOY")
  Call Data.Root.ActiveChannelGroup.Channels.Remove("SplittedUpperAmplitude")
  Call Data.Root.ActiveChannelGroup.Channels.Remove("SplittedUpperAmplitude1")
  Call ChnTMinTMaxCalc("","/" & Occ_Pos & "COMBINED_PELV_FOY")
  Call Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "COMBINED_PELV_FOY").Properties.Add("Max_time",round( maxtime*1000,1), DataTypeFloat64) 
  Call ChnTMinTMaxCalc("","/" & Occ_Pos & "PELV0000__ACYA")
  Call Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "PELV0000__ACYA").Properties.Add("Max_time",round( maxtime*1000,1), DataTypeFloat64) 
  Call ChnTMinTMaxCalc("","/" & Occ_Pos & "ILACLE00__FOYB")
  Call Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "ILACLE00__FOYB").Properties.Add("Max_time",round( maxtime*1000,1), DataTypeFloat64)  
  Call ChnTMinTMaxCalc("","/" & Occ_Pos & "ACTBLE00__FOYB")
  Call Data.Root.ActiveChannelGroup.Channels(Occ_Pos & "ACTBLE00__FOYB").Properties.Add("Max_time",round( maxtime*1000,1), DataTypeFloat64)  

  Call Data.Move(Data.Root.ActiveChannelGroup.Channels( Occ_Pos & "COMBINED_PELV_FOY"),Data.Root.ActiveChannelGroup.Channels,4)

  'Find pelvis loading time
  Dim pelvis_name, pelv_count
  pelvis_name = array("ILACLE00__FOYB", "ACTBLE00__FOYB")
  load_time = 1
  
  For pelv_count = 0 to 1
    channel_length = Data.Root.ActiveChannelGroup.Channels(Occ_Pos & pelvis_name(pelv_count)).Properties("length").Value
    channel_offset = Data.Root.ActiveChannelGroup.Channels(Occ_Pos & pelvis_name(pelv_count)).Properties("wf_start_offset").Value
    DX = Data.Root.ActiveChannelGroup.Channels(Occ_Pos & pelvis_name(pelv_count)).Properties("wf_increment").Value
    
    For i = 1 to (channel_length - 1)
        temp_value = Data.Root.ActiveChannelGroup.Channels(Occ_Pos & pelvis_name(pelv_count)).Values(i)
        If temp_value >= 50 Then
          channel_target_point = i
          Exit For
        End If
    Next
    If (channel_offset + (channel_target_point - 1)*DX) < load_time Then
      load_time = (channel_offset + (channel_target_point - 1)*DX)
    End If
  Next
  
  Call Data.Root.ActiveChannelGroup.Properties.Add("Pelvis_loading_time", 1000*load_time,DataTypeFloat64) 

  'Femur
  Dim Femur_name, Femur_count
  Femur_name = array("FEMRLELO__FOYB","FEMRLELO__MOXB", "FEMRLELO__MOYB")
    
	For Femur_count = 0 to 2
		'3ms clip
		Call ChnXMSCalc("","/" & Occ_Pos & Femur_name(Femur_count),3,0,"in one peak",1) 
		Call Data.Root.ActiveChannelGroup.Channels(Occ_Pos & Femur_name(Femur_count)).Properties.Add("Clip_3_ms", XmsAcceleration, DataTypeFloat64)
		Call Data.Root.ActiveChannelGroup.Channels(Occ_Pos & Femur_name(Femur_count)).Properties.Add("Clip_3_ms_T1", XmsTimeBegin, DataTypeFloat64)
		Call Data.Root.ActiveChannelGroup.Channels(Occ_Pos & Femur_name(Femur_count)).Properties.Add("Clip_3_ms_T2",XmsTimeEnd , DataTypeFloat64)
	Next
  '-----------------------------------------------

  'Vehicle delta V
  Call Data.Root.ChannelGroups("Vehicle").Activate

  Dim accel_location
  'Calculate delta-V
  If Occ_Pos = 11 Then
    accel_location = "VEHCRIFR"
  Else
    accel_location = "VEHCRIRE"
  End If

  'Remove data before time zero and create new channel then integrate data
  Call ChnTimeAreaCopy("[1]/Time axis","/10"  & accel_location & "__ACYC","/TimeRangeCopyX","/TimeRangeCopyY",0,1) 
  Call ChnIntegrate("/TimeRangeCopyX","/TimeRangeCopyY","/Integrated")
  Call ChnToWfChn("/TimeRangeCopyX","/Integrated",0) 
  'Take integrated acceleration and multiply by 35.316 (9.81*3.6) to convert from m/s to km/h 
  
  Call Calculate("Ch(""10" & accel_location & "__VEY"")=Ch(""/Integrated"")*9.81*3.6",NULL,NULL,"") 
  Call ChnTMinTMaxCalc("","10" & accel_location & "__VEY")
  Call Data.Root.ActiveChannelGroup.Channels("10" & accel_location & "__VEY").Properties.Add("Max_time", round(maxtime*1000,1), DataTypeFloat64) 
  
  Data.Root.ActiveChannelGroup.Channels("10" & accel_location & "__VEY").Properties("unit_string").Value = "km/h"
  Data.Root.ActiveChannelGroup.Channels("10" & accel_location & "__VEY").Properties("wf_xname").Value = "Time_" & accel_location & "__VEY"
  
  'Delete channels that aren't needed
  Call Data.Root.ActiveChannelGroup.Channels.Remove("TimeRangeCopyX")
  Call Data.Root.ActiveChannelGroup.Channels.Remove("TimeRangeCopyY")
  Call Data.Root.ActiveChannelGroup.Channels.Remove("Integrated")
End Sub

