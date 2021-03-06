'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 12/22/2010 14:55:09
'-- Author: 
'-- Comment: Determine when the max and min values occur (within HCT or 300 ms)
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.

'Calculate the Neck Force Rating
Dim Max_Head_Xg
Dim Max_Reb_Vel
Dim Max_T1Xg
Dim Max_T1XG_HeadXG
Dim Max_NIC
Dim Max_Fx, Min_Fx
Dim Max_Fz, Min_Fz
Dim Max_My, Min_My
Dim Max_NCE, Max_NCF, Max_NTE, Max_NTF
Dim Max_NEA, Max_NFA, Max_NEP, Max_NFP
Dim Max_Lower_Fx, Min_Lower_Fx
Dim Max_Lower_Fz, Min_Lower_Fz
Dim Max_Lower_My, Min_Lower_My
Dim Average_Accel

Dim HCT_end
Dim End_Time
HCT_end = Data.Root.ChannelGroups(3).Channels("HCT").Properties("HCT_end").Value

'Determine when to end evaluation (Head rebound or 300 ms)
If HCT_end < 0.3 then
    End_Time = HCT_End
    Else
    End_Time = 0.3
End if        

  'Trim data to HCT_end or 0.3 ms
  Call ChnTimeAreaCopy("","[2]/S1HEAD0000__ACXA","","[2]/S1HEAD0000__ACXA_trim",0,End_Time) '... XW,Y,E,E,TimeAreaCopyBeg,TimeAreaCopyEnd 
  Call ChnTimeAreaCopy("","[3]/AverageT1XG","","[3]/AverageT1XG_trim",0,End_Time) '... XW,Y,E,E,TimeAreaCopyBeg,TimeAreaCopyEnd 
  Call ChnTimeAreaCopy("","[3]/T1XG_HeadXG","","[3]/T1XG_HeadXG_trim",0,End_Time) '... XW,Y,E,E,TimeAreaCopyBeg,TimeAreaCopyEnd 
  Call ChnTimeAreaCopy("","[3]/NIC","","[3]/NIC_trim",0,End_Time) '... XW,Y,E,E,TimeAreaCopyBeg,TimeAreaCopyEnd 
  Call ChnTimeAreaCopy("","[2]/S1NECKUP00__FOXA","","[2]/S1NECKUP00__FOXA_trim",0,End_Time) '... XW,Y,E,E,TimeAreaCopyBeg,TimeAreaCopyEnd 
  Call ChnTimeAreaCopy("","[2]/S1NECKUP00__FOZA","","[2]/S1NECKUP00__FOZA_trim",0,End_Time) '... XW,Y,E,E,TimeAreaCopyBeg,TimeAreaCopyEnd 
  Call ChnTimeAreaCopy("","[3]/MOCy","","[3]/MOCy_trim",0,End_Time) '... XW,Y,E,E,TimeAreaCopyBeg,TimeAreaCopyEnd 
  Call ChnTimeAreaCopy("","[3]/NIJ_NTE","","[3]/NIJ_NTE_trim",0,End_Time) '... XW,Y,E,E,TimeAreaCopyBeg,TimeAreaCopyEnd 
  Call ChnTimeAreaCopy("","[3]/NIJ_NTF","","[3]/NIJ_NTF_trim",0,End_Time) '... XW,Y,E,E,TimeAreaCopyBeg,TimeAreaCopyEnd 
  Call ChnTimeAreaCopy("","[3]/NIJ_NCE","","[3]/NIJ_NCE_trim",0,End_Time) '... XW,Y,E,E,TimeAreaCopyBeg,TimeAreaCopyEnd 
  Call ChnTimeAreaCopy("","[3]/NIJ_NCF","","[3]/NIJ_NCF_trim",0,End_Time) '... XW,Y,E,E,TimeAreaCopyBeg,TimeAreaCopyEnd 
  Call ChnTimeAreaCopy("","[3]/NKM_NEA","","[3]/NKM_NEA_trim",0,End_Time) '... XW,Y,E,E,TimeAreaCopyBeg,TimeAreaCopyEnd
  Call ChnTimeAreaCopy("","[3]/NKM_NFA","","[3]/NKM_NFA_trim",0,End_Time) '... XW,Y,E,E,TimeAreaCopyBeg,TimeAreaCopyEnd
  Call ChnTimeAreaCopy("","[3]/NKM_NEP","","[3]/NKM_NEP_trim",0,End_Time) '... XW,Y,E,E,TimeAreaCopyBeg,TimeAreaCopyEnd
  Call ChnTimeAreaCopy("","[3]/NKM_NFP","","[3]/NKM_NFP_trim",0,End_Time) '... XW,Y,E,E,TimeAreaCopyBeg,TimeAreaCopyEnd
  Call ChnTimeAreaCopy("","[2]/S1NECKLO00__FOXA","","[2]/S1NECKLO00__FOXA_trim",0,End_Time) '... XW,Y,E,E,TimeAreaCopyBeg,TimeAreaCopyEnd
  Call ChnTimeAreaCopy("","[2]/S1NECKLO00__FOZA","","[2]/S1NECKLO00__FOZA_trim",0,End_Time) '... XW,Y,E,E,TimeAreaCopyBeg,TimeAreaCopyEnd
  Call ChnTimeAreaCopy("","[2]/S1NECKLO00__MOYB","","[2]/S1NECKLO00__MOYB_trim",0,End_Time) '... XW,Y,E,E,TimeAreaCopyBeg,TimeAreaCopyEnd
  

'Determine max, min and time values and then save as properties
Max_Head_Xg = Data.Root.ChannelGroups(2).Channels("S1HEAD0000__ACXA_trim").Properties("maximum").Value
Call Data.Root.ChannelGroups(2).Channels("S1HEAD0000__ACXA").Properties.Add("Max_Head_Xg", Max_Head_Xg, DataTypeFloat64)
Call ChnTMinTMaxCalc("","[2]/S1HEAD0000__ACXA_trim") '... XW,Y
Call Data.Root.ChannelGroups(2).Channels("S1HEAD0000__ACXA").Properties.Add("Max_Head_Xg_t", Maxtime, DataTypeFloat64)

Max_Reb_Vel = Data.Root.ChannelGroups(3).Channels("Head_Rebound_Velocity").Properties("maximum").Value
Call Data.Root.ChannelGroups(3).Channels("Head_Rebound_Velocity").Properties.Add("Max_Reb_Vel", Max_Reb_Vel, DataTypeFloat64)
Call ChnTMinTMaxCalc("","[3]/Head_Rebound_Velocity") '... XW,Y
Call Data.Root.ChannelGroups(3).Channels("Head_Rebound_Velocity").Properties.Add("Max_Reb_Vel_t", Maxtime, DataTypeFloat64)

Max_T1Xg = Data.Root.ChannelGroups(3).Channels("AverageT1XG_trim").Properties("maximum").Value
Call Data.Root.ChannelGroups(3).Channels("AverageT1XG").Properties.Add("Max_T1Xg", Max_T1Xg, DataTypeFloat64)
Call ChnTMinTMaxCalc("","[3]/AverageT1XG_trim") '... XW,Y
Call Data.Root.ChannelGroups(3).Channels("AverageT1XG").Properties.Add("Max_T1Xg_t", Maxtime, DataTypeFloat64)

Max_T1XG_HeadXG = Data.Root.ChannelGroups(3).Channels("T1XG_HeadXG_trim").Properties("maximum").Value
Call Data.Root.ChannelGroups(3).Channels("T1XG_HeadXG").Properties.Add("Max_T1XG_HeadXG", Max_T1XG_HeadXG, DataTypeFloat64)
Call ChnTMinTMaxCalc("","[3]/T1XG_HeadXG_trim") '... XW,Y
Call Data.Root.ChannelGroups(3).Channels("T1XG_HeadXG").Properties.Add("Max_T1XG_HeadXG_t", Maxtime, DataTypeFloat64)

Max_NIC = Data.Root.ChannelGroups(3).Channels("NIC_trim").Properties("maximum").Value
Call Data.Root.ChannelGroups(3).Channels("NIC").Properties.Add("Max_NIC", Max_NIC, DataTypeFloat64)
Call ChnTMinTMaxCalc("","[3]/NIC_trim") '... XW,Y
Call Data.Root.ChannelGroups(3).Channels("NIC").Properties.Add("Max_NIC_t", Maxtime, DataTypeFloat64)

Max_Fx = Data.Root.ChannelGroups(2).Channels("S1NECKUP00__FOXA_trim").Properties("maximum").Value
If Max_Fx < 0 THEN
  Call Data.Root.ChannelGroups(2).Channels("S1NECKUP00__FOXA").Properties.Add("Max_Fx", 0, DataTypeFloat64)
ELSE
  Call Data.Root.ChannelGroups(2).Channels("S1NECKUP00__FOXA").Properties.Add("Max_Fx", Max_Fx, DataTypeFloat64)
End IF

Call ChnTMinTMaxCalc("","[2]/S1NECKUP00__FOXA_trim") '... XW,Y
Call Data.Root.ChannelGroups(2).Channels("S1NECKUP00__FOXA").Properties.Add("Max_Fx_t", Maxtime, DataTypeFloat64)

Min_Fx = Data.Root.ChannelGroups(2).Channels("S1NECKUP00__FOXA_trim").Properties("minimum").Value
Call Data.Root.ChannelGroups(2).Channels("S1NECKUP00__FOXA").Properties.Add("Min_Fx", Min_Fx, DataTypeFloat64)
Call Data.Root.ChannelGroups(2).Channels("S1NECKUP00__FOXA").Properties.Add("Min_Fx_t", Mintime, DataTypeFloat64)

Max_Fz = Data.Root.ChannelGroups(2).Channels("S1NECKUP00__FOZA_trim").Properties("maximum").Value
Call Data.Root.ChannelGroups(2).Channels("S1NECKUP00__FOZA").Properties.Add("Tension", Max_Fz, DataTypeFloat64)
Call ChnTMinTMaxCalc("","[2]/S1NECKUP00__FOZA_trim") '... XW,Y
Call Data.Root.ChannelGroups(2).Channels("S1NECKUP00__FOZA").Properties.Add("Tension_t", Maxtime, DataTypeFloat64)

Min_Fz = -1 * Data.Root.ChannelGroups(2).Channels("S1NECKUP00__FOZA_trim").Properties("minimum").Value
Call Data.Root.ChannelGroups(2).Channels("S1NECKUP00__FOZA").Properties.Add("Compression", Min_Fz, DataTypeFloat64)
Call Data.Root.ChannelGroups(2).Channels("S1NECKUP00__FOZA").Properties.Add("Compression_t", Mintime, DataTypeFloat64)

Max_My = Data.Root.ChannelGroups(3).Channels("MOCy_trim").Properties("maximum").Value
Call Data.Root.ChannelGroups(3).Channels("MOCy").Properties.Add("Max_My", Max_My, DataTypeFloat64)
Call ChnTMinTMaxCalc("","[3]/MOCy_trim") '... XW,Y
Call Data.Root.ChannelGroups(3).Channels("MOCy").Properties.Add("Max_My_t", Maxtime, DataTypeFloat64)

Min_My = Data.Root.ChannelGroups(3).Channels("MOCy_trim").Properties("minimum").Value
Call Data.Root.ChannelGroups(3).Channels("MOCy").Properties.Add("Min_My", Min_My, DataTypeFloat64)
Call Data.Root.ChannelGroups(3).Channels("MOCy").Properties.Add("Min_My_t", Mintime, DataTypeFloat64)

Max_NTE = Data.Root.ChannelGroups(3).Channels("NIJ_NTE_trim").Properties("maximum").Value
Call Data.Root.ChannelGroups(3).Channels("NIJ_NTE").Properties.Add("Max_NTE", Max_NTE, DataTypeFloat64)
Call ChnTMinTMaxCalc("","[3]/NIJ_NTE_trim") '... XW,Y
Call Data.Root.ChannelGroups(3).Channels("NIJ_NTE").Properties.Add("Max_NTE_t", Maxtime, DataTypeFloat64)

Max_NTF = Data.Root.ChannelGroups(3).Channels("NIJ_NTF_trim").Properties("maximum").Value
Call Data.Root.ChannelGroups(3).Channels("NIJ_NTF").Properties.Add("Max_NTF", Max_NTF, DataTypeFloat64)
Call ChnTMinTMaxCalc("","[3]/NIJ_NTF_trim") '... XW,Y
Call Data.Root.ChannelGroups(3).Channels("NIJ_NTF").Properties.Add("Max_NTF_t", Maxtime, DataTypeFloat64)

Max_NCE = Data.Root.ChannelGroups(3).Channels("NIJ_NCE_trim").Properties("maximum").Value
Call Data.Root.ChannelGroups(3).Channels("NIJ_NCE").Properties.Add("Max_NCE", Max_NCE, DataTypeFloat64)
Call ChnTMinTMaxCalc("","[3]/NIJ_NCE_trim") '... XW,Y
Call Data.Root.ChannelGroups(3).Channels("NIJ_NCE").Properties.Add("Max_NCE_t", Maxtime, DataTypeFloat64)

Max_NCF = Data.Root.ChannelGroups(3).Channels("NIJ_NCF_trim").Properties("maximum").Value
Call Data.Root.ChannelGroups(3).Channels("NIJ_NCF").Properties.Add("Max_NCF", Max_NCF, DataTypeFloat64)
Call ChnTMinTMaxCalc("","[3]/NIJ_NCF_trim") '... XW,Y
Call Data.Root.ChannelGroups(3).Channels("NIJ_NCF").Properties.Add("Max_NCF_t", Maxtime, DataTypeFloat64)

Max_NEA = Data.Root.ChannelGroups(3).Channels("NKM_NEA_trim").Properties("maximum").Value
Call Data.Root.ChannelGroups(3).Channels("NKM_NEA").Properties.Add("Max_NEA", Max_NEA, DataTypeFloat64)
Call ChnTMinTMaxCalc("","[3]/NKM_NEA_trim") '... XW,Y
Call Data.Root.ChannelGroups(3).Channels("NKM_NEA").Properties.Add("Max_NEA_t", Maxtime, DataTypeFloat64)

Max_NFA = Data.Root.ChannelGroups(3).Channels("NKM_NFA_trim").Properties("maximum").Value
Call Data.Root.ChannelGroups(3).Channels("NKM_NFA").Properties.Add("Max_NFA", Max_NFA, DataTypeFloat64)
Call ChnTMinTMaxCalc("","[3]/NKM_NFA_trim") '... XW,Y
Call Data.Root.ChannelGroups(3).Channels("NKM_NFA").Properties.Add("Max_NFA_t", Maxtime, DataTypeFloat64)

Max_NEP = Data.Root.ChannelGroups(3).Channels("NKM_NEP_trim").Properties("maximum").Value
Call Data.Root.ChannelGroups(3).Channels("NKM_NEP").Properties.Add("Max_NEP", Max_NEP, DataTypeFloat64)
Call ChnTMinTMaxCalc("","[3]/NKM_NEP_trim") '... XW,Y
Call Data.Root.ChannelGroups(3).Channels("NKM_NEP").Properties.Add("Max_NEP_t", Maxtime, DataTypeFloat64)

Max_NFP = Data.Root.ChannelGroups(3).Channels("NKM_NFP_trim").Properties("maximum").Value
Call Data.Root.ChannelGroups(3).Channels("NKM_NFP").Properties.Add("Max_NFP", Max_NFP, DataTypeFloat64)
Call ChnTMinTMaxCalc("","[3]/NKM_NFP_trim") '... XW,Y
Call Data.Root.ChannelGroups(3).Channels("NKM_NFP").Properties.Add("Max_NFP_t", Maxtime, DataTypeFloat64)

Max_Lower_Fx = Data.Root.ChannelGroups(2).Channels("S1NECKLO00__FOXA_trim").Properties("maximum").Value
Call Data.Root.ChannelGroups(2).Channels("S1NECKLO00__FOXA").Properties.Add("Max_Lower_Fx", Max_Lower_Fx, DataTypeFloat64)
Call ChnTMinTMaxCalc("","[2]/S1NECKLO00__FOXA_trim") '... XW,Y
Call Data.Root.ChannelGroups(2).Channels("S1NECKLO00__FOXA").Properties.Add("Max_Lower_Fx_t", Maxtime, DataTypeFloat64)

Min_Lower_Fx = Data.Root.ChannelGroups(2).Channels("S1NECKLO00__FOXA_trim").Properties("minimum").Value
Call Data.Root.ChannelGroups(2).Channels("S1NECKLO00__FOXA").Properties.Add("Min_Lower_Fx", Min_Lower_Fx, DataTypeFloat64)
Call Data.Root.ChannelGroups(2).Channels("S1NECKLO00__FOXA").Properties.Add("Min_Lower_Fx_t", Mintime, DataTypeFloat64)

Max_Lower_Fz = Data.Root.ChannelGroups(2).Channels("S1NECKLO00__FOZA_trim").Properties("maximum").Value
Call Data.Root.ChannelGroups(2).Channels("S1NECKLO00__FOZA").Properties.Add("Tension", Max_Lower_Fz, DataTypeFloat64)
Call ChnTMinTMaxCalc("","[2]/S1NECKLO00__FOZA_trim") '... XW,Y
Call Data.Root.ChannelGroups(2).Channels("S1NECKLO00__FOZA").Properties.Add("Tension_t", Maxtime, DataTypeFloat64)

Min_Lower_Fz = -1 * Data.Root.ChannelGroups(2).Channels("S1NECKLO00__FOZA_trim").Properties("minimum").Value
Call Data.Root.ChannelGroups(2).Channels("S1NECKLO00__FOZA").Properties.Add("Compression", Min_Lower_Fz, DataTypeFloat64)
Call Data.Root.ChannelGroups(2).Channels("S1NECKLO00__FOZA").Properties.Add("Compression_t", Mintime, DataTypeFloat64)

Max_Lower_My = Data.Root.ChannelGroups(2).Channels("S1NECKLO00__MOYB_trim").Properties("maximum").Value
Call Data.Root.ChannelGroups(2).Channels("S1NECKLO00__MOYB").Properties.Add("Max_Lower_My", Max_Lower_My, DataTypeFloat64)
Call ChnTMinTMaxCalc("","[2]/S1NECKLO00__MOYB_trim") '... XW,Y
Call Data.Root.ChannelGroups(2).Channels("S1NECKLO00__MOYB").Properties.Add("Max_Lower_My_t", Maxtime, DataTypeFloat64)

Min_Lower_My = Data.Root.ChannelGroups(2).Channels("S1NECKLO00__MOYB_trim").Properties("minimum").Value
Call Data.Root.ChannelGroups(2).Channels("S1NECKLO00__MOYB").Properties.Add("Min_Lower_My", Min_Lower_My, DataTypeFloat64)
Call Data.Root.ChannelGroups(2).Channels("S1NECKLO00__MOYB").Properties.Add("Min_Lower_My_t", Mintime, DataTypeFloat64)



'Remove temporary channels
Call Data.Root.ChannelGroups(2).Channels.Remove("S1HEAD0000__ACXA_trim")
Call Data.Root.ChannelGroups(3).Channels.Remove("AverageT1XG_trim")
Call Data.Root.ChannelGroups(3).Channels.Remove("T1XG_HeadXG_trim")
Call Data.Root.ChannelGroups(3).Channels.Remove("NIC_trim")
Call Data.Root.ChannelGroups(2).Channels.Remove("S1NECKUP00__FOXA_trim")
Call Data.Root.ChannelGroups(2).Channels.Remove("S1NECKUP00__FOZA_trim")
Call Data.Root.ChannelGroups(3).Channels.Remove("MOCy_trim")
Call Data.Root.ChannelGroups(3).Channels.Remove("NIJ_NTE_trim")
Call Data.Root.ChannelGroups(3).Channels.Remove("NIJ_NTF_trim")
Call Data.Root.ChannelGroups(3).Channels.Remove("NIJ_NCE_trim")
Call Data.Root.ChannelGroups(3).Channels.Remove("NIJ_NCF_trim")
Call Data.Root.ChannelGroups(3).Channels.Remove("NKM_NEA_trim")
Call Data.Root.ChannelGroups(3).Channels.Remove("NKM_NFA_trim")
Call Data.Root.ChannelGroups(3).Channels.Remove("NKM_NEP_trim")
Call Data.Root.ChannelGroups(3).Channels.Remove("NKM_NFP_trim")
Call Data.Root.ChannelGroups(2).Channels.Remove("S1NECKLO00__FOXA_trim")
Call Data.Root.ChannelGroups(2).Channels.Remove("S1NECKLO00__FOZA_trim")
Call Data.Root.ChannelGroups(2).Channels.Remove("S1NECKLO00__MOYB_trim")
