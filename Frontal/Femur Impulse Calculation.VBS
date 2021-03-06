'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 05/10/2010 08:52:57
'-- Author: 
'-- Comment: 
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.
'UPDATED on 6/7/12 by Raul Arbelaez to reflect an update by Rupp et al. (2009 ESV paper).  
  
Function CalculateImpulse(group, channelName)
    Dim femurChannel : Set femurChannel = group.Channels(channelName)
    group.Activate
    Dim impulseData : Set impulseData = New FemurImpulseData
    
    If femurChannel.Minimum >= -4050 Then 
      impulseData.Impulse = 0
      Set CalculateImpulse = impulseData
      Exit Function
    End If
 
    Dim minIndex : minIndex = ChnFind("ValEqual(Ch(""/" & femurChannel.Name & """), " & femurChannel.Minimum & ")") 'Get the index of the maximum femur compression   
    impulseData.Point1 = ChnFindReverse("Ch(""/" & femurChannel.Name & """)>=0", minIndex) + 1 'Find index just before crossing zero
    
    impulseData.Point2 = ChnFind("Ch(""/" & femurChannel.Name & """) > -4050", minIndex) - 1 'Find the index just before crossing -4050 
    
    impulseData.Point1Time = impulseData.Point1 * femurChannel.Properties("wf_increment").Value + femurChannel.Properties("wf_start_offset").Value - femurChannel.Properties("wf_increment").Value
    impulseData.Point2Time = impulseData.Point2 * femurChannel.Properties("wf_increment").Value + femurChannel.Properties("wf_start_offset").Value - femurChannel.Properties("wf_increment").Value
    
    Dim timeRangeName : timeRangeName = "TimeRangeCopyY"
    Dim integratedName : integratedName = "Integrated"
    Call ChnTimeAreaCopy("","/" & channelName,"", "/" & timeRangeName, impulseData.Point1Time, impulseData.Point2Time) 
    Call ChnIntegrate("","/" & timeRangeName,"/" & integratedName)
  
    impulseData.Impulse = round(group.Channels(integratedName).Minimum * -1,1)
    
    Call group.Channels.Remove(timeRangeName)
    Call group.Channels.Remove(integratedName)
    
    Set CalculateImpulse = impulseData
End Function

  Class FemurImpulseData 
    Dim Point1, Point2, Impulse, Point1Time, Point2Time
  End Class
  
  Class KTHData
    Private iHipLoadRisk, iHipImpulseRisk, iFemurLoadRisk, iHipRisk, iKth, iKthLocation
    
    Public Property Get HipLoadRisk
      HipLoadRisk = iHipLoadRisk
    End Property
    
    Public Property Let HipLoadRisk(v)
      iHipLoadRisk = v
      Call SetHipRisk
    End Property
    
    Public Property Get HipImpulseRisk
      HipImpulseRisk = iHipImpulseRisk
    End Property
    
    Public Property Let HipImpulseRisk(v)
      iHipImpulseRisk = v
      Call SetHipRisk
    End Property
    
    Public Property Get HipRisk
      HipRisk = iHipRisk
    End Property
    
    Private Property Let HipRisk(v)
      iHipRisk = v
      Call SetKth
    End Property
    
    Public Property Get FemurLoadRisk
      FemurLoadRisk = iFemurLoadRisk
    End Property
    
    Public Property Let FemurLoadRisk(v)
      iFemurLoadRisk = v
      Call SetKth
    End Property
    
    Public Property Get KTH
      KTH = iKTH
    End Property
    
    Public Property Get KTHLocation
      KTHLocation = iKthLocation
    End Property
    
    Private Sub SetHipRisk
      If Not IsEmpty(iHipLoadRisk) And Not IsEmpty(iHipImpulseRisk) Then HipRisk = ValMin(iHipLoadRisk, iHipImpulseRisk)
    End Sub
    
    Private Sub SetKth
      If Not IsEmpty(iHipRisk) And Not IsEmpty(iFemurLoadRisk) Then
        If iHipRisk > iFemurLoadRisk Then
          iKTH = round(iHipRisk,3)
          iKTHLocation = "hip"
        Else
          iKTH = round(iFemurLoadRisk,3)
          iKTHLocation = "femur/knee"
        End If 
      End If 
    End Sub
  End Class
  
  Function CalculateKTHRisk(group, channelName, impulse)
    Dim femurChannel : Set femurChannel = group.Channels(ChannelName) 
    Dim femurData : Set femurData = New KTHData 
    Dim originalUnit : originalUnit = femurChannel.Properties("unit_string").Value
    
    Dim minimumForce : minimumForce = Round(Abs(ValueUnitConvert(femurChannel.Minimum, originalUnit, "kN")),2) 'Convert from N to kN and round
    group.Activate
      
    Dim x, A, B, C, D, F  
  
    If minimumForce > 9.3 Then		'The equation matches the Rupp tables up to 75% risk and goes above 100% when force exceeds 9.3 kN.
      femurData.HipLoadRisk = 1
    Else
      If minimumForce < 4 Then   'The numerical approximation method below does well for values above 4 kN.  The risk below 4 kN is essentially zero for the 50th male.
        femurData.HipLoadRisk = 0
      Else
        x = ln(minimumForce)
        A = 8.861433
        B = -61.839909
        C = 151.96960
        D = -145.644179
        F = 35.665567
        
        femurData.HipLoadRisk = E^(A*x^4 + B*x^3 + C*x^2 + D*x + F)
      End If
    End If
  
    If impulse > 210 Then
      femurData.HipImpulseRisk = 1
    Else
      If impulse = 113.6 Then '113.6 exceeds the 5% boundary (113.5) however it still rounds down to 5% the following code forces the risk to 5.1%
        femurData.HipImpulseRisk = 0.0508
      Else
        If impulse = 137.2 Then 'See previous comment
          femurData.HipImpulseRisk = 0.2506
        Else
          If impulse < 100 Then
            femurData.HipImpulseRisk = 0
          Else
          x = ln(impulse)
          A = 6.064737
          B = -114.63708
          C = 800.715836
          D = -2436.90941
          F = 2704.880196
          femurData.HipImpulseRisk = E^(A*x^4 + B*x^3 + C*x^2 + D*x + F)
          End If
        End If  
      End If
    End If
       
    If minimumForce <= 0 Then
      femurData.FemurLoadRisk = 0
    Else
      If minimumForce = 5.7 Then 'See previous comment for hip impulse risk
        femurData.FemurLoadRisk = 0.0506
      Else
        femurData.FemurLoadRisk =1-EXP(-EXP((LN(minimumForce)-2.514)/0.2611))  
      End If
    End If
    
    Set CalculateKTHRisk = femurData 
  End Function








