Attribute VB_Name = "ModMain"



Public hMixer&
Public MaxSources&
Public ProductName$

Private Destinations&

Public MCD As MIXERCONTROLDETAILS

Private ML As MIXERLINE

Type RECT
     rLeft As Long
     rTop As Long
     rRight As Long
     rBottom As Long
End Type



Type MIXERSETTINGS
     MxrChannels As Long
     MxrLeftVol As Long
     MxrRightVol As Long
     MxrVol As Long
     MxrVolID As Long
     MxrMute As Long
     MxrMuteID As Long
     MxrPeakID As Long
End Type


Public MixerState() As MIXERSETTINGS


Declare Function BitBlt& Lib "gdi32" (ByVal hDestDC&, ByVal x1&, ByVal y1&, ByVal nWidth&, ByVal nHeight&, ByVal hSrcDC&, ByVal xSrc&, ByVal ySrc&, ByVal dwRop&)
Declare Function DrawEdge& Lib "user32" (ByVal ahDc&, lpRect As RECT, ByVal nEdge&, ByVal nFlags&)
Declare Function SetRect& Lib "user32" (lpRect As RECT, ByVal x1&, ByVal y1&, ByVal x2&, ByVal y2&)

Declare Sub CopyStructFromPtr Lib "kernel32" Alias "RtlMoveMemory" (struct As Any, ByVal ptr&, ByVal cb&)
Declare Sub CopyPtrFromStruct Lib "kernel32" Alias "RtlMoveMemory" (ByVal ptr&, struct As Any, ByVal cb&)

Declare Function GlobalAlloc& Lib "kernel32" (ByVal wFlags&, ByVal dwBytes&)
Declare Function GlobalFree& Lib "kernel32" (ByVal hMem&)
Declare Function GlobalLock& Lib "kernel32" (ByVal hMem&)
Declare Function GlobalUnlock& Lib "kernel32" (ByVal hMem&)
Public Const Ttl = "Volume Control"

Public Function VolAvail() As Boolean

    #If Win32 Then
        If Not MixerPresent Then VolAvail = False: Exit Function
        If Not OpenMixer Then VolAvail = False: Exit Function
        If Not GetDeviceCapabilities Then VolAvail = False: Exit Function
        GetMixerInfo
        VolAvail = True
    #Else
        VolAvail = False
    #End If

End Function
Private Function MixerPresent() As Boolean
    If mixerGetNumDevs() Then
       MixerPresent = True
    Else
      MixerPresent = False
    End If
End Function
Private Function OpenMixer() As Boolean
    If mixerOpen(hMixer, 0, 0, 0, 0) = 0 Then
       OpenMixer = True
    Else
       OpenMixer = False
    End If
End Function
Private Function GetDeviceCapabilities() As Boolean

    Dim Msg$
    Dim MxrCaps As MIXERCAPS

    If mixerGetDevCaps(0, MxrCaps, Len(MxrCaps)) = 0 Then
       Destinations = MxrCaps.cDestinations - 1
       ProductName = Left(MxrCaps.szPname, InStr(MxrCaps.szPname, vbNullChar) - 1)
       GetDeviceCapabilities = True
    Else
       GetDeviceCapabilities = False
    End If
End Function
Private Sub GetMixerInfo()
    
    Dim Dst&, Src&
    Dim ControlID&

    For Dst = 0 To Destination
        ML.cbStruct = Len(ML)
        ML.dwDestination = Dst
        mixerGetLineInfo hMixer, ML, MIXER_GETLINEINFOF_DESTINATION

        If ML.dwComponentType = MIXERLINE_COMPONENTTYPE_DST_SPEAKERS Then

           If ML.cConnections > 10 Then
              ML.cConnections = 10
              MaxSources = 10
           Else
              MaxSources = ML.cConnections
           End If
           ReDim MixerState(MaxSources)
           MixerState(0).MxrChannels = ML.cChannels
           ControlID = GetControlID(ML.dwComponentType, MIXERCONTROL_CONTROLTYPE_VOLUME)
           If ControlID <> 0 Then
              With MCD
                  .cbDetails = 4
                  .cbStruct = 24
                  .cChannels = ML.cChannels
                  .dwControlID = ControlID
                  .item = 0
                  .paDetails = VarPtr(MixerState(0).MxrVol)
              End With
              mixerGetControlDetails hMixer, MCD, MIXER_GETCONTROLDETAILSF_VALUE
              MixerState(0).MxrVol = 65535 - MixerState(0).MxrVol
              MixerState(0).MxrVolID = MCD.dwControlID
           Else
              FrmMxr.SldrVol(0).Enabled = 0
           End If
           ControlID = GetControlID(ML.dwComponentType, MIXERCONTROL_CONTROLTYPE_MUTE)
           If ControlID <> 0 Then
              With MCD
                  .cbDetails = 4
                  .cbStruct = Len(MCD)
                  .cChannels = ML.cChannels
                  .dwControlID = ControlID
                  .item = 0
                  .paDetails = VarPtr(MixerState(0).MxrMute)
              End With
              mixerGetControlDetails hMixer, MCD, MIXER_GETCONTROLDETAILSF_VALUE
              MixerState(0).MxrMuteID = MCD.dwControlID
           Else
'              FrmMxr.ChkMute(0).Enabled = 0
           End If
          ControlID = GetControlID(ML.dwComponentType, MIXERCONTROL_CONTROLTYPE_PEAKMETER)
           If ControlID <> 0 Then
              MixerState(0).MxrPeakID = ControlID
           End If

           For Src = 0 To ML.cConnections - 1
               ML.cbStruct = Len(ML)
               ML.dwDestination = Dst
               ML.dwSource = Src
               mixerGetLineInfo hMixer, ML, MIXER_GETLINEINFOF_SOURCE
               MixerState(Src + 1).MxrChannels = ML.cChannels
'              FrmMxr.LblName(Src + 1).Caption = Left(ML.szName, InStr(ML.szName, vbNullChar) - 1)

               ControlID = GetControlID(ML.dwComponentType, MIXERCONTROL_CONTROLTYPE_VOLUME)
               If ControlID <> 0 Then
                  With MCD
                      .cbDetails = 4
                      .cbStruct = Len(MCD)
                      .cChannels = ML.cChannels
                      .dwControlID = ControlID
                      .item = 0
                      .paDetails = VarPtr(MixerState(Src + 1).MxrVol)
                  End With
                  mixerGetControlDetails hMixer, MCD, MIXER_GETCONTROLDETAILSF_VALUE
                  MixerState(Src + 1).MxrVol = 65535 - MixerState(Src + 1).MxrVol
                  MixerState(Src + 1).MxrVolID = MCD.dwControlID
               Else
'                  FrmMxr.SldrVol(Src + 1).Enabled = 0
               End If

               ControlID = GetControlID(ML.dwComponentType, MIXERCONTROL_CONTROLTYPE_MUTE)
               If ControlID <> 0 Then
                  With MCD
                      .cbDetails = 4
                      .cbStruct = Len(MCD)
                      .cChannels = ML.cChannels
                      .dwControlID = ControlID
                      .item = 0
                      .paDetails = VarPtr(MixerState(Src + 1).MxrMute)
                  End With
                  mixerGetControlDetails hMixer, MCD, MIXER_GETCONTROLDETAILSF_VALUE
                  MixerState(Src + 1).MxrMuteID = MCD.dwControlID
               Else
                  'FrmMxr.ChkMute(Src + 1).Enabled = 0
               End If

               ControlID = GetControlID(ML.dwComponentType, MIXERCONTROL_CONTROLTYPE_PEAKMETER)
               If ControlID <> 0 Then
                  MixerState(Src + 1).MxrPeakID = ControlID
               End If
           Next
           Exit For
        End If
    Next

End Sub
Public Function GetControlID&(ByVal ComponentType&, ByVal ControlType&)

   Dim hMem&
   Dim MC As MIXERCONTROL
   Dim MxrLine As MIXERLINE
   Dim MLC As MIXERLINECONTROLS

   MxrLine.cbStruct = Len(MxrLine)
   MxrLine.dwComponentType = ComponentType

   If mixerGetLineInfo(hMixer, MxrLine, MIXER_GETLINEINFOF_COMPONENTTYPE) = 0 Then
      MLC.cbStruct = Len(MLC)
      MLC.dwLineID = ML.dwLineID
      MLC.dwControl = ControlType
      MLC.cControls = 1
      MLC.cbmxctrl = Len(MC)

      hMem = GlobalAlloc(&H40, Len(MC))
      MLC.pamxctrl = GlobalLock(hMem)

      MC.cbStruct = Len(MC)

      If mixerGetLineControls(hMixer, MLC, MIXER_GETLINECONTROLSF_ONEBYTYPE) = 0 Then
         CopyStructFromPtr MC, MLC.pamxctrl, Len(MC)
         GetControlID = MC.dwControlID
      End If

      GlobalUnlock hMem
      GlobalFree hMem
   End If

End Function


'___________________________________

Public Sub AdjustOutput(InVolume As Long, Bal As Integer)
    Dim FaderVol&
    Dim PanPos&
    Dim hMem&
    Dim MCDMono As MIXERCONTROLDETAILS
    Dim MCDStereo As MIXERCONTROLDETAILS

    If MixerState(1).MxrChannels = 2 Then
       PanPos = Bal
       FaderVol = 65535 - InVolume
       If PanPos >= 0 Then
          MixerState(1).MxrRightVol = FaderVol
          MixerState(1).MxrLeftVol = FaderVol - ((PanPos / 100) * FaderVol)
       Else
          MixerState(1).MxrLeftVol = FaderVol
          MixerState(1).MxrRightVol = FaderVol + ((PanPos / 100) * FaderVol)
       End If
       MCDStereo.cbDetails = 4
       MCDStereo.cbStruct = 24
       MCDStereo.dwControlID = MixerState(1).MxrVolID
       MCDStereo.item = 0
       MCDStereo.cChannels = 2

       hMem = GlobalAlloc(&H40, 8)
       MCDStereo.paDetails = GlobalLock(hMem)
       CopyPtrFromStruct MCDStereo.paDetails, MixerState(1).MxrRightVol, 8
       CopyPtrFromStruct MCDStereo.paDetails, MixerState(1).MxrLeftVol, 8
       mixerSetControlDetails hMixer, MCDStereo, MIXER_SETCONTROLDETAILSF_VALUE
       GlobalUnlock hMem
       GlobalFree hMem
    Else
       MixerState(1).MxrVol = 65535 - InVolume
       MCDMono.cbDetails = Len(MixerState(1).MxrVol)
       MCDMono.cbStruct = Len(MCDMono)
       MCDMono.dwControlID = MixerState(1).MxrVolID
       MCDMono.item = 0
       MCDMono.cChannels = 1
       hMem = GlobalAlloc(&H40, 4)
       MCDMono.paDetails = GlobalLock(hMem)
       CopyPtrFromStruct MCDMono.paDetails, MixerState(1).MxrVol, 4
       mixerSetControlDetails hMixer, MCDMono, MIXER_SETCONTROLDETAILSF_VALUE
       GlobalUnlock hMem
       GlobalFree hMem
    End If
End Sub


Public Function CallBal(InPercent As Integer) As Integer

Dim C As Integer

If InPercent > 50 Then
  C = InPercent - 50
  CallBal = (100 * (C * 2)) / 100
End If

If InPercent < 50 Then
  C = 50 - InPercent
  CallBal = -((100 * (C * 2)) / 100)
End If

If InPercent = 50 Then
  CallBal = 0
End If

End Function

Public Function PercentE(S As Long) As Integer
  PercentE = ((S / 65535) * 100)
End Function

Public Function PercentF(S As Integer) As Long
  PercentF = ((65535 * S) / 100)
End Function


Public Sub MuteMe(State As Boolean)
    Dim hMem&
    MixerState(1).MxrMute = State
    MCD.cbStruct = Len(MCD)
    MCD.dwControlID = MixerState(1).MxrMuteID
    MCD.cbDetails = 4
    MCD.cChannels = 1
    MCD.item = 0
    hMem = GlobalAlloc(&H40, 4)
    MCD.paDetails = GlobalLock(hMem)
    CopyPtrFromStruct MCD.paDetails, MixerState(1).MxrMute, 4
    mixerSetControlDetails hMixer, MCD, MIXER_SETCONTROLDETAILSF_VALUE
    GlobalUnlock hMem
    GlobalFree hMem
End Sub



