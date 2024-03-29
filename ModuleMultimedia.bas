Attribute VB_Name = "ModuleMultimedia"

Option Explicit

'Private Declares
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

'Private types
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

'variable just in the module
Dim glo_from As Long
Dim glo_to As Long
Dim glo_AliasName As String
Dim glo_hWnd As Long

Public Function OpenMultimedia(hwnd As Long, AliasName As String, FileName As String, typeDevice As String) As String

Dim cmdToDo As String * 255
Dim dwReturn As Long
Dim ret As String * 128
Dim tmp As String * 255
Dim lenShort As Long
Dim ShortPathAndFile As String
Const WS_CHILD = &H40000000

lenShort = GetShortPathName(FileName, tmp, 255)
ShortPathAndFile = Left$(tmp, lenShort) 'cut short path from buffer


cmdToDo = "open " & ShortPathAndFile & " type " & typeDevice & " Alias " & AliasName & " parent " & hwnd & " Style " & WS_CHILD
dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)

If Not dwReturn = 0 Then  'not success
    mciGetErrorString dwReturn, ret, 128 'Get the error
    OpenMultimedia = ret: Exit Function
End If

'Success
OpenMultimedia = "Success"
End Function

Public Function PlayMultimedia(AliasName As String, from_where As String, to_where As String) As String

If from_where = vbNullString Then from_where = 0
If to_where = vbNullString Then to_where = GetTotalframes(AliasName)

'Improtant for auto repeat
If AliasName = glo_AliasName Then
    glo_from = from_where
    glo_to = to_where
End If

Dim cmdToDo As String * 128
Dim dwReturn As Long
Dim ret As String * 128

cmdToDo = "play " & AliasName & " from " & from_where & " to " & to_where

dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&) 'play

If Not dwReturn = 0 Then  'not success
    mciGetErrorString dwReturn, ret, 128 'get the error
    PlayMultimedia = ret
    Exit Function
End If

'Success
PlayMultimedia = "Success"
End Function

Public Function PlayMultimediaFullScreen(AliasName As String, from_where As String, to_where As String) As String

If from_where = vbNullString Then from_where = 0
If to_where = vbNullString Then to_where = GetTotalframes(AliasName)

'Improtant for auto repeat
If AliasName = glo_AliasName Then
    glo_from = from_where
    glo_to = to_where
End If

Dim cmdToDo As String * 128
Dim dwReturn As Long
Dim ret As String * 128

cmdToDo = "play " & AliasName & " Fullscreen from " & from_where & " to " & to_where

dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&) 'play

If Not dwReturn = 0 Then  'not success
    mciGetErrorString dwReturn, ret, 128 'get the error
    PlayMultimediaFullScreen = ret
    Exit Function
End If

'Success
PlayMultimediaFullScreen = "Success"
End Function

Public Function CloseMultimedia(AliasName As String) As String

Dim dwReturn As Long
Dim ret As String * 128
dwReturn = mciSendString("Close " & AliasName, 0&, 0&, 0&) 'close

If Not dwReturn = 0 Then  'not success
    mciGetErrorString dwReturn, ret, 128 'Get the error
    CloseMultimedia = ret
    Exit Function
End If

'Success
If AliasName = glo_AliasName Then 'if alias the same
'this mean the user close this alias then we must delete
'the timer Function
KillTimer glo_hWnd, 500
End If

CloseMultimedia = "Success"
End Function

Public Function PauseMultimedia(AliasName As String) As String

Dim dwReturn As Long
Dim ret As String * 128
dwReturn = mciSendString("Pause " & AliasName, 0&, 0&, 0&) 'pause

If Not dwReturn = 0 Then  'not success
    mciGetErrorString dwReturn, ret, 128
    PauseMultimedia = ret
    Exit Function
End If

'Success
PauseMultimedia = "Success"
End Function

Public Function StopMultimedia(AliasName As String) As String

Dim dwReturn As Long
Dim ret As String * 128
dwReturn = mciSendString("Stop " & AliasName, 0&, 0&, 0&) 'stop

If Not dwReturn = 0 Then  'not success
    mciGetErrorString dwReturn, ret, 128 'Get the error
    StopMultimedia = ret
    Exit Function
End If

'Success
StopMultimedia = "Success"
End Function

Public Function ResumeMultimedia(AliasName As String) As String

Dim dwReturn As Long
Dim ret As String * 128
dwReturn = mciSendString("Resume " & AliasName, 0&, 0&, 0&) 'Resume

If Not dwReturn = 0 Then  'not success
    mciGetErrorString dwReturn, ret, 128 'Get the error
    ResumeMultimedia = ret
    Exit Function
End If

'Success
ResumeMultimedia = "Success"
End Function

Public Function GetStatusMultimedia(AliasName As String) As String

Dim dwReturn As Long
Dim Status As String * 128
Dim ret As String * 128

dwReturn = mciSendString("status " & AliasName & " mode", Status, 128, 0&)  'Get status

If Not dwReturn = 0 Then  'not success
    GetStatusMultimedia = "ERROR"
    Exit Function
End If

'Extract just the string
Dim i As Integer
Dim CharA As String
Dim RChar As String
RChar = Right$(Status, 1)
For i = 1 To Len(Status)
    CharA = Mid(Status, i, 1)
    If CharA = RChar Then Exit For
    GetStatusMultimedia = GetStatusMultimedia + CharA
Next i
End Function

Public Function GetTotalframes(AliasName As String) As Long
Dim dwReturn As Long
Dim Total As String * 128

dwReturn = mciSendString("set " & AliasName & " time format frames", Total, 128, 0&)
dwReturn = mciSendString("status " & AliasName & " length", Total, 128, 0&)

If Not dwReturn = 0 Then  'not success
    GetTotalframes = -1
    Exit Function
End If

'Success
GetTotalframes = Val(Total)
End Function
Public Function OpenCDTray() As String
Dim dwReturn As Long
Dim Total As String * 128
 
dwReturn = mciSendString("set cd door open", 0, 0, 0)

If Not dwReturn = 0 Then  'not success
    OpenCDTray = "No"
    Exit Function
End If

OpenCDTray = "Success"
End Function

Public Function CloseCDTray() As String
Dim dwReturn As Long
Dim Total As String * 128
 
dwReturn = mciSendString("set cd door closed", 0, 0, 0)

If Not dwReturn = 0 Then  'not success
    CloseCDTray = "No"
    Exit Function
End If

CloseCDTray = "Success"
End Function



Public Function GetTotalTimeByMS(AliasName As String) As Long

Dim dwReturn As Long
Dim TotalTime As String * 128


dwReturn = mciSendString("set " & AliasName & " time format ms", TotalTime, 128, 0&)
dwReturn = mciSendString("status " & AliasName & " length", TotalTime, 128, 0&)

mciSendString "set " & AliasName & " time format frames", 0&, 128, 0& ' return focus to frames not to time

If Not dwReturn = 0 Then  'not success
    GetTotalTimeByMS = -1
    Exit Function
End If

'Success
GetTotalTimeByMS = Val(TotalTime)
End Function

Public Function MoveMultimedia(AliasName As String, to_where As Long) As String

Dim dwReturn As Long
Dim ret As String * 128

dwReturn = mciSendString("seek " & AliasName & " to " & to_where, 0&, 0&, 0&)
mciSendString "Play " & AliasName, 0&, 0&, 0&

If Not dwReturn = 0 Then  'not success
    mciGetErrorString dwReturn, ret, 128 'Get the error
    MoveMultimedia = ret
    Exit Function
End If

'Success
MoveMultimedia = "Success"
End Function

Public Function GetCurrentMultimediaPos(AliasName As String) As Long


Dim dwReturn As Long
Dim pos As String * 128

dwReturn = mciSendString("status " & AliasName & " position", pos, 128, 0&)

If Not dwReturn = 0 Then  'not success
    GetCurrentMultimediaPos = -1
    Exit Function
End If

'Success
GetCurrentMultimediaPos = Val(pos)
End Function

Public Function PutMultimedia(hwnd As Long, AliasName As String, Left As Long, Top As Long, Width As Long, Height As Long) As String
Dim dwReturn As Long
Dim ret As String * 128

If Width = 0 Or Height = 0 Then
    'Get Window Size
    Dim rec As RECT
    Call GetWindowRect(hwnd, rec)
    Width = rec.Right - rec.Left
    Height = rec.Bottom - rec.Top
End If

dwReturn = mciSendString("put " & AliasName & " window at " & Left & " " & Top & " " & Width & " " & Height, 0&, 0&, 0&)

If Not dwReturn = 0 Then  'not success
    mciGetErrorString dwReturn, ret, 128 'Get the error
    PutMultimedia = ret
    Exit Function
End If

'Success
PutMultimedia = "Success"
End Function
Public Function GetPercent(AliasName As String) As Long

On Error Resume Next
Dim TotalFrames As Long
Dim currframe As Long
TotalFrames = GetTotalframes(AliasName)
currframe = GetCurrentMultimediaPos(AliasName)

If TotalFrames = -1 Or currframe = -1 Then 'Not success
    GetPercent = -1
    Exit Function
End If

'Success
GetPercent = currframe * 100 / TotalFrames
End Function
Public Function GetFramesPerSecond(AliasName As String) As Long
Dim TotalFrames As Long
Dim TotalTime As Long
TotalTime = GetTotalTimeByMS(AliasName)
TotalFrames = GetTotalframes(AliasName)
If TotalFrames = -1 Or TotalTime = -1 Then 'Not success
    GetFramesPerSecond = -1
    Exit Function
End If

'Success
GetFramesPerSecond = TotalFrames / (TotalTime / 1000)
End Function
Public Function GetSize(AliasName As String, CxOrCy As String) As Long
If Not CxOrCy = "cx" And Not CxOrCy = "cy" Then GetSize = -1: Exit Function
Dim dwReturn As Long
Dim size As String * 128
Dim s1, s2, s3, Width, Height As Long

dwReturn = mciSendString("Where " & AliasName & " destination", size, 128, 0&)


If Not dwReturn = 0 Then  'not success
    GetSize = -1
    Exit Function
End If

s1 = InStr(1, size, " "): s2 = InStr(s1 + 1, size, " "): s1 = InStr(s2 + 1, size, " ")
Width = Mid(size, s2, s1 - s2): Height = Mid(size, s1 + 1)

'Success
If CxOrCy = "cx" Then 'get the width
GetSize = Width
ElseIf CxOrCy = "cy" Then 'Get the height
GetSize = Height
End If

End Function
Public Function CloseAll() As String

Dim dwReturn As Long
Dim ret As String * 128
dwReturn = mciSendString("Close All", 0&, 0&, 0&)

If Not dwReturn = 0 Then  'not success
    mciGetErrorString dwReturn, ret, 128 'Get the error
    CloseAll = ret
    Exit Function
End If

'Success
CloseAll = "Success"
End Function
Public Function AreMultimediaAtEnd(AliasName As String, lastFrame As Long) As Boolean

Dim currpos As Long

'if last frame is zero then get actaul last frame
If lastFrame = 0 Then lastFrame = GetTotalframes(AliasName)

currpos = Val(GetCurrentMultimediaPos(AliasName))

If currpos = -1 Or lastFrame = -1 Then 'there are an error then not resume
    AreMultimediaAtEnd = False
    Exit Function
End If
    
If lastFrame = currpos Or (lastFrame - 1) < currpos Then
AreMultimediaAtEnd = True ' ok we reach to last frame
Else
AreMultimediaAtEnd = False ' we not reach to last frame
End If
End Function
Public Function SetAutoRepeat(hwnd As Long, AliasName As String, first_frame As String, last_frame As String, autoTrueOrFalse As Boolean) As Boolean

Dim Result As String

If first_frame = vbNullString Then first_frame = 0
If last_frame = vbNullString Then last_frame = GetTotalframes(AliasName)

glo_from = first_frame 'store it in global to use it TimerFunction
glo_to = last_frame ' store it in global to use it TimerFunction

glo_hWnd = hwnd
If autoTrueOrFalse = True Then
    glo_AliasName = AliasName
    Result = SetTimer(hwnd, 500, 100, AddressOf TimerFunction)
Else
    glo_AliasName = vbNullString
    Result = KillTimer(hwnd, 500)
End If

If Result = 0 Then
    SetAutoRepeat = False
Else
    SetAutoRepeat = True
End If
End Function

Sub TimerFunction()
'Important for auto repeat
Dim currpos As Long
Dim Result As String
currpos = Val(GetCurrentMultimediaPos(glo_AliasName))
If currpos = -1 Then Exit Sub   'if  function get cuurent pos not success then exit
'
If Val(glo_to) = currpos Or (Val(glo_to) - 1) < currpos Then
    Result = PlayMultimedia(glo_AliasName, Str(glo_from), Str(glo_to))
    If Not Result = "Success" Then KillTimer glo_hWnd, 500 'if  function play not success then kill timer
End If
End Sub

Public Sub SetDefaultDevice(typeDevice As String, drvDefaultDevice As String)
Dim Res As String
Dim tmp As String * 255
Dim Windir As String
Res = GetWindowsDirectory(tmp, 255)
Windir = Left$(tmp, Res)
Res = WritePrivateProfileString("MCI", typeDevice, drvDefaultDevice, Windir & "\" & "system.ini")
End Sub

Public Function GetDefaultDevice(typeDevice As String) As String
Dim tmp As String * 255
Dim Res As String
Dim Windir As String
Res = GetWindowsDirectory(tmp, 255)
Windir = Left$(tmp, Res)
Res = GetPrivateProfileString("MCI", typeDevice, "None", tmp, 255, Windir & "\" & "system.ini")
GetDefaultDevice = Left$(tmp, Res)
End Function
