Attribute VB_Name = "vars"
Option Explicit
'Public Declare Function SipShowIM Lib "Coredll" (ByVal flags As Long) As Long
Public auth As Integer
Public clientid As Integer
Public categorieid As Integer
Public produsid As Integer

Public Const SW_HIDE = 0
Public Const SW_SHOWNORMAL = 1

Public Const SHFS_SHOWTASKBAR = &H1
Public Const SHFS_HIDETASKBAR = &H2
Public Const SHFS_SHOWSIPBUTTON = &H4
Public Const SHFS_HIDESIPBUTTON = &H8
Public Const SHFS_SHOWSTARTICON = &H10
Public Const SHFS_HIDESTARTICON = &H20
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Public Const SWP_SHOWWINDOW = &H40
Public Const SM_CXSCREEN = &H0
Public Const SM_CYSCREEN = &H1
Public Const HHTASKBARHEIGHT = 26


Declare Function GetSystemMetrics Lib "Coredll" ( _
    ByVal nIndex As Long) As Long

Declare Function SHFullScreen Lib "aygshell" ( _
    ByVal hwndRequester As Long, _
    ByVal dwState As Long) As Boolean

Declare Function MoveWindow Lib "Coredll" ( _
    ByVal hwnd As Long, _
    ByVal X As Long, _
    ByVal y As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long, _
    ByVal bRepaint As Long) As Long

Declare Function SetForegroundWindow Lib "Coredll" ( _
    ByVal hwnd As Long) As Boolean

Declare Function GetLastError Lib "Coredll" () As Long

Declare Function ShowWindow Lib "Coredll" ( _
    ByVal hwnd As Long, _
    ByVal nCmdShow As Long) As Long

Declare Function FindWindow Lib "Coredll" Alias "FindWindowW" ( _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long
    
Public Declare Sub keybd_event Lib "coredll.dll" (ByVal bVK As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)


Public Sub formx(frm)
Dim lret
Dim lHwnd As Long
lHwnd = frm.hwnd
lret = FindWindow("menu_worker", "")
If lret <> 0 Then 'window found
 ShowWindow lret, SW_HIDE
End If
lret = SetForegroundWindow(frm.hwnd)
lret = MoveWindow(frm.hwnd, 0, 0, GetSystemMetrics(SM_CXSCREEN), GetSystemMetrics(SM_CYSCREEN) + HHTASKBARHEIGHT, 0)
lret = SHFullScreen(lHwnd, SHFS_HIDESTARTICON)
'lret = SHFullScreen(lHwnd, SHFS_HIDESIPBUTTON)
lret = SHFullScreen(lHwnd, SHFS_HIDETASKBAR)
End Sub

Public Sub hidetask(frm)
Dim lret
Dim lHwnd As Long
lHwnd = frm.hwnd
lret = SHFullScreen(lHwnd, SHFS_HIDESIPBUTTON)

End Sub

Function num2rol(v As String) As String
Dim pos, pos2 As Integer
Dim lung As Integer
Dim v2 As String
pos2 = InStr(1, v, ",")
If pos2 <> 0 Then
 v2 = Mid(v, 1, pos2 - 1)
Else
 v2 = v
End If

lung = Len(v2)
For pos = 1 To lung
 num2rol = num2rol + Mid(v2, pos, 1)
 If (lung - pos) / 3 = Int((lung - pos) / 3) And pos <> lung Then
  num2rol = num2rol + "."
 End If
Next

If pos2 <> 0 Then
 num2rol = num2rol + Mid(v, pos2)
End If
End Function

Function num2ron(v As String) As String
 Dim num As Double
 Dim txt As String
 If IsNumeric(v) Then
  num = CDbl(v)
  num = num / 10000
  txt = CStr(num)
  txt = Replace(txt, ".", ",")
  num2ron = num2rol(txt)
 End If
End Function

Function rol2num(v As String) As Double
 v = Replace(v, ".", "")
 v = Replace(v, ",", ".")
 rol2num = CDbl(v)
End Function
