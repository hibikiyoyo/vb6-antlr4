Attribute VB_Name = "Module1"
Rem This module was created by keith_escalade
Rem http://www.yahpro.org
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hWndCallback As Long) As Long
Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Sub ReleaseCapture Lib "user32" ()
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Private Declare Function Getyahoostrings Lib "venky2.dll" (ByVal username As String, ByVal password As String, ByVal seed As String, ByVal result1 As String, ByVal result2 As String) As Integer
Public Declare Function venkymd5crypt Lib "venky2.dll" (ByVal pass As String, ByVal salt As String, ByVal Ret As String) As Long
Global LeftX
Global topY
Global sa
Public lRet As Long
Public Const SND_ASYNC = &H1
Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const WM_COMMAND = &H111

Public Const BM_SETCHECK = &HF1
Public Const BM_GETCHECK = &HF0

Public Const CB_GETCOUNT = &H146
Public Const CB_GETLBTEXT = &H148
Public Const CB_SETCURSEL = &H14E

Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDNEXT = 2
Public Const GW_CHILD = 5

Public Const LB_GETCOUNT = &H18B
Public Const LB_GETTEXT = &H189
Public Const LB_SETCURSEL = &H186

Public Const SW_HIDE = 0
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_NORMAL = 1
Public Const SW_SHOW = 5

Public Const VK_SPACE = &H20

Public Const WM_CHAR = &H102
Public Const WM_CLOSE = &H10
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MOVE = &HF012
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_SETTEXT = &HC
Public Const WM_SYSCOMMAND = &H112
Public Const SIZE_SE = &HF008&
Public Const RGN_AND = 1
Public Const RGN_COPY = 5
Public Const RGN_DIFF = 4
Public Const RGN_OR = 2
Public Const RGN_XOR = 3
Public Const SRCCOPY = &HCC0020
Public Function GetEncrStrings(Usrnm As String, Passwd As String, seed As String, str1 As String, str2 As String)
Rem Yahoo protocol related, don't worry about it
Dim ts As String
Dim ts2 As String
ts = Space$(24)
ts2 = Space$(24)
Dim X As Long
X = Getyahoostrings(Usrnm, Passwd, seed, ts, ts2)
str1 = ts
str2 = ts2
End Function
Public Function GetCrypt(Passwd As String)
Rem Yahoo protocol related, don't worry about it
Dim ts As String
ts = Space$(50)
Dim X As Long
Dim saltc As String
saltc = "_2S43d5f"
X = venkymd5crypt(Passwd, saltc, ts)
GetCrypt = ts
End Function
Sub YChatSend(what2say As String)
Dim imclass As Long
Dim richedit As Long
Dim Button As Long
imclass = FindWindow("imclass", vbNullString)
richedit = FindWindowEx(imclass, 0&, "richedit", vbNullString)
Call SendMessageByString(richedit, WM_SETTEXT, 0&, what2say)
imclass = FindWindow("imclass", vbNullString)
Button = FindWindowEx(imclass, 0&, "button", vbNullString)
Button = FindWindowEx(imclass, Button, "button", vbNullString)
Button = FindWindowEx(imclass, Button, "button", vbNullString)
Button = FindWindowEx(imclass, Button, "button", vbNullString)
Call SendMessageLong(Button, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(Button, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub StayOnTop(WhatForm As Form)
Rem Makes a certain form top most
Call SetWindowPos(WhatForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub
Public Sub DontStayOnTop(WhatForm As Form)
Rem If form is top most, this disables it
Call SetWindowPos(WhatForm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
End Sub
Sub Pause(interval)
Rem Freezes for specified seconds
current = Timer
Do While Timer - current < Val(interval)
DoEvents
Loop
End Sub
Sub RoundForm(WhatForm As Form)
Rem Creates a form with round, smooth edges
WhatForm.ScaleMode = 3
With WhatForm
TheMatrix = CreateRectRgn(0, 0, .ScaleWidth, .ScaleHeight)
notthematrix = CreateRectRgn(0, 0, .ScaleWidth, .ScaleHeight)
A = CreateRectRgn(10, 0, .ScaleWidth - 10, .ScaleHeight)
b = CreateRectRgn(0, 10, .ScaleWidth, .ScaleHeight - 10)
c = CreateEllipticRgn(0, 0, 20, 20)
d = CreateEllipticRgn(0, .ScaleHeight, 20, .ScaleHeight - 20)
e = CreateEllipticRgn(.ScaleWidth, 0, .ScaleWidth - 20, 20)
f = CreateEllipticRgn(.ScaleWidth, .ScaleHeight, .ScaleWidth - 20, .ScaleHeight - 20)
g = CombineRgn(TheMatrix, TheMatrix, A, 4)
g = CombineRgn(TheMatrix, TheMatrix, b, 4)
g = CombineRgn(TheMatrix, TheMatrix, c, 4)
g = CombineRgn(TheMatrix, TheMatrix, d, 4)
g = CombineRgn(TheMatrix, TheMatrix, e, 4)
g = CombineRgn(TheMatrix, TheMatrix, f, 4)
g = CombineRgn(TheMatrix, notthematrix, TheMatrix, 4)
m = SetWindowRgn(.hwnd, TheMatrix, True)
DeleteObject TheMatrix
DeleteObject notthematrix
DeleteObject A
DeleteObject b
DeleteObject c
DeleteObject d
DeleteObject e
DeleteObject f
DeleteObject g
DeleteObject m
End With
End Sub
Public Sub FormDrag(WhatForm As Form)
Rem Makes a form draggable. Ex: control_mousedown()
Rem                             FormDrag me
ReleaseCapture
Call SendMessage(WhatForm.hwnd, &HA1, 2, 0&)
End Sub
Public Sub Minimize(WhatForm As Form)
Rem Less code to type to minimize a form
WhatForm.WindowState = 1
End Sub
Public Sub Maximize(WhatForm As Form)
Rem Less code to type to maximize a form
WhatForm.WindowState = 2
End Sub
Public Sub Restore(WhatForm As Form)
Rem Less code to type to restore a form
WhatForm.WindowState = 0
End Sub
Public Sub Quit()
Rem Exits your application
End
End Sub
Sub OpenCDTray()
Rem Opens compact disc tray
lRet = mciSendString("set cdaudio door open", 0&, 0, 0)
End Sub
Sub CloseCDTray()
Rem Closes compact disc tray if already opened
lRet = mciSendString("set cdaudio door closed", 0&, 0, 0)
End Sub
Sub PlaySound(SoundFile$)
Rem I think this is for *.wav files only
sndPlaySound SoundFile, SND_ASYNC
End Sub
Function AIM_Algorithum(ByVal sUser As String, ByVal sPass As String) As String
Rem Aim protocol related, don't worry about it
On Error Resume Next
Dim sUserChar As Long, sVar As Long
DoEvents: sUser = Left(LCase(sUser), 1)
DoEvents: sUserChar = Int(Asc(sUser) - 96)
DoEvents: sVar = Int(sUserChar * 7696) + 738816
DoEvents: sBase = Int(sUserChar * 746512)
DoEvents: sVal = Int(Asc(Left(LCase(sPass), 1)) - 96) * sVar
AIM_Algorithum = Int(Int(sVal) - sVar) + Int(sBase + 71665152)
End Function
Function AIM_EncryptPW(ByVal sPass As String) As String
Rem Aim protocol related, don't worry about it
Dim vTable() As Variant, sString As String
Dim sLoop As Long, sHex As String
vTable = Array("84", "105", "99", "47", "84", "111", "99")
sString = "0x"
For sLoop = 0 To Len(sPass) - 1
sHex = Hex(Asc(Mid(sPass, sLoop + 1, 1)) Xor CLng(vTable(sLoop Mod 7)))
If CLng("&H" & sHex) < 16 Then
sString = sString & "0"
End If
sString = sString & sHex
Next
AIM_EncryptPW = LCase(sString)
End Function
Public Sub TextToList(WhatList As ListBox, TextFile$)
Rem Makes a .txt file from a listbox
On Error GoTo FartK
Dim X%
X% = FreeFile
Open TextFile$ For Input As #X
While Not EOF(X)
Input #X, sText$
DoEvents
WhatList.AddItem sText$
Wend
Close #X
Exit Sub
FartK:
Exit Sub
End Sub
Public Sub ListToText(WhatList As ListBox, TheDestination$, Append As Boolean)
Rem Opens a .txt file and loads it into a listbox
On Error GoTo ReportError
If Append = True Then
Dim X%
X% = FreeFile
Open TheDestination For Append As #X
For i = 0 To WhatList.ListCount - 1
DoEvents
Print #X, WhatList.List(i)
Next i
Close #X
End If
If Append = False Then
X% = FreeFile
Open TheDestination$ For Output As #X
For i = 0 To WhatList.ListCount - 1
DoEvents
Print #X, WhatList.List(i)
Next i
Close #X
End If
Exit Sub
ReportError:
MsgBox Error, vbCritical
End Sub
Public Sub Scramble(PB As PictureBox)
Rem Scrambles colors on a picturebox
Dim X, Y, blue As Integer
X = PB.ScaleWidth
Y = PB.ScaleHeight
For A = 0 To 300
X = X - PB.Width / 255
PB.Line (0, 0)-(X, Y), RGB(RandomNumber(255, 0), RandomNumber(255, 0), RandomNumber(255, 0)), BF
Next A
End Sub
Function RandomNumber(Number%, dec%)
Rem Generates a random number Ex: text1.text = RandomNumber(10,0), this will pick a random number from 0 to 10
Randomize
RandomNumber = Round(Rnd * Number, dec)
End Function
