VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Example Web browser by keith_escalade"
   ClientHeight    =   10335
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmBrowser.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10335
   ScaleMode       =   0  'User
   ScaleWidth      =   15240
   Begin MSComctlLib.StatusBar SB2 
      Align           =   2  'Align Bottom
      DragMode        =   1  'Automatic
      Height          =   270
      Left            =   0
      TabIndex        =   13
      Top             =   9810
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar SB1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   10080
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   "Example web browser in vb by keith_escalade"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   26406
            MinWidth        =   21167
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Stay on top"
      Height          =   255
      Left            =   11520
      TabIndex        =   11
      Top             =   0
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      TabIndex        =   10
      Text            =   "http://"
      Top             =   480
      Width           =   15015
   End
   Begin SHDocVwCtl.WebBrowser WB1 
      Height          =   8775
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   15015
      ExtentX         =   26485
      ExtentY         =   15478
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "res://C:\WINDOWS\system32\shdoclc.dll/dnserror.htm#http:///"
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Search"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8760
      TabIndex        =   14
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "History"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7680
      TabIndex        =   9
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hide controls <"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9840
      TabIndex        =   7
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "View fav."
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5520
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Add fav."
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4440
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Navigate"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6600
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Stop"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Refresh"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Forward ->"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "<- Back"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Menu Search 
      Caption         =   "Search"
      Visible         =   0   'False
      Begin VB.Menu google 
         Caption         =   "Google"
      End
      Begin VB.Menu Yahoo 
         Caption         =   "Yahoo!"
      End
      Begin VB.Menu Dogpile 
         Caption         =   "Dogpile"
      End
      Begin VB.Menu com37 
         Caption         =   "37.com"
      End
      Begin VB.Menu AskJeeves 
         Caption         =   "Ask Jeeves"
      End
      Begin VB.Menu About 
         Caption         =   "About.com"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem Please do not change this coding and call it yours
Rem This was made to show how you how a browser is made
Rem If you decide to compile this source, please don't take credit for it
Rem http://www.yahpro.org

Private Sub About_Click()
Dim A$
A$ = InputBox("What to search for", "Select a title to search for")
If A$ = "" Then: Exit Sub
Text1.Text = "About.com:" & A$
Label5_Click
End Sub
Private Sub AskJeeves_Click()
Dim A$
A$ = InputBox("What to search for", "Select a title to search for")
If A$ = "" Then: Exit Sub
Text1.Text = "Askj:" & A$
Label5_Click
End Sub
Private Sub Check1_Click()
If Check1.Value = 1 Then StayOnTop Me: Exit Sub
If Check1.Value = 0 Then DontStayOnTop Me: Exit Sub
End Sub
Private Sub com37_Click()
Dim A$
A$ = InputBox("What to search for", "Select a title to search for")
If A$ = "" Then: Exit Sub
Text1.Text = "37.com:" & A$
Label5_Click
End Sub
Private Sub Dogpile_Click()
Dim A$
A$ = InputBox("What to search for", "Select a title to search for")
If A$ = "" Then: Exit Sub
Text1.Text = "Dogpile:" & A$
Label5_Click
End Sub
Private Sub Form_Load()
On Error Resume Next
Rem Loads the positions last time the browser was exited
Open App.Path & "/BasicHeight.ini" For Input As #1
Input #1, sText$
Me.Height = sText$
Close #1
Open App.Path & "/BasicWidth.ini" For Input As #1
Input #1, sText$
Me.Width = sText$
Close #1
Open App.Path & "/BasicTop.ini" For Input As #1
Input #1, sText$
Me.Top = sText$
Close #1
Open App.Path & "/BasicLeft.ini" For Input As #1
Input #1, sText$
Me.Left = sText$
Close #1
WB1.Navigate "about:<b><body bgcolor = ""gray""><font face = ""arial"" size = ""2"" color = ""white""><center><br><br><br><br><br><br><br><br><br><br><br><br>Example web browser in vb!<br>by: keith_escalade<br><hr color = ""white""><a href = ""http://www.yahpro.org"">Yah-Pro<br><img src = ""http://www.yahpro.org/themes/XP-Silver/images/logo.gif"">"
End Sub
Private Sub Form_Resize()
Rem used to get the web browser control just right to fit the form
On Error Resume Next
If Me.Height > 11130 Then Me.Height = 11130
Text1.Width = Me.Width - 345
WB1.Width = Me.Width - 345
WB1.Height = Me.Height - 1935
PB1.Width = WB1.Width
PB1.Top = SB1.Top - 120
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Rem Saves positions at exiting
Open App.Path & "/BasicHeight.ini" For Output As #1
Print #1, Me.Height
Close #1
Open App.Path & "/BasicWidth.ini" For Output As #1
Print #1, Me.Width
Close #1
Open App.Path & "/BasicTop.ini" For Output As #1
Print #1, Me.Top
Close #1
Open App.Path & "/BasicLeft.ini" For Output As #1
Print #1, Me.Left
Close #1
Unload Me
End
End Sub
Private Sub google_Click()
Dim A$
A$ = InputBox("What to search for", "Select a title to search for")
If A$ = "" Then: Exit Sub
Text1.Text = "Google:" & A$
End Sub
Private Sub Label1_Click()
Rem Goes back to the previous page the web browser navigated
On Error GoTo DieError
WB1.GoBack
Exit Sub
DieError:
Exit Sub
End Sub
Private Sub Label10_Click()
Rem Shows the pop up menu
PopupMenu Search
End Sub
Private Sub Label2_Click()
Rem Goes to the next page
On Error GoTo DieError
WB1.GoForward
Exit Sub
DieError:
Exit Sub
End Sub
Private Sub Label3_Click()
Rem Refreshes the webpage
WB1.Refresh
End Sub
Private Sub Label4_Click()
Rem Stops all processes loading the page
WB1.Stop
End Sub
Private Sub Label5_Click()
On Error Resume Next
Rem Looks to see if the first letters are a search engine word
If Left(Text1, 7) = "Google:" Then: WB1.Navigate "http://www.google.com/search?hl=en&ie=ISO-8859-1&q=" & Right(Text1, Len(Text1) - 7): Exit Sub
If Left(Text1, 8) = "Dogpile:" Then: WB1.Navigate "http://search.dogpile.com/texis/search?q=" & Right(Text1, Len(Text1) - 8) & "&top=1": Exit Sub
If Left(Text1, 7) = "37.com:" Then: WB1.Navigate "http://search.megaspider.com/XP.html?" & Right(Text1, Len(Text1) - 7): Exit Sub
If Left(Text1, 5) = "Askj:" Then: WB1.Navigate "http://www.ask.com/main/askjeeves.asp?ask=" & Right(Text1, Len(Text1) - 5) & "&o=0": Exit Sub
If Left(Text1, 6) = "Yahoo:" Then: WB1.Navigate "http://www.search.yahoo.com/search?p=" & Right(Text1, Len(Text1) - 6): Exit Sub
If Left(Text1, 10) = "About.com:" Then: WB1.Navigate "http://www.search.about.com/fullsearch.htm?terms=" & Right(Text1, Len(Text1) - 10): Exit Sub
WB1.Navigate Text1.Text
Form3.List1.AddItem Text1.Text
ListToText Form3.List1, App.Path & "/BasicHistory.ini", False
SB2.Panels(1).Text = "0% Complete"
End Sub
Private Sub Label6_Click()
Rem Adds a favorite
Dim A$
A$ = InputBox("Input a favorite to add to your fav's list", "Input fav.", "http://")
If A$ = "" Then: Exit Sub
Form2.List1.AddItem A$
ListToText Form2.List1, App.Path & "/BasicFavs.ini", False
Open App.Path & "/BasicFavsBackup.ini" For Append As #1
Print #1, A$
Close #1
End Sub
Private Sub Label7_Click()
Me.Hide
Form2.Show
End Sub
Private Sub Label8_Click()
Rem Hides & shows the controls
If Label8.Caption = "Hide controls <" Then: Label8.Caption = "Show controls >": Label10.Visible = False: Form1.WindowState = 2: WB1.Height = 9880: WB1.Top = 360: Label8.Left = 13680: Label8.Top = 0: Label1.Visible = False: Label2.Visible = False: Label3.Visible = False: Label4.Visible = False: Label5.Visible = False: Label6.Visible = False: Label7.Visible = False: Label9.Visible = False: Text1.Visible = False: Exit Sub
If Label8.Caption = "Show controls >" Then: Label8.Caption = "Hide controls <": Label10.Visible = True: WB1.Top = 960: WB1.Height = Me.Height - 1935: Label1.Visible = True: Label8.Left = 9840: Label8.Top = 120: Label2.Visible = True: Label3.Visible = True: Label4.Visible = True: Label5.Visible = True: Label6.Visible = True: Label7.Visible = True: Label9.Visible = True: Text1.Visible = True: Exit Sub
End Sub
Private Sub Label9_Click()
Form3.Show
Form1.Hide
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
Rem If the enter key is hit then click label5
If KeyAscii = 13 Then Label5_Click
End Sub
Private Sub WB1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
Rem Notifies when page is done loading
SB2.Panels(1).Text = "Document done"
End Sub
Private Sub WB1_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
Rem Notifies you what percent of the page is done, would look nice with a progress bar
On Error Resume Next
Dim Stuff As Long, Crap As Long, LastCrap As Long
Stuff = ProgressMax / 100
Crap = Progress / Stuff
If Crap < 1 Then Exit Sub
If Crap > 100 Then Exit Sub
If Crap < LastCrap Then Exit Sub
SB2.Panels(1).Text = Crap & "% Complete"
LastCrap = Crap
End Sub
Private Sub WB1_StatusTextChange(ByVal Text As String)
Rem Description of page
If WB1.LocationURL = "about:<b><body%20bgcolor%20=%20""gray""><font%20face%20=%20""arial""%20size%20=%20""2""%20color%20=%20""white""><center><br><br><br><br><br><br><br><br><br><br><br><br>Example%20web%20browser%20in%20vb!<br>by:%20keith_escalade<br><hr%20color%20=%20""white""><a%20href%20=%20""http://www.yahpro.org"">Yah-Pro<br><img%20src%20=%20""http://www.yahpro.org/themes/XP-Silver/images/logo.gif"">" Then: Exit Sub
Me.Caption = WB1.LocationURL & " : " & WB1.LocationName
SB1.SimpleText = Text
End Sub
Private Sub WB1_TitleChange(ByVal Text As String)
Rem Changes text1 when title is changed
If WB1.LocationURL = "about:<b><body%20bgcolor%20=%20""gray""><font%20face%20=%20""arial""%20size%20=%20""2""%20color%20=%20""white""><center><br><br><br><br><br><br><br><br><br><br><br><br>Example%20web%20browser%20in%20vb!<br>by:%20keith_escalade<br><hr%20color%20=%20""white""><a%20href%20=%20""http://www.yahpro.org"">Yah-Pro<br><img%20src%20=%20""http://www.yahpro.org/themes/XP-Silver/images/logo.gif"">" Then: Exit Sub
Text1.Text = WB1.LocationURL
Text1.SelStart = Len(Text1.Text)
End Sub
Private Sub Yahoo_Click()
Dim A$
A$ = InputBox("What to search for", "Select a title to search for")
If A$ = "" Then: Exit Sub
Text1.Text = "Yahoo:" & A$
Label5_Click
End Sub
