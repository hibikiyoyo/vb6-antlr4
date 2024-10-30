VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Favorite list"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4710
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmFavs.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   4710
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   2550
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Add current page"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Navigate highlighted fav"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Load backup"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3360
      TabIndex        =   4
      ToolTipText     =   "If you accidentaly cleared you favs, click this to load your backup file"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Clear"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Remove"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Add"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   975
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem This form is for all of your favorites, it's hard to keep track of them all without this

Private Sub Form_Load()
TextToList List1, App.Path & "/BasicFavs.ini"
End Sub
Private Sub Form_Unload(Cancel As Integer)
Form1.Show
Form2.Hide
End Sub
Private Sub Label1_Click()
Dim A$
A$ = InputBox("Input a favorite to add to your fav's list", "Input fav.", "http://")
If A$ = "" Then: Exit Sub
List1.AddItem A$
ListToText List1, App.Path & "/BasicFavs.ini", False
Open App.Path & "/BasicFavsBackup.ini" For Append As #1
Print #1, A$
Close #1
End Sub
Private Sub Label2_Click()
On Error Resume Next
List1.RemoveItem List1.ListIndex
ListToText List1, App.Path & "/BasicFavs.ini", False
End Sub
Private Sub Label3_Click()
On Error Resume Next
List1.Clear
Kill App.Path & "/BasicFavs.ini"
End Sub
Private Sub Label4_Click()
On Error Resume Next
TextToList List1, App.Path & "/BasicFavsBackup.ini"
ListToText List1, App.Path & "/BasicFavs.ini", False
End Sub
Private Sub Label5_Click()
On Error Resume Next
Form1.WB1.Navigate List1.List(List1.ListIndex)
Form1.Text1.Text = List1.List(List1.ListIndex)
Form1.Show
Form2.Hide
End Sub
Private Sub Label6_Click()
List1.AddItem Form1.WB1.LocationURL
ListToText List1, App.Path & "/BasicFavs.ini", False
Open App.Path & "/BasicFavsBackup.ini" For Append As #1
Print #1, Form1.WB1.LocationURL
Close #1
End Sub
Private Sub List1_DblClick()
Label5_Click
End Sub
