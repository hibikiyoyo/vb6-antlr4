VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "History"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmHistory.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   2550
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Remove"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Clear history"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   1215
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem You can back track your history if you passed up something you liked and couldn't find

Private Sub Form_Load()
On Error Resume Next
TextToList List1, App.Path & "/BasicHistory.ini"
End Sub
Private Sub Form_Unload(Cancel As Integer)
Form1.Show
Form3.Hide
End Sub
Private Sub Label1_Click()
On Error Resume Next
List1.RemoveItem List1.ListIndex
ListToText List1, App.Path & "/BasicHistory.ini", False
End Sub
Private Sub Label3_Click()
On Error Resume Next
Msg$ = MsgBox("Are u sure you want to delete your history?", vbYesNo)
If Msg$ = vbNo Then: Exit Sub
If Msg$ = vbYes Then Kill App.Path & "/BasicHistory.ini": List1.Clear: Exit Sub
End Sub
