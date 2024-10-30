VERSION 5.00
Begin VB.Form frmRA 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Random Access by Eric Osterheldt"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   4950
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fram4 
      Caption         =   "Delete A Record"
      Height          =   975
      Left            =   2520
      TabIndex        =   17
      Top             =   2160
      Width           =   2415
      Begin VB.TextBox txtDelNum 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   23
         Text            =   "1"
         Top             =   480
         Width           =   255
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "Delete"
         Height          =   255
         Left            =   1560
         TabIndex        =   22
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lbl 
         Caption         =   "< Record Num"
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   24
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame fram3 
      Caption         =   "Search For A Record"
      Height          =   975
      Left            =   0
      TabIndex        =   16
      Top             =   2160
      Width           =   2415
      Begin VB.TextBox txtTopics 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1440
         TabIndex        =   20
         Text            =   "First Name"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         Height          =   255
         Left            =   1560
         TabIndex        =   19
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtSearch 
         Height          =   285
         Left            =   120
         TabIndex        =   18
         Text            =   "What to Search"
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lbl 
         Caption         =   "Field to Search >"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame fram2 
      Caption         =   "Get Record"
      Height          =   2055
      Left            =   2520
      TabIndex        =   6
      Top             =   0
      Width           =   2415
      Begin VB.TextBox txtGetNum 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Text            =   "1"
         Top             =   1680
         Width           =   255
      End
      Begin VB.CommandButton cmdGet 
         Caption         =   "Get"
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox txtGetB 
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox txtGetA 
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txtGetL 
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtGetF 
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lbl 
         Caption         =   " < Record Num"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   15
         Top             =   1680
         Width           =   1095
      End
   End
   Begin VB.Frame fram1 
      Caption         =   "Create Record"
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2415
      Begin VB.TextBox txtNum 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Text            =   "1"
         Top             =   1680
         Width           =   255
      End
      Begin VB.CommandButton cmdCreate 
         Caption         =   "Create"
         Height          =   255
         Left            =   1560
         TabIndex        =   5
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txtBDay 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Text            =   "Birthday"
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox txtAge 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Text            =   "Age"
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txtLName 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Text            =   "Last Name"
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtFName 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Text            =   "First Name"
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lbl 
         Caption         =   " < Record Num"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   8
         Top             =   1680
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmRA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#########################################
'#  Random Access Example                #
'#  by Eric Osterheldt (aka deep arctic) #
'#  Created on 2-23-01                   #
'#########################################

'Defining the record structure.
Private Type Records
    strFirst As String * 25
    strLast As String * 25
    intAge As Integer
    strBDay As String * 10
End Type

Private Sub cmdCreate_Click()
'Declaring the varibles.
Dim theRecords As Records
Dim intNum As Integer
'Opening the file as random.
Open "C:\Deep.rda" For Random As 1 Len = Len(theRecords)
'Putting all the data into the fields.
theRecords.strFirst = txtFName.Text
theRecords.strLast = txtLName.Text
theRecords.intAge = txtAge.Text
theRecords.strBDay = txtBDay.Text
'Makes sure the record number is a number.
intNum = Val(txtNum.Text)
'Puts the data.
Put #1, intNum, theRecords
'Closes the file.
Close #1
End Sub

Private Sub cmdDel_Click()
'Delcaring the varibles.
Dim theRecords As Records
Dim intNum As Integer
'Open the file for random access.
Open "C:\Deep.rda" For Random As 1 Len = Len(theRecords)
'Set everything to its null.
theRecords.strFirst = ""
theRecords.strLast = ""
theRecords.intAge = 0
theRecords.strBDay = ""
'Make sure the record number is a number.
intNum = Val(txtDelNum.Text)
'Put the null fields in the correct record.
Put #1, intNum, theRecords
'Close the file.
Close #1
End Sub

Private Sub cmdGet_Click()
'Declaring the varibles.
Dim theRecords As Records
'Open the file for random access.
Open "C:\Deep.rda" For Random As 1 Len = Len(theRecords)
'Make sure the record number is a number.
intNum = Val(txtGetNum.Text)
'Gets the specific data from the fields.
Get #1, intNum, theRecords
'The data in the fields is displayed in the textboxes.
txtGetF.Text = Trim(theRecords.strFirst)
txtGetL.Text = Trim(theRecords.strLast)
txtGetA.Text = Trim(theRecords.intAge)
txtGetB.Text = Trim(theRecords.strBDay)
'Closes the file.
Close #1
End Sub

Private Sub cmdSearch_Click()
'This part actually took me a few days to get,
'I kept trying to use Sequential Access within
'Random Access.  You can't do that. ;x

'Declaring the varibles.
Dim theRecords As Records
Dim strSearch As String
Dim intResult As Integer
Dim intCount As Integer
Dim intNum As Integer
'Open the file for random access.
Open "C:\Deep.rda" For Random As 1 Len = Len(theRecords)
'Gets the record count.
intCount = LOF(1) / Len(theRecords)
'Searching from the first record to the last record.
For intNum = 1 To intCount
'Get one by one.
Get #1, intNum, theRecords
'What are you searching for?  Specifies the field.
If txtTopics.Text = "First Name" Then strSearch = theRecords.strFirst
If txtTopics.Text = "Last Name" Then strSearch = theRecords.strLast
If txtTopics.Text = "Age" Then strSearch = theRecords.intAge
If txtTopics.Text = "Birthday" Then strSearch = theRecords.strBDay
'We found a match, display a msgbox of the info in all the record.
If Trim(strSearch) = txtSearch.Text Then
intResult = MsgBox(Trim(theRecords.strFirst) & " " & Trim(theRecords.strLast) & " is " & Trim(theRecords.intAge) & " years old born on " & Trim(theRecords.strBDay) & ".", vbYesNo, "Searching..")
End If
'If clicked yes then put the data in the textboxes.
If intResult = vbYes Then
txtGetF.Text = Trim(theRecords.strFirst)
txtGetL.Text = Trim(theRecords.strLast)
txtGetA.Text = Trim(theRecords.intAge)
txtGetB.Text = Trim(theRecords.strBDay)
End If
'Ends the For..Next Statement.
Next intNum
'Closes the file.
Close #1
End Sub
