VERSION 5.00
Begin VB.Form frmTTT 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tic Tac Toe"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   2070
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrTurn 
      Interval        =   1000
      Left            =   960
      Top             =   2640
   End
   Begin VB.Frame fram1 
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   -70
      Width           =   2055
      Begin VB.Label lbl9 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1380
         TabIndex        =   10
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label lbl6 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1380
         TabIndex        =   9
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lbl3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1380
         TabIndex        =   8
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lbl8 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   780
         TabIndex        =   7
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label lbl2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   780
         TabIndex        =   6
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lbl4 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   180
         TabIndex        =   5
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lbl1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   180
         TabIndex        =   4
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lbl7 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   180
         TabIndex        =   3
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Status: Your Turn"
         Height          =   255
         Left            =   0
         TabIndex        =   2
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label lbl5 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   780
         TabIndex        =   1
         Top             =   840
         Width           =   495
      End
      Begin VB.Line linHor2 
         X1              =   240
         X2              =   1800
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line linHor1 
         X1              =   240
         X2              =   1800
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line linVer2 
         X1              =   1320
         X2              =   1320
         Y1              =   240
         Y2              =   1800
      End
      Begin VB.Line linVer1 
         X1              =   720
         X2              =   720
         Y1              =   240
         Y2              =   1800
      End
   End
End
Attribute VB_Name = "frmTTT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'######################################'
'# Tic Tac Toe Example                #'
'# by Eric Osterheldt aka deep arctic #'
'# Created on 4-22-01                 #'
'# Email: deeparctic@yahoo.com        #'
'######################################'

'API declaration for a duration.
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub ResetGame()
'Resets everything by making the colors black again
'and by making all the captions null.
    lbl1.Caption = "": lbl2.Caption = "": lbl3.Caption = ""
    lbl4.Caption = "": lbl5.Caption = "": lbl6.Caption = ""
    lbl7.Caption = "": lbl8.Caption = "": lbl9.Caption = ""
    lbl1.ForeColor = vbBlack: lbl2.ForeColor = vbBlack: lbl3.ForeColor = vbBlack
    lbl4.ForeColor = vbBlack: lbl5.ForeColor = vbBlack: lbl6.ForeColor = vbBlack
    lbl7.ForeColor = vbBlack: lbl8.ForeColor = vbBlack: lbl9.ForeColor = vbBlack
End Sub

Public Function Draw() As Boolean
'If everything is full, the game is a draw.
    Dim intResult As Integer
    If Len(lbl1.Caption) = 1 And Len(lbl2.Caption) = 1 And Len(lbl3.Caption) = 1 And Len(lbl4.Caption) = 1 And Len(lbl5.Caption) = 1 And Len(lbl6.Caption) = 1 And Len(lbl7.Caption) = 1 And Len(lbl8.Caption) = 1 And Len(lbl9.Caption) = 1 Then
        intResult = MsgBox("Game Over: Draw!" & vbCrLf & vbCrLf & "Start a new game?", vbExclamation + vbYesNo, "Game")
        If intResult = vbYes Then
            Call ResetGame
        End If
    Draw = True
    End If
End Function

Public Function Combinations_X() As String
'Just checks for all the combinations for a tic tac toe.
    Dim intResult As Integer
    If lbl1.Caption = "X" And lbl2.Caption = "X" And lbl3.Caption = "X" Then
        lbl1.ForeColor = vbRed: lbl2.ForeColor = vbRed: lbl3.ForeColor = vbRed
        Combinations_X = "You won"
    ElseIf lbl4.Caption = "X" And lbl5.Caption = "X" And lbl6.Caption = "X" Then
        lbl4.ForeColor = vbRed: lbl5.ForeColor = vbRed: lbl6.ForeColor = vbRed
        Combinations_X = "You won"
    ElseIf lbl7.Caption = "X" And lbl8.Caption = "X" And lbl9.Caption = "X" Then
        lbl7.ForeColor = vbRed: lbl8.ForeColor = vbRed: lbl9.ForeColor = vbRed
        Combinations_X = "You won"
    ElseIf lbl1.Caption = "X" And lbl5.Caption = "X" And lbl9.Caption = "X" Then
        lbl1.ForeColor = vbRed: lbl5.ForeColor = vbRed: lbl9.ForeColor = vbRed
        Combinations_X = "You won"
    ElseIf lbl3.Caption = "X" And lbl5.Caption = "X" And lbl7.Caption = "X" Then
        lbl3.ForeColor = vbRed: lbl5.ForeColor = vbRed: lbl7.ForeColor = vbRed
        Combinations_X = "You won"
    ElseIf lbl1.Caption = "X" And lbl4.Caption = "X" And lbl7.Caption = "X" Then
        lbl1.ForeColor = vbRed: lbl4.ForeColor = vbRed: lbl7.ForeColor = vbRed
        Combinations_X = "You won"
    ElseIf lbl2.Caption = "X" And lbl5.Caption = "X" And lbl8.Caption = "X" Then
        lbl2.ForeColor = vbRed: lbl5.ForeColor = vbRed: lbl8.ForeColor = vbRed
        Combinations_X = "You won"
    ElseIf lbl3.Caption = "X" And lbl6.Caption = "X" And lbl9.Caption = "X" Then
        lbl3.ForeColor = vbRed: lbl6.ForeColor = vbRed: lbl9.ForeColor = vbRed
        Combinations_X = "You won"
    End If
    If Combinations_X = "You won" Then
        intResult = MsgBox("Game Over: You win!" & vbCrLf & vbCrLf & "Start a new game?", vbExclamation + vbYesNo, "Game")
        If intResult = vbYes Then
            Call ResetGame
        End If
    End If
End Function

Public Function Combinations_O()
'Just checks for all the combinations for a tic tac toe.
    Dim intResult As Integer
    If lbl1.Caption = "O" And lbl2.Caption = "O" And lbl3.Caption = "O" Then
        lbl1.ForeColor = vbRed: lbl2.ForeColor = vbRed: lbl3.ForeColor = vbRed
        Combinations_O = "You lost"
    ElseIf lbl4.Caption = "O" And lbl5.Caption = "O" And lbl6.Caption = "O" Then
        lbl4.ForeColor = vbRed: lbl5.ForeColor = vbRed: lbl6.ForeColor = vbRed
        Combinations_O = "You lost"
    ElseIf lbl7.Caption = "O" And lbl8.Caption = "O" And lbl9.Caption = "O" Then
        lbl7.ForeColor = vbRed: lbl8.ForeColor = vbRed: lbl9.ForeColor = vbRed
        Combinations_O = "You lost"
    ElseIf lbl1.Caption = "O" And lbl5.Caption = "O" And lbl9.Caption = "O" Then
        lbl1.ForeColor = vbRed: lbl5.ForeColor = vbRed: lbl9.ForeColor = vbRed
        Combinations_O = "You lost"
    ElseIf lbl3.Caption = "O" And lbl5.Caption = "O" And lbl7.Caption = "O" Then
        lbl3.ForeColor = vbRed: lbl5.ForeColor = vbRed: lbl7.ForeColor = vbRed
        Combinations_O = "You lost"
    ElseIf lbl1.Caption = "O" And lbl4.Caption = "O" And lbl7.Caption = "O" Then
        lbl1.ForeColor = vbRed: lbl4.ForeColor = vbRed: lbl7.ForeColor = vbRed
        Combinations_O = "You lost"
    ElseIf lbl2.Caption = "O" And lbl5.Caption = "O" And lbl8.Caption = "O" Then
        lbl2.ForeColor = vbRed: lbl5.ForeColor = vbRed: lbl8.ForeColor = vbRed
        Combinations_O = "You lost"
    ElseIf lbl3.Caption = "O" And lbl6.Caption = "O" And lbl9.Caption = "O" Then
        lbl3.ForeColor = vbRed: lbl6.ForeColor = vbRed: lbl9.ForeColor = vbRed
        Combinations_O = "You lost"
    End If
    If Combinations_O = "You lost" Then
        intResult = MsgBox("Game Over: You Lose!" & vbCrLf & vbCrLf & "Start a new game?", vbExclamation + vbYesNo, "Game")
        If intResult = vbYes Then
            Call ResetGame
        End If
    End If
End Function

Public Sub ComputerMove()
'Moves for the computer.
    Dim intMove As Integer
DoAgain:
    'Randomize which number to go to.
    Randomize
    intMove = Int(Rnd * 9) + 1
    If intMove = 1 Then
        'If that area is full, go back and do it again.
        If Len(lbl1.Caption) = 1 Then
            GoTo DoAgain
        End If
        'Set the caption to O.
        lbl1.Caption = "O"
        'If the computer wins, exit the sub.
        If Combinations_O = "You lost" Then
            Exit Sub
        End If
        'Notify you that its the user's turn.
        lblStatus.Caption = "Status: Your Turn"
    ElseIf intMove = 2 Then
        If Len(lbl2.Caption) = 1 Then
            GoTo DoAgain
        End If
        lbl2.Caption = "O"
        If Combinations_O = "You lost" Then
            Exit Sub
        End If
        lblStatus.Caption = "Status: Your Turn"
    ElseIf intMove = 3 Then
        If Len(lbl3.Caption) = 1 Then
            GoTo DoAgain
        End If
        lbl3.Caption = "O"
        If Combinations_O = "You lost" Then
            Exit Sub
        End If
        lblStatus.Caption = "Status: Your Turn"
    ElseIf intMove = 4 Then
        If Len(lbl4.Caption) = 1 Then
            GoTo DoAgain
        End If
        lbl4.Caption = "O"
        If Combinations_O = "You lost" Then
            Exit Sub
        End If
        lblStatus.Caption = "Status: Your Turn"
    ElseIf intMove = 5 Then
        If Len(lbl5.Caption) = 1 Then
            GoTo DoAgain
        End If
        lbl5.Caption = "O"
        If Combinations_O = "You lost" Then
            Exit Sub
        End If
        lblStatus.Caption = "Status: Your Turn"
    ElseIf intMove = 6 Then
        If Len(lbl6.Caption) = 1 Then
            GoTo DoAgain
        End If
        lbl6.Caption = "O"
        If Combinations_O = "You lost" Then
            Exit Sub
        End If
        lblStatus.Caption = "Status: Your Turn"
    ElseIf intMove = 7 Then
        If Len(lbl7.Caption) = 1 Then
            GoTo DoAgain
        End If
        lbl7.Caption = "O"
        If Combinations_O = "You lost" Then
            Exit Sub
        End If
        lblStatus.Caption = "Status: Your Turn"
    ElseIf intMove = 8 Then
        If Len(lbl8.Caption) = 1 Then
            GoTo DoAgain
        End If
        lbl8.Caption = "O"
        If Combinations_O = "You lost" Then
            Exit Sub
        End If
        lblStatus.Caption = "Status: Your Turn"
    ElseIf intMove = 9 Then
        If Len(lbl9.Caption) = 1 Then
            GoTo DoAgain
        End If
        lbl9.Caption = "O"
        If Combinations_O = "You lost" Then
            Exit Sub
        End If
        lblStatus.Caption = "Status: Your Turn"
    End If
End Sub

Private Sub Form_Load()
'Resets on the form load.
    Call ResetGame
End Sub

Private Sub lbl1_Click()
'Note: I 'm only going to explain lbl1_Click and not the
'rest of them best its all the same thing.

    'If the area is open and your turn...
    If lbl1.Caption = "" And lblStatus.Caption = "Status: Your Turn" Then
        'Sets the caption to X
        lbl1.Caption = "X"
        'If you won or there's a draw, exit the sub.
        If Combinations_X = "You won" Then
            Exit Sub
        End If
        If Draw = True Then
            Exit Sub
        End If
        'Notify that its the computer's turn.
        lblStatus.Caption = "Status: Computer's Turn"
    End If
End Sub

Private Sub lbl2_Click()
    If lbl2.Caption = "" And lblStatus.Caption = "Status: Your Turn" Then
        lbl2.Caption = "X"
        If Combinations_X = "You won" Then
            Exit Sub
        End If
        If Draw = True Then
            Exit Sub
        End If
        lblStatus.Caption = "Status: Computer's Turn"
    End If
End Sub

Private Sub lbl3_Click()
    If lbl3.Caption = "" And lblStatus.Caption = "Status: Your Turn" Then
        lbl3.Caption = "X"
        If Combinations_X = "You won" Then
            Exit Sub
        End If
        If Draw = True Then
            Exit Sub
        End If
        lblStatus.Caption = "Status: Computer's Turn"
    End If
End Sub

Private Sub lbl4_Click()
    If lbl4.Caption = "" And lblStatus.Caption = "Status: Your Turn" Then
        lbl4.Caption = "X"
        If Combinations_X = "You won" Then
            Exit Sub
        End If
        If Draw = True Then
            Exit Sub
        End If
        lblStatus.Caption = "Status: Computer's Turn"
    End If
End Sub

Private Sub lbl5_Click()
    If lbl5.Caption = "" And lblStatus.Caption = "Status: Your Turn" Then
        lbl5.Caption = "X"
        If Combinations_X = "You won" Then
            Exit Sub
        End If
        If Draw = True Then
            Exit Sub
        End If
        lblStatus.Caption = "Status: Computer's Turn"
    End If
End Sub

Private Sub lbl6_Click()
    If lbl6.Caption = "" And lblStatus.Caption = "Status: Your Turn" Then
        lbl6.Caption = "X"
        If Combinations_X = "You won" Then
            Exit Sub
        End If
        If Draw = True Then
            Exit Sub
        End If
        lblStatus.Caption = "Status: Computer's Turn"
    End If
End Sub

Private Sub lbl7_Click()
    If lbl7.Caption = "" And lblStatus.Caption = "Status: Your Turn" Then
        lbl7.Caption = "X"
        If Combinations_X = "You won" Then
            Exit Sub
        End If
        If Draw = True Then
            Exit Sub
        End If
        lblStatus.Caption = "Status: Computer's Turn"
    End If
End Sub

Private Sub lbl8_Click()
    If lbl8.Caption = "" And lblStatus.Caption = "Status: Your Turn" Then
        lbl8.Caption = "X"
        If Combinations_X = "You won" Then
            Exit Sub
        End If
        If Draw = True Then
            Exit Sub
        End If
        lblStatus.Caption = "Status: Computer's Turn"
    End If
End Sub

Private Sub lbl9_Click()
    If lbl9.Caption = "" And lblStatus.Caption = "Status: Your Turn" Then
        lbl9.Caption = "X"
        If Combinations_X = "You won" Or Draw = True Then
            Exit Sub
        End If
        If Draw = True Then
            Exit Sub
        End If
        lblStatus.Caption = "Status: Computer's Turn"
    End If
End Sub

Private Sub lblStatus_DblClick()
    'If the status bar is double clicked, it asks if they
    'want to restart.
    Dim intResult As Integer
    intResult = MsgBox("Are you sure you want to start a new game?", vbExclamation + vbYesNo, "Game")
    If intResult = vbYes Then
        Call ResetGame
    End If
End Sub

Private Sub tmrTurn_Timer()
    'This timer is only here to keep the status caption
    'at Computer's Turn.
    If lblStatus.Caption = "Status: Computer's Turn" Then
        'Pause for 2 seconds.
        Sleep (2000)
        'Lets the computer move.
        Call ComputerMove
    End If
End Sub
