VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Form1"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8880
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7485
   ScaleWidth      =   8880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "EXIT"
      Height          =   495
      Left            =   6000
      TabIndex        =   11
      Top             =   3000
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OPTIONS"
      Height          =   495
      Left            =   6000
      TabIndex        =   10
      Top             =   2280
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "RESET"
      Height          =   495
      Left            =   6000
      TabIndex        =   9
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Line ln8 
      BorderWidth     =   5
      X1              =   600
      X2              =   4920
      Y1              =   5760
      Y2              =   1440
   End
   Begin VB.Line ln7 
      BorderWidth     =   5
      X1              =   600
      X2              =   4920
      Y1              =   1440
      Y2              =   5760
   End
   Begin VB.Line ln6 
      BorderWidth     =   5
      X1              =   600
      X2              =   4920
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line ln5 
      BorderWidth     =   5
      X1              =   600
      X2              =   4920
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line ln4 
      BorderWidth     =   5
      X1              =   600
      X2              =   4920
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line ln3 
      BorderWidth     =   5
      X1              =   4200
      X2              =   4200
      Y1              =   1440
      Y2              =   5760
   End
   Begin VB.Line ln2 
      BorderWidth     =   5
      X1              =   2760
      X2              =   2760
      Y1              =   1440
      Y2              =   5760
   End
   Begin VB.Line Line8 
      X1              =   600
      X2              =   600
      Y1              =   5760
      Y2              =   1440
   End
   Begin VB.Line ln1 
      BorderWidth     =   5
      X1              =   1320
      X2              =   1320
      Y1              =   5760
      Y2              =   1440
   End
   Begin VB.Line Line7 
      X1              =   600
      X2              =   4920
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line6 
      X1              =   600
      X2              =   4920
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line Line5 
      X1              =   4920
      X2              =   4920
      Y1              =   1440
      Y2              =   5760
   End
   Begin VB.Line Line4 
      X1              =   600
      X2              =   4920
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line3 
      X1              =   600
      X2              =   4920
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line2 
      X1              =   3480
      X2              =   3480
      Y1              =   1440
      Y2              =   5760
   End
   Begin VB.Line Line1 
      X1              =   2040
      X2              =   2040
      Y1              =   1440
      Y2              =   5760
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   50.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1215
      Left            =   3600
      TabIndex        =   8
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   50.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1215
      Left            =   2160
      TabIndex        =   7
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   50.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1215
      Left            =   720
      TabIndex        =   6
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   50.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1215
      Left            =   3600
      TabIndex        =   5
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   50.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1215
      Left            =   2160
      TabIndex        =   4
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   50.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1215
      Left            =   720
      TabIndex        =   3
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   50.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1215
      Left            =   3600
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   50.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1215
      Left            =   2160
      TabIndex        =   1
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF80FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   50.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1215
      Left            =   720
      TabIndex        =   0
      Top             =   1560
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim mov As Integer




Private Sub Command1_Click()
Label1.Caption = ""
Label2.Caption = ""
Label3.Caption = ""
Label4.Caption = ""
Label5.Caption = ""
Label6.Caption = ""
Label7.Caption = ""
Label8.Caption = ""
Label9.Caption = ""
ln1.Visible = False
ln2.Visible = False
ln3.Visible = False
ln4.Visible = False
ln5.Visible = False
ln6.Visible = False
ln7.Visible = False
ln8.Visible = False

End Sub

Private Sub Form_Load()
mov = 0
ln1.Visible = False
ln2.Visible = False
ln3.Visible = False
ln4.Visible = False
ln5.Visible = False
ln6.Visible = False
ln7.Visible = False
ln8.Visible = False

End Sub

Private Sub Label1_Click()
If Label1.Caption = "" And mov = 0 Then
mov = 1
Label1.Caption = "X"
ElseIf Label1.Caption = "" And mov = 1 Then
mov = 0
Label1.Caption = "O"
End If
If Label1.Caption = "X" Then
    If Label2.Caption = "X" And Label3.Caption = "X" Then
    ln4.Visible = True
    MsgBox ("X WINS!!!")
    ElseIf Label4.Caption = "X" And Label7.Caption = "X" Then
    MsgBox ("X WINS!!!")
    ElseIf Label5.Caption = "X" And Label9.Caption = "X" Then
    MsgBox ("X WINS!!!")
    End If
ElseIf Label1.Caption = "O" Then
    If Label2.Caption = "O" And Label3.Caption = "O" Then
    MsgBox ("O WINS!!!")
    ElseIf Label4.Caption = "O" And Label7.Caption = "O" Then
    MsgBox ("O WINS!!!")
    ElseIf Label5.Caption = "O" And Label9.Caption = "O" Then
    MsgBox ("O WINS!!!")
    End If
End If
End Sub

Private Sub Label2_Click()
If Label2.Caption = "" And mov = 0 Then
mov = 1
Label2.Caption = "X"
ElseIf Label2.Caption = "" And mov = 1 Then
mov = 0
Label2.Caption = "O"
End If
If Label2.Caption = "X" Then
    If Label1.Caption = "X" And Label3.Caption = "X" Then
    MsgBox ("X WINS!!!")
    ElseIf Label5.Caption = "X" And Label8.Caption = "X" Then
    MsgBox ("X WINS!!!")
    End If
ElseIf Label2.Caption = "O" Then
    If Label1.Caption = "O" And Label3.Caption = "O" Then
    MsgBox ("O WINS!!!")
    ElseIf Label5.Caption = "O" And Label8.Caption = "O" Then
    MsgBox ("O WINS!!!")
    End If
End If

End Sub

Private Sub Label3_Click()
If Label3.Caption = "" And mov = 0 Then
mov = 1
Label3.Caption = "X"
ElseIf Label3.Caption = "" And mov = 1 Then
mov = 0
Label3.Caption = "O"
End If
If Label3.Caption = "X" Then
    If Label2.Caption = "X" And Label1.Caption = "X" Then
    MsgBox ("X WINS!!!")
    ElseIf Label5.Caption = "X" And Label7.Caption = "X" Then
    MsgBox ("X WINS!!!")
    ElseIf Label6.Caption = "X" And Label9.Caption = "X" Then
    MsgBox ("X WINS!!!")
    End If
ElseIf Label3.Caption = "O" Then
    If Label2.Caption = "O" And Label1.Caption = "O" Then
    MsgBox ("O WINS!!!")
    ElseIf Label5.Caption = "O" And Label7.Caption = "O" Then
    MsgBox ("O WINS!!!")
    ElseIf Label6.Caption = "O" And Label9.Caption = "O" Then
    MsgBox ("O WINS!!!")
    End If
End If

End Sub

Private Sub Label4_Click()
If Label4.Caption = "" And mov = 0 Then
mov = 1
Label4.Caption = "X"
ElseIf Label4.Caption = "" And mov = 1 Then
mov = 0
Label4.Caption = "O"
End If
If Label4.Caption = "X" Then
    If Label1.Caption = "X" And Label7.Caption = "X" Then
    MsgBox ("X WINS!!!")
    ElseIf Label5.Caption = "X" And Label6.Caption = "X" Then
    MsgBox ("X WINS!!!")
    End If
ElseIf Label4.Caption = "O" Then
    If Label1.Caption = "O" And Label7.Caption = "O" Then
    MsgBox ("O WINS!!!")
    ElseIf Label5.Caption = "O" And Label6.Caption = "O" Then
    MsgBox ("O WINS!!!")
    End If
End If

End Sub

Private Sub Label5_Click()
If Label5.Caption = "" And mov = 0 Then
mov = 1
Label5.Caption = "X"
ElseIf Label5.Caption = "" And mov = 1 Then
mov = 0
Label5.Caption = "O"
End If
If Label5.Caption = "X" Then
    If Label1.Caption = "X" And Label9.Caption = "X" Then
    MsgBox ("X WINS!!!")
    ElseIf Label4.Caption = "X" And Label6.Caption = "X" Then
    MsgBox ("X WINS!!!")
    ElseIf Label2.Caption = "X" And Label8.Caption = "X" Then
    MsgBox ("X WINS!!!")
    ElseIf Label3.Caption = "X" And Label7.Caption = "X" Then
    MsgBox ("X WINS!!!")
    End If
ElseIf Label5.Caption = "O" Then
    If Label1.Caption = "O" And Label9.Caption = "O" Then
    MsgBox ("O WINS!!!")
    ElseIf Label4.Caption = "O" And Label6.Caption = "O" Then
    MsgBox ("O WINS!!!")
    ElseIf Label2.Caption = "O" And Label8.Caption = "O" Then
    MsgBox ("O WINS!!!")
    ElseIf Label3.Caption = "O" And Label7.Caption = "O" Then
    MsgBox ("O WINS!!!")
    End If
End If

End Sub

Private Sub Label6_Click()
If Label6.Caption = "" And mov = 0 Then
mov = 1
Label6.Caption = "X"
ElseIf Label6.Caption = "" And mov = 1 Then
mov = 0
Label6.Caption = "O"
End If
If Label6.Caption = "X" Then
    If Label3.Caption = "X" And Label9.Caption = "X" Then
    MsgBox ("X WINS!!!")
    ElseIf Label5.Caption = "X" And Label4.Caption = "X" Then
    MsgBox ("X WINS!!!")
    End If
ElseIf Label6.Caption = "O" Then
    If Label3.Caption = "O" And Label9.Caption = "O" Then
    MsgBox ("O WINS!!!")
    ElseIf Label5.Caption = "O" And Label4.Caption = "O" Then
    MsgBox ("O WINS!!!")
    End If
End If

End Sub

Private Sub Label7_Click()
If Label7.Caption = "" And mov = 0 Then
mov = 1
Label7.Caption = "X"
ElseIf Label7.Caption = "" And mov = 1 Then
mov = 0
Label7.Caption = "O"
End If
If Label7.Caption = "X" Then
    If Label4.Caption = "X" And Label1.Caption = "X" Then
    MsgBox ("X WINS!!!")
    ElseIf Label5.Caption = "X" And Label3.Caption = "X" Then
    MsgBox ("X WINS!!!")
    ElseIf Label8.Caption = "X" And Label9.Caption = "X" Then
    MsgBox ("X WINS!!!")
    End If
ElseIf Label7.Caption = "O" Then
    If Label4.Caption = "O" And Label1.Caption = "O" Then
    MsgBox ("O WINS!!!")
    ElseIf Label5.Caption = "O" And Label3.Caption = "O" Then
    MsgBox ("O WINS!!!")
    ElseIf Label8.Caption = "O" And Label9.Caption = "O" Then
    MsgBox ("O WINS!!!")
    End If
End If

End Sub

Private Sub Label8_Click()
If Label8.Caption = "" And mov = 0 Then
mov = 1
Label8.Caption = "X"
ElseIf Label8.Caption = "" And mov = 1 Then
mov = 0
Label8.Caption = "O"
End If
If Label8.Caption = "X" Then
    If Label2.Caption = "X" And Label5.Caption = "X" Then
    MsgBox ("X WINS!!!")
    ElseIf Label7.Caption = "X" And Label9.Caption = "X" Then
    MsgBox ("X WINS!!!")
    End If
ElseIf Label8.Caption = "O" Then
    If Label2.Caption = "O" And Label5.Caption = "O" Then
    MsgBox ("O WINS!!!")
    ElseIf Label7.Caption = "O" And Label9.Caption = "O" Then
    MsgBox ("O WINS!!!")
    End If
End If

End Sub

Private Sub Label9_Click()
If Label9.Caption = "" And mov = 0 Then
mov = 1
Label9.Caption = "X"
ElseIf Label9.Caption = "" And mov = 1 Then
mov = 0
Label9.Caption = "O"
End If
If Label9.Caption = "X" Then
    If Label5.Caption = "X" And Label1.Caption = "X" Then
    MsgBox ("X WINS!!!")
    ElseIf Label8.Caption = "X" And Label7.Caption = "X" Then
    MsgBox ("X WINS!!!")
    ElseIf Label6.Caption = "X" And Label3.Caption = "X" Then
    MsgBox ("X WINS!!!")
    End If
ElseIf Label9.Caption = "O" Then
    If Label5.Caption = "O" And Label1.Caption = "O" Then
    MsgBox ("O WINS!!!")
    ElseIf Label8.Caption = "O" And Label7.Caption = "O" Then
    MsgBox ("O WINS!!!")
    ElseIf Label6.Caption = "O" And Label3.Caption = "O" Then
    MsgBox ("O WINS!!!")
    End If
End If

End Sub
