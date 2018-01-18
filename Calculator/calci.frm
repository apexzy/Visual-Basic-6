VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Calculator"
   ClientHeight    =   6060
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   ScaleHeight     =   6060
   ScaleWidth      =   7455
   Begin VB.CommandButton Command10 
      Caption         =   "MEMW"
      Height          =   495
      Left            =   6120
      TabIndex        =   33
      Top             =   4080
      Width           =   855
   End
   Begin VB.CommandButton Command9 
      Caption         =   "MEMR"
      Height          =   255
      Left            =   6240
      TabIndex        =   32
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Atn"
      Height          =   255
      Index           =   7
      Left            =   6240
      TabIndex        =   31
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "log"
      Height          =   255
      Index           =   6
      Left            =   3600
      TabIndex        =   30
      Top             =   5040
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "rnd"
      Height          =   255
      Index           =   5
      Left            =   3600
      TabIndex        =   29
      Top             =   4320
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "abs"
      Height          =   255
      Index           =   4
      Left            =   3600
      TabIndex        =   28
      Top             =   3600
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "sqrt"
      Height          =   255
      Index           =   3
      Left            =   3600
      TabIndex        =   27
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "tan"
      Height          =   255
      Index           =   2
      Left            =   2880
      TabIndex        =   26
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "cos"
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   25
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "sin"
      Height          =   255
      Index           =   0
      Left            =   1200
      TabIndex        =   24
      Top             =   2400
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   360
      Top             =   360
   End
   Begin VB.CommandButton Command7 
      Caption         =   "On"
      Height          =   255
      Left            =   4080
      TabIndex        =   22
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFF80&
      Caption         =   "off"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton command1 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   10
      Left            =   1200
      TabIndex        =   19
      Top             =   4920
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFF00&
      Caption         =   "c"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H000000FF&
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H000000FF&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H000000FF&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H000000FF&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H000000FF&
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4560
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   11
      Top             =   4920
      Width           =   615
   End
   Begin VB.CommandButton command1 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   2040
      TabIndex        =   10
      Top             =   4920
      Width           =   615
   End
   Begin VB.CommandButton command1 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   2880
      TabIndex        =   9
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton command1 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   2040
      TabIndex        =   8
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton command1 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   1200
      TabIndex        =   7
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton command1 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   2880
      TabIndex        =   6
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton command1 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   2040
      TabIndex        =   5
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton command1 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   1200
      TabIndex        =   4
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton command1 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   2880
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton command1 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   2040
      TabIndex        =   2
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton command1 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   1200
      TabIndex        =   1
      Top             =   2760
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1560
      Width           =   3855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   4320
      TabIndex        =   16
      Top             =   2880
      Width           =   1695
      Begin VB.CommandButton Command2 
         BackColor       =   &H000000FF&
         Caption         =   "-/+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1680
         Width           =   495
      End
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Scientific Calculator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   23
      Top             =   360
      Width           =   5175
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      Height          =   4695
      Left            =   840
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   6255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


' You can use this application in your program or software anyway you like.
' To download more Application programs and softwares in Visual Basic
' visit the website     www.geocities.com/t1softwares


Dim op As String
Option Explicit

Dim i As Integer
Dim exp1 As Double
Dim exp2 As Double
Dim Result As Double
Dim count1 As Integer
Dim scitype As String
Private X As Double
Dim mem As Variant
  


Private Sub Command1_Click(Index As Integer)

If count1 = 0 Then
Text1.Text = " "
MsgBox ("Calculator is not on")

End If

If count1 = 1 Then
Text1.Text = " "
count1 = count1 + 1
End If

If count1 > 1 Then
Text1.Text = Text1.Text & command1(Index).Caption
End If

End Sub

Private Sub Command10_Click()

 Text1.Text = mem
End Sub

Private Sub Command2_Click()
Text1.Text = -Val(Text1.Text)
End Sub

Private Sub Command3_Click()
If count1 > 0 Then
exp2 = Val(Text1.Text)
Select Case (op)
Case "+"
        Result = exp1 + exp2
        Text1.Text = Result
        count1 = 0
Case "-"
        Result = exp1 - exp2
        Text1.Text = Result
        count1 = 0
Case "*"
        Result = exp1 * exp2
        Text1.Text = Result
        count1 = 0
Case "/"
        Result = exp1 / exp2
        Text1.Text = Result
        count1 = 0
Case "%"
        Result = (exp1 / 100) * exp2
        Text1.Text = Result
        count1 = 0
End Select
End If
End Sub



Private Sub Command4_Click(Index As Integer)
Result = exp1
exp1 = Result + Val(Text1.Text)
Text1.Text = " "
op = Command4(Index).Caption
End Sub

Private Sub Command5_Click()
Result = 0
exp1 = 0
exp2 = 0
Text1.Text = " "
count1 = 1
End Sub

Private Sub Command6_Click()
count1 = 0
Text1.Text = ""
End Sub

Private Sub Command7_Click()
Result = 0
exp1 = 0
exp2 = 0
count1 = 1
Text1.Text = "0"
End Sub

Private Sub UpdateLog()
    Trim (Form1.Text1.Text)
    
    
End Sub

Private Sub Command8_Click(Index As Integer)
scitype = Command8(Index).Caption
Select Case (scitype)
Case "sin"
           Text1.Text = (Text1.Text * 3.14) / 180
           Text1.Text = Math.Sin(Val(Text1.Text))
           count1 = 0
Case "cos"
           Text1.Text = (Text1.Text * 3.14) / 180
           Text1.Text = Math.Cos(Val(Text1.Text))
           count1 = 0
Case "tan"
           Text1.Text = (Text1.Text * 3.14) / 180
           Text1.Text = Math.Tan(Val(Text1.Text))
           count1 = 0
Case "sqrt"
           Text1.Text = Math.Sqr(Val(Text1.Text))
           count1 = 0
Case "abs"
           Text1.Text = Math.Abs(Val(Text1.Text))
           count1 = 0
Case "rnd"
           Text1.Text = Math.Rnd(Val(Text1.Text))
           count1 = 0
Case "log"
           Text1.Text = Math.Log(Val(Text1.Text))
           count1 = 0
Case "Atn"
           Text1.Text = Math.Log(Val(Text1.Text))
           count1 = 0

End Select
End Sub

Private Sub Command9_Click()
mem = Text1.Text
End Sub

Private Sub Timer1_Timer()
Label1.BackColor = RGB(256 * Rnd, 256 * Rnd, 256 * Rnd)
End Sub
