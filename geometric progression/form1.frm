VERSION 5.00
Begin VB.Form Txt_FirstNum 
   BackColor       =   &H00404040&
   Caption         =   "Geometric Progression"
   ClientHeight    =   7335
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   ScaleHeight     =   7335
   ScaleWidth      =   9600
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Cmd_exit 
      BackColor       =   &H000000FF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5520
      Width           =   1815
   End
   Begin VB.CommandButton cmd_compute 
      BackColor       =   &H0080FF80&
      Caption         =   "Compute"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5520
      Width           =   1575
   End
   Begin VB.TextBox Txt_Terms 
      BackColor       =   &H8000000A&
      ForeColor       =   &H80000005&
      Height          =   735
      Left            =   2280
      TabIndex        =   6
      Top             =   3960
      Width           =   1815
   End
   Begin VB.TextBox Txt_CR 
      BackColor       =   &H80000018&
      Height          =   615
      Left            =   2280
      TabIndex        =   5
      Top             =   2520
      Width           =   1815
   End
   Begin VB.TextBox Txt_FirstNum 
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   2280
      TabIndex        =   4
      Top             =   1200
      Width           =   1815
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6540
      Left            =   4680
      TabIndex        =   0
      Top             =   480
      Width           =   4695
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000007&
      Caption         =   "Number of Terms"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Caption         =   "Common Ratio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "First Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
End
Attribute VB_Name = "Txt_FirstNum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmd_compute_Click()
Dim x, n, num As Integer
Dim r As Single
x = Txt_FirstNum.Text
r = Txt_CR
num = Txt_Terms.Text
List1.AddItem "n" & vbTab & "x"
List1.AddItem "___________"

n = 1
Do
x = x * r
List1.AddItem n & vbTab & x
n = n + 1
Loop Until n = num + 1

End Sub







Private Sub Cmd_exit_Click()
End
End Sub
