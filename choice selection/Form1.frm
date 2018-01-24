VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0000FFFF&
   Caption         =   "Form1"
   ClientHeight    =   3765
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   ScaleHeight     =   3765
   ScaleWidth      =   6105
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   4680
      TabIndex        =   7
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CheckBox Check3 
      Height          =   615
      Left            =   3000
      TabIndex        =   3
      Top             =   2280
      Width           =   255
   End
   Begin VB.CheckBox Check2 
      Height          =   735
      Left            =   3000
      TabIndex        =   2
      Top             =   1560
      Width           =   315
   End
   Begin VB.CheckBox Check1 
      Height          =   615
      Left            =   3000
      TabIndex        =   1
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "Sports"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Computer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Reading"
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
      Left            =   840
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5760
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Choice Selection Program"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   5295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Check1.Value = vbChecked And Check2.Value = vbChecked And Check3.Value = vbChecked Then
MsgBox ("You like Reading, Computer and Sports")

ElseIf Check1.Value = vbChecked And Check2.Value = vbChecked And Check3.Value = vbUnchecked Then

MsgBox ("You like Reading and Computer")

ElseIf Check1.Value = vbChecked And Check2.Value = vbUnchecked And Check3.Value = vbChecked Then
MsgBox ("You like Reading and Sports")

ElseIf Check1.Value = vbUnchecked And Check2.Value = vbChecked And Check3.Value = vbChecked Then
MsgBox ("You like Computer and Sports")

ElseIf Check1.Value = vbChecked And Check2.Value = vbUnchecked And Check3.Value = vbChecked Then
MsgBox ("You like Reading and Sports")

ElseIf Check1.Value = vbChecked And Check2.Value = vbUnchecked And Check3.Value = vbUnchecked Then
MsgBox ("You like Reading only ")

ElseIf Check1.Value = vbUnchecked And Check2.Value = vbChecked And Check3.Value = vbUnchecked Then
MsgBox ("You like computer only")

ElseIf Check1.Value = vbUnchecked And Check2.Value = vbUnchecked And Check3.Value = vbChecked Then
MsgBox ("You like Sports only")

Else
MsgBox ("You have no hobby")
End If
End Sub
