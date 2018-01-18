VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BMI Calculator"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Calculate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   2280
      TabIndex        =   2
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Height          =   495
      Left            =   2280
      TabIndex        =   5
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "YOUR BMI="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Weight (in kg)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Height(in meter)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
 Label4.Caption = BMI(Text1.Text, Text2.Text)
End Sub

Private Function BMI(height, weight)
      BMIValue = (weight) / (height ^ 2)
      BMI = Format(BMIValue, "0.00")
End Function
