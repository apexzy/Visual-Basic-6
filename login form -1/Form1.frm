VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "EXIT"
      Height          =   615
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C000&
      Caption         =   "LOGIN"
      Height          =   615
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "PASSWORD"
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
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "USERNAME"
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
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim username As String
Dim password As String
username = "admin"
password = "admin"
If (username = Text1.Text And password = Text2.Text) Then
MsgBox "login successful......"
Unload Me
Else
MsgBox "sorry..... login failed....try again...."
End If
End Sub

Private Sub Command2_Click()
Unload Me
End
End Sub
