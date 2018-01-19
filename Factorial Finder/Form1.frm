VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Factor Finding Program"
   ClientHeight    =   3975
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6000
   FillColor       =   &H00FFFFFF&
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
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   6000
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Find Factors"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.CommandButton Command2 
         BackColor       =   &H000000FF&
         Caption         =   "Reset"
         Height          =   615
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2280
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "Find Factors"
         Height          =   735
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   480
         Width           =   1215
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1260
         Left            =   1440
         TabIndex        =   4
         Top             =   2040
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   1440
         TabIndex        =   3
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label Label3 
         Caption         =   "Courtesy of the Open GNL Project, 2017."
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   3720
         Width           =   4335
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000007&
         Caption         =   "List of Factors"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   120
         TabIndex        =   2
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000007&
         Caption         =   "Enter A Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   615
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim N, x As Integer
N = Val(Text1.Text)
For x = 2 To N - 1
If N Mod x = 0 Then
List1.AddItem (x)
End If
Next
List1.AddItem (N)
End Sub

Private Sub Command2_Click()
List1.Clear
Text1 = ""


End Sub

