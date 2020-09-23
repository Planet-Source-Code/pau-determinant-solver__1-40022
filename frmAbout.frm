VERSION 5.00
Begin VB.Form frmAbout 
   Caption         =   "Author"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6150
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   6150
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Accept"
      Height          =   375
      Left            =   4200
      TabIndex        =   0
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   2625
      Left            =   120
      Picture         =   "frmAbout.frx":030A
      Stretch         =   -1  'True
      Top             =   240
      Width           =   3240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PAU JANSÀ PADRÓ"
      Height          =   195
      Left            =   4080
      TabIndex        =   5
      Top             =   360
      Width           =   1500
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vila-seca (tarragona)"
      Height          =   195
      Left            =   4080
      TabIndex        =   4
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2002 - 2003"
      Height          =   195
      Left            =   4320
      TabIndex        =   3
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(25 - 03 - 1984)"
      Height          =   195
      Left            =   4200
      TabIndex        =   2
      Top             =   600
      Width           =   1080
   End
   Begin VB.Label lblMailTo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "mailto:lambdero18@hotmail.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   3720
      MouseIcon       =   "frmAbout.frx":6F71
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2640
      Width           =   2280
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub lblMailTo_Click()
ShellExecute Me.hwnd, "open", "mailto:lambdero18@hotmail.com", vbNullString, "C:\", 5
End Sub
