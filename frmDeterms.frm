VERSION 5.00
Begin VB.Form frmDeterms 
   Caption         =   "Determinants Solver"
   ClientHeight    =   5040
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7845
   Icon            =   "frmDeterms.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   7845
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Author"
      Height          =   255
      Left            =   4440
      TabIndex        =   45
      Top             =   4560
      Width           =   2895
   End
   Begin VB.TextBox txt_16 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   6480
      TabIndex        =   36
      Text            =   "0"
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox txt_12 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   6480
      TabIndex        =   35
      Text            =   "0"
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox txt_8 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   6480
      TabIndex        =   34
      Text            =   "0"
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox txt_4 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   6480
      TabIndex        =   33
      Text            =   "0"
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox txt_13 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   4680
      TabIndex        =   32
      Text            =   "0"
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox txt_15 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   5880
      TabIndex        =   31
      Text            =   "0"
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox txt_14 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   5280
      TabIndex        =   30
      Text            =   "0"
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox txt_9 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   4680
      TabIndex        =   29
      Text            =   "0"
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox txt_5 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   4680
      TabIndex        =   28
      Text            =   "0"
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox txt_1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   4680
      TabIndex        =   27
      Text            =   "0"
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox txt_11 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   5880
      TabIndex        =   26
      Text            =   "0"
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox txt_10 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   5280
      TabIndex        =   25
      Text            =   "0"
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox txt_7 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   5880
      TabIndex        =   24
      Text            =   "0"
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox txt_6 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   5280
      TabIndex        =   23
      Text            =   "0"
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox txt_3 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   5880
      TabIndex        =   22
      Text            =   "0"
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox txt_2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   5280
      TabIndex        =   21
      Text            =   "0"
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox txtTotal2x2 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2760
      TabIndex        =   20
      Text            =   "0"
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton CmdCalc2x2 
      Caption         =   "Solve"
      Enabled         =   0   'False
      Height          =   375
      Left            =   480
      TabIndex        =   19
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1080
      TabIndex        =   18
      Text            =   "0"
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      TabIndex        =   17
      Text            =   "0"
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1080
      TabIndex        =   16
      Text            =   "0"
      Top             =   3360
      Width           =   615
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      TabIndex        =   15
      Text            =   "0"
      Top             =   3360
      Width           =   615
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1080
      TabIndex        =   14
      Text            =   "0"
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      TabIndex        =   13
      Text            =   "0"
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox txtTotal3x3 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2760
      TabIndex        =   12
      Text            =   "0"
      Top             =   3960
      Width           =   495
   End
   Begin VB.CommandButton cmdCalc3x3 
      Caption         =   "Solve"
      Enabled         =   0   'False
      Height          =   375
      Left            =   480
      TabIndex        =   11
      Top             =   4080
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   480
      TabIndex        =   10
      Text            =   "0"
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   480
      TabIndex        =   9
      Text            =   "0"
      Top             =   3360
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   480
      TabIndex        =   8
      Text            =   "0"
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox txt4 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1320
      TabIndex        =   7
      Text            =   "0"
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txt3 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   480
      TabIndex        =   6
      Text            =   "0"
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox txt2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Text            =   "0"
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox txt1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   480
      TabIndex        =   4
      Text            =   "0"
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton cmdCls3x3 
      Caption         =   "Clear"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton cmdCls2x2 
      Caption         =   "Clear"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   1440
      Width           =   1575
   End
   Begin VB.ComboBox cmbDimension 
      Height          =   315
      ItemData        =   "frmDeterms.frx":030A
      Left            =   3000
      List            =   "frmDeterms.frx":0317
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Cl&ose"
      Height          =   375
      Left            =   4440
      TabIndex        =   0
      Top             =   360
      Width           =   2895
   End
   Begin VB.Frame Frame4x4 
      Caption         =   "Determinant 4 x 4"
      Enabled         =   0   'False
      Height          =   3495
      Left            =   4440
      TabIndex        =   38
      Top             =   960
      Width           =   2895
      Begin VB.TextBox txtTotal4x4 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   840
         TabIndex        =   41
         Text            =   "0"
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton cmdCls4x4 
         Caption         =   "Clear"
         Enabled         =   0   'False
         Height          =   375
         Left            =   600
         TabIndex        =   40
         Top             =   1560
         Width           =   1575
      End
      Begin VB.CommandButton cmdCalc4x4 
         Caption         =   "Solve"
         Enabled         =   0   'False
         Height          =   375
         Left            =   600
         TabIndex        =   39
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   1200
         X2              =   1440
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   1200
         X2              =   1440
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Label lblResult4x4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   600
         TabIndex        =   42
         Top             =   1320
         Width           =   45
      End
   End
   Begin VB.Frame Frame3x3 
      Caption         =   "Determinantt3 x 3"
      Enabled         =   0   'False
      Height          =   1935
      Left            =   240
      TabIndex        =   37
      Top             =   2880
      Width           =   3975
      Begin VB.Line Line1 
         Index           =   1
         X1              =   2160
         X2              =   2400
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line3 
         X1              =   2160
         X2              =   2400
         Y1              =   1320
         Y2              =   1320
      End
   End
   Begin VB.Frame Frame2x2 
      Caption         =   "Determinant 2 x 2"
      Enabled         =   0   'False
      Height          =   1695
      Left            =   240
      TabIndex        =   43
      Top             =   960
      Width           =   3975
      Begin VB.Line Line1 
         Index           =   0
         X1              =   2160
         X2              =   2400
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line2 
         X1              =   2160
         X2              =   2400
         Y1              =   1320
         Y2              =   1320
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Introduce the determinant dimension:"
      Height          =   195
      Left            =   240
      TabIndex        =   44
      Top             =   360
      Width           =   2610
   End
End
Attribute VB_Name = "frmDeterms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
    frmAbout.Show
End Sub
Private Sub cmbDimension_Click()
    Select Case cmbDimension.List(cmbDimension.ListIndex)
    
        Case "2x2"
            CmdCalc2x2.Enabled = True '''''''
            Frame2x2.Enabled = True         '
            txt1.Enabled = True             '
            txt2.Enabled = True             'Enable True Determinante 2x2
            txt3.Enabled = True             '
            txt4.Enabled = True             '
            cmdCls2x2.Enabled = True        '
            txtTotal2x2.Enabled = True ''''''
            
            
            Text1.Enabled = False '''''''''''
            Text2.Enabled = False           '
            Frame3x3.Enabled = False        '
            Text3.Enabled = False           '
            Text4.Enabled = False           '
            Text5.Enabled = False           '
            Text6.Enabled = False           'Enable False Determinante 3x3
            Text7.Enabled = False           '
            Text8.Enabled = False           '
            Text9.Enabled = False           '
            cmdCalc3x3.Enabled = False      '
            cmdCls3x3.Enabled = False       '
            txtTotal3x3.Enabled = False     '
            
            
            txt_1.Enabled = False '''''''''''
            txt_2.Enabled = False           '
            txt_3.Enabled = False           '
            txt_4.Enabled = False           '
            txt_5.Enabled = False           '
            txt_6.Enabled = False           '
            txt_7.Enabled = False           '
            txt_8.Enabled = False           '
            txt_9.Enabled = False           'Enable False Determinante 4x4
            txt_10.Enabled = False          '
            txt_11.Enabled = False          '
            txt_12.Enabled = False          '
            txt_13.Enabled = False          '
            txt_14.Enabled = False          '
            txt_15.Enabled = False          '
            txt_16.Enabled = False          '
            cmdCalc4x4.Enabled = False      '
            cmdCls4x4.Enabled = False       '
            txtTotal4x4.Enabled = False     '
            Frame4x4.Enabled = False        '
            
        Case "3x3"
            CmdCalc2x2.Enabled = False '''''''
            Frame2x2.Enabled = False         '
            txt1.Enabled = False             '
            txt2.Enabled = False             'Enable False Determinante 2x2
            txt3.Enabled = False             '
            txt4.Enabled = False             '
            cmdCls2x2.Enabled = False        '
            txtTotal2x2.Enabled = False ''''''
            
            
            Text1.Enabled = True '''''''''''
            Text2.Enabled = True           '
            Text3.Enabled = True           '
            Text4.Enabled = True           '
            Text5.Enabled = True           '
            Text6.Enabled = True           'Enable True Determinante 3x3
            Text7.Enabled = True           '
            Text8.Enabled = True           '
            Text9.Enabled = True           '
            cmdCalc3x3.Enabled = True      '
            cmdCls3x3.Enabled = True       '
            txtTotal3x3.Enabled = True     '
            Frame3x3.Enabled = True        '
            
            
            txt_1.Enabled = False '''''''''''
            txt_2.Enabled = False           '
            txt_3.Enabled = False           '
            txt_4.Enabled = False           '
            txt_5.Enabled = False           '
            txt_6.Enabled = False           '
            txt_7.Enabled = False           '
            txt_8.Enabled = False           '
            txt_9.Enabled = False           'Enable False Determinante 4x4
            txt_10.Enabled = False          '
            txt_11.Enabled = False          '
            txt_12.Enabled = False          '
            txt_13.Enabled = False          '
            txt_14.Enabled = False          '
            txt_15.Enabled = False          '
            txt_16.Enabled = False          '
            cmdCalc4x4.Enabled = False      '
            cmdCls4x4.Enabled = False       '
            txtTotal4x4.Enabled = False     '
            Frame4x4.Enabled = False        '
            
        Case "4x4"
            CmdCalc2x2.Enabled = False '''''''
            Frame2x2.Enabled = False         '
            txt1.Enabled = False             '
            txt2.Enabled = False             'Enable False Determinante 2x2
            txt3.Enabled = False             '
            txt4.Enabled = False             '
            cmdCls2x2.Enabled = False        '
            txtTotal2x2.Enabled = False ''''''
            
            
            Text1.Enabled = False '''''''''''
            Text2.Enabled = False           '
            Text3.Enabled = False           '
            Text4.Enabled = False           '
            Text5.Enabled = False           '
            Text6.Enabled = False           'Enable False Determinante 3x3
            Text7.Enabled = False           '
            Text8.Enabled = False           '
            Text9.Enabled = False           '
            cmdCalc3x3.Enabled = False      '
            cmdCls3x3.Enabled = False       '
            txtTotal3x3.Enabled = False     '
            Frame3x3.Enabled = False        '
            
            
            txt_1.Enabled = True '''''''''''
            txt_2.Enabled = True           '
            txt_3.Enabled = True           '
            txt_4.Enabled = True           '
            txt_5.Enabled = True           '
            txt_6.Enabled = True           '
            txt_7.Enabled = True           '
            txt_8.Enabled = True           '
            txt_9.Enabled = True           'Enable True Determinante 4x4
            txt_10.Enabled = True          '
            txt_11.Enabled = True          '
            txt_12.Enabled = True          '
            txt_13.Enabled = True          '
            txt_14.Enabled = True          '
            txt_15.Enabled = True          '
            txt_16.Enabled = True          '
            cmdCalc4x4.Enabled = True      '
            cmdCls4x4.Enabled = True       '
            txtTotal4x4.Enabled = True     '
            Frame4x4.Enabled = True ''''''''
    End Select
End Sub

Private Sub cmdCalc4x4_Click()
        txtTotal4x4.Text = Determinante4x4(Val(txt_1), Val(txt_2), Val(txt_3), Val(txt_4), Val(txt_5), Val(txt_6), Val(txt_7), Val(txt_8), _
        Val(txt_9), Val(txt_10), Val(txt_11), Val(txt_12), Val(txt_13), Val(txt_14), Val(txt_15), Val(txt_16))
End Sub
Private Sub CmdCalc2x2_Click()
txtTotal2x2.Text = Determinante2x2(Val(txt1), Val(txt2), Val(txt3), Val(txt4))
End Sub

Private Sub cmdCalc3x3_Click()
    txtTotal3x3.Text = Determinante3x3(Val(Text1), Val(Text2), Val(Text3), Val(Text4), Val(Text5), Val(Text6), Val(Text7), Val(Text8), Val(Text9))

End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdCls2x2_click()
    txtTotal2x2.Text = "0"
    txt1.Text = "0"
    txt2.Text = "0"
    txt3.Text = "0"
    txt4.Text = "0"
    lblResult2x2.Caption = ""
End Sub
Private Sub cmdCls3x3_Click()
    txtTotal3x3.Text = "0"
    Text1.Text = "0"
    Text2.Text = "0"
    Text3.Text = "0"
    Text4.Text = "0"
    Text5.Text = "0"
    Text6.Text = "0"
    Text7.Text = "0"
    Text8.Text = "0"
    Text9.Text = "0"
End Sub

Private Sub cmdCls4x4_Click()
    txtTotal4x4.Text = "0"
    txt_1 = "0"
    txt_2 = "0"
    txt_3 = "0"
    txt_4 = "0"
    txt_5 = "0"
    txt_6 = "0"
    txt_7 = "0"
    txt_8 = "0"
    txt_9 = "0"
    txt_10 = "0"
    txt_11 = "0"
    txt_12 = "0"
    txt_13 = "0"
    txt_14 = "0"
    txt_15 = "0"
    txt_16 = "0"
End Sub

Private Sub Intro2x2_Click()
    With Form2
        .List1.AddItem "(# SOL" & Space$(1) & "[" & "Det2x2(" & CStr(txt1) & ", " & CStr(txt2) & "; " & CStr(txt3) & ", " & CStr(txt4) & ")" & "]" & Space$(1) & ":) = " & txtTotal2x2.Text
    End With
End Sub

Private Sub Intro3x3_Click()
    With Form2
        .List1.AddItem "(# SOL" & Space$(1) & "[" & "Det3x3(" & CStr(Text1) & ", " & CStr(Text2) & ", " & CStr(Text3) & "; " & CStr(Text4) & ", " & CStr(Text5) & ", " & CStr(Text6) & "; " & CStr(Text7) & ", " & CStr(Text8) & ", " & CStr(Text9) & ")" & "]" & Space$(1) & ":) = " & txtTotal3x3.Text
    End With
End Sub
Private Sub Intro4x4_Click()
    With Form2
        .List1.AddItem "(# SOL" & Space$(1) & "[" & "Det4x4(" & CStr(txt_1) & ", " & CStr(txt_2) & ", " & CStr(txt_3) & ", " & CStr(txt_4) & "; " & CStr(txt_5) & ", " & CStr(txt_6) & ", " & CStr(txt_7) & ", " & CStr(txt_8) & "; " & CStr(txt_9) & ", " & CStr(txt_10) & ", " & CStr(txt_11) & ", " & CStr(txt_12) & "; " & CStr(txt_13) & ", " & CStr(txt_14) & ", " & CStr(txt_15) & ", " & CStr(txt_16) & ")" & "]" & Space$(1) & ":) = " & txtTotal4x4.Text
    End With
End Sub



Private Sub Form_Unload(Cancel As Integer)
    MsgBox "Vote me please!!, thanks!!  ;-)", vbExclamation, App.Title
End Sub
