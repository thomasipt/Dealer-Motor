VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form RP003 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LAPORAN KEUANGAN"
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7125
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   7125
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport crpt 
      Left            =   1050
      Top             =   1050
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cMDeXIT 
      Caption         =   "KELUAR"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2872
      TabIndex        =   0
      Top             =   945
      Width           =   1110
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      Caption         =   "LAPORAN NERACA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   275
      Left            =   67
      TabIndex        =   1
      Top             =   105
      Width           =   3405
   End
   Begin VB.OptionButton Option3 
      Alignment       =   1  'Right Justify
      Caption         =   "LAPORAN RUGI/LABA KUMULATIF"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   275
      Left            =   67
      TabIndex        =   2
      Top             =   465
      Width           =   3405
   End
   Begin VB.OptionButton Option4 
      Alignment       =   1  'Right Justify
      Caption         =   "LAPORAN TRIAL BALANCE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   275
      Left            =   3652
      TabIndex        =   4
      Top             =   465
      Width           =   3405
   End
   Begin VB.OptionButton Option2 
      Alignment       =   1  'Right Justify
      Caption         =   "LAPORAN TRIAL BALANCE HARIAN"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3652
      TabIndex        =   3
      Top             =   105
      Width           =   3405
   End
   Begin VB.Line Line2 
      X1              =   -300
      X2              =   7140
      Y1              =   825
      Y2              =   825
   End
End
Attribute VB_Name = "RP003"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cMDeXIT_Click()
Unload Me
End Sub

Private Sub Form_Load()
If Operator = "SERVICE SPAREPART" Then
    Option3.Caption = "LAPORAN AKUMULASI L/R BENGKEL"
    Me.BackColor = &H80C0FF
    Option1.BackColor = &H80C0FF
    Option2.BackColor = &H80C0FF
    Option3.BackColor = &H80C0FF
    Option4.BackColor = &H80C0FF
End If
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
crpt.ReportFileName = App.Path + "\ReportD\Neraca.rpt"
crpt.WindowState = crptMaximized
crpt.WindowMaxButton = True
crpt.WindowMinButton = True
crpt.Action = 1
End If
Option1.Value = False
End Sub

Private Sub Option3_Click()
If Option3.Value = True Then
crpt.ReportFileName = App.Path + "\ReportD\Laba.rpt"
crpt.WindowState = crptMaximized
crpt.WindowMaxButton = True
crpt.WindowMinButton = True
crpt.Action = 1
End If
Option3.Value = False
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
crpt.ReportFileName = App.Path + "\ReportD\TrialBal.rpt"
crpt.WindowState = crptMaximized
crpt.WindowMaxButton = True
crpt.WindowMinButton = True
crpt.Action = 1
End If
Option2.Value = False
End Sub

Private Sub Option4_Click()
If Option4.Value = True Then
crpt.ReportFileName = App.Path + "\ReportD\TrialBalance.rpt"
crpt.WindowState = crptMaximized
crpt.WindowMaxButton = True
crpt.WindowMinButton = True
crpt.Action = 1
End If
Option4.Value = False
End Sub


