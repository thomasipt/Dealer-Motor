VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form RP005A 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LAPORAN PERSEDIAAN"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4695
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "CETAK"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3675
      TabIndex        =   8
      Top             =   2250
      Width           =   885
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   135
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2250
      Width           =   3435
   End
   Begin VB.OptionButton Option3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080C0FF&
      Caption         =   "LAPORAN SERVICE SPAREPART"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   135
      TabIndex        =   5
      Top             =   675
      Width           =   4425
   End
   Begin VB.OptionButton Option2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080C0FF&
      Caption         =   "LAPORAN SALDO SPAREPART"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   135
      TabIndex        =   2
      Top             =   1035
      Width           =   4425
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080C0FF&
      Caption         =   "LAPORAN LABA SPAREPART"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   135
      TabIndex        =   1
      Top             =   1395
      Width           =   4425
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
      Left            =   1785
      TabIndex        =   0
      Top             =   2790
      Width           =   1110
   End
   Begin Crystal.CrystalReport CRPT 
      Left            =   0
      Top             =   5040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   330
      Left            =   2445
      TabIndex        =   3
      Top             =   90
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   582
      _Version        =   393216
      Format          =   54853633
      CurrentDate     =   39531
   End
   Begin VB.Line Line3 
      X1              =   127
      X2              =   4552
      Y1              =   2655
      Y2              =   2655
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080C0FF&
      Caption         =   "LAPORAN MUTASI SPAREPART"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   967
      TabIndex        =   6
      Top             =   1980
      Width           =   2760
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   4545
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "TANGGAL TRANSAKSI"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   435
      TabIndex        =   4
      Top             =   158
      Width           =   1815
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4545
      Y1              =   1845
      Y2              =   1845
   End
End
Attribute VB_Name = "RP005A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private RDel, RCari, RCari2, RPLY, RSPR As rdoResultset
Private SDel, SCari, SCari2, SPLY, SSPR As String

Private RSave, RSave2, RSave3, RSave4, RSave5, RSave6, RSave7 As rdoResultset
Private SSave, SSave2, SSave3, SSave4, SSave5, SSave6, SSave7 As String

Private RJual1, RJual2, RJual3, RJual4 As rdoResultset
Private SJual1, SJual2, SJual3, SJual4 As String

Private RKAS, RKAS2, RKAS3 As rdoResultset
Private SKAS, SKAS2, SKAS3 As String

Private RKREDIT, RKREDIT2, RKREDIT3 As rdoResultset
Private SKREDIT, SKREDIT2, SKREDIT3 As String

Private RLABA, RLABA2 As rdoResultset
Private SLABA, SLABA2 As String

Private CEKKODE, SGLPART, SGLJASA, SGLSEDIA

Private KKODES, NNAMAS, BBIAYAS
Private NoUrutTrans As String

Private Y, M, D As Integer

Private Sub Command1_Click()
crpt.ReportFileName = App.Path + "\ReportD\HisBarang.rpt"
crpt.SelectionFormula = "{B003.KODE_JNS} = '" + Trim(Combo1) + "'"
crpt.WindowState = crptMaximized
crpt.WindowMaxButton = False
crpt.WindowMinButton = False
crpt.Action = 1
Option1.Value = False
End Sub

Private Sub cMDeXIT_Click()
Unload Me
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=DEALER", rdDriverNoPrompt, False, CN)
DTPicker1 = Tanggal
Call IsiCombo
End Sub

Private Sub IsiCombo()
SSPR = "Select * From B003 where KODE_IND = '153' order by KODE_JNS"
Set RSPR = RDCO.OpenResultset(SSPR, rdOpenDynamic, rdOpenKeyset)
RSPR.MoveFirst
Do While Not RSPR.EOF
    Combo1.AddItem RSPR("KODE_JNS")
RSPR.MoveNext
Loop
RSPR.Close
Set RSPR = Nothing

Combo1.ListIndex = 0
End Sub

Private Sub Option1_Click()
D = Day(DTPicker1)
M = Month(DTPicker1)
Y = Year(DTPicker1)
crpt.ReportFileName = App.Path + "\ReportD\MUTASI.rpt"
crpt.SelectionFormula = "{B005.TANGGAL} in date (" & Y & "," & M & "," & D & ") to date (" & Y & "," & M & "," & D & ")"
crpt.WindowState = crptMaximized
crpt.WindowMaxButton = False
crpt.WindowMinButton = False
crpt.Action = 1
Option1.Value = False
End Sub

Private Sub Option2_Click()
D = Day(DTPicker1)
M = Month(DTPicker1)
Y = Year(DTPicker1)
crpt.ReportFileName = App.Path + "\ReportD\SALDOSEDIAPART.rpt"
crpt.SelectionFormula = "{B003.KODE_IND} = '153'"
crpt.WindowState = crptMaximized
crpt.WindowMaxButton = False
crpt.WindowMinButton = False
crpt.Action = 1
Option2.Value = False
End Sub

Private Sub Option3_Click()
D = Day(DTPicker1)
M = Month(DTPicker1)
Y = Year(DTPicker1)
crpt.ReportFileName = App.Path + "\ReportD\NOTAALL.rpt"
crpt.SelectionFormula = "{S003.TGL_TRANS} in date (" & Y & "," & M & "," & D & ") to date (" & Y & "," & M & "," & D & ")"
crpt.WindowState = crptMaximized
crpt.WindowMaxButton = False
crpt.WindowMinButton = False
crpt.Action = 1
Option3.Value = False
crpt.Reset
End Sub

Private Sub Option4_Click()
D = Day(DTPicker1)
M = Month(DTPicker1)
Y = Year(DTPicker1)
crpt.ReportFileName = App.Path + "\ReportD\NOTAALL_NKSG.rpt"
crpt.SelectionFormula = "{S003.TGL_TRANS} in date (" & Y & "," & M & "," & D & ") to date (" & Y & "," & M & "," & D & ")"
crpt.WindowState = crptMaximized
crpt.WindowMaxButton = False
crpt.WindowMinButton = False
crpt.Action = 1
Option4.Value = False
crpt.Reset
End Sub



