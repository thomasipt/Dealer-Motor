VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form RP004 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   330
      Left            =   2449
      TabIndex        =   5
      Top             =   367
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   582
      _Version        =   393216
      Format          =   54591489
      CurrentDate     =   39531
   End
   Begin Crystal.CrystalReport crpt 
      Left            =   3120
      Top             =   4560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.OptionButton Option3 
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
      Height          =   300
      Left            =   683
      TabIndex        =   3
      Top             =   1785
      Width           =   3315
   End
   Begin VB.OptionButton Option2 
      Caption         =   "LAPORAN LABA / RUGI"
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
      Left            =   683
      TabIndex        =   2
      Top             =   1335
      Width           =   3315
   End
   Begin VB.OptionButton Option1 
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
      Height          =   300
      Left            =   683
      TabIndex        =   1
      Top             =   885
      Width           =   3315
   End
   Begin VB.CommandButton Command1 
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
      Left            =   1853
      TabIndex        =   0
      Top             =   2535
      Width           =   1095
   End
   Begin VB.Label Label1 
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
      Height          =   285
      Left            =   432
      TabIndex        =   4
      Top             =   390
      Width           =   1815
   End
End
Attribute VB_Name = "RP004"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private RMiyabi As rdoResultset
Private SMiyabi As String

Private Y, M, D As Integer

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=DEALER", rdDriverNoPrompt, False, CN)

DTPicker1 = DateAdd("d", -1, Tanggal)
End Sub

Private Sub Option1_Click()
D = Day(DTPicker1)
M = Month(DTPicker1)
Y = Year(DTPicker1)
If Option1.Value = True Then
    crpt.ReportFileName = App.Path + "\ReportD\NERACA1.rpt"
    crpt.SelectionFormula = "{G003.tanggal} in date (" & Y & "," & M & "," & D & ") to date (" & Y & "," & M & "," & D & ")"
    crpt.WindowState = crptMaximized
    crpt.WindowMaxButton = False
    crpt.WindowMinButton = False
    crpt.Action = 1
End If
Option1.Value = False
End Sub

Private Sub Option2_Click()
D = Day(DTPicker1)
M = Month(DTPicker1)
Y = Year(DTPicker1)
If Option2.Value = True Then
    crpt.ReportFileName = App.Path + "\ReportD\LABA1.rpt"
    crpt.SelectionFormula = "{G003.tanggal} in date (" & Y & "," & M & "," & D & ") to date (" & Y & "," & M & "," & D & ")"
    crpt.WindowState = crptMaximized
    crpt.WindowMaxButton = False
    crpt.WindowMinButton = False
    crpt.Action = 1
End If
Option2.Value = False
End Sub

Private Sub Option3_Click()
D = Day(DTPicker1)
M = Month(DTPicker1)
Y = Year(DTPicker1)
If Option3.Value = True Then
    crpt.ReportFileName = App.Path + "\ReportD\trialbalance1.rpt"
    crpt.SelectionFormula = "{G003.tanggal} in date (" & Y & "," & M & "," & D & ") to date (" & Y & "," & M & "," & D & ")"
    crpt.WindowState = crptMaximized
    crpt.WindowMaxButton = False
    crpt.WindowMinButton = False
    crpt.Action = 1
End If
Option3.Value = False
End Sub

Private Sub DTPicker1_LostFocus()
'SMiyabi = "Select Tanggal from G003A where Tanggal = '" + Trim(DTPicker1) + "'"
'Set RMiyabi = RDCO.OpenResultset(SMiyabi, rdOpenDynamic, rdConcurRowVer)
'If RMiyabi.RowCount <> 0 Then
'    If DateValue(DTPicker1) = DateValue(Tanggal) Then
'        DTPicker1.SetFocus
'        MsgBox "LAPORAN UNTUK TANGGAL HARI INI SILAHKAN LIHAT DI MENU LAPORAN KEUANGAN", vbInformation, "LAPORAN KEUANGAN PER TANGGAL INFO"
'        Exit Sub
'    End If
'Else
'    MsgBox "TANGGAL TIDAK ADA", vbInformation, "WARNING"
'End If
If DateValue(DTPicker1) = DateValue(Tanggal) Then
    DTPicker1.SetFocus
    DTPicker1 = DateAdd("d", -1, Tanggal)
    MsgBox "LAPORAN UNTUK TANGGAL HARI INI SILAHKAN LIHAT DI MENU LAPORAN KEUANGAN", vbInformation, "LAPORAN KEUANGAN PER TANGGAL INFO"
    Exit Sub
End If
'RMiyabi.Close
'Set RMiyabi = Nothing

End Sub
