VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form RP003A 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LAPORAN JURNAL HARIAN"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   6855
   StartUpPosition =   2  'CenterScreen
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
      Top             =   1485
      Width           =   1110
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080C0FF&
      Caption         =   "JURNAL HARIAN"
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
      Left            =   180
      TabIndex        =   1
      Top             =   645
      Width           =   3135
   End
   Begin VB.OptionButton Option3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080C0FF&
      Caption         =   "JURNAL PENDAPATAN PENJUALAN"
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
      Left            =   180
      TabIndex        =   2
      Top             =   1005
      Width           =   3135
   End
   Begin VB.OptionButton Option4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080C0FF&
      Caption         =   "JURNAL PERSEDIAAN SPAREPART"
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
      Left            =   3540
      TabIndex        =   4
      Top             =   1005
      Width           =   3135
   End
   Begin VB.OptionButton Option2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080C0FF&
      Caption         =   "JURNAL PENDAPATAN SERVICE"
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
      Left            =   3540
      TabIndex        =   3
      Top             =   645
      Width           =   3135
   End
   Begin Crystal.CrystalReport crpt 
      Left            =   840
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   330
      Left            =   3525
      TabIndex        =   5
      Top             =   90
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   582
      _Version        =   393216
      Format          =   54591489
      CurrentDate     =   39531
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
      Left            =   1515
      TabIndex        =   6
      Top             =   165
      Width           =   1815
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   7440
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Line Line2 
      X1              =   -300
      X2              =   7140
      Y1              =   1365
      Y2              =   1365
   End
End
Attribute VB_Name = "RP003A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private RCari As rdoResultset
Private SCari As String

Private D, M, M1, Y, Hari
Private Tahun, TglAng

Private Sub cMDeXIT_Click()
Unload Me
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=DEALER", rdDriverNoPrompt, False, CN)

DTPicker1 = Tanggal
End Sub

Private Sub Montok()
D = Day(DTPicker1)
M = Month(DTPicker1)
M1 = M + 1
Y = Year(DTPicker1)

Tahun = Format(DateSerial(Y, M1, D), "DD/MM/YYYY")
TglAng = Format(DateSerial(Y, M, D), "DD/MM/YYYY")
Hari = DateDiff("d", TglAng, Tahun)

End Sub

Private Sub Option1_Click()
Call Montok
If Option1.Value = True Then
    crpt.ReportFileName = App.Path + "\ReportD\HisGL.rpt"
    crpt.SelectionFormula = "{G005.Codesl} = '" + Trim(1001113) + "' and {G005.tanggal} in date (" & Y & "," & M & "," & D & ") to date (" & Y & "," & M & "," & D & ") and {G005.UserCode} = 'SERVICE'"
    crpt.WindowState = crptMaximized
    crpt.WindowMaxButton = False
    crpt.WindowMinButton = False
    crpt.Action = 1
    crpt.Reset
End If
    Option1.Value = False
End Sub

Private Sub Option3_Click()
Call Montok

If Option3.Value = True Then
    crpt.ReportFileName = App.Path + "\ReportD\HisGL.rpt"
    crpt.SelectionFormula = "{G005.Codesl} = '" + Trim(8002101) + "' and {G005.tanggal} in date (" & Y & "," & M & "," & D & ") to date (" & Y & "," & M & "," & Hari & ") and {G005.UserCode} = 'SERVICE'"
    crpt.WindowState = crptMaximized
    crpt.WindowMaxButton = False
    crpt.WindowMinButton = False
    crpt.Action = 1
    crpt.Reset
End If
    Option3.Value = False
End Sub

Private Sub Option2_Click()
Call Montok

If Option2.Value = True Then
    crpt.ReportFileName = App.Path + "\ReportD\HisGL.rpt"
    crpt.SelectionFormula = "{G005.Codesl} = '" + Trim(8002201) + "' and {G005.tanggal} in date (" & Y & "," & M & "," & D & ") to date (" & Y & "," & M & "," & Hari & ")  "
    crpt.WindowState = crptMaximized
    crpt.WindowMaxButton = False
    crpt.WindowMinButton = False
    crpt.Action = 1
    crpt.Reset
End If
    Option2.Value = False
End Sub

Private Sub Option4_Click()
Call Montok

If Option4.Value = True Then
    crpt.ReportFileName = App.Path + "\ReportD\HisGL.rpt"
    crpt.SelectionFormula = "{G005.Codesl} = '" + Trim(1001153) + "' and {G005.tanggal} in date (" & Y & "," & M & "," & D & ") to date (" & Y & "," & M & "," & Hari & ")  "
    crpt.WindowState = crptMaximized
    crpt.WindowMaxButton = False
    crpt.WindowMinButton = False
    crpt.Action = 1
    crpt.Reset
End If
    Option4.Value = False
End Sub


