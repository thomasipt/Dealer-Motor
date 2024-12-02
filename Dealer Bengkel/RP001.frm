VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form RP001 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LAPORAN SALDO"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7440
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   7440
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "LAPORAN KEUNTUNGAN PENJUALAN"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1935
      TabIndex        =   10
      Top             =   2430
      Width           =   3570
   End
   Begin VB.OptionButton Option5 
      Alignment       =   1  'Right Justify
      Caption         =   "LAPORAN SALDO HUTANG"
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
      Left            =   3803
      TabIndex        =   5
      Top             =   630
      Width           =   3480
   End
   Begin VB.OptionButton Option3 
      Alignment       =   1  'Right Justify
      Caption         =   "LAPORAN STOCK BARANG"
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
      Left            =   158
      TabIndex        =   4
      Top             =   120
      Width           =   3480
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "LAPORAN PENJUALAN MOTOR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2340
      TabIndex        =   3
      Top             =   1935
      Width           =   2760
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
      Left            =   3240
      TabIndex        =   0
      Top             =   3045
      Width           =   960
   End
   Begin VB.OptionButton Option4 
      Alignment       =   1  'Right Justify
      Caption         =   "LAPORAN STOCK KENDARAAN"
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
      Left            =   158
      TabIndex        =   1
      Top             =   653
      Width           =   3480
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      Caption         =   "LAPORAN SALDO PIUTANG"
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
      Left            =   3803
      TabIndex        =   2
      Top             =   135
      Width           =   3480
   End
   Begin Crystal.CrystalReport CRPT 
      Left            =   1755
      Top             =   4050
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   330
      Left            =   1410
      TabIndex        =   6
      Top             =   1380
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   582
      _Version        =   393216
      Format          =   54788097
      CurrentDate     =   39531
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   330
      Left            =   5235
      TabIndex        =   9
      Top             =   1380
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   582
      _Version        =   393216
      Format          =   54788097
      CurrentDate     =   39531
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "TGL AKHIR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   4200
      TabIndex        =   8
      Top             =   1440
      Width           =   990
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "TGL AWAL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   420
      TabIndex        =   7
      Top             =   1440
      Width           =   990
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      BorderWidth     =   3
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   1680
      Left            =   -90
      Top             =   1170
      Width           =   7665
   End
End
Attribute VB_Name = "RP001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private RCari As rdoResultset
Private SCari As String

Private D, M, Y
Private D1, M1, Y1
Private D2, M2, Y2

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=DEALER", rdDriverNoPrompt, False, CN)

Call Montok

End Sub

Private Sub Montok()
D = 1
M = Month(Tanggal)
Y = Year(Tanggal)

DTPicker1 = Format(DateSerial(Y, M, D), "DD/MM/YYYY")
DTPicker2 = Tanggal

End Sub

Private Sub Option1_Click()
    If Option1.Value = True Then
    crpt.ReportFileName = App.Path + "\ReportD\SALDO_PIUT.rpt"
    crpt.WindowState = crptMaximized
    crpt.WindowMaxButton = True
    crpt.WindowMinButton = True
    crpt.Action = 1
    End If
    Option1.Value = False
    crpt.Reset
End Sub

Private Sub Option2_Click()
Dim A As Integer
A = Month(Tanggal)

D1 = Day(DTPicker1)
M1 = Month(DTPicker1)
Y1 = Year(DTPicker1)

D2 = Day(DTPicker2)
M2 = Month(DTPicker2)
Y2 = Year(DTPicker2)

If Option2.Value = True Then
    crpt.ReportFileName = App.Path + "\ReportD\SALDO_JUAL.rpt"
    crpt.SelectionFormula = "{M001.TGL_JUAL} in date (" & Y1 & "," & M1 & "," & D1 & ") to date (" & Y2 & "," & M2 & "," & D2 & ") and {M001.STS_Jual} = '1'"
    crpt.WindowState = crptMaximized
    crpt.WindowMaxButton = True
    crpt.WindowMinButton = True
    crpt.Action = 1
End If
    Option2.Value = False
    crpt.Reset
        
End Sub

Private Sub Option3_Click()
    If Option3.Value = True Then
    crpt.ReportFileName = App.Path + "\ReportD\SALDOSEDIA.rpt"
    crpt.WindowState = crptMaximized
    crpt.WindowMaxButton = True
    crpt.WindowMinButton = True
    crpt.Action = 1
    End If
    Option3.Value = False
End Sub

Private Sub Option4_Click()
    If Option4.Value = True Then
    crpt.ReportFileName = App.Path + "\ReportD\SALDOSEDIAMTR.rpt"
    crpt.SelectionFormula = "{M001.STS_Jual} = '0'"
    crpt.WindowState = crptMaximized
    crpt.WindowMaxButton = True
    crpt.WindowMinButton = True
    crpt.Action = 1
    End If
    Option4.Value = False
    crpt.Reset
End Sub

Private Sub Option5_Click()
    If Option5.Value = True Then
    crpt.ReportFileName = App.Path + "\ReportD\SALDO_HUT.rpt"
    crpt.SelectionFormula = "{H002.STATUS} = '0'"
    crpt.WindowState = crptMaximized
    crpt.WindowMaxButton = True
    crpt.WindowMinButton = True
    crpt.Action = 1
    End If
    Option5.Value = False
    crpt.Reset
End Sub

Private Sub Option6_Click()
Dim A As Integer
A = Month(Tanggal)

D1 = Day(DTPicker1)
M1 = Month(DTPicker1)
Y1 = Year(DTPicker1)

D2 = Day(DTPicker2)
M2 = Month(DTPicker2)
Y2 = Year(DTPicker2)

If Option6.Value = True Then
    crpt.ReportFileName = App.Path + "\ReportD\SALDO_JUAL2.rpt"
    crpt.SelectionFormula = "{M001.TGL_JUAL} in date (" & Y1 & "," & M1 & "," & D1 & ") to date (" & Y2 & "," & M2 & "," & D2 & ") and {M001.STS_Jual} = '1'"
    crpt.WindowState = crptMaximized
    crpt.WindowMaxButton = True
    crpt.WindowMinButton = True
    crpt.Action = 1
End If
    Option6.Value = False
    crpt.Reset
End Sub
