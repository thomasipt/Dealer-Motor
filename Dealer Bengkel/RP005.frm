VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form RP005 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LAPORAN MASTER DATA"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4695
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   4695
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
      Left            =   1792
      TabIndex        =   0
      Top             =   1800
      Width           =   1110
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
      Left            =   142
      TabIndex        =   4
      Top             =   1200
      Width           =   4425
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
      Left            =   142
      TabIndex        =   3
      Top             =   435
      Width           =   4425
   End
   Begin VB.OptionButton Option2 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Left            =   142
      TabIndex        =   2
      Top             =   840
      Width           =   4425
   End
   Begin VB.OptionButton Option3 
      Alignment       =   1  'Right Justify
      Caption         =   "LAPORAN DATA CUSTUMER / SUPPLIER / CABANG"
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
      Left            =   142
      TabIndex        =   1
      Top             =   98
      Width           =   4425
   End
   Begin Crystal.CrystalReport CRPT 
      Left            =   2805
      Top             =   3135
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4545
      Y1              =   1695
      Y2              =   1695
   End
End
Attribute VB_Name = "RP005"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub cMDeXIT_Click()
Unload Me
End Sub

Private Sub Form_Load()
Option1.Value = False
Option4.Value = False
End Sub

Private Sub Option1_Click()
    crpt.ReportFileName = App.Path + "\ReportD\SALDO_PIUT.rpt"
    crpt.WindowState = crptMaximized
    crpt.WindowMaxButton = False
    crpt.WindowMinButton = False
    crpt.Action = 1
    Option4.Value = False
End Sub

Private Sub Option2_Click()
    crpt.ReportFileName = App.Path + "\ReportD\SALDO_JUAL.rpt"
    crpt.WindowState = crptMaximized
    crpt.WindowMaxButton = False
    crpt.WindowMinButton = False
    crpt.Action = 1
    Option2.Value = False
End Sub

Private Sub Option3_Click()
    crpt.ReportFileName = App.Path + "\ReportD\C012.rpt"
    crpt.WindowState = crptMaximized
    crpt.WindowMaxButton = False
    crpt.WindowMinButton = False
    crpt.Action = 1
    Option2.Value = False
End Sub

Private Sub Option4_Click()
    crpt.ReportFileName = App.Path + "\ReportD\SALDOSEDIAMTR.rpt"
    crpt.SelectionFormula = "{M001.STS_Jual} = '0'"
    crpt.WindowState = crptMaximized
    crpt.WindowMaxButton = False
    crpt.WindowMinButton = False
    crpt.Action = 1
    Option4.Value = False
    crpt.Reset
End Sub


