VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form MAINSUPER 
   BackColor       =   &H00FF8080&
   Caption         =   "MENU SUPERVISOR     |     SUZUKI BEDAGAN SEMARANG"
   ClientHeight    =   6075
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   Picture         =   "MAINSUPER.frx":0000
   ScaleHeight     =   405
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   674
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      TabIndex        =   3
      Top             =   5640
      Width           =   10110
      _ExtentX        =   17833
      _ExtentY        =   767
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   5900
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   5900
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   5900
         EndProperty
      EndProperty
   End
   Begin Crystal.CrystalReport crpt 
      Left            =   480
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   7620
      TabIndex        =   2
      Top             =   11730
      Width           =   11475
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   7620
      TabIndex        =   1
      Top             =   10365
      Width           =   11475
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   450
      TabIndex        =   0
      Top             =   12345
      Width           =   5295
   End
   Begin VB.Menu TS 
      Caption         =   "TABEL SISTEM"
      Index           =   10
      Begin VB.Menu TSS 
         Caption         =   "TABEL KENDARAAN"
         Index           =   11
      End
      Begin VB.Menu TSS 
         Caption         =   "TABEL SPAREPART"
         Index           =   12
         Visible         =   0   'False
      End
      Begin VB.Menu TSS 
         Caption         =   "TABEL CUSTOMER"
         Index           =   13
      End
      Begin VB.Menu TSS 
         Caption         =   "TABEL CABANG"
         Index           =   14
      End
      Begin VB.Menu TSS 
         Caption         =   "TABEL HUTANG"
         Index           =   15
         Visible         =   0   'False
      End
   End
   Begin VB.Menu SB 
      Caption         =   "PEMBELIAN"
      Index           =   20
      Begin VB.Menu SBB 
         Caption         =   "DATA PEMBELIAN KENDARAAN"
         Index           =   21
      End
      Begin VB.Menu SBB 
         Caption         =   "DATA PEBELIAN SPAREPART"
         Index           =   22
         Visible         =   0   'False
      End
   End
   Begin VB.Menu P 
      Caption         =   "PENJUALAN"
      Index           =   30
      Begin VB.Menu PP 
         Caption         =   "DATA PENJUALAN KENDARAAN"
         Index           =   31
      End
      Begin VB.Menu PP 
         Caption         =   "DATA PENJUALAN SPAREPART"
         Index           =   32
         Visible         =   0   'False
      End
   End
   Begin VB.Menu L 
      Caption         =   "LAPORAN"
      Index           =   90
      Begin VB.Menu LL 
         Caption         =   "LAPORAN KEUANGAN"
         Index           =   91
      End
      Begin VB.Menu LL 
         Caption         =   "STATEMENT SUB GL"
         Index           =   92
      End
   End
   Begin VB.Menu KL 
      Caption         =   "KELUAR"
      Index           =   200
      Begin VB.Menu KLL 
         Caption         =   "EXIT"
         Index           =   201
      End
      Begin VB.Menu KLL 
         Caption         =   "PASSWORD"
         Index           =   202
      End
   End
End
Attribute VB_Name = "MAINSUPER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'Label2 = N_CCAB
'Label3 = N_ALAMAT
'Label1 = Time

MAINSALE.Top = 0
MAINSALE.Left = 0


With StatusBar1.Panels
    .Item(1).Text = "NAMA USER: " & Operator
    
    .Item(2).Text = "TANGGAL SYSTEM  : " & Tanggal
    
    .Item(3).Text = "Copyrighted © 2008 EDP IPT"
End With
End Sub

Private Sub KLL_Click(Index As Integer)
Select Case Index
    Case 201
        End
    Case 202
        Unload Me
        PASS.Show 1
End Select
End Sub

Private Sub LL_Click(Index As Integer)
Select Case Index
    Case 91
        RP003.Show 1   'LAPORAN KEUANGAN
    Case 92
        RP011A.Show 1   'LAPORAN STATEMENT SUB GL
End Select
End Sub

Private Sub PP_Click(Index As Integer)
Select Case Index
    Case 31
        CRPT.ReportFileName = App.Path + "\ReportD\SALDO_JUAL.rpt"
        CRPT.SelectionFormula = "{M001.STS_Jual} = '1'"
        CRPT.WindowState = crptMaximized
        CRPT.WindowMaxButton = True
        CRPT.WindowMinButton = True
        CRPT.Action = 1
End Select
End Sub

Private Sub SBB_Click(Index As Integer)
Select Case Index
    Case 21
        CRPT.ReportFileName = App.Path + "\ReportD\SALDOSEDIAMTR.rpt"
        CRPT.SelectionFormula = "{M001.STS_Jual} = '0'"
        CRPT.WindowState = crptMaximized
        CRPT.WindowMaxButton = True
        CRPT.WindowMinButton = True
        CRPT.Action = 1
End Select
End Sub

Private Sub TSS_Click(Index As Integer)
Select Case Index
    Case 11
        CRPT.ReportFileName = App.Path + "\ReportD\B003AA.rpt"
        CRPT.WindowState = crptMaximized
        CRPT.WindowMaxButton = False
        CRPT.WindowMinButton = False
        CRPT.Action = 1
    Case 12

    Case 13
        CRPT.ReportFileName = App.Path + "\ReportD\C012A.rpt"
        CRPT.SelectionFormula = "{C012.TIPE}='200'"
        CRPT.WindowState = crptMaximized
        CRPT.WindowMaxButton = False
        CRPT.WindowMinButton = False
        CRPT.Action = 1
    Case 14
        CRPT.ReportFileName = App.Path + "\ReportD\C012B.rpt"
        CRPT.SelectionFormula = "{C012.TIPE}='300'"
        CRPT.WindowState = crptMaximized
        CRPT.WindowMaxButton = False
        CRPT.WindowMinButton = False
        CRPT.Action = 1
    Case 15
End Select
End Sub


