VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MAINSERVICE 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MENU SERVICE & SPAREPART KENDARAAN"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   10110
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   405
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   674
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      TabIndex        =   0
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
      TabIndex        =   3
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
      TabIndex        =   2
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
      TabIndex        =   1
      Top             =   12345
      Width           =   5295
   End
   Begin VB.Menu TS 
      Caption         =   "TABEL SISTEM"
      Index           =   10
      Begin VB.Menu TSS 
         Caption         =   "TABEL KODE PELAYANAN"
         Index           =   11
      End
      Begin VB.Menu TSS 
         Caption         =   "TABEL KODE SPAREPART"
         Index           =   12
      End
      Begin VB.Menu TSS 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu TSS 
         Caption         =   "PENCARIAN SPAREPART"
         Index           =   14
      End
      Begin VB.Menu TSS 
         Caption         =   "TABEL KODE PIUTANG"
         Index           =   15
         Visible         =   0   'False
      End
   End
   Begin VB.Menu SB 
      Caption         =   "PEMBELIAN"
      Index           =   20
      Begin VB.Menu SBB 
         Caption         =   "FAKTUR"
         Index           =   21
         Visible         =   0   'False
      End
      Begin VB.Menu SBB 
         Caption         =   "PURCHASE ORDER"
         Index           =   22
         Visible         =   0   'False
      End
      Begin VB.Menu SBB 
         Caption         =   "ENTRY DATA MOTOR"
         Index           =   23
         Visible         =   0   'False
      End
      Begin VB.Menu SBB 
         Caption         =   "-"
         Index           =   24
         Visible         =   0   'False
      End
      Begin VB.Menu SBB 
         Caption         =   "TABEL DATA KENDARAAN"
         Index           =   25
         Visible         =   0   'False
      End
      Begin VB.Menu SBB 
         Caption         =   "STOCK MOTOR"
         Index           =   26
         Visible         =   0   'False
      End
      Begin VB.Menu SBB 
         Caption         =   "SPAREPART"
         Index           =   27
      End
   End
   Begin VB.Menu P 
      Caption         =   "SERVICE SPAREPART"
      Index           =   30
      Begin VB.Menu PP 
         Caption         =   "TRANSAKSI PENJUALAN"
         Index           =   31
         Visible         =   0   'False
      End
      Begin VB.Menu PP 
         Caption         =   "-"
         Index           =   32
         Visible         =   0   'False
      End
      Begin VB.Menu PP 
         Caption         =   "CETAK DO"
         Index           =   33
         Visible         =   0   'False
      End
      Begin VB.Menu PP 
         Caption         =   "-"
         Index           =   34
         Visible         =   0   'False
      End
      Begin VB.Menu PP 
         Caption         =   "MUTASI KENDARAAN"
         Index           =   35
         Visible         =   0   'False
      End
      Begin VB.Menu PP 
         Caption         =   "SERVICE"
         Index           =   36
      End
      Begin VB.Menu PP 
         Caption         =   "DATA SERVICE"
         Index           =   37
         Visible         =   0   'False
      End
      Begin VB.Menu PP 
         Caption         =   "-"
         Index           =   38
         Visible         =   0   'False
      End
      Begin VB.Menu PP 
         Caption         =   "TABEL SERVICE SPAREPART"
         Index           =   39
         Visible         =   0   'False
      End
   End
   Begin VB.Menu T 
      Caption         =   "TRANSAKSI"
      Index           =   40
      Begin VB.Menu TT 
         Caption         =   "TRANSAKSI ANTAR SUB GL"
         Index           =   41
      End
   End
   Begin VB.Menu L 
      Caption         =   "LAPORAN"
      Index           =   90
      Begin VB.Menu LL 
         Caption         =   "JURNAL PELAYANAN"
         Index           =   90
      End
      Begin VB.Menu LL 
         Caption         =   "-"
         Index           =   91
      End
      Begin VB.Menu LL 
         Caption         =   "JURNAL HARIAN"
         Index           =   92
      End
      Begin VB.Menu LL 
         Caption         =   "JURNAL PERSEDIAAN"
         Index           =   93
      End
      Begin VB.Menu LL 
         Caption         =   "-"
         Index           =   94
      End
      Begin VB.Menu LL 
         Caption         =   "LAPORAN KEUANGAN"
         Index           =   95
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
Attribute VB_Name = "MAINSERVICE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private RSaveP, RSave, RCari As rdoResultset
Private SSaveP, SSave, SCari As String


Private Sub Command1_Click()
SCari = "Select * from S004"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
Do Until RCari.EOF
    RCari.EDIT
        RCari("Diskon_Beli") = 0
        RCari("Laba") = 0
    RCari.Update
RCari.MoveNext
Loop
RCari.Close
Set RCari = Nothing

MsgBox "SELESAI BUNG.......!!!!   MERDEKA..............!!!"

End Sub

Private Sub Command2_Click()
S003A.Show
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=DEALER", rdDriverNoPrompt, False, CN)

'Label2 = N_CCAB
'Label3 = N_ALAMAT
'Label1 = Time

Me.Width = Screen.Width
Me.Height = 2000
Me.Top = 0
Me.Left = 0

With StatusBar1.Panels
    .Item(1).Text = "NAMA USER: " & Operator
    
    .Item(2).Text = "TANGGAL SYSTEM  : " & Tanggal
    
    .Item(3).Text = "Copyright © 2008 IPT"
End With
End Sub

Private Sub KLL_Click(Index As Integer)
Select Case Index
    Case 201
        End
    Case 202
        Unload Me
        PASS.Show
End Select
End Sub

Private Sub LL_Click(Index As Integer)
Select Case Index
    Case 90
        SH001.Show 1         'HISTORY SEMUA
    Case 92
        RP003A.Show 1       'LAPORAN KEUANGAN
    Case 93
        RP005A.Show 1       'LAPORAN BARANG
    Case 95
        RP003.Show  'LAPORAN KEUANGAN
End Select
End Sub

Private Sub PHPP_Click(Index As Integer)
Select Case Index
'    Case 51
'        H002.Show       'HUTANG
    Case 53
        P002.Show       'PIUTANG
    Case 54
        P003.Show       'ANGS PIUTANG
End Select
End Sub

Private Sub PP_Click(Index As Integer)
Select Case Index
    Case 31
        JL003.Show      'PENJUALAN
    Case 33
        DO1.Show        'CETAK DO
    Case 35
        JL005.Show      'MUTASI MOTOR
    Case 36
        'KSG.Show
        S003.Show 1     'SERVICE
    Case 37
        S002.Show       'TABEL SERVICE
    Case 39
        S004.Show       'TABEL NOTA
End Select
End Sub

Private Sub PSS_Click(Index As Integer)
Select Case Index
    Case 101
        MsgBox "PASTIKAN SISTEM DI SEMUA KOMPUTER TELAH DITUTUP !!!!"
        E001.Show   'PROSES EOD
End Select
End Sub

Private Sub SBB_Click(Index As Integer)
Select Case Index
    Case 21
        F001.Show   'FAKTUR
    Case 22
        PO1.Show    'ENTRI PO
    Case 23
        M001.Show   'ENTRI MOTOR
    Case 25
        M002.Show   'TABEL M001
    Case 26
        M003.Show   'TABEL MOTOR
    Case 27
        BL01.Show   'BELI SPAREPART
End Select
End Sub

Private Sub TSGLL_Click(Index As Integer)
Select Case Index
    Case 81
        G003.Show   'TRANSAKSI SUB GL
    Case 83
        B004.Show   'TRANSAKSI BARANG
End Select
End Sub

Private Sub TSS_Click(Index As Integer)
Select Case Index
    Case 11
        S002.Show   'TABEL KODE SERVICE
    Case 12
        B003SS.Show   'TABEL KODE BAHAN  (B003)
    Case 13
        C012.Show   'TABEL CUSTOMER / SUPPLIER / DLL
    Case 14
        B003A.Show   'TABEL HUTANG
    Case 15
        P001.Show   'TABEL PIUTANG
End Select
End Sub

Private Sub TT_Click(Index As Integer)
Select Case Index
    Case 41
        G003.Show   'TRANSAKSI SUB GL
End Select
End Sub
