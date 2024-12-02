VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MAINSALE 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MENU PENJUALAN & PEMBELIAN KENDARAAN"
   ClientHeight    =   8685
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   11880
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "MAINSALE.frx":0000
   ScaleHeight     =   768
   ScaleMode       =   0  'User
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   8250
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   767
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   6932
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   6932
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   6932
         EndProperty
      EndProperty
   End
   Begin VB.Menu TS 
      Caption         =   "TABEL SISTEM"
      Index           =   10
      Begin VB.Menu TSS 
         Caption         =   "TABEL KODE INDUK BAHAN"
         Index           =   11
      End
      Begin VB.Menu TSS 
         Caption         =   "TABEL KODE JENIS BAHAN"
         Index           =   12
      End
      Begin VB.Menu TSS 
         Caption         =   "DATA CUSTOMER / CABANG / SUPPLIER"
         Index           =   13
      End
      Begin VB.Menu TSS 
         Caption         =   "TABEL KODE HUTANG"
         Index           =   14
         Visible         =   0   'False
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
      End
   End
   Begin VB.Menu P 
      Caption         =   "PENJUALAN"
      Index           =   30
      Begin VB.Menu PP 
         Caption         =   "TRANSAKSI PENJUALAN"
         Index           =   31
      End
      Begin VB.Menu PP 
         Caption         =   "-"
         Index           =   32
      End
      Begin VB.Menu PP 
         Caption         =   "DAFTAR PENJUALAN"
         Index           =   33
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
   End
   Begin VB.Menu PHP 
      Caption         =   "PENCAIRAN H/P"
      Index           =   50
      Begin VB.Menu PHPP 
         Caption         =   "TABEL HUTANG"
         Index           =   51
      End
      Begin VB.Menu PHPP 
         Caption         =   "-"
         Index           =   52
      End
      Begin VB.Menu PHPP 
         Caption         =   "TABEL PIUTANG"
         Index           =   53
      End
      Begin VB.Menu PHPP 
         Caption         =   "TRANSAKSI PIUTANG"
         Index           =   54
         Visible         =   0   'False
      End
   End
   Begin VB.Menu TSGL 
      Caption         =   "TRANSAKSI"
      Index           =   80
      Begin VB.Menu TSGLL 
         Caption         =   "TRANSAKSI ANTAR SUB GL"
         Index           =   81
      End
      Begin VB.Menu TSGLL 
         Caption         =   "-"
         Index           =   82
         Visible         =   0   'False
      End
      Begin VB.Menu TSGLL 
         Caption         =   "DEBET / CREDIT BARANG"
         Index           =   83
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
         Caption         =   "LAPORAN KEUANGAN PER TANGGAL"
         Index           =   92
      End
      Begin VB.Menu LL 
         Caption         =   "-"
         Index           =   93
      End
      Begin VB.Menu LL 
         Caption         =   "STATEMENT SUB GL"
         Index           =   94
      End
      Begin VB.Menu LL 
         Caption         =   "LAPORAN SALDO"
         Index           =   95
      End
      Begin VB.Menu LL 
         Caption         =   "-"
         Index           =   96
         Visible         =   0   'False
      End
      Begin VB.Menu LL 
         Caption         =   "LAPORAN DATA"
         Index           =   97
         Visible         =   0   'False
      End
   End
   Begin VB.Menu PS 
      Caption         =   "PROSES"
      Index           =   100
      Begin VB.Menu PSS 
         Caption         =   "END OFF DAY"
         Index           =   101
      End
      Begin VB.Menu PSS 
         Caption         =   "-"
         Index           =   102
      End
      Begin VB.Menu PSS 
         Caption         =   "BACKUP RESTORE"
         Index           =   103
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
Attribute VB_Name = "MAINSALE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type tAtur
    sIPT As String
End Type

Dim tSet As tAtur

Private Sub LoadSaveAtur(bLoad As Boolean)
    Dim nFile As String
    Dim ff As Integer

    nFile = App.Path & "\NOVI.dat"
    ff = FreeFile

    If bLoad = True Then
    
        Open nFile For Binary Access Read As #ff
        Get #ff, , tSet
        Close #ff
        
        With tSet
            Text1.Text = .sIPT
        End With

    Else

        With tSet
            .sIPT = Tanggal
        End With
        
        If Dir(nFile, 1 Or 2 Or 4 Or 32) <> "" Then Kill nFile
        
        Open nFile For Binary Access Read Write As #ff
        Put #ff, , tSet
        Close #ff

    End If
End Sub

Private Sub Form_Load()
'Label2 = N_CCAB
'Label3 = N_ALAMAT
'Label1 = Time


MAINSALE.Top = 0
MAINSALE.Left = 0


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
    Case 91
        RP003.Show  'LAPORAN KEUANGAN
    Case 92
        RP004.Show  'LAPORAN KEUANGAN PER TANGGAL
    Case 94
        RP011.Show  'LAPORAN STATEMENT SUB GL
    Case 95
        RP001.Show 1 'LAPORAN SALDO PERSEDIAAN
    Case 97
        RP005.Show  'LAPORAN JATUH TEMPO
End Select
End Sub

Private Sub PHPP_Click(Index As Integer)
Select Case Index
    Case 51
        H002.Show       'HUTANG
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
        JL003A.Show     'BPKB + STNK
    Case 35
        JL005.Show      'MUTASI MOTOR
End Select
End Sub

Private Sub PSS_Click(Index As Integer)
Select Case Index
    Case 101
        MsgBox "PASTIKAN SISTEM DI SEMUA KOMPUTER TELAH DITUTUP !!!!"
        E001.Show   'PROSES EOD
    Case 103
        LoadSaveAtur False
        ReturnValue = Shell("C:\Program Files\P_DEALER\BACKUP.EXE", 1)
        AppActivate ReturnValue
        End
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
        M003.Show 1   'TABEL MOTOR
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
        B001.Show 1   'TABEL KODE GOLONGAN (B001)
    Case 12
        B003.Show 1   'TABEL KODE BAHAN  (B003)
    Case 13
        C012.Show 1   'TABEL CUSTOMER / SUPPLIER / DLL
    Case 14
        H001.Show 1   'TABEL HUTANG
    Case 15
        P001.Show 1   'TABEL PIUTANG
End Select
End Sub

