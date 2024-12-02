VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form BL01 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TRANSAKSI PEMBELIAN SPAREPART"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14445
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   14445
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text11 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   11325
      TabIndex        =   23
      Text            =   "Text11"
      Top             =   2475
      Width           =   2310
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   9090
      TabIndex        =   5
      Text            =   "Text10"
      Top             =   2475
      Width           =   990
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   11325
      TabIndex        =   22
      Text            =   "Text7"
      Top             =   1830
      Width           =   2310
   End
   Begin VB.CommandButton cmdBL003 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   14085
      MaskColor       =   &H8000000F&
      TabIndex        =   6
      Top             =   2790
      Width           =   165
   End
   Begin VB.CommandButton Command2 
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
      Height          =   540
      Left            =   12015
      TabIndex        =   21
      Top             =   7695
      Width           =   2220
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SIMPAN"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   180
      TabIndex        =   20
      Top             =   7695
      Width           =   2220
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1335
      Width           =   7470
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1830
      Width           =   7470
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   11325
      TabIndex        =   4
      Text            =   "Text6"
      Top             =   1335
      Width           =   2310
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   9165
      TabIndex        =   13
      Text            =   "1,000,000.00"
      Top             =   60
      Width           =   5190
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   1702
      MaxLength       =   15
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   1950
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   1702
      TabIndex        =   1
      Text            =   "Text4"
      Top             =   645
      Width           =   1950
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   4515
      Left            =   150
      TabIndex        =   14
      Top             =   3105
      Width           =   14130
      _ExtentX        =   24924
      _ExtentY        =   7964
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   65280
      BackColorBkg    =   16777152
      GridColor       =   0
      TextStyle       =   3
      TextStyleFixed  =   3
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "LABA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9945
      TabIndex        =   25
      Top             =   2520
      Width           =   1305
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "DISKON PEMBELIAN"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5775
      TabIndex        =   24
      Top             =   2520
      Width           =   3240
   End
   Begin VB.Label Label52 
      Caption         =   "(dd/mm/yyyy)"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3780
      TabIndex        =   19
      Top             =   675
      Width           =   1020
   End
   Begin VB.Label Label4 
      Caption         =   "KODE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   330
      TabIndex        =   18
      Top             =   1335
      Width           =   2070
   End
   Begin VB.Label Label5 
      Caption         =   "NAMA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   330
      TabIndex        =   17
      Top             =   1830
      Width           =   2070
   End
   Begin VB.Label Label6 
      Caption         =   "JUMLAH"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9090
      TabIndex        =   16
      Top             =   1380
      Width           =   2070
   End
   Begin VB.Label Label7 
      Caption         =   "HARGA SATUAN"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9090
      TabIndex        =   15
      Top             =   1875
      Width           =   2070
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7605
      TabIndex        =   12
      Top             =   270
      Width           =   1320
   End
   Begin VB.Label Label2 
      Caption         =   "NO. TRANSAKSI"
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
      Left            =   6165
      TabIndex        =   11
      Top             =   270
      Width           =   1365
   End
   Begin VB.Label Label1 
      Caption         =   "NO. FAKTUR"
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
      Left            =   300
      TabIndex        =   10
      Top             =   255
      Width           =   1140
   End
   Begin VB.Label Label8 
      Caption         =   "TGL PEMBELIAN"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   300
      TabIndex        =   9
      Top             =   645
      Width           =   1365
   End
   Begin VB.Label Label9 
      Caption         =   "TGL TRANSAKSI"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6150
      TabIndex        =   8
      Top             =   645
      Width           =   1320
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7605
      TabIndex        =   7
      Top             =   675
      Width           =   1320
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      Height          =   1080
      Left            =   210
      Top             =   60
      Width           =   8850
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      FillColor       =   &H8000000F&
      FillStyle       =   0  'Solid
      Height          =   1785
      Left            =   150
      Top             =   1230
      Width           =   14130
   End
End
Attribute VB_Name = "BL01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim A, Isi As String

Private RDOE As rdoEnvironment
Private RDCO As rdoConnection
Private RSLNO As rdoResultset

Private RDel, RSL, RSLUser, RCari, RCari2, RCari3, RCari4, RCari44, RCari5, RCari6, RCari7, RCari8, RCari9, RCari10, RCari11, RCari12, RCari13, RCari14, RCari15, RCari16, RSave, RSave2, RSave3, RSave4, RSave5, RSave6, RSave7, RSave8, RSave9, RSave10, RSave11, RSave12, REdit As rdoResultset
Private SDel, SQL, SQLUser, SCari, SCari2, SCari3, SCari4, SCari44, SCari5, SCari6, SCari7, SCari8, SCari9, SCari10, SCari11, SCari12, SCari13, SCari14, SCari15, SCari16, SSave, SSave2, SSave3, SSave4, SSave5, SSave6, SSave7, SSave8, SSave9, SSave10, SSave11, SSave12, SEdit As String

Private RKTG, RSTN, RSPL, RPBR As rdoResultset
Private SKTG, SSTN, SSPL, SPBR As String

Private RKAS As rdoResultset
Private SKAS As String

Private SqlNo As String
Private IndekKode, NoNo, GLCREDIT, NoFuckU

Private Sub cmdBL003_Click()
Dim Tanya
If Text1 = "" Or Text4 = "" Or Combo1 = "" Or Combo2 = "" Or Text6 = "" Or Text7 = "" Then
    MsgBox "MASIH ADA DATA YANG KOSONG", vbSystemModal, "KONFIRMASI"
    Combo1.SetFocus
    Exit Sub
Else
    Tanya = MsgBox("TAMBAH DATA PEMBELIAN", vbSystemModal, "KONFIRMASI")
    If Tanya = vbOK Then
        Text8 = Format(CCur(Text8) + (CCur(Text6) * CCur(Text7)), "##,###.00")
        Call SimpanBL001
        Call SiapkanGrid
        Call IsiGrid
        Text6 = ""
        Text7 = ""
        Text10 = ""
        Text11 = ""
        Combo1.SetFocus
        NoNo = NoNo + 1
        Exit Sub
    End If
End If
End Sub

Private Sub SimpanBL001()
SSave = "Select * From BL01"
Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
RSave.AddNew
    RSave("No_Trans") = Label3
    RSave("No_Fak") = Trim(Text1)
    RSave("No_Urut") = NoNo
    RSave("Kode_JNS") = Combo1
    RSave("Nama_JNS") = Combo2
    RSave("Jml_Beli") = CCur(Text6)
    RSave("Harga_PCS") = CCur(Text7)
    RSave("Jml_Harga") = CCur(Text7) * CCur(Text6)
    RSave("User_Code") = Operator
    RSave("Tanggal") = Tanggal
    
    RSave("DISKON_BELI") = CCur(Text10)
    RSave("LABA") = CCur(Text11)
   
RSave.Update
RSave.Close
Set RSave = Nothing
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Combo1_LostFocus()

If Combo1 = "" Then Exit Sub
SCari = "Select * From B003A where Kode_JNS='" + Trim(Combo1) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    Combo2 = RCari("Nama_JNS")
    Text7 = Format(RCari("HBeli"), "##,###.00")
    Text6.SetFocus
Else
    MsgBox "KODE BELUM TERDAFTAR", vbSystemModal, "KONFIRMASI"
    Combo1.SetFocus
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub Combo2_LostFocus()

If Combo2 = "" Then Exit Sub
SCari2 = "Select * From B003A where Nama_JNS='" + Trim(Combo2) + "'"
Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenDynamic, rdConcurRowVer)
If RCari2.RowCount <> 0 Then
    Combo1 = RCari2("Kode_JNS")
    Text7 = Format(RCari2("HBeli"), "##,###.00")
    Text6.SetFocus
Else
    MsgBox "NAMA BELUM TERDAFTAR", vbSystemModal, "KONFIRMASI"
    Combo2.SetFocus
End If
RCari2.Close
Set RCari2 = Nothing
End Sub

Private Sub Command1_Click()
SKAS = "Select * From G003 where CodeSL = '1001113'"
Set RKAS = RDCO.OpenResultset(SKAS, rdOpenKeyset, rdConcurRowVer)
If CCur(Text8) > RKAS("SALDO") Then
    BL01.Hide
    MsgBox "KAS TERSEDIA SEBESAR RP " + Trim(Format(RKAS("SALDO"), "##,###.00")), vbSystemModal, "KONFIRMASI"
    MsgBox "LAKUKAN TRANSAKSI GENERAL LEDGER TERLEBIH DAHULU", vbSystemModal, "KONFIRMASI"
    
    Exit Sub
Else
    Tanya = MsgBox("ANDA YAKIN MELAKUKAN TRANSAKSI PEMBELIAN", vbSystemModal, "KONFIRMASI")
        If Tanya = vbOK Then
            Call JurnalDEBET
            Call JurnalCredit
            Call UpdateBahanHIS
            Call DelBL001
            Call NoBukti
        End If
End If

Unload Me
BL01.Show

End Sub

Private Sub JurnalDEBET()
SCari7 = "Select * From G003 where codesl ='" + Trim(1001153) + "'"
Set RCari7 = RDCO.OpenResultset(SCari7, rdOpenKeyset, rdConcurRowVer)
    MMUTASID = RCari7("mutasid") + CCur(Text8)
    SSALDO = RCari7("saldo") + CCur(Text8)
    NNAMA = RCari7("NamaSL")
    RCari7.EDIT
        RCari7("mutasid") = CCur(MMUTASID)
        RCari7("saldo") = CCur(SSALDO)

    SCari8 = "Select * From G005"
    Set RCari8 = RDCO.OpenResultset(SCari8, rdOpenKeyset, rdConcurRowVer)
    RCari8.AddNew
        RCari8("codecab") = CodeCab
        RCari8("codesl") = 1001153
        RCari8("namasl") = NNAMA
        RCari8("nobukti") = Label3
        RCari8("keterangan") = "BL.SP." + Text1
        RCari8("nominald") = CCur(Text8)
        RCari8("nominalc") = 0
        RCari8("saldo") = SSALDO
        RCari8("tanggal") = Tanggal
        RCari8("jam") = Date
        RCari8("usercode") = Operator
    RCari8.Update
    RCari8.Close
    Set RCari8 = Nothing

RCari7.Update
RCari7.Close
Set RCari7 = Nothing
End Sub

Private Sub JurnalCredit()
SCari9 = "Select * From C013 where Nama = '" + Trim(Operator) + "'"
Set RCari9 = RDCO.OpenResultset(SCari9, rdOpenKeyset, rdConcurRowVer)
GLCREDIT = RCari9("GCredit")
NoFuckU = CCur(RCari9("NoBeli"))
   
RCari9.EDIT
RCari9("NoBeli") = NoFuckU + 1
    
    SCariCari = "Select * From G003 where codesl='" + Trim(GLCREDIT) + "'"
    Set RCariCari = RDCO.OpenResultset(SCariCari, rdOpenKeyset, rdConcurRowVer)
        MMUTASIC = RCariCari("mutasic") + CCur(Text8)
        SSALDO = RCariCari("saldo") - CCur(Text8)
        NNAMA = RCariCari("NamaSL")
        RCariCari.EDIT
            RCariCari("mutasic") = CCur(MMUTASIC)
            RCariCari("saldo") = CCur(SSALDO)
    
        SCariCari2 = "Select * From G005"
        Set RCariCari2 = RDCO.OpenResultset(SCariCari2, rdOpenKeyset, rdConcurRowVer)
        RCariCari2.AddNew
            RCariCari2("codecab") = CodeCab
            RCariCari2("codesl") = GLCREDIT
            RCariCari2("namasl") = NNAMA
            RCariCari2("nobukti") = Label3
            RCariCari2("keterangan") = "BL.SP." + Text1
            RCariCari2("nominald") = 0
            RCariCari2("nominalc") = CCur(Text8)
            RCariCari2("saldo") = SSALDO
            RCariCari2("tanggal") = Tanggal
            RCariCari2("jam") = Date
            RCariCari2("usercode") = Operator
        RCariCari2.Update
        RCariCari2.Close
        Set RCariCari2 = Nothing
    
    RCariCari.Update
    RCariCari.Close
    Set RCariCari = Nothing

RCari9.Update
RCari9.Close
Set RCari9 = Nothing


End Sub

Private Sub UpdateBahanHIS()
SCari12 = "Select * From BL01"
Set RCari12 = RDCO.OpenResultset(SCari12, rdOpenKeyset, rdConcurRowVer)
RCari12.MoveFirst
Do While Not RCari12.EOF
    KODEJNS = RCari12("Kode_JNS")

    SCari13 = "Select * From B003 where Kode_JNS = '" + Trim(KODEJNS) + "'"
    Set RCari13 = RDCO.OpenResultset(SCari13, rdOpenKeyset, rdConcurRowVer)
    If RCari13.RowCount <> 0 Then
        JMLDBT = CCur(RCari13("JML_DBT")) + CCur(RCari12("JML_BELI"))
        JMLAKHIR = CCur(RCari13("JML_AKHIR")) + CCur(RCari12("JML_BELI"))
    Else
        MsgBox "KODE BARANG TIDAK TERDAFTAR", vbCritical, "WARNING"
        Exit Sub
    End If
    
    SCari13A = "Select * From B003A where Kode_JNS = '" + Trim(KODEJNS) + "'"
    Set RCari13A = RDCO.OpenResultset(SCari13A, rdOpenKeyset, rdConcurRowVer)
    
        MUTASIDBT = CCur(RCari13A("mutasid")) + CCur(RCari12("JML_HARGA"))
        SALDOAKHIR = CCur(RCari13A("saldo")) + CCur(RCari12("JML_HARGA"))
        
        SCari14 = "Select * From B004"
        Set RCari14 = RDCO.OpenResultset(SCari14, rdOpenKeyset, rdConcurRowVer)
            
            'INFO HARGA JUAL PCS'
            RCari14.AddNew
            RCari14("NO_TRANS") = Label3
            RCari14("TGL_BELI") = Text4
            RCari14("KODE_JNS") = KODEJNS
            RCari14("NAMA_JNS") = RCari12("NAMA_JNS")
            RCari14("JML_SATUAN") = RCari12("JML_BELI")
            RCari14("HBELI_PCS") = RCari12("HARGA_PCS")
            RCari14("HARGA_BELI") = RCari12("JML_HARGA")
            RCari14("JML_SALDO") = RCari12("JML_BELI")
            RCari14("NOM_SALDO") = RCari12("JML_HARGA")
            
            RCari14("DISKON_BELI") = RCari12("DISKON_BELI")
            RCari14("LABA") = RCari12("LABA")
        
        'UPDATE MUTASI PEMBELIAN TABEL MASTER BAHAN '
        RCari13.EDIT
        RCari13("JML_DBT") = CCur(JMLDBT)
        RCari13("JML_AKHIR") = CCur(JMLAKHIR)
        
        RCari13A.EDIT
        RCari13A("MUTASID") = CCur(MUTASIDBT)
        RCari13A("SALDO") = CCur(SALDOAKHIR)
        
            If CCur(HJUAL) < RCari13A("HJUAL") Then
                HHJUAL = RCari13A("HJUAL")
            ElseIf CCur(HJUAL) = RCari13A("HJUAL") Then
                HHJUAL = RCari13A("HJUAL")
            Else
                HHJUAL = CCur(HJUAL)
            End If
        
        RCari13A("HJUAL") = CCur(HHJUAL)
            
    RCari13.Update
    RCari13.Close
    Set RCari13 = Nothing
        
    RCari13A.Update
    RCari13A.Close
    Set RCari13A = Nothing
        
        RCari14.Update
        RCari14.Close
        Set RCari14 = Nothing

            'UPDATE HISTORY TRANSAKSI'
            SCari15 = "Select * From B005"
            Set RCari15 = RDCO.OpenResultset(SCari15, rdOpenKeyset, rdConcurRowVer)
            RCari15.AddNew
            RCari15("KODE_TRANS") = "BL"
            RCari15("KODE_JNS") = RCari12("Kode_JNS")
            RCari15("NAMA_JNS") = RCari12("Nama_JNS")
            RCari15("NO_FAKTUR") = Trim(Text1)
            RCari15("NO_BUKTI") = Label3
            RCari15("KETERANGAN") = "BL.SP.NO." + Text1
            RCari15("JML_DBT") = RCari12("Jml_Beli")
            RCari15("JML_CRD") = 0
            RCari15("JML_AKHIR") = CCur(JMLAKHIR)
            RCari15("MUTASI_DBT") = RCari12("Jml_Harga")
            RCari15("MUTASI_CRT") = 0
            RCari15("SALDO_AKHIR") = CCur(SALDOAKHIR)
            RCari15("H_POKOK") = RCari12("Harga_PCS")
            RCari15("KAS") = CCur(SALDOAKHIR)
            RCari15("TANGGAL") = Tanggal
            RCari15("TGL_TRANS") = Label10
            RCari15.Update
            RCari15.Close
            Set RCari15 = Nothing
        
RCari12.MoveNext
Loop
RCari12.Close
Set RCari12 = Nothing

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=DEALER", rdDriverNoPrompt, False, CN)

Call DelBL001
Call NoBukti
Call SiapkanGrid
Call IsiGrid
Call Combo
'Call Auto

Text8 = 0
ClearTextBoxes BL01
Label10 = Tanggal
Text4 = Tanggal
'Combo1 = ""
'Combo2 = ""

NoNo = 1

End Sub

Private Sub DelBL001()
SDel = "Delete * From BL01"
Set RDel = RDCO.OpenResultset(SDel, rdOpenDynamic, rdConcurRowVer)
RDel.Close
Set RDel = Nothing
End Sub

Private Sub NoBukti()
Dim No As Double
SqlNo = "Select * from C013 where nama = '" + Operator + "'"
Set RSLNO = RDCO.OpenResultset(SqlNo, rdOpenDynamic, rdConcurRowVer)

No = Val(RSLNO("NoBeli")) + 1
Label3 = "SP." + Trim(Digit(5, No))
RSLNO.Close
Set RSLNO = Nothing
End Sub

Private Sub SiapkanGrid()
With grid
    .Row = 0
    .Cols = 8
    .Col = 0: .ColWidth(0) = 750: .Text = "NO": .CellAlignment = 4
    .Col = 1: .ColWidth(1) = 2000: .Text = "KODE": .CellAlignment = 4
    .Col = 2: .ColWidth(2) = 3400: .Text = "NAMA BARANG": .CellAlignment = 4
    .Col = 3: .ColWidth(3) = 1000: .Text = "JUMLAH": .CellAlignment = 4
    .Col = 4: .ColWidth(4) = 1500: .Text = "HARGA PCS": .CellAlignment = 4
    .Col = 5: .ColWidth(5) = 2000: .Text = "SUB TOTAL": .CellAlignment = 4
    .Col = 6: .ColWidth(6) = 1500: .Text = "DISKON (%)": .CellAlignment = 4
    .Col = 7: .ColWidth(7) = 1500: .Text = "LABA": .CellAlignment = 4
End With
End Sub

Private Sub IsiGrid()
SCari4 = "Select * From BL01"
Set RCari4 = RDCO.OpenResultset(SCari4, rdOpenKeyset, rdConcurReadOnly)
If RCari4.RowCount <> 0 Then
   RCari4.MoveFirst
   B = 1
   Do Until RCari4.EOF
      grid.Rows = B + 1
      grid.Row = B
         With grid
              .Col = 0: .Text = RCari4("No_Urut"): .CellAlignment = 4
              .Col = 1: .Text = RCari4("Kode_JNS"): .CellAlignment = 4
              .Col = 2: .Text = RCari4("Nama_JNS")
              .Col = 3: .Text = RCari4("Jml_Beli"): .CellAlignment = 4
              .Col = 4: .Text = Format(RCari4("Harga_PCS"), "##,###.00")
              .Col = 5: .Text = Format(RCari4("Jml_Harga"), "##,###.00")
              .Col = 6: .Text = Format(RCari4("DISKON_BELI"), "##,###.00")
              .Col = 7: .Text = Format(RCari4("LABA"), "##,###.00")
         End With
         
      B = B + 1
      RCari4.MoveNext
   Loop
End If
RCari4.Close
Set RCari4 = Nothing
End Sub

Private Sub Combo()

SKTG = "Select * From B003 where KODE_IND = '153' order by KODE_JNS"
Set RKTG = RDCO.OpenResultset(SKTG, rdOpenDynamic, rdOpenKeyset)
RKTG.MoveFirst
Do While Not RKTG.EOF
    Combo1.AddItem RKTG("Kode_JNS")
RKTG.MoveNext
Loop
RKTG.Close
Set RKTG = Nothing

SSTN = "Select * From B003 where KODE_IND = '153' order by NAMA_JNS"
Set RSTN = RDCO.OpenResultset(SSTN, rdOpenDynamic, rdOpenKeyset)
RSTN.MoveFirst
Do While Not RSTN.EOF
    Combo2.AddItem RSTN("Nama_JNS")
RSTN.MoveNext
Loop
RSTN.Close
Set RSTN = Nothing

End Sub

Private Sub text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text8 = "0,00"
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text1_LostFocus()
Text1 = Format(Text1, ">")
End Sub

Private Sub text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys vbTab
End If
End Sub

Private Sub Text4_LostFocus()
If Text4 = "" Then Exit Sub
If Not IsDate(Text4) Then
    Text4.SetFocus
    Text4 = Tanggal
    MsgBox "TYPE TANGGAL SALAH", vbCritical, "WARNING"
End If
Text4 = Format(Text4, "DD/MM/YYYY")
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub text7_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If

If KeyAscii = 13 Then
    SendKeys "{TAB}"
    Text7 = Format(Text7, "##,###.00")
End If

End Sub

Private Sub Text10_LostFocus()
If Text10 = "" Then
    Text10 = 0
    Text11 = 0
    Exit Sub
End If

If Text10 = "" Then Exit Sub
    If Not IsNumeric(Text10) Then
        MsgBox "HARUS DALAM FORMAT ANGKA", vbCritical, "TYPE DATA SALAH"
        Text10.SetFocus
    Exit Sub
End If

Text11 = Format(CCur(Text7) * CCur(Text10) / 100, "##,###.00")

End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub
