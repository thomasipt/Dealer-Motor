VERSION 5.00
Begin VB.Form P003 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PENERIMAAN ANGSURAN PIUTANG"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5805
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   5805
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   1710
      TabIndex        =   10
      Text            =   "Text4"
      Top             =   2250
      Width           =   1905
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
      Height          =   420
      Left            =   4402
      TabIndex        =   9
      Top             =   3060
      Width           =   960
   End
   Begin VB.CommandButton Tmb_Save 
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
      Height          =   375
      Left            =   442
      TabIndex        =   3
      Top             =   3105
      Width           =   1000
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   1710
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   1725
      Width           =   1905
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   1710
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   645
      Width           =   1320
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   1710
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   150
      Width           =   1320
   End
   Begin VB.Label Label2 
      Caption         =   "INTENSIF"
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
      Left            =   120
      TabIndex        =   11
      Top             =   2250
      Width           =   1410
   End
   Begin VB.Line Line1 
      X1              =   90
      X2              =   5715
      Y1              =   2910
      Y2              =   2910
   End
   Begin VB.Label Label7 
      Caption         =   "PIUTANG"
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
      Left            =   120
      TabIndex        =   8
      Top             =   1725
      Width           =   1410
   End
   Begin VB.Label Label6 
      Caption         =   "NAMA DEBITUR"
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
      Left            =   120
      TabIndex        =   7
      Top             =   1185
      Width           =   1410
   End
   Begin VB.Label Label5 
      Caption         =   "NO. PIUTANG"
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
      Left            =   120
      TabIndex        =   6
      Top             =   690
      Width           =   1410
   End
   Begin VB.Label Label3 
      Caption         =   "NO. BUKTI"
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
      Left            =   120
      TabIndex        =   5
      Top             =   150
      Width           =   1410
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
      Height          =   300
      Left            =   1710
      TabIndex        =   4
      Top             =   1140
      Width           =   3975
   End
End
Attribute VB_Name = "P003"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private RSGL, RBukti, RHutang As rdoResultset
Private RCari, RCari2, RCari3, RCari4, RCari5, RCari6 As rdoResultset
Private ROyen, ROyen2, ROyen3, ROyen4, ROyen5, ROyen6 As rdoResultset

Private SSGL, SBukti, SHutang As String
Private SCari, SCari2, SCari3, SCari4, SCari5, SCari6 As String
Private SOyen, SOyen2, SOyen3, SOyen4 As String

Private Baki

Private Sub Command2_Click()
Unload Me
P002.Show
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=DEALER", rdDriverNoPrompt, False, CN)

Call Kosong
End Sub

Private Sub Kosong()
Label1 = ""
Text1 = ""
Text2 = ""
Text3 = 0
Text4 = 0
Text2 = NoPinjaman
End Sub

Private Sub text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text1_LostFocus()
If Text1 = "" Then Exit Sub
If Not IsNumeric(Text1) Then
    Text1.SetFocus
    MsgBox "NOMOR BUKTI HARUS ANGKA", vbCritical, "TYPE DATA SALAH"
    Exit Sub
End If
Text1 = Digit(10, Text1)
SBukti = "Select NoBukti from G004 where Nobukti = '" + Trim(Text1) + "'"
Set RBukti = RDCO.OpenResultset(SBukti, rdOpenDynamic, rdConcurRowVer)
If RBukti.RowCount <> 0 Then
    Text1.SetFocus
    MsgBox "NOMOR BUKTI TRANSAKSI TELAH DIGUNAKAN", vbInformation, "NO. BUKTI SUDAH TERPAKAI"
End If
RBukti.Close
Set RBukti = Nothing
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text2_LostFocus()
If Text2 = "" Then Exit Sub
'Dim Awal
'Dim Akhir As Double
'Awal = Mid(Text2, 1, 3)
'Akhir = Mid(Text2, 4, 7)
'Text2 = Trim(Awal) + "." + Digit(7, Akhir)
SHutang = "Select * from P002 where Nomor_Pin = '" + Trim(Text2) + "'"
Set RHutang = RDCO.OpenResultset(SHutang, rdOpenDynamic, rdConcurRowVer)
Do While Not RHutang.EOF

    Label1 = RHutang("Nama_Nas")
    Baki = RHutang("Baki_Debet")
    Ints = RHutang("Intensif")
    Text3 = Format(Baki, "##,###.00")
    Text4 = Format(Ints, "##,###.00")
    
RHutang.MoveNext
Loop

RHutang.Close
Set RHutang = Nothing
Text2 = Format(Text2, ">")
End Sub

Private Sub Text3_GotFocus()
If Text3 = 0 Then Text3 = ""
End Sub

Private Sub text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text3_LostFocus()
If Text3 = "" Then Text3 = 0
If Not IsNumeric(Text3) Then
    Text3.SetFocus
    MsgBox "NOMINAL ANGSURAN POKOK HARUS ANGKA", vbCritical, "TYPE DATA SALAH"
    Exit Sub
End If
Text3 = Format(Text3, "##,###.00")
If CCur(Text3) > CCur(Baki) Then
    Text3.SetFocus
    MsgBox "NOMINAL ANGSURAN POKOK TIDAK BOLEH LEBIH BESAR DARI OUTSTANDING HUTANG", vbCritical, "ANGSURAN POKOK LEBIH BESAR DARI SISA HUTANG"
    Exit Sub
End If
End Sub

Private Sub Tmb_Save_Click()
Dim Tanya

If Text1 = "" Or Text2 = "" Then
    MsgBox "MASIH ADA DATA YANG KOSONG", vbCritical, "DATA MASIH KOSONG"
    Exit Sub
End If

Tanya = MsgBox("YAKIN PROSES TRANSAKSI PIUTANG ?", vbOKCancel, "SIMPAN TRANSAKSI PENJUALAN")
If Tanya = vbCancel Then Exit Sub

Call PIUTANG
Call PIUTANG2
Call TUNAI
Call TUNAI2

Call Kosong
Text1.SetFocus
Unload Me
P003.Show
End Sub

Private Sub PIUTANG()
SCari = "Select * From P002 where NOMOR_PIN = '" + Trim(Text2) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset, rdConcurRowVer)
    BAKIDEBET = RCari("BAKI_DEBET") - CCur(Text3)
    KODEPIN = RCari("KODE_PIN")
    
    SCari4 = "Select * From P001 where KODE_PIN ='" + Trim(KODEPIN) + "'"
    Set RCari4 = RDCO.OpenResultset(SCari4, rdOpenKeyset, rdConcurRowVer)
    SGLPIUTANG = RCari4("SGL_PIN")
    SGLPDPT = RCari4("SGL_PDPT")
    
        SCari3 = "Select * From G003 where codesl='" + Trim(SGLPIUTANG) + "'"
        Set RCari3 = RDCO.OpenResultset(SCari3, rdOpenKeyset, rdConcurRowVer)
        MMUTASIC = RCari3("MutasiC") + CCur(Text3)
        SSALDO = RCari3("Saldo") - CCur(Text3)
        
                
        RCari3.EDIT
        RCari3("MutasiC") = CCur(MMUTASIC)
        RCari3("Saldo") = CCur(SSALDO)
        
            SCari5 = "Select * From G004"
            Set RCari5 = RDCO.OpenResultset(SCari5, rdOpenKeyset, rdConcurRowVer)
            RCari5.AddNew
            RCari5("codecab") = CodeCab
            RCari5("codesl") = SGLPIUTANG
            RCari5("namasl") = RCari3("NamaSL")
            RCari5("nobukti") = Trim(Text1)
            RCari5("keterangan") = "TR.PIUTANG NO." + Text2
            RCari5("nominald") = 0
            RCari5("nominalc") = CCur(Text3)
            RCari5("tanggal") = Tanggal
            RCari5("usercode") = Operator
            RCari5("jam") = Time
            RCari5.Update
            RCari5.Close
            Set RCari5 = Nothing
         
                SCari6 = "Select * From G005"
                Set RCari6 = RDCO.OpenResultset(SCari6, rdOpenKeyset, rdConcurRowVer)
                RCari6.AddNew
                RCari6("codecab") = CodeCab
                RCari6("codesl") = SGLPIUTANG
                RCari6("namasl") = RCari3("NamaSL")
                RCari6("nobukti") = Trim(Text1)
                RCari6("keterangan") = "TRANSAKSI PIUTANG NO." + Text2
                RCari6("nominald") = 0
                RCari6("nominalc") = CCur(Text3)
                RCari6("saldo") = CCur(SSALDO)
                RCari6("tanggal") = Tanggal
                RCari6("usercode") = Operator
                RCari6("jam") = Time
                RCari6.Update
                RCari6.Close
                Set RCari6 = Nothing
                
        RCari3.Update
        RCari3.Close
        Set RCari3 = Nothing
    RCari.EDIT
    RCari("BAKI_DEBET") = CCur(BAKIDEBET)
    RCari("status") = 1
    
    SCari2 = "Select * From P003"
    Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenKeyset, rdConcurRowVer)
    RCari2.AddNew
    RCari2("Nomor_Pin") = Trim(Text2)
    RCari2("Nama_Nas") = Trim(Label1)
    RCari2("No_Bukti") = Trim(Text1)
    RCari2("Keterangan") = "TRANSAKSI PIUTANG NO." + Text2
    RCari2("Pokok") = CCur(Text3)
    RCari2("Bunga") = 0
    RCari2("Denda") = 0
    RCari2("Baki_Debet") = CCur(BAKIDEBET)
    RCari2("Tanggal") = Tanggal
    RCari2("User_Code") = Operator
    RCari2.Update
    RCari2.Close
    Set RCari2 = Nothing
    
        
        
RCari.Update
RCari.Close
Set RCari = Nothing

End Sub

Private Sub PIUTANG2()
SCari = "Select * From P002 where NOMOR_PIN = '" + Trim(Text2) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset, rdConcurRowVer)
    BAKIDEBET = RCari("BAKI_DEBET") - CCur(Text4)
    KODEPIN = RCari("KODE_PIN")
    
    SCari4 = "Select * From P001 where KODE_PIN ='" + Trim(KODEPIN) + "'"
    Set RCari4 = RDCO.OpenResultset(SCari4, rdOpenKeyset, rdConcurRowVer)
    SGLPIUTANG = RCari4("SGL_INTS")
    SGLPDPT = RCari4("SGL_PDPT")
    
        SCari3 = "Select * From G003 where codesl='" + Trim(SGLPIUTANG) + "'"
        Set RCari3 = RDCO.OpenResultset(SCari3, rdOpenKeyset, rdConcurRowVer)
        MMUTASIC = RCari3("MutasiC") + CCur(Text4)
        SSALDO = RCari3("Saldo") - CCur(Text4)
        
                
        RCari3.EDIT
        RCari3("MutasiC") = CCur(MMUTASIC)
        RCari3("Saldo") = CCur(SSALDO)
        
            SCari5 = "Select * From G004"
            Set RCari5 = RDCO.OpenResultset(SCari5, rdOpenKeyset, rdConcurRowVer)
            RCari5.AddNew
            RCari5("codecab") = CodeCab
            RCari5("codesl") = SGLPIUTANG
            RCari5("namasl") = RCari3("NamaSL")
            RCari5("nobukti") = Trim(Text1)
            RCari5("keterangan") = "TR.INTS NO." + Text2
            RCari5("nominald") = 0
            RCari5("nominalc") = CCur(Text4)
            RCari5("tanggal") = Tanggal
            RCari5("usercode") = Operator
            RCari5("jam") = Time
            RCari5.Update
            RCari5.Close
            Set RCari5 = Nothing
         
                SCari6 = "Select * From G005"
                Set RCari6 = RDCO.OpenResultset(SCari6, rdOpenKeyset, rdConcurRowVer)
                RCari6.AddNew
                RCari6("codecab") = CodeCab
                RCari6("codesl") = SGLPIUTANG
                RCari6("namasl") = RCari3("NamaSL")
                RCari6("nobukti") = Trim(Text1)
                RCari6("keterangan") = "TRANSAKSI INTENSIF NO." + Text2
                RCari6("nominald") = 0
                RCari6("nominalc") = CCur(Text4)
                RCari6("saldo") = CCur(SSALDO)
                RCari6("tanggal") = Tanggal
                RCari6("usercode") = Operator
                RCari6("jam") = Time
                RCari6.Update
                RCari6.Close
                Set RCari6 = Nothing
                
        RCari3.Update
        RCari3.Close
        Set RCari3 = Nothing
    RCari.EDIT
    RCari("BAKI_DEBET") = CCur(BAKIDEBET)
    RCari("status") = 1
        
RCari.Update
RCari.Close
Set RCari = Nothing

End Sub

Private Sub TUNAI()
'SOyen = "Select * From C013 where UserCode = '" + Trim(Operator) + "'"
'Set ROyen = RDCO.OpenResultset(SOyen, rdOpenKeyset, rdConcurRowVer)
'    SGLKAS = ROyen("GDEBET")
     
     SGLKAS = "1001114"
    
    SOyen2 = "Select * From G003 where codesl = '" + Trim(SGLKAS) + "'"
    Set ROyen2 = RDCO.OpenResultset(SOyen2, rdOpenKeyset, rdConcurRowVer)
    MMUTASID = ROyen2("MutasiD") + CCur(Text3)
    SSALDO = ROyen2("Saldo") + CCur(Text3)
    
    ROyen2.EDIT
    ROyen2("MutasiD") = CCur(MMUTASID)
    ROyen2("Saldo") = CCur(SSALDO)
    
        SOyen3 = "Select * From G004"
        Set ROyen3 = RDCO.OpenResultset(SOyen3, rdOpenKeyset, rdConcurRowVer)
        ROyen3.AddNew
        ROyen3("codecab") = CodeCab
        ROyen3("codesl") = SGLKAS
        ROyen3("namasl") = ROyen3("NamaSL")
        ROyen3("nobukti") = Trim(Text1)
        ROyen3("keterangan") = "TRANSAKSI PIUTANG NO." + Text2
        ROyen3("nominald") = CCur(Text3)
        ROyen3("nominalc") = 0
        ROyen3("tanggal") = Tanggal
        ROyen3("usercode") = Operator
        ROyen3("jam") = Time
        ROyen3.Update
        ROyen3.Close
        Set ROyen3 = Nothing

            SOyen4 = "Select * From G005"
            Set ROyen4 = RDCO.OpenResultset(SOyen4, rdOpenKeyset, rdConcurRowVer)
            ROyen4.AddNew
            ROyen4("codecab") = CodeCab
            ROyen4("codesl") = SGLKAS
            ROyen4("namasl") = ROyen4("NamaSL")
            ROyen4("nobukti") = Trim(Text1)
            ROyen4("keterangan") = "TRANSAKSI PIUTANG NO." + Text2
            ROyen4("nominald") = CCur(Text3)
            ROyen4("nominalc") = 0
            ROyen4("saldo") = CCur(SSALDO)
            ROyen4("tanggal") = Tanggal
            ROyen4("usercode") = Operator
            ROyen4("jam") = Time
            ROyen4.Update
            ROyen4.Close
            Set ROyen4 = Nothing

    ROyen2.Update
    ROyen2.Close
    Set ROyen2 = Nothing

'ROyen.Close
'Set ROyen = Nothing

End Sub

Private Sub TUNAI2()
    SGLKAS = "1001114"
    
    SOyen2 = "Select * From G003 where codesl = '" + Trim(SGLKAS) + "'"
    Set ROyen2 = RDCO.OpenResultset(SOyen2, rdOpenKeyset, rdConcurRowVer)
    MMUTASID = ROyen2("MutasiD") + CCur(Text4)
    SSALDO = ROyen2("Saldo") + CCur(Text4)
    
    ROyen2.EDIT
    ROyen2("MutasiD") = CCur(MMUTASID)
    ROyen2("Saldo") = CCur(SSALDO)
    
        SOyen3 = "Select * From G004"
        Set ROyen3 = RDCO.OpenResultset(SOyen3, rdOpenKeyset, rdConcurRowVer)
        ROyen3.AddNew
        ROyen3("codecab") = CodeCab
        ROyen3("codesl") = SGLKAS
        ROyen3("namasl") = ROyen3("NamaSL")
        ROyen3("nobukti") = Trim(Text1)
        ROyen3("keterangan") = "TRANSAKSI INTENSIF NO." + Text2
        ROyen3("nominald") = CCur(Text4)
        ROyen3("nominalc") = 0
        ROyen3("tanggal") = Tanggal
        ROyen3("usercode") = Operator
        ROyen3("jam") = Time
        ROyen3.Update
        ROyen3.Close
        Set ROyen3 = Nothing

            SOyen4 = "Select * From G005"
            Set ROyen4 = RDCO.OpenResultset(SOyen4, rdOpenKeyset, rdConcurRowVer)
            ROyen4.AddNew
            ROyen4("codecab") = CodeCab
            ROyen4("codesl") = SGLKAS
            ROyen4("namasl") = ROyen4("NamaSL")
            ROyen4("nobukti") = Trim(Text1)
            ROyen4("keterangan") = "TRANSAKSI INTENSIF NO." + Text2
            ROyen4("nominald") = CCur(Text4)
            ROyen4("nominalc") = 0
            ROyen4("saldo") = CCur(SSALDO)
            ROyen4("tanggal") = Tanggal
            ROyen4("usercode") = Operator
            ROyen4("jam") = Time
            ROyen4.Update
            ROyen4.Close
            Set ROyen4 = Nothing

    ROyen2.Update
    ROyen2.Close
    Set ROyen2 = Nothing
End Sub
