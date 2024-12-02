VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form G003 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TRANSAKSI GL TO GL"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   ControlBox      =   0   'False
   DrawStyle       =   2  'Dot
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   5895
   StartUpPosition =   2  'CenterScreen
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
      Left            =   4515
      TabIndex        =   15
      Top             =   3555
      Width           =   960
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   1683
      MaxLength       =   10
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   1320
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   1683
      MaxLength       =   7
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   570
      Width           =   1005
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   1683
      MaxLength       =   7
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   1425
      Width           =   1005
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   1683
      TabIndex        =   3
      Text            =   "Text4"
      Top             =   2325
      Width           =   1950
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   1683
      MaxLength       =   35
      TabIndex        =   4
      Text            =   "Text5"
      Top             =   2775
      Width           =   3885
   End
   Begin VB.CommandButton TmbSave 
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
      Height          =   420
      Left            =   420
      TabIndex        =   5
      Top             =   3555
      Width           =   960
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      TabIndex        =   16
      Top             =   7155
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   767
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   2566
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   2566
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   2566
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   2566
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Height          =   690
      Left            =   -75
      ScaleHeight     =   630
      ScaleWidth      =   5985
      TabIndex        =   21
      Top             =   3420
      Width           =   6045
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   1429
      MaxLength       =   10
      TabIndex        =   17
      Text            =   "Text6"
      Top             =   4545
      Width           =   1320
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   4406
      MaxLength       =   10
      TabIndex        =   18
      Text            =   "Text7"
      Top             =   4545
      Width           =   1320
   End
   Begin VB.Label Label11 
      Caption         =   "NO. MESIN"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3139
      TabIndex        =   20
      Top             =   4545
      Width           =   1230
   End
   Begin VB.Label Label10 
      Caption         =   "NO. RANGKA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   169
      TabIndex        =   19
      Top             =   4545
      Width           =   1230
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
      Height          =   360
      Left            =   1710
      TabIndex        =   14
      Top             =   1020
      Width           =   3885
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label2"
      Height          =   360
      Left            =   1710
      TabIndex        =   13
      Top             =   1875
      Width           =   3885
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
      Height          =   360
      Left            =   300
      TabIndex        =   12
      Top             =   120
      Width           =   1230
   End
   Begin VB.Label Label4 
      Caption         =   "SGL DEBET"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   300
      TabIndex        =   11
      Top             =   570
      Width           =   1230
   End
   Begin VB.Label Label5 
      Caption         =   "NAMA SGL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   300
      TabIndex        =   10
      Top             =   1020
      Width           =   1230
   End
   Begin VB.Label Label6 
      Caption         =   "SGL CREDIT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   300
      TabIndex        =   9
      Top             =   1470
      Width           =   1230
   End
   Begin VB.Label Label7 
      Caption         =   "NAMA SGL "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   300
      TabIndex        =   8
      Top             =   1920
      Width           =   1230
   End
   Begin VB.Label Label8 
      Caption         =   "NOMINAL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   300
      TabIndex        =   7
      Top             =   2370
      Width           =   1230
   End
   Begin VB.Label Label9 
      Caption         =   "KETERANGAN"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   300
      TabIndex        =   6
      Top             =   2820
      Width           =   1230
   End
End
Attribute VB_Name = "G003"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Dim SAS, Posisi, Saldo As String

Private RCredit2, RDebet2, RSLNO, RSL, RSLUser, RSimpan, RSimpan2, RSimpan3, RSimpan4, RSimpan5, RSave, RSave2, RSave3, RSave4, RSave5, REdit As rdoResultset
Private SCredit2, SDebet2, SqlNo, SQL, SQLUser, SSimpan, SSimpan2, SSimpan3, SSimpan4, SSimpan5, SSave, SSave2, SSave3, SSave4, SSave5, SEdit As String

Private Kirim, RBukti, RDebet, RCredit, RLock, RVal As rdoResultset
Private SBukti, SDebet, SCredit, SLock, SVal As String

Private a_sl, a_nsl, a_md, a_salsl, gdebet

Private Sub Command1_Click()
STS_Biaya = 0
Unload Me
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=DEALER", rdDriverNoPrompt, False, CN)

Call Kosong
Call AutoNo

If Operator = "SERVICE SPAREPART" Then
    Me.BackColor = &H80C0FF
    Label3.BackColor = &H80C0FF
    Label4.BackColor = &H80C0FF
    Label5.BackColor = &H80C0FF
    Label6.BackColor = &H80C0FF
    Label7.BackColor = &H80C0FF
    Label8.BackColor = &H80C0FF
    Label9.BackColor = &H80C0FF
End If

Me.Height = 5310

SCari = "Select * from G003 where CODESL = '1001111'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    S_AWAL = RCari("SALDOAWAL")
    M_DBT = RCari("MUTASID")
    M_CRD = RCari("MUTASIC")
    S_AKHIR = RCari("SALDO")
End If
RCari.Close
Set RCari = Nothing

With StatusBar1.Panels
    .Item(1).Text = Format(S_AWAL, "##,###")
    .Item(2).Text = Format(M_DBT, "##,###")
    .Item(3).Text = Format(M_CRD, "##,###")
    .Item(4).Text = Format(S_AKHIR, "##,###")
End With

If STS_Biaya = 1 Then
    Me.Height = 6240
    Call Auto_Biaya
    SendKeys "{TAB}"
    SendKeys "{TAB}"
    SendKeys "{TAB}"
    Text4.BackColor = &HC0C0FF
Else
    Exit Sub
End If
End Sub

Private Sub Auto_Biaya()
Text2 = G_DEBET
Text3 = G_CREDIT
Text6 = STS_Rangka
Text7 = STS_Mesin
Text5 = "EDIT BIAYA R." + Trim(Text6) + " M." + Trim(Text7)

Text2_LostFocus
Text3_LostFocus

Me.Caption = STS_Nama

End Sub

Private Sub AutoNo()
Dim No As Double
SCari = "Select NoBukti from G004 order by NoBukti desc"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    No = Val(RCari("NoBukti")) + 1
    NoStr = Digit(10, No)
    Text1 = NoStr
Else
    No = 1
    NoStr = Digit(10, No)
    Text1 = NoStr
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub Kosong()
Label1 = ""
Label2 = ""
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
End Sub

Private Sub text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text1_LostFocus()
If Text1 = "" Then Exit Sub
Text1 = Digit(10, Text1)
SBukti = "Select NoBukti from G004 where NoBukti = '" + Trim(Text1) + "'"
Set RBukti = RDCO.OpenResultset(SBukti, rdOpenDynamic, rdConcurRowVer)
If RBukti.RowCount <> 0 Then
    Text1.SetFocus
    MsgBox "NOMOR VOUCHER TELAH DIGUNAKAN", vbInformation, "NO. BUKTI"
    Text1 = ""
    Exit Sub
End If
RBukti.Close
Set RBukti = Nothing
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text2_LostFocus()
If Text2 = "" Then Exit Sub

If Text2 = "0000000" Then
    MsgBox "SGL TIDAK UNTUK TRANSAKSI", vbCritical, "WARNING"
    Text2 = ""
    Text2.SetFocus
    Exit Sub
End If

If Trim(Text2) = Trim(Text3) Then
    Text2.SetFocus
    MsgBox "SGL DEBET TIDAK BOLEH SAMA DENGAN SGL CREDIT", vbCritical, "SGL SAMA"
    Text2 = ""
    Exit Sub
End If

SDebet = "Select NamaSl from G003 where CodeSL = '" + Trim(Text2) + "'"
Set RDebet = RDCO.OpenResultset(SDebet, rdOpenDynamic, rdConcurRowVer)
If RDebet.RowCount <> 0 Then
    Label1 = RDebet("NamaSL")
Else
    Text2.SetFocus
    MsgBox "NOMOR SUB GL BELUM TERDAFTAR", vbInformation, "SGL DEBET"
    Text2 = ""
    Exit Sub
End If
RDebet.Close
Set RDebet = Nothing
End Sub

Private Sub text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text3_LostFocus()
If Text3 = "" Then Exit Sub

If Text3 = "0000000" Then
    MsgBox "SGL TIDAK UNTUK TRANSAKSI", vbCritical, "WARNING"
    Text3 = ""
    Text3.SetFocus
    Exit Sub
End If

If Trim(Text3) = Trim(Text2) Then
    Text3.SetFocus
    MsgBox "SGL CREDIT TIDAK BOLEH SAMA DENGAN SGL DEBET", vbCritical, "SGL SAMA"
    Text3 = ""
    Exit Sub
End If

SCredit = "Select NamaSl from G003 where CodeSL = '" + Trim(Text3) + "'"
Set RCredit = RDCO.OpenResultset(SCredit, rdOpenDynamic, rdConcurRowVer)
If RCredit.RowCount <> 0 Then
    Label2 = RCredit("NamaSL")
Else
    Text3.SetFocus
    MsgBox "NOMOR SUB GL BELUM TERDAFTAR", vbInformation, "SGL CREDIT"
    Text3 = ""
    Exit Sub
End If
RCredit.Close
Set RCredit = Nothing
End Sub

Private Sub text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text4_LostFocus()
If Text4 = "" Then Exit Sub
If Not IsNumeric(Text4) Then
    Text4.SetFocus
    MsgBox "NOMINAL HARUS TYPE ANGKA", vbCritical, "TYPE DATA SALAH"
    Text4 = ""
    Exit Sub
End If
Text4 = Format(Text4, "##,###.00")
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text5_LostFocus()
Text5 = Format(Text5, ">")
End Sub

Private Sub TmbClose_Click()
Unload Me
End Sub

Private Sub TmbSave_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Then
    MsgBox "MASIH ADA DATA YANG KOSONG", vbCritical, "DATA TIDAK BOLEH KOSONG"
    Exit Sub
End If

If STS_Biaya = 1 Then
    Call Simpan
    Call Simpan2
Else
    Call Simpan
End If


'Call Cetak
Call Kosong
Text1.SetFocus

STS_Biaya = 0

Unload Me
G003.Show 1
End Sub

Private Sub CETAK()
SVal = "Select NoBukti from G004 where NoBukti = '" + Text1 + "'"
Set RVal = RDCO.OpenResultset(SVal, rdOpenDynamic, rdConcurRowVer)
If RVal.RowCount <> 0 Then
    Printer.CurrentX = 3000
    Printer.CurrentY = 100
    Printer.Print Tab(20); "D : "; Text2; "  ("; Label1; ")"
    Printer.Print Tab(20); "C : "; Text3; "  ("; Label2; ")"
    Printer.Print Tab(20); Text5; "  Rp. "; Text4
    Printer.Print Tab(20); Operator; "  "; Tanggal
    Printer.EndDoc
End If
RVal.Close
Set RVal = Nothing
End Sub

Private Sub Simpan()
Dim Tanya
Tanya = MsgBox("YAKIN PROSES TRANSAKSI GL ?", vbOKCancel, "KONFIRMASI")
If Tanya = vbCancel Then Exit Sub
    Call JurnalD
    Call JurnalC
    Call LabaRugi
End Sub

Private Sub Simpan2()
If Me.Caption = "INPUT BIAYA BBN" Then
    SSave = "Select * From M001 where RANGKA = '" + Trim(Text6) + "' and MESIN = '" + Trim(Text7) + "'"
    Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
    RSave.EDIT
        RSave("BBN") = CCur(Text4)
    RSave.Update
    RSave.Close
    Set RSave = Nothing
ElseIf Me.Caption = "INPUT BIAYA KACAB" Then
    SSave = "Select * From M001 where RANGKA = '" + Trim(Text6) + "' and MESIN = '" + Trim(Text7) + "'"
    Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
    RSave.EDIT
        RSave("KACAB") = CCur(Text4)
    RSave.Update
    RSave.Close
    Set RSave = Nothing
ElseIf Me.Caption = "INPUT BIAYA BROKER" Then
    SSave = "Select * From M001 where RANGKA = '" + Trim(Text6) + "' and MESIN = '" + Trim(Text7) + "'"
    Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
    RSave.EDIT
        RSave("BROKER") = CCur(Text4)
    RSave.Update
    RSave.Close
    Set RSave = Nothing
End If
End Sub

Private Sub JurnalD()
Dim Enak

SSimpan2 = "Select * From G003 where codesl = '" + Trim(Text2) + "'"
Set RSimpan2 = RDCO.OpenResultset(SSimpan2, rdOpenDynamic, rdConcurRowVer)
Enak = RSimpan2("Posisi")
    If RSimpan2("Posisi") = "D" Then
        RSimpan2.EDIT
        A = CCur(RSimpan2("mutasid")) + CCur(Text4)
        RSimpan2("mutasid") = A
        RSimpan2("saldo") = CCur(RSimpan2("SaldoAwal")) + CCur(RSimpan2("mutasid")) - CCur(RSimpan2("mutasic"))
    ElseIf RSimpan2("Posisi") = "C" Then
        RSimpan2.EDIT
        A = CCur(RSimpan2("mutasid")) + CCur(Text4)
        RSimpan2("mutasid") = A
        RSimpan2("saldo") = CCur(RSimpan2("SaldoAwal")) - CCur(RSimpan2("mutasid")) + CCur(RSimpan2("mutasic"))
    End If
    
        SSave2 = "Select * From G004"
        Set RSave2 = RDCO.OpenResultset(SSave2, rdOpenDynamic, rdConcurRowVer)
        RSave2.AddNew
            RSave2("CodeCab") = CodeCab
            RSave2("Codesl") = Trim(Text2)
            RSave2("NamaSL") = Trim(Label1)
            RSave2("NoBukti") = Trim(Text1)
            RSave2("Keterangan") = Trim(Text5)
            RSave2("NominalD") = CCur(Text4)
            RSave2("NominalC") = 0
            RSave2("Tanggal") = Tanggal
            RSave2("UserCode") = Operator
            RSave2("Jam") = Time
        RSave2.Update
        RSave2.Close
        Set RSave2 = Nothing
        
RSimpan2.Update

            SDebet2 = "Select * From G005"
            Set RDebet2 = RDCO.OpenResultset(SDebet2, rdOpenKeyset, rdConcurRowVer)
                RDebet2.AddNew
                RDebet2("codecab") = CodeCab
                RDebet2("codesl") = Trim(Text2)
                RDebet2("namasl") = Label1
                RDebet2("nobukti") = Trim(Text1)
                RDebet2("keterangan") = Trim(Text5)
                RDebet2("nominald") = CCur(Text4)
                RDebet2("nominalc") = 0
                RDebet2("saldo") = RSimpan2("SALDO")
                RDebet2("tanggal") = Tanggal
                RDebet2("jam") = Time
                RDebet2("usercode") = Operator
            RDebet2.Update
            RDebet2.Close
            Set RDebet2 = Nothing
            
RSimpan2.Close
Set RSimpan2 = Nothing
End Sub

Private Sub JurnalC()
SSimpan3 = "Select * From G003 where codesl = '" + Trim(Text3) + "'"
Set RSimpan3 = RDCO.OpenResultset(SSimpan3, rdOpenDynamic, rdConcurRowVer)
    If RSimpan3("Posisi") = "D" Then
        RSimpan3.EDIT
        A = CCur(RSimpan3("mutasic")) + CCur(Text4)
        RSimpan3("mutasic") = A
        RSimpan3("saldo") = CCur(RSimpan3("SaldoAwal")) + CCur(RSimpan3("mutasid")) - CCur(RSimpan3("mutasic"))
    ElseIf RSimpan3("Posisi") = "C" Then
        RSimpan3.EDIT
        A = CCur(RSimpan3("mutasic")) + CCur(Text4)
        RSimpan3("mutasic") = A
        RSimpan3("saldo") = CCur(RSimpan3("SaldoAwal")) - CCur(RSimpan3("mutasid")) + CCur(RSimpan3("mutasic"))
    End If
       
        SSave3 = "Select * From G004"
        Set RSave3 = RDCO.OpenResultset(SSave3, rdOpenDynamic, rdConcurRowVer)
        RSave3.AddNew
            RSave3("CodeCab") = CodeCab
            RSave3("Codesl") = Trim(Text3)
            RSave3("Namasl") = Label1
            RSave3("NoBukti") = Trim(Text1)
            RSave3("Keterangan") = Trim(Text5)
            RSave3("NominalD") = 0
            RSave3("NominalC") = CCur(Text4)
            RSave3("Tanggal") = Tanggal
            RSave3("UserCode") = Operator
            RSave3("Jam") = Time
        RSave3.Update
        RSave3.Close
        Set RSave3 = Nothing
        
RSimpan3.Update

            SCredit2 = "Select * From G005"
            Set RCredit2 = RDCO.OpenResultset(SCredit2, rdOpenKeyset, rdConcurRowVer)
            RCredit2.AddNew
                RCredit2("codecab") = CodeCab
                RCredit2("codesl") = Trim(Text3)
                RCredit2("namasl") = Label2
                RCredit2("nobukti") = Trim(Text1)
                RCredit2("keterangan") = Trim(Text5)
                RCredit2("nominald") = 0
                RCredit2("nominalc") = CCur(Text4)
                RCredit2("saldo") = RSimpan3("SALDO")
                RCredit2("tanggal") = Tanggal
                RCredit2("jam") = Time
                RCredit2("usercode") = Operator
            RCredit2.Update
            RCredit2.Close
            Set RCredit2 = Nothing


RSimpan3.Close
Set RSimpan3 = Nothing
End Sub

Private Sub LabaRugi()
SSimpan5 = "Select * From LabaRugi"
Set RSimpan5 = RDCO.OpenResultset(SSimpan5, rdOpenDynamic, rdConcurRowVer)
    SaldoD = CCur(RSimpan5("sumofmutasid"))
    SaldoC = CCur(RSimpan5("sumofmutasic"))
    
    SSave5 = "Select * From G003 where Posisi = 'L'"
    Set RSave5 = RDCO.OpenResultset(SSave5, rdOpenKeyset, rdConcurRowVer)
    Saldo = RSave5("saldoawal")
    RSave5.EDIT
        RSave5("mutasid") = SaldoD
        RSave5("mutasic") = SaldoC
        RSave5("saldo") = CCur(RSave5("SaldoAwal")) - CCur(RSave5("mutasid")) + CCur(RSave5("mutasic"))
    RSave5.Update
    RSave5.Close
    Set RSave5 = Nothing

RSimpan5.Close
Set RSimpan5 = Nothing
End Sub


