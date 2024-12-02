VERSION 5.00
Begin VB.Form G003LAMA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TRANSAKSI GL TO GL"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   5895
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   1522
      MaxLength       =   10
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   113
      Width           =   1320
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   1522
      MaxLength       =   7
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   563
      Width           =   1005
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   1522
      MaxLength       =   7
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   1418
      Width           =   1005
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   1522
      TabIndex        =   3
      Text            =   "Text4"
      Top             =   2318
      Width           =   1950
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   1522
      MaxLength       =   35
      TabIndex        =   4
      Text            =   "Text5"
      Top             =   2768
      Width           =   4200
   End
   Begin VB.CommandButton TmbSave 
      Caption         =   "SAVE"
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
      Left            =   2467
      TabIndex        =   5
      Top             =   3465
      Width           =   960
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
      Height          =   360
      Left            =   1552
      TabIndex        =   14
      Top             =   1013
      Width           =   4200
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label2"
      Height          =   360
      Left            =   1552
      TabIndex        =   13
      Top             =   1868
      Width           =   4200
   End
   Begin VB.Label Label3 
      Caption         =   "NO. BUKTI"
      Height          =   360
      Left            =   142
      TabIndex        =   12
      Top             =   113
      Width           =   1230
   End
   Begin VB.Label Label4 
      Caption         =   "SGL DEBET"
      Height          =   360
      Left            =   142
      TabIndex        =   11
      Top             =   563
      Width           =   1230
   End
   Begin VB.Label Label5 
      Caption         =   "NAMA SGL"
      Height          =   360
      Left            =   142
      TabIndex        =   10
      Top             =   1013
      Width           =   1230
   End
   Begin VB.Label Label6 
      Caption         =   "SGL CREDIT"
      Height          =   360
      Left            =   142
      TabIndex        =   9
      Top             =   1463
      Width           =   1230
   End
   Begin VB.Label Label7 
      Caption         =   "NAMA SGL "
      Height          =   360
      Left            =   142
      TabIndex        =   8
      Top             =   1913
      Width           =   1230
   End
   Begin VB.Label Label8 
      Caption         =   "NOMINAL"
      Height          =   360
      Left            =   142
      TabIndex        =   7
      Top             =   2363
      Width           =   1230
   End
   Begin VB.Label Label9 
      Caption         =   "KETERANGAN"
      Height          =   360
      Left            =   142
      TabIndex        =   6
      Top             =   2813
      Width           =   1230
   End
   Begin VB.Line Line1 
      X1              =   127
      X2              =   5707
      Y1              =   3308
      Y2              =   3308
   End
End
Attribute VB_Name = "G003LAMA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Dim SAS, Posisi As String

Private RSLNO, RSL, RSLUser, RSimpan, RSimpan2, RSimpan3, RSimpan4, RSimpan5, RSave, RSave2, RSave3, RSave4, RSave5, REdit As rdoResultset
Private SQLNO, SQL, SQLUser, SSimpan, SSimpan2, SSimpan3, SSimpan4, SSimpan5, SSave, SSave2, SSave3, SSave4, SSave5, SEdit As String

Private Kirim, RBukti, RDebet, RCredit, RLock, RVal As rdoResultset
Private SBukti, SDebet, SCredit, SLock, SVal As String

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=DEALER", rdDriverNoPrompt, False, CN)

Call Kosong
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

Private Sub Text1_KeyPress(KeyAscii As Integer)
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

Private Sub text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text2_LostFocus()
If Text2 = "" Then Exit Sub
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

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text3_LostFocus()
If Text3 = "" Then Exit Sub
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

Private Sub Text4_KeyPress(KeyAscii As Integer)
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
Call Simpan
'Call Cetak
Call Kosong
Text1.SetFocus
End Sub

Private Sub Cetak()
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
    'Call JurnalC
    'Call LabaRugi
End Sub

Private Sub JurnalD()
SSimpan2 = "Select * From G003 where codesl = '" + Trim(Text2) + "'"
Set RSimpan2 = RDCO.OpenResultset(SSimpan2, rdOpenKeyset, rdConcurRowVer)
    If RSimpan2("Posisi") = "D" Then
MsgBox "POSISI D", vbCritical, ""
'        SAS = CCur(RSimpan2("mutasid")) + CCur(Text4)
'        RSimpan.Edit
'        RSimpan2("mutasid") = SAS
'        RSimpan2("saldo") = CCur(RSimpan2("SaldoAwal")) + CCur(RSimpan2("mutasid")) - CCur(RSimpan2("mutasic"))
    ElseIf RSimpan2("Posisi") = "C" Then
MsgBox "POSISI C", vbCritical, ""
'        SAS = CCur(RSimpan2("mutasid")) + CCur(Text4)
'        RSimpan2.Edit
'        RSimpan2("mutasid") = SAS
'        RSimpan2("saldo") = CCur(RSimpan2("SaldoAwal")) - CCur(RSimpan2("mutasid")) + CCur(RSimpan2("mutasic"))
    End If
    
'        SSave2 = "Select * From G004"
'        Set RSave2 = RDCO.OpenResultset(SSave2, rdOpenDynamic, rdConcurRowVer)
'        RSave2.AddNew
'            RSave2("CodeCab") = CodeCab
'            RSave2("Codesl") = Trim(Text1)
'            RSave2("NamaSL") = Trim(Label1)
'            RSave2("NoBukti") = Trim(Text1)
'            RSave2("Keterangan") = Trim(Text5)
'            RSave2("NominalD") = CCur(Text4)
'            RSave2("NominalC") = 0
'            RSave2("Tanggal") = Tanggal
'            RSave2("UserCode") = Operator
'            RSave2("Jam") = Time
'        RSave2.Update
'        RSave2.Close
'        Set RSave2 = Nothing
        
'RSimpan2.Update
RSimpan2.Close
Set RSimpan2 = Nothing
End Sub

Private Sub JurnalC()
SSimpan3 = "Select * From G003 where codesl = '" + Trim(Text3) + "'"
Set RSimpan3 = RDCO.OpenResultset(SSimpan3, rdOpenDynamic, rdConcurRowVer)
    If RSimpan3("S150") = "D" Then
        RSimpan3.Edit
        A = CCur(RSimpan3("mutasic")) + CCur(Text3)
        RSimpan3("mutasic") = A
        RSimpan3("saldo") = CCur(RSimpan3("SaldoAwal")) + CCur(RSimpan3("mutasid")) - CCur(RSimpan3("mutasic"))
    ElseIf RSimpan3("S150") = "C" Then
        RSimpan3.Edit
        A = CCur(RSimpan3("mutasic")) + CCur(Text3)
        RSimpan3("mutasic") = A
        RSimpan3("saldo") = CCur(RSimpan3("SaldoAwal")) - CCur(RSimpan3("mutasid")) + CCur(RSimpan3("mutasic"))
    End If
       
        SSave3 = "Select * From G004"
        Set RSave3 = RDCO.OpenResultset(SSave3, rdOpenDynamic, rdConcurRowVer)
        RSave3.AddNew
            RSave3("CodeCab") = CodeCab
            RSave3("Codesl") = Trim(Text3)
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
RSimpan3.Close
Set RSimpan3 = Nothing
End Sub

Private Sub LabaRugi()
SSimpan5 = "Select * From LabaRugi"
Set RSimpan5 = RDCO.OpenResultset(SSimpan5, rdOpenDynamic, rdConcurRowVer)
    SaldoD = CCur(RSimpan5("sumofmutasid"))
    SaldoC = CCur(RSimpan5("sumofmutasic"))
    
    SSave5 = "Select * From G003 where S150 = 'LR'"
    Set RSave5 = RDCO.OpenResultset(SSave5, rdOpenDynamic, rdConcurRowVer)
    Saldo = RSave5("saldoawal")
    RSave5.Edit
        RSave5("mutasid") = SaldoD
        RSave5("mutasic") = SaldoC
        RSave5("saldo") = CCur(RSave5("SaldoAwal")) - CCur(RSave5("mutasid")) + CCur(RSave5("mutasic"))
    RSave5.Update
    RSave5.Close
    Set RSave5 = Nothing

RSimpan5.Close
Set RSimpan5 = Nothing
End Sub


