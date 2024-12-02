VERSION 5.00
Begin VB.Form E001 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "END OFF DAY PROCESS"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3540
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   3540
   StartUpPosition =   2  'CenterScreen
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
      Left            =   2273
      TabIndex        =   8
      Top             =   2235
      Width           =   1005
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Height          =   360
      Left            =   1943
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1065
      Width           =   1230
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PROSES"
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
      Left            =   263
      TabIndex        =   0
      Top             =   2265
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "END OFF DAY "
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
      Left            =   188
      TabIndex        =   7
      Top             =   165
      Width           =   1365
   End
   Begin VB.Label Label2 
      Caption         =   "BEGIN OFF DAY"
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
      Left            =   188
      TabIndex        =   6
      Top             =   1695
      Width           =   1320
   End
   Begin VB.Label Label3 
      Caption         =   "TANGGAL"
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
      Left            =   548
      TabIndex        =   5
      Top             =   570
      Width           =   1140
   End
   Begin VB.Label Label4 
      Caption         =   "S/D TANGGAL"
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
      Left            =   548
      TabIndex        =   4
      Top             =   1065
      Width           =   1140
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label5"
      Height          =   330
      Left            =   1943
      TabIndex        =   3
      Top             =   570
      Width           =   1230
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1943
      TabIndex        =   2
      Top             =   1695
      Width           =   1230
   End
   Begin VB.Line Line1 
      X1              =   188
      X2              =   3338
      Y1              =   1605
      Y2              =   1605
   End
End
Attribute VB_Name = "E001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private RSL, RSLUser, RCari, RCari2, RCari3, RCari4, RCari5, RSave, RSave2, RSave3, RSave4, RSave5, REdit As rdoResultset
Private SQL, SQLUser, SCari, SCari2, SCari3, SCari4, SCari5, SSave, SSave2, SSave3, SSave4, SSave5, SEdit As String

Private RCOPY, RCOPY2, RCOPY3, RCOPY4, RCOPY5, RCOPY6, RCOPY7, RCOPY8, RCOPY9, RCOPY10 As rdoResultset
Private SCOPY, SCOPY2, SCOPY3, SCOPY4, SCOPY5, SCOPY6, SCOPY7, SCOPY8, SCOPY9, SCOPY10 As String

Private Sub Command1_Click()
Dim Tanya

Tanya = MsgBox("ANDA YAKIN PROSES EOD S/D TANGGAL " + Text1 + "?", vbOKCancel, "END OFF DAY PROCESS")
If Tanya = vbCancel Then Exit Sub

If Day(Label5) = 31 And Month(Label5) = 12 Then
    
    Call LabaRugi
    Call COPYG003
    Call UPDATE_B003
    Call UPDATE_B003A
    Call UPDATE_G003
    Call UPDATE_A001
    Call UPDATE_C013
    
    MsgBox "JALANKAN PROSES TUTUP TAHUN", vbInformation, "END OFF YEAR PROCESS"
    Call LABATAHUNLALU

Else

    Call LabaRugi
    Call COPYG003
    Call UPDATE_B003
    Call UPDATE_B003A
    Call UPDATE_G003
    Call UPDATE_A001
    Call UPDATE_C013
    
End If



MsgBox "PROSES BOD SELESAI", vbInformation, "BEGIN OFF DAY PROCESS"
End
End Sub

Private Sub Command2_Click()
Unload Me
LOGIN.Show
End Sub

Private Sub Form_Load()
MAINSALE.Hide

Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=APOTIK", rdDriverNoPrompt, False, CN)

Label5 = Tanggal
Text1 = Tanggal
Label6 = DateAdd("d", 1, Text1)

End Sub

Private Sub text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text1_LostFocus()
If Text1 = "" Then
    Text1.SetFocus
    MsgBox "TANGGAL AKHIR TIDAK BOLEH KOSONG", vbCritical, "DATA KOSONG"
    Exit Sub
End If

If Not IsDate(Text1) Then
    Text1.SetFocus
    MsgBox "DATA BUKAN TANGGAL", vbCritical, "TYPE DATA SALAH"
    Exit Sub
End If

If DateValue(Text1) < DateValue(Label5) Then
    Text1.SetFocus
    MsgBox "TANGGAL AKHIR HARUS LEBIH BESAR atau SAMA DENGAN TANGGAL AWAL", vbCritical, "TANGGAL AKHIR TIDAK BOLEH LEBIH BESAR"
    Exit Sub
End If

Label6 = DateAdd("d", 1, Text1)
End Sub

Private Sub LabaRugi()
SCari5 = "Select * From LabaRugi"
Set RCari5 = RDCO.OpenResultset(SCari5, rdOpenDynamic, rdConcurRowVer)
    SaldoD = CCur(RCari5("SumOfmutasid"))
    SaldoC = CCur(RCari5("SumOfmutasic"))

    SSave5 = "Select * From G003 where Posisi = 'L'"
    Set RSave5 = RDCO.OpenResultset(SSave5, rdOpenDynamic, rdConcurRowVer)
    Saldo = RSave5("saldoawal")
    RSave5.EDIT
        RSave5("mutasid") = SaldoD
        RSave5("mutasic") = SaldoC
        RSave5("saldo") = CCur(RSave5("SaldoAwal")) - CCur(RSave5("mutasid")) + CCur(RSave5("mutasic"))
    RSave5.Update
    RSave5.Close
    Set RSave5 = Nothing

RCari5.Close
Set RCari5 = Nothing
End Sub

Private Sub COPYG003()
SCOPY = "Select * From G003"
Set RCOPY = RDCO.OpenResultset(SCOPY, rdOpenDynamic, rdConcurRowVer)
RCOPY.MoveFirst
Do While Not RCOPY.EOF

    SCOPY2 = "Select * From G003A"
    Set RCOPY2 = RDCO.OpenResultset(SCOPY2, rdOpenDynamic, rdConcurRowVer)
    RCOPY2.AddNew
        RCOPY2("CodeCab") = RCOPY("CodeCab")
        RCOPY2("CodeSGL") = RCOPY("CodeSGL")
        RCOPY2("CodeSL") = RCOPY("CodeSL")
        RCOPY2("NamaSL") = RCOPY("NamaSL")
        RCOPY2("SaldoAwal") = RCOPY("SaldoAwal")
        RCOPY2("MutasiD") = RCOPY("MutasiD")
        RCOPY2("MutasiC") = RCOPY("MutasiC")
        RCOPY2("Saldo") = RCOPY("Saldo")
        RCOPY2("Tanggal") = Tanggal
    RCOPY2.Update
    RCOPY2.Close
    Set RCOPY2 = Nothing

RCOPY.MoveNext
Loop
RCOPY.Close
Set RCOPY = Nothing
End Sub

Private Sub UPDATE_B003()
SCOPY3 = "Select * From B003"
Set RCOPY3 = RDCO.OpenResultset(SCOPY3, rdOpenDynamic, rdConcurRowVer)
RCOPY3.MoveFirst
Do While Not RCOPY3.EOF
    RCOPY3.EDIT
    RCOPY3("JML_AWAL") = RCOPY3("JML_AKHIR")
    RCOPY3("JML_DBT") = 0
    RCOPY3("JML_CRD") = 0
    RCOPY3.Update

RCOPY3.MoveNext
Loop
RCOPY3.Close
Set RCOPY3 = Nothing
End Sub

Private Sub UPDATE_B003A()
SCOPY3 = "Select * From B003A"
Set RCOPY3 = RDCO.OpenResultset(SCOPY3, rdOpenDynamic, rdConcurRowVer)
RCOPY3.MoveFirst
Do While Not RCOPY3.EOF
    RCOPY3.EDIT
    RCOPY3("SALDOAWAL") = RCOPY3("SALDO")
    RCOPY3("MUTASID") = 0
    RCOPY3("MUTASIC") = 0
    RCOPY3.Update

RCOPY3.MoveNext
Loop
RCOPY3.Close
Set RCOPY3 = Nothing
End Sub

Private Sub UPDATE_G003()
SCOPY4 = "Select * From G003"
Set RCOPY4 = RDCO.OpenResultset(SCOPY4, rdOpenDynamic, rdConcurRowVer)
RCOPY4.MoveFirst
Do While Not RCOPY4.EOF
    RCOPY4.EDIT
    RCOPY4("SaldoAwal") = RCOPY4("Saldo")
    RCOPY4("MutasiD") = 0
    RCOPY4("MutasiC") = 0
    RCOPY4("Tanggal") = DateValue(Label6)
    RCOPY4.Update
RCOPY4.MoveNext
Loop
RCOPY4.Close
Set RCOPY4 = Nothing
End Sub

Private Sub UPDATE_A001()
SCOPY5 = "Select * From A001"
Set RCOPY5 = RDCO.OpenResultset(SCOPY5, rdOpenDynamic, rdConcurRowVer)
RCOPY5.EDIT
RCOPY5("Tanggal") = DateValue(Label6)
RCOPY5("SEOD") = 0
RCOPY5.Update
RCOPY5.Close
Set RCOPY5 = Nothing
End Sub

Private Sub UPDATE_C013()
SCOPY5 = "Select * From C013"
Set RCOPY5 = RDCO.OpenResultset(SCOPY5, rdOpenDynamic, rdConcurRowVer)
RCOPY5.EDIT
RCOPY5("STATUS") = 0
RCOPY5("NOBELI") = 0
RCOPY5("NOJUAL") = 0
RCOPY5("NOPELAYANAN") = 0
RCOPY5.Update
RCOPY5.Close
Set RCOPY5 = Nothing
End Sub

Private Sub LABATAHUNLALU()
SCari = "Select * from G003 where SCAR2 = '2'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
Do Until RCari.EOF
    
    SSave = "Select * From G005"
    Set RSave = RDCO.OpenResultset(SSave, rdOpenKeyset, rdConcurRowVer)
    RSave.AddNew
        RSave("codecab") = CodeCab
        RSave("codesl") = RCari("CODESL")
        RSave("namasl") = RCari("NAMASL")
        RSave("nobukti") = "KONVERT"
        RSave("keterangan") = "KONVERT LABA KE TAHUN LALU"
        
            If RCari("POSISI") = "C" Then
                RSave("nominald") = CCur(RCari("SALDO"))
                RSave("nominalc") = 0
                RSave("saldo") = 0
            ElseIf RCari("POSISI") = "D" Then
                RSave("nominald") = 0
                RSave("nominalc") = CCur(RCari("SALDO"))
                RSave("saldo") = 0
            End If
        
        RSave("tanggal") = Tanggal
        RSave("jam") = Date
        RSave("usercode") = "SYSTEM"
    RSave.Update
    RSave.Close
    Set RSave = Nothing

    RCari.EDIT
        RCari("saldoawal") = 0
        RCari("mutasid") = 0
        RCari("mutasic") = 0
        RCari("saldo") = 0
    RCari.Update
    
RCari.MoveNext
Loop
RCari.Close
Set RCari = Nothing

SCari = "Select * from G003 where CODESL = '7001001'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)

    SSave = "Select * From G003 where CODESL = '7002001'"
    Set RSave = RDCO.OpenResultset(SSave, rdOpenKeyset, rdConcurRowVer)
    RSave.EDIT
        RSave("saldoawal") = CCur(RCari("SALDOAWAL")) + CCur(RSave("saldo"))
        RSave("mutasid") = CCur(RCari("MUTASID")) + CCur(RSave("mutasid"))
        RSave("mutasic") = CCur(RCari("MUTASIC")) + CCur(RSave("mutasic"))
        RSave("saldo") = CCur(RCari("SALDO")) + CCur(RSave("saldo"))
    RSave.Update
    RSave.Close
    Set RSave = Nothing

RCari.EDIT
    RCari("saldoawal") = 0
    RCari("mutasid") = 0
    RCari("mutasid") = 0
    RCari("saldo") = 0

RCari.Update
RCari.Close
Set RCari = Nothing

End Sub
