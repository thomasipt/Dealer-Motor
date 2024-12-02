VERSION 5.00
Begin VB.Form H003 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PENCAIRAN HUTANG"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7635
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   7635
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   1747
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2925
      Width           =   1320
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   1747
      TabIndex        =   4
      Text            =   "Text4"
      Top             =   2115
      Width           =   1860
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   1747
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   173
      Width           =   1320
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   1747
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   630
      Width           =   1320
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   1747
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   1125
      Width           =   1320
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   1747
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   1620
      Width           =   5730
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
      Left            =   412
      TabIndex        =   6
      Top             =   3735
      Width           =   1000
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
      Left            =   6262
      TabIndex        =   7
      Top             =   3712
      Width           =   960
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   780
      Left            =   -90
      Top             =   2655
      Width           =   8070
   End
   Begin VB.Label Label8 
      Caption         =   "SGL NON TUNAI"
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
      Left            =   157
      TabIndex        =   15
      Top             =   2932
      Width           =   1410
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
      Height          =   300
      Left            =   3225
      TabIndex        =   14
      Top             =   2925
      Width           =   4245
   End
   Begin VB.Label Label4 
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
      Height          =   300
      Left            =   157
      TabIndex        =   13
      Top             =   1650
      Width           =   1410
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label2"
      Height          =   300
      Left            =   3195
      TabIndex        =   12
      Top             =   180
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "KODE"
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
      Left            =   157
      TabIndex        =   11
      Top             =   180
      Width           =   1410
   End
   Begin VB.Label Label5 
      Caption         =   "NO. HUTANG"
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
      Left            =   157
      TabIndex        =   10
      Top             =   660
      Width           =   1410
   End
   Begin VB.Label Label6 
      Caption         =   "NAMA"
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
      Left            =   157
      TabIndex        =   9
      Top             =   1155
      Width           =   1410
   End
   Begin VB.Label Label7 
      Caption         =   "PLAFON"
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
      Left            =   157
      TabIndex        =   8
      Top             =   2145
      Width           =   1410
   End
End
Attribute VB_Name = "H003"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private RDebet, RDebet2, RDebet3, RGele, R3GP, RPin, RToket, RToge As rdoResultset
Private SDebet, SDebet2, SDebet3, SGele, S3GP, SPin, SToket, SToge As String

Private RBahan, RBahan2, RBahan3 As rdoResultset
Private SBahan, SBahan2, SBahan3 As String

Private RBahannya, RBahannya2, RBahannya3 As rdoResultset
Private SBahannya, SBahannya2, SBahannya3 As String

Private RCBiaya, RCBiaya2, RCBiaya3 As rdoResultset
Private SCBiaya, SCBiaya2, SCBiaya3 As String

Private RDBank2, RDBank3 As rdoResultset
Private SDBank2, SDBank3 As String

Private RPDPT, RPDPT2, RPDPT3 As rdoResultset
Private SPDPT, SPDPT2, SPDPT3 As String

Private ROyen, RNovi, RUhAh As rdoResultset
Private SOyen, DNovi, SUhAh As String

Private Cash, Bank, KHutang

Private Sub Combo1_Click()
Call Combo1_LostFocus
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Combo1_LostFocus()
If Combo1 = "" Then Exit Sub

SCari = "Select * From H001 where KODE = '" + Trim(Combo1) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    Label2 = RCari("NAMA")
End If
RCari.Close
Set RCari = Nothing

End Sub

Private Sub Combo2_Click()
Call Combo2_LostFocus
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Combo2_LostFocus()
If Combo2 = "" Then Exit Sub

SCari = "Select * From G003 where CODESL = '" + Trim(Combo2) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    Label1 = RCari("NAMASL")
End If
RCari.Close
Set RCari = Nothing

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=DEALER", rdDriverNoPrompt, False, CN)

SGol = "Select * From H001 order by KODE Asc"
Set RGol = RDCO.OpenResultset(SGol, rdOpenDynamic, rdConcurRowVer)
If RGol.RowCount <> 0 Then
    RGol.MoveFirst
    Do While Not RGol.EOF
        Combo1.AddItem RGol("KODE")
    RGol.MoveNext
    Loop
End If
RGol.Close
Set RGol = Nothing
Combo1.ListIndex = 0

SSGL = "Select * From G003 where CODESGL = '1001110' Order by CODESL"
Set RSGL = RDCO.OpenResultset(SSGL, rdOpenDynamic, rdConcurRowVer)
If RSGL.RowCount <> 0 Then
    RSGL.MoveFirst
    Do While Not RSGL.EOF
        Combo2.AddItem RSGL("CODESL")
    RSGL.MoveNext
    Loop
End If

RSGL.Close
Set RSGL = Nothing
Combo2.ListIndex = 0

ClearTextBoxes Me
Label1 = ""
Label2 = ""

End Sub

Private Sub text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text1_LostFocus()
Dim Tanya
If Text1 = "" Then Exit Sub
SKode = "Select * From H002 where NOMOR_HUTANG = '" + Text1 + "'"
Set RKode = RDCO.OpenResultset(SKode, rdOpenDynamic, rdOpenKeyset)
If RKode.RowCount <> 0 Then
    Tanya = MsgBox("NOMOR HUTANG TELAH TERDAFTAR", vbCritical, "KONFIRMASI")
    Text1 = ""
    Text1.SetFocus
End If
RKode.Close
Set RKode = Nothing
Text1 = Format(Text1, ">")
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text2_LostFocus()
Text2 = Format(Text2, ">")
End Sub

Private Sub text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text3_LostFocus()
Text3 = Format(Text3, ">")
End Sub

Private Sub text4_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Text4_LostFocus()
Text4 = Format(Text4, "##,###.00")
End Sub

Private Sub Tmb_Save_Click()
Dim Tanya

If Label1 = "" Or Label2 = "" Or Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Then
    MsgBox "MASIH ADA DATA YANG KOSONG", vbSystemModal, "KONFIRMASI"
    Exit Sub
Else
    Tanya = MsgBox("ANDA YAKIN MELAKUKAN TRANSAKSI HUTANG", vbOKCancel, "KONFIRMASI")
        If Tanya = vbOK Then
            Call JurnalKas
            Call JurnalHutang
        ElseIf Tanya = vbCancel Then
            Exit Sub
        End If
End If

Unload Me
H002.Show

End Sub

Private Sub JurnalKas()
SDBank2 = "Select * From G003 where CODESL='" + Trim(Combo2) + "'"
Set RDBank2 = RDCO.OpenResultset(SDBank2, rdOpenKeyset, rdConcurRowVer)

MMUTASID = RDBank2("mutasid") + CCur(Text4)
SSALDO = RDBank2("saldo") + CCur(Text4)
RDBank2.EDIT
    RDBank2("mutasid") = CCur(MMUTASID)
    RDBank2("saldo") = CCur(SSALDO)
        SDBank3 = "Select * From G005"
        Set RDBank3 = RDCO.OpenResultset(SDBank3, rdOpenKeyset, rdConcurRowVer)
        RDBank3.AddNew
            RDBank3("codecab") = CodeCab
            RDBank3("codesl") = Trim(Combo2)
            RDBank3("namasl") = RDBank2("NamaSL")
            RDBank3("nobukti") = Trim(Text1)
            RDBank3("keterangan") = Trim(Combo1) + "." + Trim(Text2) + "." + Trim(Text3)
            RDBank3("nominald") = CCur(Text4)
            RDBank3("nominalc") = 0
            RDBank3("saldo") = SSALDO
            RDBank3("tanggal") = Tanggal
            RDBank3("jam") = Time
            RDBank3("usercode") = Operator
        RDBank3.Update
        RDBank3.Close
        Set RDBank3 = Nothing
RDBank2.Update
RDBank2.Close
Set RDBank2 = Nothing
End Sub

Private Sub JurnalHutang()
SPDPT = "Select * From H001 where KODE ='" + Trim(Combo1) + "'"
Set RPDPT = RDCO.OpenResultset(SPDPT, rdOpenKeyset, rdConcurRowVer)
    GPIN = RPDPT("SGL_HUTANG")
    
    SPDPT2 = "Select * From G003 where CODESL='" + Trim(GPIN) + "'"
    Set RPDPT2 = RDCO.OpenResultset(SPDPT2, rdOpenKeyset, rdConcurRowVer)
    
    MMUTASID = RPDPT2("mutasic") + CCur(Text4)
    SSALDO = RPDPT2("saldo") + CCur(Text4)
    RPDPT2.EDIT
        RPDPT2("mutasic") = CCur(MMUTASID)
        RPDPT2("saldo") = CCur(SSALDO)
        
            SPDPT3 = "Select * From H002"
            Set RPDPT3 = RDCO.OpenResultset(SPDPT3, rdOpenKeyset, rdConcurRowVer)
            RPDPT3.AddNew
                RPDPT3("kode_hutang") = Trim(Combo1)
                RPDPT3("nomor_hutang") = Trim(Text1)
                RPDPT3("nama") = Trim(Text2)
                RPDPT3("keterangan") = Trim(Text3)
                RPDPT3("plafon") = CCur(Text4)
                RPDPT3("SGL_NONTUNAI") = Trim(Combo2)
                RPDPT3("tanggal") = Tanggal
                RPDPT3("user_code") = Operator
            RPDPT3.Update
            RPDPT3.Close
            Set RPDPT3 = Nothing
            
            SDBank3 = "Select * From G005"
            Set RDBank3 = RDCO.OpenResultset(SDBank3, rdOpenKeyset, rdConcurRowVer)
            RDBank3.AddNew
                RDBank3("codecab") = CodeCab
                RDBank3("codesl") = GPIN
                RDBank3("namasl") = RPDPT2("NamaSL")
                RDBank3("nobukti") = Trim(Text1)
                RDBank3("keterangan") = Trim(Combo1) + "." + Trim(Text2) + "." + Trim(Text3)
                RDBank3("nominald") = 0
                RDBank3("nominalc") = CCur(Text4)
                RDBank3("saldo") = SSALDO
                RDBank3("tanggal") = Tanggal
                RDBank3("jam") = Time
                RDBank3("usercode") = Operator
            RDBank3.Update
            RDBank3.Close
            Set RDBank3 = Nothing
        
    RPDPT2.Update
    RPDPT2.Close
    Set RPDPT2 = Nothing
End Sub
