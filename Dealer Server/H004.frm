VERSION 5.00
Begin VB.Form H004 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PELUNASAN HUTANG"
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
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0C0FF&
      Height          =   360
      Left            =   1747
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "Text6"
      Top             =   2895
      Width           =   1320
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00C0C0FF&
      Height          =   360
      Left            =   1747
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "Text5"
      Top             =   150
      Width           =   1320
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
      TabIndex        =   1
      Top             =   3712
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
      Left            =   412
      TabIndex        =   0
      Top             =   3735
      Width           =   1000
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0C0FF&
      Height          =   360
      Left            =   1747
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   1620
      Width           =   5730
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0C0FF&
      Height          =   360
      Left            =   1747
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   1125
      Width           =   1320
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0FF&
      Height          =   360
      Left            =   1747
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   630
      Width           =   1320
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0FF&
      Height          =   360
      Left            =   1747
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "Text4"
      Top             =   2115
      Width           =   1860
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
      TabIndex        =   13
      Top             =   2145
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
      TabIndex        =   12
      Top             =   1155
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
      TabIndex        =   11
      Top             =   660
      Width           =   1410
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
      TabIndex        =   10
      Top             =   180
      Width           =   1410
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label2"
      Height          =   300
      Left            =   3195
      TabIndex        =   9
      Top             =   180
      Width           =   3255
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
      TabIndex        =   8
      Top             =   1650
      Width           =   1410
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
      Height          =   300
      Left            =   3225
      TabIndex        =   7
      Top             =   2925
      Width           =   4245
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
      TabIndex        =   6
      Top             =   2932
      Width           =   1410
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   780
      Left            =   -218
      Top             =   2700
      Width           =   8070
   End
End
Attribute VB_Name = "H004"
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

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=DEALER", rdDriverNoPrompt, False, CN)

Text1 = NoHutang
Text2 = NamaHutang

Call Cari

End Sub

Private Sub Cari()
SCari = "Select * from V_SALDOHUTANG where NOMOR_HUTANG ='" + Trim(Text1) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    Text5 = Format(RCari("KODE"), ">")
    Label2 = Format(RCari("NAMA"), ">")
    Text1 = Format(RCari("NOMOR_HUTANG"), ">")
    Text3 = Format(RCari("KETERANGAN"), ".")
    Text4 = Format(RCari("PLAFON"), "##,###.00")
    Text6 = Format(RCari("SGL_NONTUNAI"), ">")
    
    SCari2 = "Select * from G003 where CODESL ='" + Trim(Text6) + "'"
    Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenKeyset, rdConcurRowVer)
        Label1 = Format(RCari2("NamaSL"), ">")
    RCari2.Close
    Set RCari2 = Nothing
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub Tmb_Save_Click()
Dim Tanya
Tanya = MsgBox("ANDA YAKIN MELAKUKAN TRANSAKSI HUTANG", vbOKCancel, "KONFIRMASI")
    
    If Tanya = vbOK Then
        Call JurnalKas
        Call JurnalHutang
    ElseIf Tanya = vbCancel Then
        Exit Sub
    End If

Unload Me
H002.Show

End Sub

Private Sub JurnalKas()
SDBank2 = "Select * From G003 where CODESL='" + Trim(Text6) + "'"
Set RDBank2 = RDCO.OpenResultset(SDBank2, rdOpenKeyset, rdConcurRowVer)

MMUTASIC = RDBank2("mutasic") + CCur(Text4)
SSALDO = RDBank2("saldo") - CCur(Text4)
RDBank2.EDIT
    RDBank2("mutasic") = CCur(MMUTASIC)
    RDBank2("saldo") = CCur(SSALDO)
        SDBank3 = "Select * From G005"
        Set RDBank3 = RDCO.OpenResultset(SDBank3, rdOpenKeyset, rdConcurRowVer)
        RDBank3.AddNew
            RDBank3("codecab") = CodeCab
            RDBank3("codesl") = Trim(Text6)
            RDBank3("namasl") = RDBank2("NamaSL")
            RDBank3("nobukti") = "P." + Trim(Text1)
            RDBank3("keterangan") = "P." + Trim(Combo1) + "." + Trim(Text2) + "." + Trim(Text3)
            RDBank3("nominald") = 0
            RDBank3("nominalc") = CCur(Text4)
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
SPDPT = "Select * From H001 where KODE ='" + Trim(Text5) + "'"
Set RPDPT = RDCO.OpenResultset(SPDPT, rdOpenKeyset, rdConcurRowVer)
    GPIN = RPDPT("SGL_HUTANG")
    
    SPDPT2 = "Select * From G003 where CODESL='" + Trim(GPIN) + "'"
    Set RPDPT2 = RDCO.OpenResultset(SPDPT2, rdOpenKeyset, rdConcurRowVer)
    
    MMUTASID = RPDPT2("mutasid") + CCur(Text4)
    SSALDO = RPDPT2("saldo") - CCur(Text4)
    RPDPT2.EDIT
        RPDPT2("mutasid") = CCur(MMUTASID)
        RPDPT2("saldo") = CCur(SSALDO)
        
            SPDPT3 = "Select * From H002 where NOMOR_HUTANG = '" + Trim(Text1) + "'"
            Set RPDPT3 = RDCO.OpenResultset(SPDPT3, rdOpenKeyset, rdConcurRowVer)
            RPDPT3.EDIT
                RPDPT3("STATUS") = "1"
            RPDPT3.Update
            RPDPT3.Close
            Set RPDPT3 = Nothing
            
            SDBank3 = "Select * From G005"
            Set RDBank3 = RDCO.OpenResultset(SDBank3, rdOpenKeyset, rdConcurRowVer)
            RDBank3.AddNew
                RDBank3("codecab") = CodeCab
                RDBank3("codesl") = GPIN
                RDBank3("namasl") = RPDPT2("NamaSL")
                RDBank3("nobukti") = "P." + Trim(Text1)
                RDBank3("keterangan") = "P." + Trim(Combo1) + "." + Trim(Text2) + "." + Trim(Text3)
                RDBank3("nominald") = CCur(Text4)
                RDBank3("nominalc") = 0
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

