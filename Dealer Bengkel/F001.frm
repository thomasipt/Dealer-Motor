VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form F001 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ENTRI FAKTUR"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   8790
   StartUpPosition =   2  'CenterScreen
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
      Left            =   5880
      TabIndex        =   8
      Top             =   4635
      Width           =   1000
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   2569
      TabIndex        =   0
      Text            =   "Text6"
      Top             =   210
      Width           =   2370
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      Caption         =   "Data Sub GL (Pembayaran Non Tunai)"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1320
      Left            =   4095
      TabIndex        =   26
      Top             =   3045
      Width           =   4590
      Begin VB.TextBox Text10 
         BackColor       =   &H00C0C0C0&
         Height          =   360
         Left            =   1530
         TabIndex        =   7
         Text            =   "Text10"
         Top             =   360
         Width           =   1050
      End
      Begin VB.CommandButton InfoGL 
         Caption         =   "Info Code Sub GL"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2700
         TabIndex        =   9
         Top             =   360
         Width           =   1725
      End
      Begin VB.Label InfoGL2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Info Kode Sub GL"
         Height          =   360
         Left            =   2700
         TabIndex        =   30
         Top             =   360
         Width           =   1725
      End
      Begin VB.Label Label28 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label28"
         Height          =   330
         Left            =   1530
         TabIndex        =   29
         Top             =   810
         Width           =   2085
      End
      Begin VB.Label Label14 
         Caption         =   "KODE SUB GL"
         Height          =   330
         Left            =   135
         TabIndex        =   28
         Top             =   405
         Width           =   1455
      End
      Begin VB.Label Label23 
         Caption         =   "NAMA SUB GL"
         Height          =   375
         Left            =   135
         TabIndex        =   27
         Top             =   810
         Width           =   1140
      End
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      Caption         =   "Cara Pembayaran"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2115
      Left            =   105
      TabIndex        =   19
      Top             =   3045
      Width           =   3855
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   360
         Left            =   1320
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   810
         Width           =   1860
      End
      Begin VB.TextBox Text14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   360
         Left            =   1320
         TabIndex        =   6
         Text            =   "Text14"
         Top             =   1275
         Width           =   1860
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label7"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1830
         TabIndex        =   25
         Top             =   1710
         Width           =   1860
      End
      Begin VB.Label Label18 
         Caption         =   "TOTAL DIBAYAR"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   105
         TabIndex        =   24
         Top             =   1800
         Width           =   1410
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label17"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1830
         TabIndex        =   23
         Top             =   360
         Width           =   1860
      End
      Begin VB.Label Label15 
         Caption         =   "TOTAL PEMBELIAN"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   105
         TabIndex        =   22
         Top             =   405
         Width           =   1590
      End
      Begin VB.Label Label11 
         Caption         =   "> TUNAI"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   180
         TabIndex        =   21
         Top             =   855
         Width           =   1230
      End
      Begin VB.Label Label13 
         Caption         =   "> NON TUNAI"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   180
         TabIndex        =   20
         Top             =   1320
         Width           =   1230
      End
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   1763
      MaxLength       =   65
      TabIndex        =   4
      Text            =   "Text5"
      Top             =   2520
      Width           =   5835
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      Left            =   1763
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   1890
      Width           =   3105
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   1763
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   1365
      Width           =   1110
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   2569
      TabIndex        =   1
      Text            =   "Text4"
      Top             =   630
      Width           =   1110
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   2445
      Left            =   90
      TabIndex        =   32
      ToolTipText     =   "Klik untuk edit"
      Top             =   5250
      Width           =   8610
      _ExtentX        =   15187
      _ExtentY        =   4313
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   65280
      BackColorBkg    =   16777152
      GridColor       =   0
      Enabled         =   -1  'True
      TextStyle       =   3
      TextStyleFixed  =   3
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
   Begin MSFlexGridLib.MSFlexGrid GridGL 
      Height          =   2085
      Left            =   4935
      TabIndex        =   31
      Top             =   1260
      Width           =   3750
      _ExtentX        =   6615
      _ExtentY        =   3678
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Label Label6 
      Caption         =   "KETERANGAN"
      Height          =   270
      Left            =   210
      TabIndex        =   18
      Top             =   2565
      Width           =   1365
   End
   Begin VB.Label Label5 
      Caption         =   "NOMINAL            Rp."
      Height          =   270
      Left            =   210
      TabIndex        =   17
      Top             =   1935
      Width           =   1890
   End
   Begin VB.Label Label4 
      Caption         =   "JUMLAH                                                UNIT"
      Height          =   270
      Left            =   210
      TabIndex        =   16
      Top             =   1410
      Width           =   4095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "Label3"
      Height          =   300
      Left            =   6822
      TabIndex        =   15
      Top             =   255
      Width           =   1320
   End
   Begin VB.Label Label2 
      Caption         =   "NO. SYSTEM"
      Height          =   225
      Left            =   5382
      TabIndex        =   14
      Top             =   300
      Width           =   1365
   End
   Begin VB.Label Label1 
      Caption         =   "NO. FAKTUR"
      Height          =   225
      Left            =   649
      TabIndex        =   13
      Top             =   300
      Width           =   1140
   End
   Begin VB.Label Label8 
      Caption         =   "TGL PEMBELIAN"
      Height          =   225
      Left            =   649
      TabIndex        =   12
      Top             =   720
      Width           =   1365
   End
   Begin VB.Label Label9 
      Caption         =   "TGL SISTEM"
      Height          =   225
      Left            =   5382
      TabIndex        =   11
      Top             =   720
      Width           =   1320
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "Label10"
      Height          =   300
      Left            =   6822
      TabIndex        =   10
      Top             =   675
      Width           =   1320
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C00000&
      Height          =   1695
      Left            =   135
      Top             =   1260
      Width           =   8520
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C00000&
      Height          =   960
      Left            =   398
      Shape           =   4  'Rounded Rectangle
      Top             =   105
      Width           =   7995
   End
End
Attribute VB_Name = "F001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private RKode, RSetSTS, RDEBET1, RDebet2, RDebet3, RCREDIT1, RCredit2, RCREDIT3, RFuck, RNo, RSGL, RCari As rdoResultset
Private SKode, SSetSTS, SDEBET1, SDebet2, SDebet3, SCREDIT1, SCredit2, SCREDIT3, SFuck, SNo, SSGL, SCari As String

Private SGLNonKas

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=DEALER", rdDriverNoPrompt, False, CN)

Call NoTrans
Call Kosong
Call Cari_SubGL

Call SiapkanGrid
Call IsiGrid

End Sub

Private Sub NoTrans()
Dim Nomor As Double
Dim InfoNomor As Double

SCari = "Select Top 1 No_System From F001 order by No_System Desc"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    Nomor = Val(RCari("No_System")) + 1
    InfoNomor = Digit(5, Val(RCari("No_System")))
    If Pesan = True Then
        MsgBox "NOMOR TERSIMPAN " + Trim(InfoNomor), vbOKOnly, "DATA TERSIMPAN"
    End If
    Label3 = Digit(7, Nomor)
Else
    Label3 = "0000001"
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub Kosong()
Text1 = 0
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Text14 = 0
Text10 = ""
Label10 = Tanggal
Text4 = Tanggal
Label17 = ""
Label7 = ""
Label28 = ""
Call InfoGL2_Click
Call FrameGLNonAktif
End Sub

Private Sub TblSave_Click()
Call Kosong
Text1.SetFocus
Unload Me
F001.Show
End Sub

Private Sub Text1_GotFocus()
If CCur(Text1) = 0 Then Text1 = ""
End Sub

Private Sub text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text1_LostFocus()
If Text1 = "" Then Text1 = 0
If Not IsNumeric(Text1) Then
    Text1.SetFocus
    MsgBox "NOMINAL PEMBAYARAN HARUS ANGKA", vbCritical, "TYPE DATA SALAH"
    Exit Sub
End If
Text1 = Format(Text1, "##,###.00")
Label7 = Format(CCur(Text1) + CCur(Text14), "##,###.00")
If CCur(Label7) > CCur(Label17) Then
    Text1.SetFocus
    MsgBox "TOTAL PEMBAYARAN TIDAK BOLEH MELEBIHI TOTAL PEMBELIAN", vbCritical, "TOTAL PEMBAYARAN SALAH"
End If
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text10_LostFocus()
If Text10 = "" Then Exit Sub
SSGL = "Select CodeSL, NamaSL From G003 where CodeSL = '" + Trim(Text10) + "'"
Set RSGL = RDCO.OpenResultset(SSGL, rdOpenDynamic, rdConcurRowVer)
If RSGL.RowCount <> 0 Then
    Label28 = RSGL("NamaSL")
Else
    Text10.SetFocus
    MsgBox "KODE SUB GENERAL LEDGER BELUM TERDAFTAR", vbInformation, "KODE GL BELUM TERDAFTAR"
End If
RSGL.Close
Set RSGL = Nothing
End Sub

Private Sub Text14_GotFocus()
If Text14 = 0 Then Text14 = ""
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text14_LostFocus()
If Text14 = "" Then Text14 = 0
If Not IsNumeric(Text14) Then
    Text14.SetFocus
    MsgBox "NOMINAL PEMBAYARAN HARUS ANGKA", vbCritical, "TYPE DATA SALAH"
    Exit Sub
End If
Text14 = Format(Text14, "##,###.00")

Label7 = Format(CCur(Text1) + CCur(Text14), "##,###.00")
If CCur(Label7) > CCur(Label17) Then
    Text14.SetFocus
    MsgBox "TOTAL PEMBAYARAN TIDAK BOLEH MELEBIHI TOTAL PEMBELIAN", vbCritical, "TOTAL PEMBAYARAN SALAH"
    Exit Sub
End If

If CCur(Text14) > 0 Then
    Call FrameGLAktif
    Text10.SetFocus
Else
    Call FrameGLNonAktif
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text3_LostFocus()
If Text3 = "" Then Text3 = 0
Text3 = Format(Text3, "##,###.00")
Label17 = Text3
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text4_LostFocus()
If Text4 = "" Then Exit Sub
If Not IsDate(Text4) Then
    Text4.SetFocus
    MsgBox "TYPE DATA HARUS TANGGAL", vbCritical, "TYPE DATA SALAH"
    Text4 = Tanggal
    Exit Sub
End If
    Text4 = Format(Text4, "DD/MM/YYYY")
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys vbTab
    Text5 = Format(Text5, ">")
End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub FrameGLNonAktif()
Frame2.Enabled = False
Label14.Enabled = False
Text10.Enabled = False
Text10.BackColor = &HC0C0C0
InfoGL.Enabled = False
Label28.Enabled = False
Text10 = ""
Label28 = ""
End Sub

Private Sub FrameGLAktif()
Frame2.Enabled = True
Label14.Enabled = True
Text10.Enabled = True
Text10.BackColor = &HFFFFC0
InfoGL.Enabled = True
Label28.Enabled = True
End Sub

Private Sub Cari_SubGL()
Dim Baris As Integer
With GridGL
    .Row = 0
    .Cols = 2
    .Col = 0: .ColWidth(0) = 1000: .Text = "KODE": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 9
    .Col = 1: .ColWidth(1) = 2500: .Text = "NAMA SUPPLIER": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 9
End With
Baris = 1
SSGL = "Select CodeSL, NamaSL From G003 where CodeSGL = '" + Trim("1001120") + "' order by codeSL"
Set RSGL = RDCO.OpenResultset(SSGL, rdOpenDynamic, rdConcurRowVer)
If RSGL.RowCount <> 0 Then
RSGL.MoveFirst
Do Until RSGL.EOF
    With GridGL
        .Rows = Baris + 1
        .Row = Baris
        .Col = 0: .Text = RSGL("CodeSL"): .CellAlignment = 4
        .Col = 1: .Text = RSGL("NamaSL")
        Baris = Baris + 1
    End With
RSGL.MoveNext
Loop
End If
RSGL.Close
Set RSGL = Nothing
End Sub

Private Sub GridGL_DblClick()
Call InfoGL2_Click
Text10 = GridGL.TextMatrix(GridGL.Row, 0)
Label28 = GridGL.TextMatrix(GridGL.Row, 1)
Text10.SetFocus
End Sub

Private Sub GridGL_LostFocus()
GridGL.Visible = False
Call InfoGL2_Click
End Sub

Private Sub InfoGL_Click()
InfoGL.Visible = False
InfoGL2.Visible = True
GridGL.Visible = True
GridGL.SetFocus
End Sub

Private Sub InfoGL2_Click()
InfoGL.Visible = True
InfoGL2.Visible = False
GridGL.Visible = False
End Sub

Private Sub TmbSave_Click()
If Text6 = "" Or Text4 = "" Then
    MsgBox "FAKTUR PEMBELIAN / TANGGAL PEMBELIAN MASIH KOSONG", vbCritical, "DATA TIDAK BOLEH KOSONG"
    Text1.SetFocus
    Exit Sub
End If

If CCur(Label7) <= 0 Then
    Text5.SetFocus
    MsgBox "NOMINAL PEMBAYARAN MASIH KOSONG", vbCritical, "DATA TIDAK BOLEH KOSONG"
    Exit Sub
End If

If CCur(Text14) > 0 And Text10 = "" Then
    MsgBox "DATA GENERAL LEDGER PEMBAYARAN NON TUNAI MASIH KOSONG", vbCritical, "DATA TIDAK BOLEH KOSONG"
    Text10.SetFocus
    Exit Sub
End If

If CCur(Label7) <> CCur(Label17) Then
    MsgBox "TOTAL DIBAYAR PEMBAYARAN HARUS SAMA DENGAN TOTAL PEMBELIAN", vbCritical, "NOMINAL PEMBAYARAN SALAH"
    Text1.SetFocus
    Exit Sub
End If

'Jika Pembayaran Non Tunai
If CCur(Text14) > 0 Then
    SGLNonKas = Text10
Else
    SGLNonKas = "-"
End If

Tanya = MsgBox("YAKIN PROSES TRANSAKSI PEMBELIAN ?", vbOKCancel, "PROSES TRANSAKSI PEMBELIAN ?")
If Tanya = vbCancel Then Exit Sub

Call BELI_BAHAN3


Call JurnalDebet

If Text10 = "" Then
    Call JurnalCredit
Else
    Call JurnalCredit
    Call JurnalCREDIT2
End If

Call Kosong
Text1.SetFocus

Unload Me
F001.Show
End Sub

Private Sub BELI_BAHAN3()
SFuck = "Select * From F001"
Set RFuck = RDCO.OpenResultset(SFuck, rdOpenKeyset, rdConcurRowVer)
RFuck.AddNew
    RFuck("NO_SYSTEM") = Label3
    RFuck("TGL_SYSTEM") = Tanggal
    RFuck("NO_FAK") = Trim(Text6)
    RFuck("JUMLAH") = CCur(Text2)
    RFuck("KAS") = CCur(Text1)
    RFuck("NON_KAS") = CCur(Text14)
    RFuck("GL_NONKAS") = SGLNonKas
    RFuck("KETERANGAN") = Trim(Text5)
    RFuck("STATUS") = 0
    RFuck("H_JUMLAH") = CCur(Text3)
    RFuck("TGL_BELI") = Trim(Text4)
    RFuck("STS_FAK") = 0
    RFuck("STS_M001") = 0
RFuck.Update
RFuck.Close
Set RFuck = Nothing
End Sub

Private Sub JurnalDebet()
SDebet2 = "Select * From G003 where codesl = '" + Trim(F_DEBET) + "'"
Set RDebet2 = RDCO.OpenResultset(SDebet2, rdOpenKeyset, rdConcurRowVer)
    MMUTASID = RDebet2("mutasid") + CCur(Text3)
    SSALDO = RDebet2("saldo") + CCur(Text3)
    RDebet2.EDIT
        RDebet2("mutasid") = CCur(MMUTASID)
        RDebet2("saldo") = CCur(SSALDO)

    SDebet3 = "Select * From G005"
    Set RDebet3 = RDCO.OpenResultset(SDebet3, rdOpenKeyset, rdConcurRowVer)
    RDebet3.AddNew
        RDebet3("codecab") = CodeCab
        RDebet3("codesl") = F_DEBET
        RDebet3("namasl") = RDebet2("NamaSL")
        RDebet3("nobukti") = Label3
        RDebet3("keterangan") = "BL. No Fak." + Text6
        RDebet3("nominald") = CCur(Text1)
        RDebet3("nominalc") = 0
        RDebet3("saldo") = SSALDO
        RDebet3("tanggal") = Tanggal
        RDebet3("jam") = Time
        RDebet3("usercode") = Operator
    RDebet3.Update
    RDebet3.Close
    Set RDebet3 = Nothing

RDebet2.Update
RDebet2.Close
Set RDebet2 = Nothing
End Sub

Private Sub JurnalCredit()
SCredit2 = "Select * From G003 where codesl='" + Trim(G_CREDIT) + "'"
Set RCredit2 = RDCO.OpenResultset(SCredit2, rdOpenKeyset, rdConcurRowVer)
    MMUTASIC = CCur(RCredit2("mutasiC")) + CCur(Text1)
    SSSALDO = RCredit2("saldo") - CCur(Text1)
    
    RCredit2.EDIT
        RCredit2("mutasic") = CCur(MMUTASIC)
        RCredit2("saldo") = CCur(SSSALDO)

    SCREDIT3 = "Select * From G005"
    Set RCREDIT3 = RDCO.OpenResultset(SCREDIT3, rdOpenKeyset, rdConcurRowVer)
    RCREDIT3.AddNew
        RCREDIT3("codecab") = CodeCab
        RCREDIT3("codesl") = G_CREDIT
        RCREDIT3("namasl") = RCredit2("NamaSL")
        RCREDIT3("nobukti") = Label3
        RCREDIT3("keterangan") = "BL. No Fak." + Text6
        RCREDIT3("nominald") = 0
        RCREDIT3("nominalc") = CCur(Text1)
        RCREDIT3("saldo") = SSSALDO
        RCREDIT3("tanggal") = Tanggal
        RCREDIT3("jam") = Time
        RCREDIT3("usercode") = Operator
    RCREDIT3.Update
    RCREDIT3.Close
    Set RCREDIT3 = Nothing

RCredit2.Update
RCredit2.Close
Set RCredit2 = Nothing
End Sub

Private Sub JurnalCREDIT2()
SCredit2 = "Select * From G003 where codesl='" + Trim(SGLNonKas) + "'"
Set RCredit2 = RDCO.OpenResultset(SCredit2, rdOpenKeyset, rdConcurRowVer)
    MMUTASIC = CCur(RCredit2("mutasiC")) + CCur(Text14)
    SSSALDO = RCredit2("saldo") - CCur(Text14)
    
    RCredit2.EDIT
        RCredit2("mutasic") = CCur(MMUTASIC)
        RCredit2("saldo") = CCur(SSSALDO)

    SCREDIT3 = "Select * From G005"
    Set RCREDIT3 = RDCO.OpenResultset(SCREDIT3, rdOpenKeyset, rdConcurRowVer)
    RCREDIT3.AddNew
        RCREDIT3("codecab") = CodeCab
        RCREDIT3("codesl") = SGLNonKas
        RCREDIT3("namasl") = RCredit2("NamaSL")
        RCREDIT3("nobukti") = Label3
        RCREDIT3("keterangan") = "BL. No Fak." + Label3
        RCREDIT3("nominald") = 0
        RCREDIT3("nominalc") = CCur(Text14)
        RCREDIT3("saldo") = SSSALDO
        RCREDIT3("tanggal") = Tanggal
        RCREDIT3("jam") = Time
        RCREDIT3("usercode") = CodeBag
    RCREDIT3.Update
    RCREDIT3.Close
    Set RCREDIT3 = Nothing

RCredit2.Update
RCredit2.Close
Set RCredit2 = Nothing
End Sub

Private Sub SiapkanGrid()
With grid
    .Row = 0
    .Cols = 6
    .Col = 0: .ColWidth(0) = 1200: .Text = "TGL": .CellAlignment = 4: .CellFontBold = True
    .Col = 1: .ColWidth(1) = 1500: .Text = "NO. FAK": .CellAlignment = 4: .CellFontBold = True
    .Col = 2: .ColWidth(2) = 1000: .Text = "UNIT": .CellAlignment = 4: .CellFontBold = True
    .Col = 3: .ColWidth(3) = 1500: .Text = "TUNAI": .CellAlignment = 4: .CellFontBold = True
    .Col = 4: .ColWidth(4) = 1500: .Text = "NON TUNAI": .CellAlignment = 4: .CellFontBold = True
    .Col = 5: .ColWidth(5) = 1500: .Text = "JUMLAH": .CellAlignment = 4: .CellFontBold = True
    
End With
End Sub

Private Sub IsiGrid()
Dim Brs
Brs = 1
SKode = "Select * From F001 order by No_System Desc"
Set RKode = RDCO.OpenResultset(SKode, rdOpenDynamic, rdOpenKeyset)
If RKode.RowCount <> 0 Then
RKode.MoveFirst
Do Until RKode.EOF
    With grid
        .Rows = Brs + 1
        .Row = Brs
        .Col = 0: .Text = RKode("TGL_SYSTEM"): .CellAlignment = 4
        .Col = 1: .Text = RKode("NO_FAK"): .CellAlignment = 4
        .Col = 2: .Text = RKode("JUMLAH"): .CellAlignment = 4
        .Col = 3: .Text = Format(RKode("KAS"), "##,###.00"): .CellAlignment = 4
        .Col = 4: .Text = Format(RKode("NON_KAS"), "##,###.00"): .CellAlignment = 4
        .Col = 5: .Text = Format(RKode("H_JUMLAH"), "##,###.00"): .CellAlignment = 4
        Brs = Brs + 1
    End With
RKode.MoveNext
Loop
End If
RKode.Close
Set RKode = Nothing
End Sub
