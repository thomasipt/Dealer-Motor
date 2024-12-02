VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form JL004KWT 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "KWITANSI"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1103
      TabIndex        =   12
      Text            =   "1"
      Top             =   1575
      Width           =   1410
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1103
      TabIndex        =   11
      Text            =   "2"
      Top             =   1830
      Width           =   1410
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3623
      TabIndex        =   10
      Text            =   "3"
      Top             =   1575
      Width           =   1410
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3623
      TabIndex        =   9
      Text            =   "4"
      Top             =   1830
      Width           =   1410
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1733
      TabIndex        =   8
      Text            =   "5"
      Top             =   60
      Width           =   4110
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   1733
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "JL004KWT.frx":0000
      Top             =   330
      Width           =   4110
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1733
      TabIndex        =   6
      Text            =   "7"
      Top             =   1170
      Width           =   4110
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3908
      TabIndex        =   5
      Text            =   "8"
      Top             =   2160
      Width           =   1905
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1058
      TabIndex        =   4
      Text            =   "9"
      Top             =   2610
      Width           =   1560
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3128
      TabIndex        =   3
      Text            =   "TTD"
      Top             =   3015
      Width           =   2715
   End
   Begin VB.TextBox Text11 
      Height          =   375
      Left            =   6165
      TabIndex        =   2
      Text            =   "Text11"
      Top             =   -15
      Width           =   1635
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
      Left            =   4335
      TabIndex        =   1
      Top             =   3540
      Width           =   1410
   End
   Begin VB.CommandButton cmdKWT 
      Caption         =   "CETAK"
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
      Left            =   165
      TabIndex        =   0
      Top             =   3540
      Width           =   1410
   End
   Begin Crystal.CrystalReport crpt 
      Left            =   8145
      Top             =   585
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.PictureBox Picture1 
      Height          =   765
      Left            =   -225
      ScaleHeight     =   705
      ScaleWidth      =   6225
      TabIndex        =   23
      Top             =   3450
      Width           =   6285
   End
   Begin VB.Label Label3 
      Caption         =   "Telah terima dari     ..........................................................................................."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   68
      TabIndex        =   22
      Top             =   45
      Width           =   5775
   End
   Begin VB.Label Label1 
      Caption         =   "Uang sebanyak"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   68
      TabIndex        =   21
      Top             =   315
      Width           =   1770
   End
   Begin VB.Label Label2 
      Caption         =   "Untuk pembayaran"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   68
      TabIndex        =   20
      Top             =   1155
      Width           =   1770
   End
   Begin VB.Label Label4 
      Caption         =   "No. Rangka"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   68
      TabIndex        =   19
      Top             =   1560
      Width           =   1005
   End
   Begin VB.Label Label5 
      Caption         =   "No. Mesin"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   68
      TabIndex        =   18
      Top             =   1815
      Width           =   1005
   End
   Begin VB.Label Label8 
      Caption         =   "Semarang,......................................................................"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3128
      TabIndex        =   17
      Top             =   2235
      Width           =   2715
   End
   Begin VB.Label Label9 
      Caption         =   "Terbilang Rp."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   68
      TabIndex        =   16
      Top             =   2595
      Width           =   960
   End
   Begin VB.Label Label10 
      Caption         =   ".............................................................................................."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3128
      TabIndex        =   15
      Top             =   3105
      Width           =   2715
   End
   Begin VB.Label Label6 
      Caption         =   "Warna"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2618
      TabIndex        =   14
      Top             =   1560
      Width           =   1005
   End
   Begin VB.Label Label7 
      Caption         =   "Tahun"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2618
      TabIndex        =   13
      Top             =   1815
      Width           =   1005
   End
End
Attribute VB_Name = "JL004KWT"
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
Private Hari, Bulan, Tahun, TglOK

Private Sub cmdKWT_Click()
Dim Tanya

If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Or Text7 = "" Or Text8 = "" Or Text9 = "" Or Text10 = "" Then
    MsgBox "MASIH ADA YANG KOSONG", vbCritical, "KONFIRMASI"
    Exit Sub
End If

If OYEN = 1 Then
    SDel = "Delete From M001_KWITANSI where NO_FAK = '" + Trim(Text11) + "'"
    Set RDel = RDCO.OpenResultset(SDel, rdOpenDynamic, rdConcurRowVer)
End If

SSave = "Select * From M001_KWITANSI"
Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
RSave.AddNew
    RSave("NO_FAK") = NO_FAKTUR
    RSave("TANGGAL") = Format(Trim(Text8), ">")
    RSave("NAMA") = Format(Trim(Text5), ">")
    RSave("TERBILANG") = Format(Trim(Text6), ">")
    RSave("RANGKA") = Trim(Text1)
    RSave("MESIN") = Trim(Text2)
    RSave("WARNA") = Trim(Text3)
    RSave("TAHUN") = Trim(Text4)
    RSave("NOMINAL") = CCur(Text9)
    RSave("TTD") = Format(Trim(Text10), ">")
    RSave("KETERANGAN") = Format(Trim(Text7), ">")
RSave.Update
RSave.Close
Set RSave = Nothing

Tanya = MsgBox("CETAK KWITANSI", vbOKCancel, "KONFIRMASI")
    If Tanya = vbOK Then
        crpt.ReportFileName = App.Path + "\ReportD\KWT.rpt"
        crpt.SelectionFormula = "{M001_KWITANSI.NO_FAK} = '" + Trim(Text11) + "'"
        crpt.WindowState = crptMaximized
        crpt.WindowMaxButton = False
        crpt.WindowMinButton = False
        crpt.Action = 1
    Else
        Unload Me
        JL003A.Show
    End If
End Sub

Private Sub Command2_Click()
Unload Me
JL003A.Show
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=DEALER", rdDriverNoPrompt, False, CN)

Hari = Format(Day(Tanggal), "00")
Bulan = BulanStr(Month(Tanggal))
Tahun = Year(Tanggal)
TglOK = Trim(Hari) + " " + Trim(Bulan) + " " + Trim(Tahun)

Text11 = NO_FAKTUR
OYEN = 0

Call CekData

End Sub

Private Sub CekData()
SCari = "Select * From M001_KWITANSI where NO_FAK = '" + Trim(Text11) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdOpenKeyset)
If RCari.RowCount <> 0 Then
    Call CariData
Else
    Call CariData2
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub CariData()
SToket = "Select * From M001_KWITANSI where NO_FAK = '" + Trim(Text11) + "'"
Set RToket = RDCO.OpenResultset(SToket, rdOpenDynamic, rdOpenKeyset)
If RToket.RowCount <> 0 Then
    Text5 = Format(RToket("NAMA"), ">")
    Text1 = Format(RToket("RANGKA"), ">")
    Text2 = Format(RToket("MESIN"), ">")
    Text3 = Format(RToket("WARNA"), ">")
    Text4 = Format(RToket("TAHUN"), ">")
    Text10 = Format(RToket("TTD"), ">")
    Text8 = Format(RToket("TANGGAL"), ">")
    Text9 = Format(RToket("NOMINAL"), "##,###.00")
    Text6 = Format(RToket("TERBILANG"), ">")
    Text7 = Format(RToket("KETERANGAN"), ">")
End If
RToket.Close
Set RToket = Nothing
OYEN = 1
End Sub

Private Sub CariData2()
SToket = "Select * From M001 where NO_FAK = '" + Trim(Text11) + "'"
Set RToket = RDCO.OpenResultset(SToket, rdOpenDynamic, rdOpenKeyset)
If RToket.RowCount <> 0 Then
    Text5 = Format(RToket("NAMA_PEMBELI"), ">")
    Text25 = Format(RToket("ALAMAT_1"), ">") + " , " + Format(RToket("ALAMAT_2"), ">")
    Text1 = Format(RToket("RANGKA"), ">")
    Text2 = Format(RToket("MESIN"), ">")
    Text3 = Format(RToket("WARNA"), ">")
    Text19 = Format(RToket("TYPE"), ">")
    Text4 = Format(RToket("TAHUN"), ">")
    Text7 = "FAKTUR KENDARAAN " + Format(RToket("TYPE"), ">")
    Text8 = TglOK
    Text9 = Format(RToket("H_OTR"), "##,###.00")
End If
RToket.Close
Set RToket = Nothing
    Text6 = Terbilang(Text9)
    OYEN = 0
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text7 = Format(Text7, ">")
    cmdKWT.SetFocus
End If
End Sub

Private Sub Text9_GotFocus()
If CCur(Text9) = 0 Then Text9 = ""
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdKWT.SetFocus
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text10 = Format(Text10, ">")
    cmdKWT.SetFocus
End If
End Sub

Private Sub Text9_LostFocus()
If Text9 = "" Then Text9 = 0
If Not IsNumeric(Text9) Then
    Text9.SetFocus
    MsgBox "NOMINAL MENGGUNAKAN ANGKA", vbCritical, "TYPE DATA SALAH"
    Exit Sub
End If
Text9 = Format(Text9, "##,###.00")
Text6 = Terbilang(Text9)
End Sub


