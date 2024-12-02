VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form JL004SPF 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SURAT PENERIMAAN FAKTUR"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8340
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   8340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text18 
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
      Left            =   1133
      TabIndex        =   44
      Text            =   "18"
      Top             =   1245
      Width           =   3945
   End
   Begin VB.TextBox Text100 
      Height          =   435
      Left            =   8565
      TabIndex        =   43
      Text            =   "100"
      Top             =   390
      Width           =   1320
   End
   Begin VB.CommandButton cmdSPF 
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
      Left            =   540
      TabIndex        =   41
      Top             =   4065
      Width           =   1410
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
      Left            =   6390
      TabIndex        =   40
      Top             =   4065
      Width           =   1410
   End
   Begin VB.TextBox Text17 
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
      Left            =   5543
      TabIndex        =   36
      Text            =   "TTD 2"
      Top             =   3420
      Width           =   2715
   End
   Begin VB.TextBox Text16 
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
      Left            =   2348
      TabIndex        =   34
      Text            =   "TTD 1"
      Top             =   3435
      Width           =   2715
   End
   Begin VB.TextBox Text15 
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
      Height          =   450
      Left            =   1163
      MultiLine       =   -1  'True
      TabIndex        =   31
      Text            =   "JL004SPF.frx":0000
      Top             =   2415
      Width           =   7095
   End
   Begin VB.TextBox Text14 
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
      Left            =   1163
      TabIndex        =   30
      Text            =   "14"
      Top             =   2160
      Width           =   7095
   End
   Begin VB.TextBox Text13 
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
      Left            =   3653
      TabIndex        =   25
      Text            =   "13"
      Top             =   1770
      Width           =   1410
   End
   Begin VB.TextBox Text12 
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
      Left            =   3653
      TabIndex        =   24
      Text            =   "12"
      Top             =   1515
      Width           =   1410
   End
   Begin VB.TextBox Text11 
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
      Left            =   1133
      TabIndex        =   23
      Text            =   "11"
      Top             =   1770
      Width           =   1410
   End
   Begin VB.TextBox Text10 
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
      Left            =   1133
      TabIndex        =   22
      Text            =   "10"
      Top             =   1515
      Width           =   1410
   End
   Begin VB.TextBox Text9 
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
      Left            =   5618
      TabIndex        =   21
      Text            =   "9"
      Top             =   870
      Width           =   2640
   End
   Begin VB.TextBox Text8 
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
      Left            =   5618
      TabIndex        =   20
      Text            =   "8"
      Top             =   615
      Width           =   2640
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
      Left            =   5618
      TabIndex        =   19
      Text            =   "7"
      Top             =   330
      Width           =   2640
   End
   Begin VB.TextBox Text6 
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
      Left            =   3743
      TabIndex        =   18
      Text            =   "6"
      Top             =   870
      Width           =   1410
   End
   Begin VB.TextBox Text5 
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
      Left            =   3743
      TabIndex        =   17
      Text            =   "5"
      Top             =   615
      Width           =   1410
   End
   Begin VB.TextBox Text4 
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
      Left            =   3743
      TabIndex        =   16
      Text            =   "4"
      Top             =   330
      Width           =   1410
   End
   Begin VB.TextBox Text3 
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
      Left            =   1163
      TabIndex        =   15
      Text            =   "3"
      Top             =   870
      Width           =   2070
   End
   Begin VB.TextBox Text2 
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
      Left            =   1163
      TabIndex        =   14
      Text            =   "2"
      Top             =   615
      Width           =   2070
   End
   Begin VB.TextBox Text1 
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
      Left            =   1163
      TabIndex        =   13
      Text            =   "1"
      Top             =   330
      Width           =   2070
   End
   Begin VB.PictureBox Picture1 
      Height          =   765
      Left            =   -150
      ScaleHeight     =   705
      ScaleWidth      =   8640
      TabIndex        =   42
      Top             =   3975
      Width           =   8700
   End
   Begin Crystal.CrystalReport crpt 
      Left            =   90
      Top             =   3195
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label24 
      Caption         =   "Type"
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
      Left            =   105
      TabIndex        =   45
      Top             =   1230
      Width           =   1005
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      Caption         =   "Pemohon,"
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
      Left            =   5550
      TabIndex        =   39
      Top             =   2985
      Width           =   2715
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      Caption         =   "Mengetahui,"
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
      Left            =   2355
      TabIndex        =   38
      Top             =   2955
      Width           =   2715
   End
   Begin VB.Label Label21 
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
      Left            =   5550
      TabIndex        =   37
      Top             =   3510
      Width           =   2715
   End
   Begin VB.Label Label20 
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
      Left            =   2355
      TabIndex        =   35
      Top             =   3525
      Width           =   2715
   End
   Begin VB.Label Label19 
      Caption         =   "Alamat"
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
      Left            =   105
      TabIndex        =   33
      Top             =   2400
      Width           =   1005
   End
   Begin VB.Label Label18 
      Caption         =   "Nama"
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
      Left            =   105
      TabIndex        =   32
      Top             =   2145
      Width           =   1005
   End
   Begin VB.Label Label17 
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
      Left            =   2655
      TabIndex        =   29
      Top             =   1755
      Width           =   1005
   End
   Begin VB.Label Label16 
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
      Left            =   2655
      TabIndex        =   28
      Top             =   1500
      Width           =   1005
   End
   Begin VB.Label Label15 
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
      Left            =   105
      TabIndex        =   27
      Top             =   1755
      Width           =   1005
   End
   Begin VB.Label Label14 
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
      Left            =   105
      TabIndex        =   26
      Top             =   1500
      Width           =   1005
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "a.n."
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
      Left            =   5183
      TabIndex        =   12
      Top             =   855
      Width           =   420
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "a.n."
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
      Left            =   5183
      TabIndex        =   11
      Top             =   600
      Width           =   420
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "a.n."
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
      Left            =   5183
      TabIndex        =   10
      Top             =   315
      Width           =   420
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "Tgl."
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
      Left            =   3278
      TabIndex        =   9
      Top             =   855
      Width           =   420
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Tgl."
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
      Left            =   3278
      TabIndex        =   8
      Top             =   600
      Width           =   420
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Tgl."
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
      Left            =   3278
      TabIndex        =   7
      Top             =   315
      Width           =   420
   End
   Begin VB.Label Label7 
      Caption         =   "No."
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
      Left            =   728
      TabIndex        =   6
      Top             =   855
      Width           =   420
   End
   Begin VB.Label Label6 
      Caption         =   "No."
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
      Left            =   728
      TabIndex        =   5
      Top             =   600
      Width           =   420
   End
   Begin VB.Label Label5 
      Caption         =   "No."
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
      Left            =   728
      TabIndex        =   4
      Top             =   315
      Width           =   420
   End
   Begin VB.Label Label4 
      Caption         =   "S.J."
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
      Left            =   83
      TabIndex        =   3
      Top             =   870
      Width           =   645
   End
   Begin VB.Label Label2 
      Caption         =   "S.P."
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
      Left            =   83
      TabIndex        =   2
      Top             =   600
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "D.O."
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
      Left            =   83
      TabIndex        =   1
      Top             =   315
      Width           =   645
   End
   Begin VB.Label Label3 
      Caption         =   "Berdasarkan"
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
      Left            =   83
      TabIndex        =   0
      Top             =   45
      Width           =   1005
   End
End
Attribute VB_Name = "JL004SPF"
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

Private Sub cmdSPF_Click()
Dim Tanya

If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Or Text7 = "" Or Text8 = "" Or Text9 = "" Or Text16 = "" Or Text17 = "" Then
    MsgBox "MASIH ADA YANG KOSONG", vbCritical, "KONFIRMASI"
    Exit Sub
End If

If OYEN = 1 Then
    SDel = "Delete From M001_SPF where NO_FAK = '" + Trim(Text100) + "'"
    Set RDel = RDCO.OpenResultset(SDel, rdOpenDynamic, rdConcurRowVer)
End If

SSave = "Select * From M001_SPF"
Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
RSave.AddNew
    RSave("NO_FAK") = NO_FAKTUR
    RSave("TANGGAL") = Tanggal
    
    RSave("DO_NO") = Format(Trim(Text1), ">")
    RSave("SP_NO") = Format(Trim(Text2), ">")
    RSave("SJ_NO") = Format(Trim(Text3), ">")
    RSave("DO_TGL") = Trim(Text4)
    RSave("SP_TGL") = Trim(Text5)
    RSave("SJ_TGL") = Trim(Text6)
    
    RSave("DO_NAMA") = Format(Trim(Text7), ">")
    RSave("SP_NAMA") = Format(Trim(Text8), ">")
    RSave("SJ_NAMA") = Format(Trim(Text9), ">")
    
    
    RSave("NAMA") = Format(Trim(Text14), ">")
    RSave("ALAMAT") = Format(Trim(Text15), ">")
    
    RSave("TYPE") = Trim(Text18)
    RSave("RANGKA") = Trim(Text10)
    RSave("MESIN") = Trim(Text11)
    RSave("WARNA") = Trim(Text12)
    RSave("TAHUN") = Trim(Text13)
    
    RSave("TTD_1") = Format(Trim(Text16), ">")
    RSave("TTD_2") = Format(Trim(Text17), ">")
    
RSave.Update
RSave.Close
Set RSave = Nothing

Tanya = MsgBox("CETAK PENERIMAAN FAKTUR", vbOKCancel, "KONFIRMASI")
    If Tanya = vbOK Then
        crpt.ReportFileName = App.Path + "\ReportD\SPF.rpt"
        crpt.SelectionFormula = "{M001_SPF.NO_FAK} = '" + Trim(Text100) + "'"
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

Text100 = NO_FAKTUR
OYEN = 0

Call CekData

End Sub

Private Sub CekData()
SCari = "Select * From M001_SPF where NO_FAK = '" + Trim(Text100) + "'"
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
SToket = "Select * From M001_SPF where NO_FAK = '" + Trim(Text100) + "'"
Set RToket = RDCO.OpenResultset(SToket, rdOpenDynamic, rdOpenKeyset)
If RToket.RowCount <> 0 Then
    Text1 = Format(RToket("DO_NO"), ">")
    Text4 = Format(RToket("DO_TGL"), ">")
    Text7 = Format(RToket("DO_NAMA"), ">")
    
    Text2 = Format(RToket("SP_NO"), ">")
    Text5 = Format(RToket("SP_TGL"), ">")
    Text8 = Format(RToket("SP_NAMA"), ">")
    
    Text3 = Format(RToket("SJ_NO"), ">")
    Text6 = Format(RToket("SJ_TGL"), ">")
    Text9 = Format(RToket("SJ_NAMA"), ">")
    
    Text10 = Format(RToket("RANGKA"), ">")
    Text11 = Format(RToket("MESIN"), ">")
    Text12 = Format(RToket("WARNA"), ">")
    Text13 = Format(RToket("TAHUN"), ">")
    Text18 = Format(RToket("TYPE"), ">")
    
    Text14 = Format(RToket("NAMA"), ">")
    Text15 = Format(RToket("ALAMAT"), ">")
    
    Text16 = Format(RToket("TTD_1"), ">")
    Text17 = Format(RToket("TTD_2"), ">")
End If
RToket.Close
Set RToket = Nothing
    OYEN = 1
End Sub

Private Sub CariData2()
SToket = "Select * From M001 where NO_FAK = '" + Trim(Text100) + "'"
Set RToket = RDCO.OpenResultset(SToket, rdOpenDynamic, rdOpenKeyset)
If RToket.RowCount <> 0 Then
    Text1 = "-"
    Text4 = "-"
    Text7 = "-"
    
    Text2 = "-"
    Text5 = "-"
    Text8 = "-"
    
    Text3 = "-"
    Text6 = "-"
    Text9 = "-"
    
    Text10 = Format(RToket("RANGKA"), ">")
    Text11 = Format(RToket("MESIN"), ">")
    Text12 = Format(RToket("WARNA"), ">")
    Text13 = Format(RToket("TAHUN"), ">")
    Text18 = Format(RToket("TYPE"), ">")
    
    Text14 = Format(RToket("NAMA_PEMBELI"), ">")
    Text15 = Format(RToket("ALAMAT_1"), ">") + " , " + Format(RToket("ALAMAT_2"), ">")
    
    Text16 = "TTD 1"
    Text17 = "TTD 2"
End If
RToket.Close
Set RToket = Nothing
    OYEN = 0
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text1 = Format(Text1, ">")
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text16 = Format(Text16, ">")
    cmdSPF.SetFocus
End If
End Sub

Private Sub Text17_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text17 = Format(Text17, ">")
    cmdSPF.SetFocus
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text2 = Format(Text2, ">")
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text3 = Format(Text3, ">")
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text4_LostFocus()
If Text4 = "" Or Text4 = "-" Then
    Text4 = "-"
    Exit Sub
End If
If Not IsDate(Text4) Then
    Text4 = ""
    Text4.SetFocus
    MsgBox "DATA MENGGUNAKAN FORMAT TANGGAL DD/MM/YYYY", vbCritical, "TYPE DATA SALAH"
    Exit Sub
End If
Text4 = Format(Text4, "DD/MM/YYYY")
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text4 = Format(Text4, ">")
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text5_LostFocus()
If Text5 = "" Or Text5 = "-" Then
    Text5 = "-"
    Exit Sub
End If
If Not IsDate(Text5) Then
    Text5 = ""
    Text5.SetFocus
    MsgBox "DATA MENGGUNAKAN FORMAT TANGGAL DD/MM/YYYY", vbCritical, "TYPE DATA SALAH"
    Exit Sub
End If
Text5 = Format(Text5, "DD/MM/YYYY")
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text5 = Format(Text5, ">")
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text6_LostFocus()
If Text6 = "" Or Text6 = "-" Then
    Text6 = "-"
    Exit Sub
End If
If Not IsDate(Text6) Then
    Text6 = ""
    Text6.SetFocus
    MsgBox "DATA MENGGUNAKAN FORMAT TANGGAL DD/MM/YYYY", vbCritical, "TYPE DATA SALAH"
    Exit Sub
End If
Text6 = Format(Text6, "DD/MM/YYYY")
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text6 = Format(Text6, ">")
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text7 = Format(Text7, ">")
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text8 = Format(Text8, ">")
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text9 = Format(Text9s, ">")
    SendKeys "{TAB}"
End If
End Sub
