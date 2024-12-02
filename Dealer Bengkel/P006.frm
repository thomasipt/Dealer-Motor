VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form P006 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INFORMASI DETAIL PIUTANG"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6570
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   6570
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
      Height          =   450
      Left            =   5115
      TabIndex        =   22
      Top             =   4470
      Width           =   1275
   End
   Begin Crystal.CrystalReport crpt 
      Left            =   1665
      Top             =   4635
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PELUNASAN"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   120
      TabIndex        =   21
      Top             =   4500
      Width           =   1275
   End
   Begin VB.CommandButton TmbMutasi 
      Caption         =   "MUTASI"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2648
      TabIndex        =   0
      Top             =   4500
      Width           =   1275
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label22"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1650
      TabIndex        =   24
      Top             =   3285
      Width           =   1545
   End
   Begin VB.Label Label15 
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
      Height          =   315
      Left            =   120
      TabIndex        =   23
      Top             =   3285
      Width           =   1455
   End
   Begin VB.Line Line1 
      X1              =   105
      X2              =   6440
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Label Label21 
      Caption         =   "JTH TEMPO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3480
      TabIndex        =   20
      Top             =   2835
      Width           =   960
   End
   Begin VB.Label Label20 
      Caption         =   "TGL MULAI"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3480
      TabIndex        =   19
      Top             =   2385
      Width           =   960
   End
   Begin VB.Label Label19 
      Caption         =   "SYARAT PEMBAYARAN"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   18
      Top             =   3645
      Width           =   1455
   End
   Begin VB.Label Label18 
      Caption         =   "BAKI DEBET"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   17
      Top             =   2835
      Width           =   1455
   End
   Begin VB.Label Label17 
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
      Height          =   315
      Left            =   120
      TabIndex        =   16
      Top             =   2385
      Width           =   1455
   End
   Begin VB.Label Label16 
      Caption         =   "NO. TELEPON"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   15
      Top             =   1935
      Width           =   1455
   End
   Begin VB.Label Label14 
      Caption         =   "ALAMAT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   14
      Top             =   1035
      Width           =   1455
   End
   Begin VB.Label Label13 
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
      Height          =   315
      Left            =   120
      TabIndex        =   13
      Top             =   585
      Width           =   1455
   End
   Begin VB.Label Label12 
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
      Height          =   315
      Left            =   120
      TabIndex        =   12
      Top             =   135
      Width           =   1455
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label11"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1650
      TabIndex        =   11
      Top             =   3735
      Width           =   2085
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label10"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4560
      TabIndex        =   10
      Top             =   2835
      Width           =   915
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4560
      TabIndex        =   9
      Top             =   2385
      Width           =   915
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label8"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1650
      TabIndex        =   8
      Top             =   2835
      Width           =   1545
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1650
      TabIndex        =   7
      Top             =   2385
      Width           =   1545
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1650
      TabIndex        =   6
      Top             =   1935
      Width           =   2130
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1650
      TabIndex        =   5
      Top             =   1485
      Width           =   3615
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1650
      TabIndex        =   4
      Top             =   1035
      Width           =   4785
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2805
      TabIndex        =   3
      Top             =   585
      Width           =   3615
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1650
      TabIndex        =   2
      Top             =   585
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1650
      TabIndex        =   1
      Top             =   135
      Width           =   2895
   End
End
Attribute VB_Name = "P006"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private RMumet, RSPL As rdoResultset

Private SMumet, SSPL As String

Private Sub Command1_Click()
If Label8 > 0 Then
    NoPinjaman = ""
    NoPinjaman = Label1
    P003.Show
    Unload Me
    Exit Sub
Else
    MsgBox "PIUTANG SUDAH SELESAI", vbCritical, "OUTSTANDING"
End If
End Sub


Private Sub Command2_Click()
Unload Me
P002.Show
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=DEALER", rdDriverNoPrompt, False, CN)

Call Kosong
Label1 = NoPinjaman
Call NoPin
End Sub

Private Sub Kosong()
Label1 = ""
Label2 = ""
Label3 = ""
Label4 = ""
Label5 = ""
Label6 = ""
Label7 = ""
Label8 = ""
Label9 = ""
Label10 = ""
Label11 = ""
Label22 = ""
End Sub

Private Sub NoPin()

SMumet = "Select * from P002 where Nomor_Pin = '" + Trim(Label1) + "'"
Set RMumet = RDCO.OpenResultset(SMumet, rdOpenKeyset, rdConcurReadOnly)
Do While Not RMumet.EOF
    Label2 = RMumet("Nomor_Nas")
    Label3 = RMumet("Nama_Nas")
    Label7 = Format(RMumet("PLafon"), "##,###.00")
    Label8 = Format(RMumet("Baki_Debet"), "##,###.00")
    Label9 = RMumet("Tgl_Mulai")
    Label10 = RMumet("Tgl_Jatuh")
    Label11 = RMumet("Syarat_Byr")
    Label22 = Format(RMumet("Intensif"), "##,###.00")

    SSPL = "Select * From C012 where Nama = '" + Trim(Label3) + "'"
    Set RSPL = RDCO.OpenResultset(SSPL, rdOpenDynamic, rdConcurRowVer)
    If RSPL.RowCount <> 0 Then
        Label4 = RSPL("Alamat1")
        Label5 = RSPL("Kota")
        Label6 = RSPL("Telpon")
    End If
    RSPL.Close
    Set RSPL = Nothing
RMumet.MoveNext
Loop
RMumet.Close
Set RMumet = Nothing
End Sub

Private Sub TmbMutasi_Click()
crpt.ReportFileName = App.Path + "\ReportD\HisPiutang.rpt"
crpt.SelectionFormula = "{P002.Nomor_Pin} = '" + Trim(Label1) + "'"
crpt.WindowState = crptMaximized
crpt.WindowMaxButton = False
crpt.WindowMinButton = False
crpt.Action = 1
End Sub

Private Sub TmbUdah_Click()
Unload Me
End Sub

