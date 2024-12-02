VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form JL004KWT_JUAL 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "KWITANSI PENJUALAN"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6045
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text200 
      Height          =   285
      Left            =   6255
      TabIndex        =   52
      Text            =   "200"
      Top             =   3750
      Width           =   1455
   End
   Begin VB.TextBox Text100 
      Height          =   240
      Left            =   6255
      TabIndex        =   51
      Text            =   "100"
      Top             =   3105
      Width           =   1455
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
      Left            =   232
      TabIndex        =   0
      Top             =   6975
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
      Left            =   4402
      TabIndex        =   1
      Top             =   6975
      Width           =   1410
   End
   Begin VB.TextBox Text21 
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
      Left            =   97
      TabIndex        =   47
      Text            =   "TTD 1"
      Top             =   6375
      Width           =   2715
   End
   Begin VB.TextBox Text20 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Left            =   1102
      TabIndex        =   45
      Text            =   "20"
      Top             =   5325
      Width           =   1560
   End
   Begin VB.TextBox Text19 
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
      Left            =   3217
      TabIndex        =   43
      Text            =   "TTD"
      Top             =   6375
      Width           =   2715
   End
   Begin VB.TextBox Text18 
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
      Left            =   3622
      TabIndex        =   42
      Text            =   "18"
      Top             =   5775
      Width           =   1905
   End
   Begin VB.TextBox Text17 
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
      Left            =   1102
      TabIndex        =   40
      Text            =   "17"
      Top             =   5025
      Width           =   4845
   End
   Begin VB.TextBox Text16 
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
      Left            =   1785
      TabIndex        =   37
      Text            =   "16"
      Top             =   4710
      Width           =   4155
   End
   Begin VB.TextBox Text15 
      Alignment       =   1  'Right Justify
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
      Left            =   1785
      TabIndex        =   34
      Text            =   "15"
      Top             =   4440
      Width           =   1440
   End
   Begin VB.TextBox Text14 
      Alignment       =   1  'Right Justify
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
      Left            =   1785
      TabIndex        =   31
      Text            =   "14"
      Top             =   4170
      Width           =   1440
   End
   Begin VB.TextBox Text13 
      Alignment       =   1  'Right Justify
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
      Left            =   1785
      TabIndex        =   28
      Text            =   "13"
      Top             =   3900
      Width           =   1440
   End
   Begin VB.TextBox Text12 
      Alignment       =   1  'Right Justify
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
      Left            =   1785
      TabIndex        =   25
      Text            =   "12"
      Top             =   3630
      Width           =   1440
   End
   Begin VB.TextBox Text11 
      Alignment       =   1  'Right Justify
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
      Left            =   1785
      TabIndex        =   22
      Text            =   "11"
      Top             =   3360
      Width           =   1440
   End
   Begin VB.TextBox Text6 
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
      Height          =   480
      Left            =   1522
      MultiLine       =   -1  'True
      TabIndex        =   20
      Text            =   "JL004KWT_JUAL.frx":0000
      Top             =   2745
      Width           =   1665
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
      Left            =   4282
      TabIndex        =   18
      Text            =   "10"
      Top             =   3015
      Width           =   1665
   End
   Begin VB.TextBox Text9 
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
      Left            =   4282
      TabIndex        =   16
      Text            =   "9"
      Top             =   2730
      Width           =   1665
   End
   Begin VB.TextBox Text8 
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
      Left            =   4282
      TabIndex        =   11
      Text            =   "8"
      Top             =   2460
      Width           =   1665
   End
   Begin VB.TextBox Text7 
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
      Left            =   4282
      TabIndex        =   10
      Text            =   "7"
      Top             =   2175
      Width           =   1665
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
      Left            =   1522
      TabIndex        =   9
      Text            =   "5"
      Top             =   2460
      Width           =   1665
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
      Left            =   1522
      TabIndex        =   8
      Text            =   "4"
      Top             =   2175
      Width           =   1665
   End
   Begin VB.TextBox Text3 
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
      Height          =   735
      Left            =   1522
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "JL004KWT_JUAL.frx":0002
      Top             =   1260
      Width           =   4410
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
      Height          =   735
      Left            =   1522
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "JL004KWT_JUAL.frx":0004
      Top             =   405
      Width           =   4410
   End
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
      Left            =   1522
      TabIndex        =   2
      Text            =   "1"
      Top             =   120
      Width           =   4410
   End
   Begin VB.PictureBox Picture1 
      Height          =   765
      Left            =   -120
      ScaleHeight     =   705
      ScaleWidth      =   6225
      TabIndex        =   50
      Top             =   6885
      Width           =   6285
   End
   Begin Crystal.CrystalReport crpt 
      Left            =   5145
      Top             =   3690
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label28 
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
      Left            =   90
      TabIndex        =   49
      Top             =   6465
      Width           =   2715
   End
   Begin VB.Label Label27 
      Alignment       =   2  'Center
      Caption         =   "Disetujui Pembeli"
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
      Left            =   90
      TabIndex        =   48
      Top             =   5760
      Width           =   2715
   End
   Begin VB.Label Label26 
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
      Left            =   90
      TabIndex        =   46
      Top             =   5325
      Width           =   960
   End
   Begin VB.Label Label25 
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
      Left            =   3210
      TabIndex        =   44
      Top             =   6465
      Width           =   2715
   End
   Begin VB.Label Label22 
      Caption         =   "Keterangan"
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
      Left            =   90
      TabIndex        =   41
      Top             =   5010
      Width           =   1395
   End
   Begin VB.Label Label21 
      Caption         =   "Rp."
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
      Left            =   1515
      TabIndex        =   39
      Top             =   4695
      Width           =   240
   End
   Begin VB.Label Label20 
      Caption         =   "Diangsur / Dibayar"
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
      Left            =   90
      TabIndex        =   38
      Top             =   4695
      Width           =   1395
   End
   Begin VB.Label Label19 
      Caption         =   "Rp."
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
      Left            =   1515
      TabIndex        =   36
      Top             =   4425
      Width           =   240
   End
   Begin VB.Label Label18 
      Caption         =   "Sisa Pinjaman"
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
      Left            =   90
      TabIndex        =   35
      Top             =   4425
      Width           =   1395
   End
   Begin VB.Label Label17 
      Caption         =   "Rp."
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
      Left            =   1515
      TabIndex        =   33
      Top             =   4155
      Width           =   240
   End
   Begin VB.Label Label13 
      Caption         =   "Uang Muka"
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
      Left            =   90
      TabIndex        =   32
      Top             =   4155
      Width           =   1395
   End
   Begin VB.Label Label12 
      Caption         =   "Rp."
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
      Left            =   1515
      TabIndex        =   30
      Top             =   3885
      Width           =   240
   End
   Begin VB.Label Label11 
      Caption         =   "Jumlah"
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
      Left            =   90
      TabIndex        =   29
      Top             =   3885
      Width           =   1395
   End
   Begin VB.Label Label10 
      Caption         =   "Rp."
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
      Left            =   1515
      TabIndex        =   27
      Top             =   3615
      Width           =   240
   End
   Begin VB.Label Label9 
      Caption         =   "Administrasi"
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
      Left            =   90
      TabIndex        =   26
      Top             =   3615
      Width           =   1395
   End
   Begin VB.Label Label8 
      Caption         =   "Rp."
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
      Left            =   1515
      TabIndex        =   24
      Top             =   3345
      Width           =   240
   End
   Begin VB.Label Label7 
      Caption         =   "Dengan harga jadi"
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
      Left            =   90
      TabIndex        =   23
      Top             =   3345
      Width           =   1395
   End
   Begin VB.Label Label6 
      Caption         =   "Type / Tahun"
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
      Left            =   90
      TabIndex        =   21
      Top             =   2730
      Width           =   1005
   End
   Begin VB.Label Label5 
      Caption         =   "No. BPKB"
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
      Left            =   3240
      TabIndex        =   19
      Top             =   3000
      Width           =   1005
   End
   Begin VB.Label Label4 
      Caption         =   "No. Polisi"
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
      Left            =   3255
      TabIndex        =   17
      Top             =   2730
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
      Left            =   3255
      TabIndex        =   15
      Top             =   2445
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
      Left            =   3255
      TabIndex        =   14
      Top             =   2160
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
      Left            =   90
      TabIndex        =   13
      Top             =   2445
      Width           =   1005
   End
   Begin VB.Label Label24 
      Caption         =   "Merk"
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
      Left            =   90
      TabIndex        =   12
      Top             =   2160
      Width           =   1005
   End
   Begin VB.Label Label2 
      Caption         =   "Uang Sebanyak"
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
      Left            =   90
      TabIndex        =   7
      Top             =   1245
      Width           =   1395
   End
   Begin VB.Label Label1 
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
      Left            =   97
      TabIndex        =   5
      Top             =   390
      Width           =   1395
   End
   Begin VB.Label Label3 
      Caption         =   "Telah terima dari"
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
      Left            =   97
      TabIndex        =   3
      Top             =   105
      Width           =   1395
   End
End
Attribute VB_Name = "JL004KWT_JUAL"
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

If Text16 = "" Or Text17 = "" Or Text18 = "" Or Text21 = "" Or Text19 = "" Then
    MsgBox "MASIH ADA YANG KOSONG", vbCritical, "KONFIRMASI"
    Exit Sub
End If

If OYEN = 1 Then
    SDel = "Delete From M001_KWTJUAL where NO_FAK = '" + Trim(Text100) + "'"
    Set RDel = RDCO.OpenResultset(SDel, rdOpenDynamic, rdConcurRowVer)
End If

SSave = "Select * From M001_KWTJUAL"
Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
RSave.AddNew
    RSave("NO_FAK") = Trim(Text100)
    RSave("NAMA") = Trim(Text1)
    RSave("ALAMAT") = Trim(Text2)
    RSave("TERBILANG") = Trim(Text3)
    
    RSave("TYPE") = Trim(Text4)
    RSave("WARNA") = Trim(Text5)
    RSave("TAHUN") = Trim(Text6)
    RSave("MESIN") = Trim(Text7)
    RSave("RANGKA") = Trim(Text8)
    RSave("POLISI") = Trim(Text9)
    RSave("BPKB") = Trim(Text10)
    
    RSave("H_JADI") = CCur(Text11)
    RSave("ADM") = CCur(Text12)
    RSave("JUMLAH") = CCur(Text13)
    RSave("DP") = CCur(Text14)
    RSave("SISA") = CCur(Text15)
    RSave("ANGSUR") = Trim(Text16)
    RSave("KETERANGAN") = Trim(Text17)
    RSave("TANGGAL") = Trim(Text18)
    RSave("NOMINAL") = CCur(Text20)
    
    RSave("TTD_1") = Trim(Text21)
    RSave("TTD_2") = Trim(Text19)
RSave.Update
RSave.Close
Set RSave = Nothing

Tanya = MsgBox("CETAK KWITANSI", vbOKCancel, "KONFIRMASI")
    If Tanya = vbOK Then
        crpt.ReportFileName = App.Path + "\ReportD\KWT_JUAL.rpt"
        crpt.SelectionFormula = "{M001_KWTJUAL.NO_FAK} = '" + Trim(Text100) + "'"
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

Text8 = TglOK

Call CekData

End Sub

Private Sub CekData()
SCari = "Select * From M001_KWTJUAL where NO_FAK = '" + Trim(Text100) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdOpenKeyset)
If RCari.RowCount <> 0 Then
    Call CariData
Else
    Call CariData2
End If
RCari.Close
Set RCari = Nothing
    
    Text3 = Terbilang(Text20)
    
End Sub

Private Sub CariData()
SToket = "Select * From M001_KWTJUAL where NO_FAK = '" + Trim(Text100) + "'"
Set RToket = RDCO.OpenResultset(SToket, rdOpenDynamic, rdOpenKeyset)
If RToket.RowCount <> 0 Then
    Text1 = Format(RToket("NAMA"), ">")
    Text2 = Format(RToket("ALAMAT"), ">")

    Text4 = Format(RToket("TYPE"), ">")
    Text5 = Format(RToket("WARNA"), ">")
    Text6 = Format(RToket("TAHUN"), ">")
    Text7 = Format(RToket("MESIN"), ">")
    Text8 = Format(RToket("RANGKA"), ">")
    Text9 = Format(RToket("POLISI"), ">")
    Text10 = Format(RToket("BPKB"), ">")
    
    Text11 = Format(RToket("H_JADI"), "##,###.00")
    Text12 = Format(RToket("ADM"), "##,###.00")
    Text13 = Format(RToket("JUMLAH"), "##,###.00")
    Text14 = Format(RToket("DP"), "##,###.00")
    Text15 = Format(RToket("SISA"), "##,###.00")
    
    Text16 = Format(RToket("ANGSUR"), ">")
    
    Text200 = ""
    
    Text17 = Format(RToket("KETERANGAN"), ">")
    Text20 = Format(RToket("NOMINAL"), "##,###.00")
        
    Text18 = TglOK
    Text21 = Format(RToket("TTD_1"), ">")
    Text19 = Format(RToket("TTD_2"), ">")
    
End If
RToket.Close
Set RToket = Nothing
    OYEN = 1
End Sub

Private Sub CariData2()
SToket = "Select * From M001 where NO_FAK = '" + Trim(Text100) + "'"
Set RToket = RDCO.OpenResultset(SToket, rdOpenDynamic, rdOpenKeyset)
If RToket.RowCount <> 0 Then
    Text1 = Format(RToket("NAMA_PEMBELI"), ">")
    Text2 = Format(RToket("ALAMAT_1"), ">") + " , " + Format(RToket("ALAMAT_2"), ">")

    Text4 = "SUZUKI"
    Text5 = Format(RToket("WARNA"), ">")
    Text6 = Format(RToket("TYPE"), ">") + " / " + Format(RToket("TAHUN"), ">")
    Text7 = Format(RToket("MESIN"), ">")
    Text8 = Format(RToket("RANGKA"), ">")
    Text9 = Format(RToket("NO_STNK"), ">")
    Text10 = Format(RToket("NO_BPKB"), ">")
    
    Text11 = Format(RToket("H_OTR"), "##,###.00")
    Text12 = "0,00"
    Text13 = "0,00"
    Text14 = Format(RToket("TUNAI"), "##,###.00")
    Text15 = Format(RToket("PIUTANG"), "##,###.00")
    
    Text200 = Format(RToket("NO_HUTANG"), ">")
    
    If Text200 <> "0" Then
        SToge = "Select * From P002 where NOMOR_PIN = '" + Trim(Text200) + "'"
        Set RToge = RDCO.OpenResultset(SToge, rdOpenDynamic, rdOpenKeyset)
        If RToge.RowCount <> 0 Then
            Text16 = Format(RToge("SYARAT_BYR"), ">") + " BY " + Format(RToge("NAMA_NAS"), ">")
        End If
        RToge.Close
        Set RToge = Nothing
    End If
   
    Text17 = "MOTOR BARU LENGKAP DENGAN PERLENGKAPANNYA"
    Text20 = Format(RToket("TUNAI"), "##,###.00")
        
    Text18 = TglOK
    Text21 = "TTD 1"
    Text19 = "TTD 2"
End If
RToket.Close
Set RToket = Nothing
    OYEN = 0
End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text16 = Format(Text16, ">")
    cmdKWT.SetFocus
End If
End Sub

Private Sub Text17_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text17 = Format(Text17, ">")
    cmdKWT.SetFocus
End If
End Sub

Private Sub Text18_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text18 = Format(Text18, ">")
    cmdKWT.SetFocus
End If
End Sub

Private Sub Text19_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text19 = Format(Text19, ">")
    cmdKWT.SetFocus
End If
End Sub

Private Sub Text21_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text21 = Format(Text21, ">")
    cmdKWT.SetFocus
End If
End Sub
