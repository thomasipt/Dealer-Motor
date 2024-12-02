VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form JL004KWT_KIRIM 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DELIVERY ORDER"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9225
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   9225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   120
      TabIndex        =   51
      Text            =   "21"
      Top             =   2625
      Width           =   1680
   End
   Begin VB.Frame Frame4 
      Caption         =   "DENGAN HURUF"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   60
      TabIndex        =   52
      Top             =   2370
      Width           =   1815
   End
   Begin VB.TextBox Text100 
      Height          =   420
      Left            =   -3390
      TabIndex        =   50
      Text            =   "Text100"
      Top             =   5175
      Width           =   2910
   End
   Begin VB.TextBox Text20 
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
      Left            =   3255
      TabIndex        =   44
      Text            =   "TTD 2"
      Top             =   7875
      Width           =   2715
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
      Left            =   60
      TabIndex        =   39
      Text            =   "TTD 1"
      Top             =   7890
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
      Left            =   6450
      TabIndex        =   38
      Text            =   "TTD 3"
      Top             =   7875
      Width           =   2715
   End
   Begin VB.Frame Frame3 
      Caption         =   "URAIAN"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5040
      Left            =   6180
      TabIndex        =   31
      Top             =   1275
      Width           =   2985
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
         Height          =   915
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   37
         Text            =   "JL004KWT_KIRIM.frx":0000
         Top             =   4050
         Width           =   2850
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
         Height          =   930
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         Text            =   "JL004KWT_KIRIM.frx":0010
         Top             =   2745
         Width           =   2850
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "URAIAN"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5040
      Left            =   2025
      TabIndex        =   15
      Top             =   1275
      Width           =   3990
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
         Height          =   915
         Left            =   75
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   36
         Text            =   "JL004KWT_KIRIM.frx":0016
         Top             =   4080
         Width           =   3795
      End
      Begin VB.TextBox Text14 
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
         Height          =   915
         Left            =   75
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   35
         Text            =   "JL004KWT_KIRIM.frx":002C
         Top             =   2760
         Width           =   3795
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
         Left            =   1140
         TabIndex        =   30
         Text            =   "13"
         Top             =   2115
         Width           =   2730
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
         Left            =   1140
         TabIndex        =   29
         Text            =   "12"
         Top             =   1860
         Width           =   2730
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
         Left            =   1140
         TabIndex        =   28
         Text            =   "11"
         Top             =   1605
         Width           =   2730
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
         Left            =   1140
         TabIndex        =   27
         Text            =   "10"
         Top             =   1350
         Width           =   2730
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
         Left            =   1140
         TabIndex        =   26
         Text            =   "9"
         Top             =   1095
         Width           =   2730
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
         Left            =   1140
         TabIndex        =   25
         Text            =   "7"
         Top             =   840
         Width           =   2730
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
         Height          =   195
         Left            =   1140
         TabIndex        =   17
         Text            =   "6"
         Top             =   585
         Width           =   2730
      End
      Begin VB.Label Label15 
         Caption         =   "Perlengkapan lain :"
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
         Left            =   75
         TabIndex        =   34
         Top             =   3780
         Width           =   2040
      End
      Begin VB.Label Label14 
         Caption         =   "Dengan Perlengkapan sbb :"
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
         Left            =   75
         TabIndex        =   33
         Top             =   2505
         Width           =   2040
      End
      Begin VB.Label Label13 
         Caption         =   "Kondisi"
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
         Left            =   75
         TabIndex        =   24
         Top             =   2100
         Width           =   1020
      End
      Begin VB.Label Label12 
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
         Left            =   75
         TabIndex        =   23
         Top             =   1845
         Width           =   1020
      End
      Begin VB.Label Label11 
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
         Left            =   75
         TabIndex        =   22
         Top             =   1590
         Width           =   1020
      End
      Begin VB.Label Label10 
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
         Left            =   75
         TabIndex        =   21
         Top             =   1335
         Width           =   1020
      End
      Begin VB.Label Label9 
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
         Left            =   75
         TabIndex        =   20
         Top             =   1080
         Width           =   1020
      End
      Begin VB.Label Label7 
         Caption         =   "Type / Jenis"
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
         Left            =   75
         TabIndex        =   19
         Top             =   825
         Width           =   1020
      End
      Begin VB.Label Label6 
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
         Left            =   75
         TabIndex        =   18
         Top             =   570
         Width           =   1020
      End
      Begin VB.Label Label5 
         Caption         =   "Kendaraan"
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
         Left            =   75
         TabIndex        =   16
         Top             =   300
         Width           =   1020
      End
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
      Left            =   120
      TabIndex        =   13
      Text            =   "5"
      Top             =   1530
      Width           =   1290
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
      Left            =   7260
      TabIndex        =   10
      Text            =   "4"
      Top             =   825
      Width           =   1905
   End
   Begin VB.TextBox Text3 
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
      Left            =   7260
      TabIndex        =   8
      Text            =   "3"
      Top             =   525
      Width           =   1905
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
      Left            =   6465
      TabIndex        =   6
      Text            =   "8"
      Top             =   75
      Width           =   1905
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
      Left            =   495
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "JL004KWT_KIRIM.frx":0036
      Top             =   285
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
      Left            =   495
      TabIndex        =   3
      Text            =   "1"
      Top             =   75
      Width           =   4410
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
      Left            =   7162
      TabIndex        =   1
      Top             =   8505
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
      Left            =   652
      TabIndex        =   0
      Top             =   8505
      Width           =   1410
   End
   Begin VB.PictureBox Picture1 
      Height          =   765
      Left            =   -113
      ScaleHeight     =   705
      ScaleWidth      =   9390
      TabIndex        =   2
      Top             =   8400
      Width           =   9450
   End
   Begin VB.Frame Frame1 
      Caption         =   "KWANTUM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   60
      TabIndex        =   12
      Top             =   1275
      Width           =   1815
      Begin VB.Label Label4 
         Caption         =   "UNIT"
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
         Left            =   1410
         TabIndex        =   14
         Top             =   255
         Width           =   360
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C0C0C0&
      Height          =   5250
      Left            =   -180
      ScaleHeight     =   5190
      ScaleWidth      =   9570
      TabIndex        =   49
      Top             =   1170
      Width           =   9630
   End
   Begin Crystal.CrystalReport crpt 
      Left            =   2670
      Top             =   6630
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      Caption         =   "keadaan baik,"
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
      Left            =   6450
      TabIndex        =   48
      Top             =   6870
      Width           =   2715
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      Caption         =   "Barang diterima dalam"
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
      Left            =   6450
      TabIndex        =   47
      Top             =   6660
      Width           =   2715
   End
   Begin VB.Label Label17 
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
      Left            =   3255
      TabIndex        =   46
      Top             =   7965
      Width           =   2715
   End
   Begin VB.Label Label16 
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
      Left            =   3255
      TabIndex        =   45
      Top             =   7170
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
      Left            =   60
      TabIndex        =   43
      Top             =   7980
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
      Left            =   6450
      TabIndex        =   42
      Top             =   7965
      Width           =   2715
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      Caption         =   "Hormat kami,"
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
      Left            =   60
      TabIndex        =   41
      Top             =   7185
      Width           =   2715
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      Caption         =   "Penerima,"
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
      Left            =   6450
      TabIndex        =   40
      Top             =   7215
      Width           =   2715
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Tanggal :"
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
      Left            =   5760
      TabIndex        =   11
      Top             =   795
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Sales Order Nomor :"
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
      Left            =   5760
      TabIndex        =   9
      Top             =   510
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "Tanggal :"
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
      Left            =   5760
      TabIndex        =   7
      Top             =   60
      Width           =   675
   End
   Begin VB.Label Label3 
      Caption         =   "YTH :"
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
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   405
   End
End
Attribute VB_Name = "JL004KWT_KIRIM"
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

If Text5 = "" Or Text21 = "" Or Text8 = "" Or Text3 = "" Or Text4 = "" Or Text14 = "" Or Text15 = "" Or Text17 = "" Or Text16 = "" Or Text19 = "" Or Text20 = "" Or Text18 = "" Then
    MsgBox "MASIH ADA YANG KOSONG", vbCritical, "KONFIRMASI"
    Exit Sub
End If

If OYEN = 1 Then
    SDel = "Delete From M001_KWTKIRIM where NO_FAK = '" + Trim(Text100) + "'"
    Set RDel = RDCO.OpenResultset(SDel, rdOpenDynamic, rdConcurRowVer)
End If

SSave = "Select * From M001_KWTKIRIM"
Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
RSave.AddNew
    RSave("NO_FAK") = Trim(Text100)
    RSave("TANGGAL") = Trim(Text8)
    RSave("NOMOR_DO") = Trim(Text3)
    RSave("TANGGAL_DO") = Trim(Text4)
    
    RSave("KWANTUM") = Trim(Text5)
    RSave("KWANTUM_TERBILANG") = Trim(Text21)
    
    RSave("NAMA") = Trim(Text1)
    RSave("ALAMAT") = Trim(Text2)
    
    RSave("MERK") = Trim(Text6)
    RSave("TYPE") = Trim(Text7)
    RSave("RANGKA") = Trim(Text9)
    RSave("MESIN") = Trim(Text10)
    RSave("WARNA") = Trim(Text11)
    RSave("TAHUN") = Trim(Text12)
    RSave("KONDISI") = Trim(Text13)
    
    RSave("PERLENGKAPAN_1") = Trim(Text14)
    RSave("PERLENGKAPAN_2") = Trim(Text15)
    
    RSave("STS_1") = Trim(Text17)
    RSave("STS_2") = Trim(Text16)
   
    RSave("TTD_1") = Trim(Text19)
    RSave("TTD_2") = Trim(Text20)
    RSave("TTD_3") = Trim(Text18)
RSave.Update
RSave.Close
Set RSave = Nothing

Tanya = MsgBox("CETAK KWITANSI", vbOKCancel, "KONFIRMASI")
    If Tanya = vbOK Then
        crpt.ReportFileName = App.Path + "\ReportD\KWT_KIRIM.rpt"
        crpt.SelectionFormula = "{M001_KWTKIRIM.NO_FAK} = '" + Trim(Text100) + "'"
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
Text4 = TglOK

Call CekData

End Sub

Private Sub CekData()
SCari = "Select * From M001_KWTKIRIM where NO_FAK = '" + Trim(Text100) + "'"
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
SToket = "Select * From M001_KWTKIRIM where NO_FAK = '" + Trim(Text100) + "'"
Set RToket = RDCO.OpenResultset(SToket, rdOpenDynamic, rdOpenKeyset)
If RToket.RowCount <> 0 Then
    Text100 = RToket("NO_FAK")
    Text8 = RToket("TANGGAL")
    Text3 = RToket("NOMOR_DO")
    Text4 = RToket("TANGGAL_DO")
    
    Text5 = RToket("KWANTUM")
    Text21 = RToket("KWANTUM_TERBILANG")
    
    Text1 = RToket("NAMA")
    Text2 = RToket("ALAMAT")
    
    Text6 = RToket("MERK")
    Text7 = RToket("TYPE")
    Text9 = RToket("RANGKA")
    Text10 = RToket("MESIN")
    Text11 = RToket("WARNA")
    Text12 = RToket("TAHUN")
    Text13 = RToket("KONDISI")
    
    Text14 = RToket("PERLENGKAPAN_1")
    Text15 = RToket("PERLENGKAPAN_2")
    
    Text17 = RToket("STS_1")
    Text16 = RToket("STS_2")
    
    Text19 = RToket("TTD_1")
    Text20 = RToket("TTD_2")
    Text18 = RToket("TTD_3")
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
    Text3 = "-"
    
    Text5 = "1"
    Text21 = "SATU"

    Text6 = "SUZUKI"
    Text7 = Format(RToket("TYPE"), ">")
    Text9 = Format(RToket("RANGKA"), ">")
    Text10 = Format(RToket("MESIN"), ">")
    Text11 = Format(RToket("WARNA"), ">")
    Text12 = Format(RToket("TAHUN"), ">")

    Text13 = "BAIK DAN BARU 100%"
    
End If
RToket.Close
Set RToket = Nothing
    OYEN = 0
End Sub

Private Sub Text14_LostFocus()
    Text14 = Format(Text14, ">")
End Sub

Private Sub Text15_LostFocus()
    Text15 = Format(Text15, ">")
End Sub

Private Sub Text16_LostFocus()
    Text16 = Format(Text16, ">")
End Sub

Private Sub Text17_LostFocus()
    Text17 = Format(Text17, ">")
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

Private Sub Text20_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text20 = Format(Text20, ">")
    cmdKWT.SetFocus
End If
End Sub

Private Sub Text21_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text21 = Format(Text21, ">")
    cmdKWT.SetFocus
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text3 = Format(Text3, ">")
    cmdKWT.SetFocus
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If Text5 = "" Then Text5 = 1
If Not IsNumeric(Text5) Then
    Text5.SetFocus
    MsgBox "JUMLAH MENGGUNAKAN ANGKA", vbCritical, "TYPE DATA SALAH"
    Exit Sub
End If
End Sub
