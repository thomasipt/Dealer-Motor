VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form BL01A 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TRANSAKSI PEMBELIAN SPAREPART"
   ClientHeight    =   8370
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   11865
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   1717
      MaxLength       =   15
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   195
      Width           =   1950
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   1717
      TabIndex        =   1
      Text            =   "Text4"
      Top             =   600
      Width           =   1950
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6945
      Left            =   97
      TabIndex        =   20
      Top             =   1155
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   12250
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "DAFTAR BARANG"
      TabPicture(0)   =   "BL01A.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label19"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Grid"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Grid3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Grid2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "CmdOK"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Text7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Text3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Text2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Combo1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Frame3"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "CARA PEMBAYARAN"
      TabPicture(1)   =   "BL01A.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame5"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame6"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "TblSave"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame2"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "TblClose"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Frame4"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "TJatuh"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "TMulai"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Frame7"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "GridGL"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).ControlCount=   11
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         Caption         =   "INFORMASI TOTAL BIAYA PEMBELIAN"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   645
         Left            =   75
         TabIndex        =   73
         Top             =   315
         Width           =   11490
         Begin VB.TextBox Text16 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   360
            Left            =   7110
            TabIndex        =   2
            Text            =   "Text16"
            Top             =   180
            Width           =   1905
         End
         Begin VB.Label Label4 
            Caption         =   "INFORMASI BIAYA PEMBELIAN (UNTUK MENGHITUNG HARGA POKOK PEMBELIAN) Rp. "
            Height          =   285
            Left            =   135
            TabIndex        =   74
            Top             =   270
            Width           =   7080
         End
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   2130
         TabIndex        =   4
         Top             =   1050
         Width           =   2955
      End
      Begin VB.Frame Frame1 
         Caption         =   "PENTING"
         Height          =   870
         Left            =   -69015
         TabIndex        =   68
         Top             =   585
         Visible         =   0   'False
         Width           =   1995
         Begin VB.TextBox Text8 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   360
            Left            =   180
            TabIndex        =   69
            Text            =   "Text8"
            Top             =   315
            Width           =   510
         End
         Begin VB.Label Label16 
            Caption         =   "DISCOUNT"
            Height          =   285
            Left            =   675
            TabIndex        =   70
            Top             =   360
            Width           =   1590
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
         Height          =   2850
         Left            =   -74820
         TabIndex        =   58
         Top             =   1845
         Width           =   5640
         Begin VB.TextBox Text5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   360
            Left            =   1530
            TabIndex        =   10
            Text            =   "Text5"
            Top             =   810
            Width           =   1860
         End
         Begin VB.TextBox Text14 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   360
            Left            =   1530
            TabIndex        =   12
            Text            =   "Text14"
            Top             =   1800
            Width           =   1860
         End
         Begin VB.TextBox Text15 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   360
            Left            =   1530
            TabIndex        =   11
            Text            =   "Text15"
            Top             =   1305
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
            Left            =   3510
            TabIndex        =   65
            Top             =   2340
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
            Left            =   945
            TabIndex        =   64
            Top             =   2430
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
            Left            =   3510
            TabIndex        =   63
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
            Left            =   945
            TabIndex        =   62
            Top             =   405
            Width           =   1590
         End
         Begin VB.Label Label11 
            Caption         =   "> TUNAI"
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   180
            TabIndex        =   61
            Top             =   855
            Width           =   1230
         End
         Begin VB.Label Label12 
            Caption         =   "> HUTANG"
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   180
            TabIndex        =   60
            Top             =   1350
            Width           =   1230
         End
         Begin VB.Label Label13 
            Caption         =   "> NON TUNAI"
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   180
            TabIndex        =   59
            Top             =   1845
            Width           =   1230
         End
      End
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         Caption         =   "Data Debitur"
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
         Height          =   4245
         Left            =   -69060
         TabIndex        =   38
         Top             =   450
         Width           =   5550
         Begin VB.TextBox Text12 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   1755
            TabIndex        =   44
            Text            =   "Text12"
            Top             =   2070
            Width           =   1185
         End
         Begin VB.TextBox Text13 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   1755
            TabIndex        =   43
            Text            =   "Text13"
            Top             =   2475
            Width           =   1185
         End
         Begin VB.TextBox Text11 
            BackColor       =   &H00FFFFC0&
            Height          =   360
            Left            =   1755
            TabIndex        =   42
            Text            =   "Text11"
            Top             =   3420
            Width           =   1680
         End
         Begin VB.TextBox Text9 
            BackColor       =   &H00FFFFC0&
            Height          =   360
            Left            =   1755
            MaxLength       =   12
            TabIndex        =   41
            Text            =   "Text9"
            Top             =   1170
            Width           =   1410
         End
         Begin VB.CommandButton InfoSPL2 
            Caption         =   "LAMA"
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
            Left            =   3330
            TabIndex        =   40
            Top             =   1170
            WhatsThisHelpID =   870
            Width           =   825
         End
         Begin VB.CommandButton EntriSPL2 
            Caption         =   "BARU"
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
            Left            =   4275
            TabIndex        =   39
            Top             =   1170
            WhatsThisHelpID =   870
            Width           =   915
         End
         Begin VB.Label Label29 
            Caption         =   "NO. SUPPLIER"
            Height          =   315
            Left            =   180
            TabIndex        =   57
            Top             =   1170
            Width           =   1230
         End
         Begin VB.Label Label30 
            Caption         =   "NAMA SUPPLIER"
            Height          =   315
            Left            =   180
            TabIndex        =   56
            Top             =   1575
            Width           =   1365
         End
         Begin VB.Line Line4 
            X1              =   180
            X2              =   5400
            Y1              =   1980
            Y2              =   1980
         End
         Begin VB.Label Label32 
            Caption         =   "TANGGAL MULAI"
            Height          =   315
            Left            =   180
            TabIndex        =   55
            Top             =   2115
            Width           =   1545
         End
         Begin VB.Label Label33 
            Caption         =   "JATUH TEMPO"
            Height          =   315
            Left            =   180
            TabIndex        =   54
            Top             =   2520
            Width           =   1455
         End
         Begin VB.Label Label34 
            Caption         =   "NOMINAL"
            Height          =   315
            Left            =   180
            TabIndex        =   53
            Top             =   2970
            Width           =   1365
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Label35"
            Height          =   315
            Left            =   1755
            TabIndex        =   52
            Top             =   2925
            Width           =   1950
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Label5"
            Height          =   315
            Left            =   1755
            TabIndex        =   51
            Top             =   765
            Width           =   1455
         End
         Begin VB.Label Label37 
            Caption         =   "NO. HUTANG"
            Height          =   330
            Left            =   180
            TabIndex        =   50
            Top             =   765
            Width           =   1230
         End
         Begin VB.Label Label38 
            Caption         =   "SYARAT PEMBAYARAN"
            Height          =   495
            Left            =   180
            TabIndex        =   49
            Top             =   3285
            Width           =   1230
         End
         Begin VB.Label Label39 
            Caption         =   "KODE HUTANG"
            Height          =   315
            Left            =   180
            TabIndex        =   48
            Top             =   405
            Width           =   1275
         End
         Begin VB.Label Label40 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "Label40"
            Height          =   315
            Left            =   1755
            TabIndex        =   47
            Top             =   360
            Width           =   645
         End
         Begin VB.Label Label41 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Label41"
            Height          =   315
            Left            =   2475
            TabIndex        =   46
            Top             =   360
            Width           =   2940
         End
         Begin VB.Label Label24 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Label24"
            Height          =   315
            Left            =   1755
            TabIndex        =   45
            Top             =   1620
            Width           =   3660
         End
      End
      Begin VB.CommandButton TblSave 
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
         Left            =   -66225
         TabIndex        =   13
         Top             =   6345
         Width           =   1230
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
         Left            =   -74820
         TabIndex        =   31
         Top             =   4815
         Width           =   5640
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
            TabIndex        =   33
            Top             =   360
            Width           =   1725
         End
         Begin VB.TextBox Text10 
            BackColor       =   &H00C0C0C0&
            Height          =   360
            Left            =   1530
            TabIndex        =   32
            Text            =   "Text10"
            Top             =   360
            Width           =   1050
         End
         Begin VB.Label InfoGL2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Info Kode Sub GL"
            Height          =   360
            Left            =   2700
            TabIndex        =   37
            Top             =   360
            Width           =   1725
         End
         Begin VB.Label Label28 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Label28"
            Height          =   330
            Left            =   1530
            TabIndex        =   36
            Top             =   810
            Width           =   3975
         End
         Begin VB.Label Label14 
            Caption         =   "KODE SUB GL"
            Height          =   330
            Left            =   135
            TabIndex        =   35
            Top             =   405
            Width           =   1455
         End
         Begin VB.Label Label23 
            Caption         =   "NAMA SUB GL"
            Height          =   375
            Left            =   135
            TabIndex        =   34
            Top             =   810
            Width           =   1140
         End
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   5130
         TabIndex        =   5
         Text            =   "Text2"
         Top             =   1050
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   6135
         TabIndex        =   6
         Text            =   "Text3"
         Top             =   1050
         Width           =   1050
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   660
         TabIndex        =   3
         Text            =   "Text7"
         Top             =   1050
         Width           =   1415
      End
      Begin VB.CommandButton TblClose 
         Caption         =   "CLOSE"
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
         Left            =   -64695
         TabIndex        =   30
         Top             =   6345
         Width           =   1230
      End
      Begin VB.CommandButton CmdOK 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8670
         TabIndex        =   7
         Top             =   1020
         Width           =   1095
      End
      Begin VB.Frame Frame4 
         Caption         =   "Data Supplier"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   -74820
         TabIndex        =   27
         Top             =   450
         Width           =   5640
         Begin VB.TextBox Text17 
            BackColor       =   &H00FFFFC0&
            Height          =   360
            Left            =   1530
            TabIndex        =   9
            Text            =   "Text17"
            Top             =   802
            Width           =   4035
         End
         Begin VB.TextBox Text6 
            BackColor       =   &H00FFFFC0&
            Height          =   360
            Left            =   1530
            TabIndex        =   8
            Text            =   "Text6"
            Top             =   375
            Width           =   1410
         End
         Begin VB.Label Label20 
            Caption         =   "NO. SUPPLIER"
            Height          =   300
            Left            =   135
            TabIndex        =   29
            Top             =   405
            Width           =   1365
         End
         Begin VB.Label Label21 
            Caption         =   "NAMA SUPPLIER"
            Height          =   345
            Left            =   135
            TabIndex        =   28
            Top             =   810
            Width           =   1365
         End
      End
      Begin VB.CommandButton TJatuh 
         Caption         =   "..."
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
         Left            =   -66045
         TabIndex        =   26
         Top             =   2925
         Width           =   330
      End
      Begin VB.CommandButton TMulai 
         Caption         =   "..."
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
         Left            =   -66045
         TabIndex        =   25
         Top             =   2520
         Width           =   330
      End
      Begin VB.Frame Frame7 
         Appearance      =   0  'Flat
         Caption         =   "Tujuan Barang (Gudang / Toko)"
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
         Left            =   -69060
         TabIndex        =   22
         Top             =   4815
         Width           =   5550
         Begin VB.OptionButton OptionGD 
            Alignment       =   1  'Right Justify
            Caption         =   "GUDANG (Barang Masuk Gudang)"
            Enabled         =   0   'False
            Height          =   300
            Left            =   225
            TabIndex        =   24
            Top             =   405
            Width           =   4800
         End
         Begin VB.OptionButton OptionTK 
            Alignment       =   1  'Right Justify
            Caption         =   "TOKO (Barang Masuk Langsung Toko)"
            Height          =   300
            Left            =   225
            TabIndex        =   23
            Top             =   810
            Width           =   4800
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GridGL 
         Height          =   2085
         Left            =   -72120
         TabIndex        =   21
         Top             =   3060
         Width           =   3750
         _ExtentX        =   6615
         _ExtentY        =   3678
         _Version        =   393216
         Appearance      =   0
      End
      Begin MSFlexGridLib.MSFlexGrid Grid2 
         Height          =   360
         Left            =   90
         TabIndex        =   66
         Top             =   6030
         Width           =   11490
         _ExtentX        =   20267
         _ExtentY        =   635
         _Version        =   393216
         Rows            =   1
         FixedCols       =   0
         BackColorFixed  =   14737632
         ForeColorFixed  =   0
         Appearance      =   0
      End
      Begin MSFlexGridLib.MSFlexGrid Grid3 
         Height          =   390
         Left            =   90
         TabIndex        =   67
         Top             =   1020
         Width           =   9690
         _ExtentX        =   17092
         _ExtentY        =   688
         _Version        =   393216
         Rows            =   1
         FixedRows       =   0
         FixedCols       =   0
         GridColor       =   0
         Appearance      =   0
      End
      Begin MSFlexGridLib.MSFlexGrid Grid 
         Height          =   4485
         Left            =   90
         TabIndex        =   75
         Top             =   1440
         Width           =   11490
         _ExtentX        =   20267
         _ExtentY        =   7911
         _Version        =   393216
         FixedRows       =   2
         BackColorFixed  =   14737632
         ForeColorFixed  =   0
         BackColorBkg    =   16777215
         MergeCells      =   2
         Appearance      =   0
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
      Begin VB.Label Label6 
         Caption         =   "> F5 (Pembayaran Transaksi)"
         Height          =   285
         Left            =   270
         TabIndex        =   72
         Top             =   6480
         Width           =   2175
      End
      Begin VB.Label Label19 
         Caption         =   "> Double Click pada baris transaksi jika melakukan edit"
         Height          =   285
         Left            =   3465
         TabIndex        =   71
         Top             =   6525
         Width           =   4020
      End
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label3"
      Height          =   300
      Left            =   7342
      TabIndex        =   19
      Top             =   195
      Width           =   1320
   End
   Begin VB.Label Label2 
      Caption         =   "NO. TRANSAKSI"
      Height          =   300
      Left            =   5902
      TabIndex        =   18
      Top             =   195
      Width           =   1365
   End
   Begin VB.Label Label1 
      Caption         =   "NO. FAKTUR"
      Height          =   330
      Left            =   322
      TabIndex        =   17
      Top             =   240
      Width           =   1140
   End
   Begin VB.Label Label8 
      Caption         =   "TGL PEMBELIAN"
      Height          =   375
      Left            =   322
      TabIndex        =   16
      Top             =   645
      Width           =   1365
   End
   Begin VB.Label Label9 
      Caption         =   "TGL TRANSAKSI"
      Height          =   375
      Left            =   5902
      TabIndex        =   15
      Top             =   600
      Width           =   1320
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label10"
      Height          =   300
      Left            =   7342
      TabIndex        =   14
      Top             =   600
      Width           =   1320
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C00000&
      Height          =   960
      Left            =   97
      Top             =   105
      Width           =   11670
   End
End
Attribute VB_Name = "BL01A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private RCari, RJenis, RToko, RTrans, RNo, RGrid, RSPL, RBeli, RLaba, RSGL, RHutang, RDel, RGudang As rdoResultset
Private SCari, SToko, SJenis, STrans, SNo, SGrid, SSPL, SBeli, SDel As String

Private RBL, RBL2 As rdoResultset
Private SBL, SBL2 As String

Private SSGL, SHutang, SGudang As String
Private SLaba, STSSimpan As String

Private Persen, IsiGrup, LSatuan, LGrup, MTLABA
Private KodeHutang, NoHutang, SyaratByr, SGLNonKas

Private NoNo

Private Pesan As Boolean
Private HPokok As Currency
Private TGL

Private Sub DaftarBrg()
SCari = "Select Nama_JNS From B003 where KODE_IND = '153' order by Nama_Jns"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    RCari.MoveFirst
    Do While Not RCari.EOF
        Combo1.AddItem RCari("Nama_JNS")
    RCari.MoveNext
    Loop
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub Combo1_GotFocus()
SendKeys "{F4}"
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF5
        SSTab1.Tab = 1
        Text6.SetFocus
End Select
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Combo1_LostFocus()
If Combo1.Text = "" Then Exit Sub
SJenis = "Select * from B003 where NAMA_JNS = '" + Trim(Combo1.Text) + "'"
Set RJenis = RDCO.OpenResultset(SJenis, rdOpenDynamic, rdConcurRowVer)
If RJenis.RowCount <> 0 Then
    Text7 = RJenis("KODE_JNS")
'    MTLABA = RJenis("MT_LABA")
'    Persen = RJenis("PERSEN")
    Text2 = 0
    Text3 = 0
    Text5 = 0
    Text6 = 0
Else
    Combo1.SetFocus
    MsgBox "NAMA JENIS BARANG BELUM TERDAFTAR. DAFTARKAN DAHULU LEWAT MENU SYSTEM", vbInformation, "NAMA JENIS BARANG BELUM TERDAFTAR"
    SendKeys "{F4}"
End If
RJenis.Close
Set RJenis = Nothing
End Sub

Private Sub CmdOK_Click()
STSSimpan = "1"
Call SimpanDaftar
End Sub

Private Sub CmdOK_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF5
        SSTab1.Tab = 1
        Text6.SetFocus
End Select
End Sub

Private Sub EntriSPL2_Click()
C012.Show 1
Text9.SetFocus
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=DEALER", rdDriverNoPrompt, False, CN)

Pesan = True
Call DaftarBrg
Call Kosong
Call Cari_SubGL
Call Cari_Trans
'Call PindahText
Text4 = Tanggal
NoNo = 1

    SDel = "Delete From BL01"
    Set RDel = RDCO.OpenResultset(SDel, rdOpenDynamic, rdConcurRowVer)
    RDel.Close
    Set RDel = Nothing
    
End Sub

Private Sub Cari_Trans()
Dim Nomor As Double
STrans = "Select No_Trans from BL01 where user_code = '" + Operator + "'"
Set RTrans = RDCO.OpenResultset(STrans, rdOpenDynamic, rdConcurRowVer)
If RTrans.RowCount <> 0 Then
    If Pesan = True Then MsgBox "MASIH ADA TRANSAKSI PEMBELIAN YANG BELUM SELESAI", vbInformation, "DAFTAR TRANSAKSI PEMBELIAN TERSIMPAN"
    Label3 = RTrans("No_Trans")
    Call Isi_Grid
    Text16 = Format(CCur(grid2.TextMatrix(0, 8)) - CCur(Label17), "##,###.00")
Else
    SNo = "Select NOBeli from C013 where NAMA = '" + Operator + "'"
    Set RNo = RDCO.OpenResultset(SNo, rdOpenDynamic, rdConcurRowVer)
    If RNo.RowCount <> 0 Then
        Nomor = Val(RNo("NoBeli"))
        Label3 = Trim("1.") + Digit(2, Trim(Status)) + "." + Digit(7, Nomor)
    End If
    RNo.Close
    Set RNo = Nothing
End If
RTrans.Close
Set RTrans = Nothing
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

Private Sub Kosong()
InfoGL.Enabled = False
Label3 = ""
Label5 = ""
Label7 = 0
Label10 = Tanggal
Label17 = 0
Label22 = ""
Label24 = "<No Name>"
Label28 = ""
Label31 = ""
Label35 = ""
Label36 = ""
Label40 = ""
OptionGD.Value = False
OptionTK.Value = True
Text1 = ""
Text4 = ""
Text5 = 0
Text6 = ""
Text8 = 0
Text9 = ""
Text10 = ""
Text10.BackColor = &HC0C0C0
Text10.Enabled = False
Text11 = ""
Text12 = Tanggal
Text13 = ""
Text14 = 0
Text15 = 0
Text16 = 0
Text17 = ""
Call InfoGL2_Click
Call HutangNonAktif
Call FrameGLNonAktif
Call Batal
Call Siap
SSTab1.Tab = 0
End Sub

Private Sub Batal()
Combo1.Text = ""
Text2 = ""
Text3 = ""
Text6 = ""
Text7 = ""
End Sub
Private Sub FrameGLAktif()
Frame2.Enabled = True
Label14.Enabled = True
Text10.Enabled = True
Text10.BackColor = &HFFFFC0
InfoGL.Enabled = True
Label28.Enabled = True
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

Private Sub HutangAktif()
Frame6.Enabled = True
Label29.Enabled = True
Label30.Enabled = True
Label31 = Label24
Label32.Enabled = True
Label33.Enabled = True
Label34.Enabled = True
Label35 = Text15
Label37.Enabled = True
Label38.Enabled = True
Label39.Enabled = True
Label40 = ""
Text9.BackColor = &HFFFFC0
Text12.BackColor = &HFFFFC0
Text13.BackColor = &HFFFFC0
Text11.BackColor = &HFFFFC0
End Sub

Private Sub HutangNonAktif()
Frame6.Enabled = False
Label5 = ""
Label29.Enabled = False
Label30.Enabled = False
Label31 = ""
Label32.Enabled = False
Label33.Enabled = False
Label34.Enabled = False
Label35 = ""
Label36 = ""
Label37.Enabled = False
Label38.Enabled = False
Label39.Enabled = False
Label40 = ""
Label41 = ""
Text9.BackColor = &HC0C0C0
Text12.BackColor = &HC0C0C0
Text13.BackColor = &HC0C0C0
Text11.BackColor = &HC0C0C0
End Sub

Private Sub Siap()
With grid
    .Row = 0
    .Cols = 11
    .Col = 0: .ColWidth(0) = 550: .Text = "NO.": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 9
    .Col = 1: .ColWidth(1) = 1450: .Text = "KODE": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 9
    .Col = 2: .ColWidth(2) = 3000: .Text = "NAMA BARANG": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 9
    .Col = 3: .ColWidth(3) = 35: .CellFontBold = True: .CellFontSize = 9
    .Col = 4: .ColWidth(4) = 1000: .Text = "HARGA BELI": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 9
    .Col = 5: .ColWidth(5) = 1100: .Text = "HARGA BELI": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 9
    .Col = 6: .ColWidth(6) = 1400: .Text = "HARGA BELI": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 9
    .Col = 7: .ColWidth(7) = 35: .CellFontBold = True: .CellFontSize = 9: .CellFontBold = True: .CellFontSize = 9
    .Col = 8: .ColWidth(8) = 1100: .Text = "H. POKOK PEMBELIAN": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 9
    .Col = 9: .ColWidth(9) = 1400: .Text = "H. POKOK PEMBELIAN": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 9
    .Col = 10: .ColWidth(10) = 35: .CellFontBold = True: .CellFontSize = 9: .CellFontBold = True: .CellFontSize = 9
    
    .Row = 1: .CellFontBold = True: .CellFontSize = 9
    .Col = 0: .ColWidth(0) = 550: .Text = "NO.": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 9
    .Col = 1: .ColWidth(1) = 1450: .Text = "KODE": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 9
    .Col = 2: .ColWidth(2) = 3000: .Text = "NAMA BARANG": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 9
    .Col = 3: .ColWidth(3) = 35: .CellFontBold = True: .CellFontSize = 9
    .Col = 4: .ColWidth(4) = 1000: .Text = "JUMLAH": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 9
    .Col = 5: .ColWidth(5) = 1100: .Text = "H. SATUAN": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 9
    .Col = 6: .ColWidth(6) = 1400: .Text = "TOTAL HARGA": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 9
    .Col = 7: .ColWidth(7) = 35: .CellFontBold = True: .CellFontSize = 9
    .Col = 8: .ColWidth(8) = 1100: .Text = "HP SATUAN": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 9
    .Col = 9: .ColWidth(9) = 1400: .Text = "HP TOTAL": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 9
    .Col = 10: .ColWidth(10) = 35: .CellFontBold = True: .CellFontSize = 9
    
    .MergeCol(0) = True
    .MergeCol(1) = True
    .MergeCol(2) = True
    .MergeCol(3) = True
    .MergeCol(4) = True
    .MergeCol(5) = True
    .MergeCol(6) = True
    .MergeCol(7) = True
    .MergeCol(8) = True
    .MergeCol(9) = True
    .MergeRow(0) = True
    .MergeRow(1) = False
End With

With grid2
    .Row = 0
    .Cols = 10
    .Col = 0: .ColWidth(0) = 550
    .Col = 1: .ColWidth(1) = 4450: .Text = "TOTAL PEMBELIAN": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 9
    .Col = 2: .ColWidth(2) = 35
    .Col = 3: .ColWidth(3) = 1000
    .Col = 4: .ColWidth(4) = 1100
    .Col = 5: .ColWidth(5) = 1400
    .Col = 6: .ColWidth(6) = 35
    .Col = 7: .ColWidth(7) = 1100
    .Col = 8: .ColWidth(8) = 1400
    .Col = 9: .ColWidth(9) = 35
    
End With

With grid3
    .RowHeight(0) = 360
    .Row = 0
    .Cols = 10
    .Col = 0: .ColWidth(0) = 535
    .Col = 1: .ColWidth(1) = 1450
    .Col = 2: .ColWidth(2) = 3000
    .Col = 3: .ColWidth(3) = 35
    .Col = 4: .ColWidth(4) = 1000
    .Col = 5: .ColWidth(5) = 1100
    .Col = 6: .ColWidth(6) = 1400
    .Col = 7: .ColWidth(7) = 35
    .Col = 8: .ColWidth(8) = 1050
    .Col = 9: .ColWidth(9) = 35
End With
End Sub
Private Sub PindahText()
grid3.Move grid.Left + 75, grid.CellTop + 1525

Text7.Left = grid3.Left + 515
Text7.Top = grid3.Top + 50
Combo1.Left = grid3.Left + 1965
Combo1.Top = grid3.Top + 35
Text2.Left = grid3.Left + 5000
Text2.Top = grid3.Top + 50
Text3.Left = grid3.Left + 6000
Text3.Top = grid3.Top + 50
CmdOK.Left = grid3.Left + 8500
CmdOK.Top = grid3.Top + 15
grid3.Left = grid3.Left - 60
With grid3
    .Row = 0: .Col = 0: .Text = grid.Rows - 1: .CellAlignment = 4
End With

End Sub


Private Sub Form_Unload(Cancel As Integer)
Call TblClose_Click
End Sub

Private Sub grid_dblClick()
Call Batal
If grid.Rows = 2 Then
    MsgBox "TIDAK ADA DATA PEMBELIAN YANG DIEDIT", vbCritical, "DATA KOSONG"
    Exit Sub
End If
NoUrutTrans = ""
NoTrans = ""
ByBeli = 0
NoUrutTrans = grid.TextMatrix(grid.Row, 0)
NoTrans = Label3
ByBeli = Text16
BL02.Show 1
grid.Clear
Call Siap
'Pesan = False
Call Isi_Grid
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF5
        SSTab1.Tab = 1
        Text6.SetFocus
End Select
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

Private Sub InfoSPL2_Click()
NoNas = ""
NamaNas = ""
C013.Show 1
Text9 = NoNas
Label24 = NamaNas
Text9.SetFocus
End Sub

Private Sub OptionGD_Click()
If OptionGD.Value = True Then
    OptionGD.FontBold = True
    OptionTK.FontBold = False
End If
End Sub

Private Sub OptionTK_Click()
If OptionTK.Value = True Then
    OptionTK.FontBold = True
    OptionGD.FontBold = False
End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 1 Then Text6.SetFocus
End Sub

Private Sub TblClose_Click()
Dim Tanya
If CCur(Label17) <> 0 Then
    Tanya = MsgBox("YAKIN AKAN KELUAR TRANSAKSI DAN HAPUS DAFTAR PEMBELIAN DI FORM INI ?", vbOKCancel, "TRANSAKSI SUDAH TERDAFTAR")
Else
    Unload Me
End If
If Tanya = vbOK Then
    SDel = "Delete From BL01"
    Set RDel = RDCO.OpenResultset(SDel, rdOpenDynamic, rdConcurRowVer)
    RDel.Close
    Set RDel = Nothing
    Unload Me
End If

End Sub

Private Sub SimpanDaftar()
Dim KDJns, NmJns, JmlBeli, HargaPCS, JmlHarga

If STSSimpan = "1" Then
    If Combo1.Text = "" Or Val(Text2) = 0 Or Val(Text3) = 0 Then
        MsgBox "MASIH ADA DATA YANG KOSONG", vbCritical, "DATA TIDAK BOLEH KOSONG"
        Exit Sub
    End If
    'KDJns = Text7
    'NmJns = Combo1.Text
    'JmlBeli = Text2
    'HargaPCS = Text3
    JmlHarga = grid3.TextMatrix(0, 6)
End If

If STSSimpan = "2" Then
    Text2 = 0
    Text3 = 0
    JmlHarga = 0
End If

SBL = "Select * From B003 where NAMA_JNS = '" + Trim(Combo1) + "'"
Set RBL = RDCO.OpenResultset(SBL, rdOpenDynamic, rdConcurRowVer)
    KODEIND = RBL("Kode_IND")
    
    SBL2 = "Select * From BL01"
    Set RBL2 = RDCO.OpenResultset(SBL2, rdOpenDynamic, rdConcurRowVer)
    RBL2.AddNew
        RBL2("No_Trans") = Trim(Label3)
        RBL2("No_Urut") = NoNo
        RBL2("Kode_Ind") = KODEIND
        RBL2("Kode_JNS") = Trim(Text7)
        RBL2("Nama_JNS") = Trim(Combo1)
        RBL2("JML_Beli") = CCur(Text2)
        RBL2("Harga_PCS") = CCur(Text3)
        RBL2("JML_Harga") = CCur(JmlHarga)
        RBL2("HPBELI_PCS") = 0
        RBL2("HPBELI_TOTAL") = 0
        RBL2("User_Code") = Operator
        RBL2("Tanggal") = Tanggal
    
    NoNo = NoNo + 1
    RBL2.Update
    RBL2.Close
    Set RBL2 = Nothing
    
RBL.Close
Set RBL = Nothing

grid3.Clear
Call Isi_Grid
Call Batal
Text7.SetFocus
End Sub

Private Sub Isi_Grid()
Dim Brs
Brs = 2
Label17 = 0
HPokok = 0
SGrid = "Select * from BL01 order by no_urut asc"
Set RGrid = RDCO.OpenResultset(SGrid, rdOpenDynamic, rdConcurRowVer)
If RGrid.RowCount <> 0 Then
    RGrid.MoveFirst
    Do While Not RGrid.EOF
        With grid
        .Rows = Brs + 1
        .Row = Brs
        .RowHeight(Brs) = 300
        .Col = 0: .Text = RGrid("NO_URUT"): .CellAlignment = 4
        .Col = 1: .Text = RGrid("KODE_JNS"): .CellAlignment = 4
        .Col = 2: .Text = RGrid("NAMA_JNS")
        .Col = 4: .Text = RGrid("JML_BELI"): .CellAlignment = 4
        .Col = 5: .Text = Format(RGrid("HARGA_PCS"), "##,###.00")
        .Col = 6: .Text = Format(RGrid("JML_HARGA"), "##,###.00")
        .Col = 8: .Text = Format(RGrid("HPBELI_PCS"), "##,###.00")
        .Col = 9: .Text = Format(RGrid("HPBELI_TOTAL"), "##,###.00")
        Label17 = Format(CCur(Label17) + CCur(RGrid("JML_HARGA")), "##,###.00")
        HPokok = Format(CCur(HPokok) + CCur(RGrid("HPBELI_TOTAL")), "##,###.00")
        End With
        Brs = Brs + 1
    RGrid.MoveNext
    Loop
'    Call PindahText
'Else
'    Unload Me
'    BL01.Show
End If
'RGrid.Close
'Set RGrid = Nothing

With grid2
    .Row = 0
    .Col = 5: .Text = Label17
    .Col = 8: .Text = Format(HPokok, "##,###.00")
End With
Label17 = Format(CCur(Label17) - CCur(Text8), "##,###.00")

With grid3
    .Row = 0: .Col = 0: .Text = grid.Rows - 1: .CellAlignment = 4
End With
End Sub

Private Sub TblSave_Click()
Dim Nomor As Double

If Text1 = "" Or Text4 = "" Then
    MsgBox "FAKTUR PEMBELIAN / TANGGAL PEMBELIAN MASIH KOSONG", vbCritical, "DATA TIDAK BOLEH KOSONG"
    Text1.SetFocus
    Exit Sub
End If

If Val(Label17) < 1 Then
    MsgBox "DAFTAR BAHAN YANG DIBELI MASIH KOSONG", vbCritical, "DATA TIDAK BOLEH KOSONG"
    SSTab1.Tab = 0
    Text7.SetFocus
    Exit Sub
End If

If CCur(Label7) <= 0 Then
    Text5.SetFocus
    MsgBox "NOMINAL PEMBAYARAN MASIH KOSONG", vbCritical, "DATA TIDAK BOLEH KOSONG"
    Exit Sub
End If

If CCur(Text15) > 0 And (Text9 = "" Or Text11 = "" Or Text12 = "" Or Text13 = "") Then
    MsgBox "DATA DEBITUR MASIH KOSONG KOSONG", vbCritical, "DATA TIDAK BOLEH KOSONG"
    Text15.SetFocus
    Exit Sub
End If

If CCur(Text14) > 0 And Text10 = "" Then
    MsgBox "DATA GENERAL LEDGER PEMBAYARAN NON TUNAI MASIH KOSONG", vbCritical, "DATA TIDAK BOLEH KOSONG"
    Text10.SetFocus
    Exit Sub
End If

If CCur(Label7) <> CCur(Label17) Then
    MsgBox "NOMINAL PEMBAYARAN HARUS SAMA DENGAN TOTAL PEMBELIAN", vbCritical, "NOMINAL PEMBAYARAN SALAH"
    Text5.SetFocus
    Exit Sub
End If

'Jika Pembayaran Hutang
If CCur(Text15) > 0 Then
    KodeSPL = Text9
    KodeHutang = Label40
    NoHutang = Label5
    TglMulai = Text12
    TglJatuh = Text13
    SyaratByr = Text11
Else
    KodeSPL = "00000"
    KodeHutang = "-"
    NoHutang = "-"
    TglMulai = Tanggal
    TglJatuh = Tanggal
    SyaratByr = "-"
End If

'Jika Pembayaran Non Tunai
If CCur(Text14) > 0 Then
    SGLNonKas = Text10
Else
    SGLNonKas = "-"
End If

If OptionGD.Value = False And OptionTK.Value = False Then
    MsgBox "TUJUAN BARANG BELUM DITENTUKAN", vbCritical, "TUJUAN BARANG TIDAK BOLEH KOSONG"
    Exit Sub
End If

Tanya = MsgBox("YAKIN PROSES TRANSAKSI PEMBELIAN ?", vbOKCancel, "PROSES TRANSAKSI PEMBELIAN ?")
If Tanya = vbCancel Then Exit Sub

NoTrans = Label3
Call Kosong
Call Isi_Grid
SSTab1.Tab = 0
Text1.SetFocus

'Tampil No. Transaksi
SNo = "Select NoBeli from C013 where usercode = '" + Operator + "'"
Set RNo = RDCO.OpenResultset(SNo, rdOpenDynamic, rdConcurRowVer)
If RNo.RowCount <> 0 Then
    Nomor = RNo("NoBeli")
    Label3 = Trim("1.") + Digit(2, Trim(NoUser)) + "." + Digit(7, Nomor)
End If
RNo.Close
Set RNo = Nothing

'MsgBox "SIAPKAN VALIDASI KE PRINTER"
'crpt.ReportFileName = "c:\windows\FCReport\transbeli.rpt"
'crpt.SelectionFormula = "{B005.No_Bukti} = '" + Trim(NoTrans) + "'"
'crpt.WindowState = crptMaximized
'crpt.WindowMaxButton = False
'crpt.WindowMinButton = False
'crpt.Action = 1

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF5
        SSTab1.Tab = 1
        Text6.SetFocus
End Select
End Sub

Private Sub text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text1_LostFocus()
Text1 = Format(Text1, ">")
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

Private Sub Text11_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TEXT12_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text12_LostFocus()
If Text12 = "" Then Exit Sub
If Not IsDate(Text12) Then
    Text12.SetFocus
    MsgBox "TANGGAL MULAI KREDIT HARUS TYPE TANGGAL", vbCritical, "TYPE DATA SALAH"
End If
Text12 = Format(Text12, "DD/MM/YYYY")
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text13_LostFocus()
If Text13 = "" Then Exit Sub
If Not IsDate(Text13) Then
    Text13.SetFocus
    MsgBox "TANGGAL JATUH TEMPO HARUS TYPE TANGGAL", vbCritical, "TYPE DATA SALAH"
End If
Text13 = Format(Text13, "DD/MM/YYYY")
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

Label7 = Format(CCur(Text5) + CCur(Text14) + CCur(Text15), "##,###.00")
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

Private Sub Text15_GotFocus()
If Text15 = 0 Then Text15 = ""
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text15_LostFocus()
Dim UrutBeli As Double
If Text15 = "" Then Text15 = 0
If Not IsNumeric(Text15) Then
    Text15.SetFocus
    MsgBox "NOMINAL PEMBAYARAN HARUS ANGKA", vbCritical, "TYPE DATA SALAH"
    Exit Sub
End If
Text15 = Format(Text15, "##,###.00")
Label7 = Format(CCur(Text5) + CCur(Text14) + CCur(Text15), "##,###.00")
If CCur(Label7) > CCur(Label17) Then
    Text15.SetFocus
    MsgBox "TOTAL PEMBAYARAN TIDAK BOLEH MELEBIHI TOTAL PEMBELIAN", vbCritical, "TOTAL PEMBAYARAN SALAH"
    Exit Sub
End If

If CCur(Text15) > 0 Then
    Call HutangAktif
    Text9.SetFocus
    SHutang = "Select Top 1 Kode_Hutang, Nama_Hutang from H001 order by Kode_Hutang"
    Set RHutang = RDCO.OpenResultset(SHutang, rdOpenDynamic, rdConcurRowVer)
    If RHutang.RowCount <> 0 Then
        Label40 = RHutang("Kode_Hutang")
        Label41 = RHutang("Nama_Hutang")
    End If
    RHutang.Close
    Set RHutang = Nothing
    UrutBeli = Val(Right(Label3, 6))
    UrutBeli = Trim(Mid(Label3, 1, 1)) + Trim(NoUser) + Trim(UrutBeli)
    Label5 = Trim(Label40) + "." + Digit(7, UrutBeli)
Else
    Call HutangNonAktif
End If
End Sub


Private Sub Text16_GotFocus()
If Text16 = 0 Then Text16 = ""
End Sub

Private Sub Text16_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF5
        SSTab1.Tab = 1
        Text6.SetFocus
End Select
End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text16_LostFocus()
If Text16 = "" Then Text16 = 0
If Not IsNumeric(Text16) Then
    Text16.SetFocus
    MsgBox "TOTAL BIAYA PEMBELIAN HARUS ANGKA", vbCritical, "TYPE DATA SALAH"
    Exit Sub
End If
Text16 = Format(Text16, "##,###.00")
If grid.Rows = 2 Then
    Exit Sub
End If
STSSimpan = "2"
Call SimpanDaftar
End Sub

Private Sub Text17_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text17 = Format(Text17, ">")
    SendKeys vbTab
End If
End Sub

Private Sub Text17_LostFocus()
If Text17 = "" Then Exit Sub
End Sub

Private Sub Text2_GotFocus()
If Val(Text2) = 0 Then Text2 = ""
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF5
        SSTab1.Tab = 1
        Text6.SetFocus
End Select
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text2_LostFocus()
If Text2 = "" Then Text2 = 0
If Not IsNumeric(Text2) Then
    Text2.SetFocus
    MsgBox "DATA PEMBELIAN HARUS ANGKA", vbCritical, "TIPE DATA SALAH"
    Exit Sub
End If
If Text2 = "" Or Text3 = "" Then Exit Sub
With grid3
    .Col = 6: .Text = Format(CCur(Text2) * CCur(Text3), "##,###.00")
End With
End Sub

Private Sub Text3_GotFocus()
If Val(Text3) = 0 Then Text3 = ""
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF5
        SSTab1.Tab = 1
        Text6.SetFocus
End Select
End Sub

Private Sub text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text3_LostFocus()
If Text3 = "" Then Text3 = "0"
If Not IsNumeric(Text3) Then
    Text3.SetFocus
    MsgBox "HARGA PEMBELIAN HARUS ANGKA", vbCritical, "TIPE DATA SALAH"
    Exit Sub
End If
Text3 = Format(Text3, "##,###.00")
If Text2 = "" Then Exit Sub
With grid3
    .Col = 6: .Text = Format(CCur(Text2) * CCur(Text3), "##,###.00")
End With
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF5
        SSTab1.Tab = 1
        Text6.SetFocus
End Select
End Sub

Private Sub text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text4_LostFocus()
If Text4 = "" Then Exit Sub
If Not IsDate(Text4) Then
    Text4.SetFocus
    MsgBox "TYPE DATA HARUS TANGGAL", vbCritical, "TYPE DATA SALAH"
    Exit Sub
End If
Text4 = Format(Text4, "DD/MM/YYYY")
End Sub

Private Sub Text5_GotFocus()
If CCur(Text5) = 0 Then Text5 = ""
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text5_LostFocus()
If Text5 = "" Then Text5 = 0
If Not IsNumeric(Text5) Then
    Text5.SetFocus
    MsgBox "NOMINAL PEMBAYARAN HARUS ANGKA", vbCritical, "TYPE DATA SALAH"
    Exit Sub
End If
Text5 = Format(Text5, "##,###.00")
Label7 = Format(CCur(Text5) + CCur(Text14) + CCur(Text15), "##,###.00")
If CCur(Label7) > CCur(Label17) Then
    Text5.SetFocus
    MsgBox "TOTAL PEMBAYARAN TIDAK BOLEH MELEBIHI TOTAL PEMBELIAN", vbCritical, "TOTAL PEMBAYARAN SALAH"
End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text6 = Format(Text6, ">")
    SendKeys vbTab
End If
End Sub

Private Sub Text6_LostFocus()
If Text6 = "" Then Exit Sub

'Text6 = Digit(5, Text6)
'SSPL = "Select Nama From C012 where NoNas = '" + Text6 + "'"
'Set RSPL = RDCO.OpenResultset(SSPL, rdOpenDynamic, rdConcurRowVer)
'If RSPL.RowCount <> 0 Then
'    Label22 = RSPL("Nama")
'Else
'    Text6.SetFocus
'    MsgBox "NOMOR SUPPLIER BELUM TERDAFTAR, KLIK TOMBOL BARU UNTUK ENTRI SUPPLIER", vbInformation, "NOMOR SUPPLIER BELUM TERDAFTAR"
'End If
'RSPL.Close
'Set RSPL = Nothing
End Sub

Private Sub Text7_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF5
        SSTab1.Tab = 1
        Text6.SetFocus
End Select
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text7_LostFocus()
If Text7 = "" Then Exit Sub
SJenis = "Select * from B003 where KODE_JNS = '" + Trim(Text7) + "'"
Set RJenis = RDCO.OpenResultset(SJenis, rdOpenDynamic, rdConcurRowVer)
If RJenis.RowCount <> 0 Then
    Combo1.Text = RJenis("NAMA_JNS")
'    MTLABA = RJenis("MT_LABA")
'    Persen = RJenis("PERSEN")
    Text2 = 0
    Text3 = 0
    Text5 = 0
    Text6 = 0
    Text2.SetFocus
Else
    Text7.SetFocus
    MsgBox "KODE JENIS BAHAN BELUM TERDAFTAR. DAFTARKAN DAHULU LEWAT MENU SYSTEM", vbInformation, "KODE JENIS BAHAN BELUM TERDAFTAR"
End If
RJenis.Close
Set RJenis = Nothing
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text8_LostFocus()
If Text8 = "" Then Text8 = 0
If Not IsNumeric(Text8) Then
    Text8.SetFocus
    MsgBox "NOMINAL DISCOUNT HARUS ANGKA", vbCritical, "TYPE DATA SALAH"
End If
Text8 = Format(Text8, "##,###.00")
Label7 = Format(CCur(Label17) - CCur(Text8), "##,###.00")
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text9_LostFocus()
If Text9 = "" Then Exit Sub
Text9 = Digit(5, Text9)
SSPL = "Select Nama From C012 where NoNas = '" + Text9 + "'"
Set RSPL = RDCO.OpenResultset(SSPL, rdOpenDynamic, rdConcurRowVer)
If RSPL.RowCount <> 0 Then
    Label24 = RSPL("Nama")
Else
    Text9.SetFocus
    MsgBox "KODE SUPPLIER BELUM TERDAFTAR", vbInformation, "INFO DATA SUPPLIER"
End If
RSPL.Close
Set RSPL = Nothing
End Sub


