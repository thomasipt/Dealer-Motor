VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form S003 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TRANSAKSI SERVICE KENDARAAN"
   ClientHeight    =   9255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14130
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9255
   ScaleMode       =   0  'User
   ScaleWidth      =   14130
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cari"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3105
      TabIndex        =   53
      Top             =   8775
      Width           =   735
   End
   Begin VB.TextBox Text19 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Height          =   360
      Left            =   1110
      TabIndex        =   51
      Text            =   "Text19"
      Top             =   8775
      Width           =   1950
   End
   Begin VB.TextBox Text18 
      BackColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   135
      MaxLength       =   7
      TabIndex        =   48
      Text            =   "Text18"
      Top             =   11250
      Width           =   1005
   End
   Begin VB.TextBox Text17 
      BackColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   135
      MaxLength       =   7
      TabIndex        =   47
      Text            =   "Text17"
      Top             =   10845
      Width           =   1005
   End
   Begin VB.CommandButton Command7 
      Caption         =   "KELUAR"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   11655
      TabIndex        =   16
      Top             =   5092
      Width           =   2295
   End
   Begin VB.CommandButton TmbSave 
      Caption         =   "SIMPAN"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   6795
      TabIndex        =   15
      Top             =   5092
      Width           =   2295
   End
   Begin VB.TextBox Text13 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   10290
      TabIndex        =   43
      Text            =   "Text13"
      Top             =   3840
      Width           =   3735
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   7890
      TabIndex        =   13
      Text            =   "Text11"
      Top             =   3840
      Width           =   1575
   End
   Begin VB.TextBox Text200 
      Height          =   540
      Left            =   7140
      TabIndex        =   40
      Text            =   "Text200"
      Top             =   12240
      Width           =   1485
   End
   Begin VB.TextBox Text100 
      Height          =   435
      Left            =   7110
      TabIndex        =   39
      Text            =   "Text100"
      Top             =   11655
      Width           =   1485
   End
   Begin VB.TextBox Text15 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   10290
      TabIndex        =   14
      Text            =   "Text15"
      Top             =   4387
      Width           =   3735
   End
   Begin VB.TextBox Text16 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6765
      Locked          =   -1  'True
      TabIndex        =   37
      Text            =   "Text16"
      Top             =   5820
      Width           =   1425
   End
   Begin VB.TextBox Text12 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   990
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   35
      Text            =   "1,000,000.00"
      Top             =   5790
      Width           =   7305
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tabel Sparepart Service"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   3720
      Left            =   6705
      TabIndex        =   31
      Top             =   90
      Width           =   7305
      Begin VB.TextBox Text14 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Height          =   390
         Left            =   105
         TabIndex        =   11
         Text            =   "Text14"
         Top             =   960
         Width           =   1035
      End
      Begin VB.CommandButton Command1 
         Caption         =   "TAMBAH"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   6165
         TabIndex        =   12
         Top             =   960
         Width           =   1035
      End
      Begin VB.TextBox Text10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   360
         Left            =   5250
         TabIndex        =   10
         Text            =   "Text10"
         Top             =   285
         Width           =   1950
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   105
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   308
         Width           =   5100
      End
      Begin MSFlexGridLib.MSFlexGrid grid 
         Height          =   1770
         Left            =   105
         TabIndex        =   33
         Top             =   1365
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   3122
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         BackColor       =   16777215
         BackColorFixed  =   65280
         BackColorBkg    =   12648384
         GridColor       =   0
         TextStyle       =   3
         TextStyleFixed  =   3
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid grid3 
         Height          =   390
         Left            =   105
         TabIndex        =   34
         Top             =   3255
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   688
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         FixedCols       =   0
         BackColor       =   16777215
         ForeColor       =   0
         BackColorFixed  =   65280
         ForeColorFixed  =   0
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         GridColor       =   0
         TextStyle       =   3
         TextStyleFixed  =   3
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid gridgrid 
         Height          =   1770
         Left            =   105
         TabIndex        =   41
         Top             =   1365
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   3122
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         BackColor       =   16777215
         BackColorFixed  =   65280
         BackColorBkg    =   12648384
         GridColor       =   0
         TextStyle       =   3
         TextStyleFixed  =   3
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Caption         =   "Label14"
         Height          =   285
         Left            =   105
         TabIndex        =   32
         Top             =   645
         Width           =   7095
      End
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   1380
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   510
      Width           =   1005
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   1380
      MaxLength       =   15
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   105
      Width           =   1950
   End
   Begin VB.Frame Frame1 
      Caption         =   "Info Kendaraan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3930
      Left            =   105
      TabIndex        =   17
      Top             =   990
      Width           =   6540
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         Height          =   360
         Left            =   2310
         TabIndex        =   2
         Text            =   "Text3"
         Top             =   315
         Width           =   4050
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00C0FFFF&
         Height          =   1200
         Left            =   2310
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   3
         Text            =   "S003.frx":0000
         Top             =   705
         Width           =   4050
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00C0FFFF&
         Height          =   360
         Left            =   2310
         TabIndex        =   4
         Text            =   "Text5"
         Top             =   1965
         Width           =   1950
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Height          =   360
         Left            =   2310
         TabIndex        =   5
         Text            =   "Text6"
         Top             =   2340
         Width           =   1950
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Height          =   360
         Left            =   2310
         TabIndex        =   6
         Text            =   "Text7"
         Top             =   2715
         Width           =   1950
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Height          =   360
         Left            =   2310
         TabIndex        =   7
         Text            =   "Text8"
         Top             =   3090
         Width           =   1950
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H00C0FFFF&
         Height          =   360
         Left            =   2310
         TabIndex        =   8
         Text            =   "Text9"
         Top             =   3465
         Width           =   1950
      End
      Begin VB.Label Label4 
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
         Height          =   225
         Left            =   210
         TabIndex        =   30
         Top             =   390
         Width           =   1980
      End
      Begin VB.Label Label5 
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
         Height          =   225
         Left            =   210
         TabIndex        =   29
         Top             =   810
         Width           =   1980
      End
      Begin VB.Label Label6 
         Caption         =   "TELEPON"
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
         Left            =   210
         TabIndex        =   28
         Top             =   2040
         Width           =   1980
      End
      Begin VB.Label Label7 
         Caption         =   "NO. POLISI"
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
         Left            =   210
         TabIndex        =   27
         Top             =   2415
         Width           =   1980
      End
      Begin VB.Label Label11 
         Caption         =   "NO. RANGKA"
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
         Left            =   210
         TabIndex        =   26
         Top             =   2790
         Width           =   1980
      End
      Begin VB.Label Label12 
         Caption         =   "NO. MESIN"
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
         Left            =   210
         TabIndex        =   25
         Top             =   3165
         Width           =   1980
      End
      Begin VB.Label Label13 
         Caption         =   "TIPE / WARNA / TAHUN"
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
         Left            =   210
         TabIndex        =   24
         Top             =   3540
         Width           =   1980
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GridInfo 
      Height          =   1830
      Left            =   90
      TabIndex        =   44
      ToolTipText     =   "Klik untuk edit"
      Top             =   6840
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   3228
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
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label20 
      BackColor       =   &H0080C0FF&
      Caption         =   "NO. POLISI"
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
      Left            =   135
      TabIndex        =   52
      Top             =   8850
      Width           =   945
   End
   Begin VB.Label Label18 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label18"
      Height          =   360
      Left            =   1260
      TabIndex        =   50
      Top             =   11250
      Width           =   3885
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label17"
      Height          =   360
      Left            =   1260
      TabIndex        =   49
      Top             =   10845
      Width           =   3885
   End
   Begin VB.Label Label15 
      BackColor       =   &H0080C0FF&
      Caption         =   "STATUS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   105
      TabIndex        =   46
      Top             =   4995
      Visible         =   0   'False
      Width           =   6540
   End
   Begin VB.Label Label200 
      BackColor       =   &H0080C0FF&
      Caption         =   "HISTORY SERVICE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   105
      TabIndex        =   45
      Top             =   6570
      Width           =   3030
   End
   Begin VB.Label Label100 
      BackColor       =   &H0080C0FF&
      Caption         =   "DISKON                         %"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6720
      TabIndex        =   42
      Top             =   3930
      Width           =   4815
   End
   Begin VB.Label Label19 
      BackColor       =   &H0080C0FF&
      Caption         =   "PEMBAYARAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6720
      TabIndex        =   38
      Top             =   4470
      Width           =   3030
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080C0FF&
      Caption         =   "Label16"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   990
      Left            =   105
      TabIndex        =   36
      Top             =   5790
      Width           =   6435
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5190
      TabIndex        =   23
      Top             =   540
      Width           =   1320
   End
   Begin VB.Label Label9 
      Caption         =   "TGL TRANSAKSI"
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
      Left            =   3735
      TabIndex        =   22
      Top             =   548
      Width           =   1320
   End
   Begin VB.Label Label8 
      Caption         =   "JAM "
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
      Left            =   195
      TabIndex        =   21
      Top             =   548
      Width           =   1140
   End
   Begin VB.Label Label1 
      Caption         =   "MEKANIK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   195
      TabIndex        =   20
      Top             =   165
      Width           =   1140
   End
   Begin VB.Label Label2 
      Caption         =   "NO. TRANSAKSI"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3750
      TabIndex        =   19
      Top             =   180
      Width           =   1365
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5190
      TabIndex        =   18
      Top             =   135
      Width           =   1320
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      Height          =   855
      Left            =   105
      Top             =   75
      Width           =   6540
   End
End
Attribute VB_Name = "S003"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private RDel, RDel2, RDel3, RCari, RCari2, RPLY, RSPR As rdoResultset
Private SDel, SDel2, SDel3, SCari, SCari2, SPLY, SSPR As String

Private RSave, RSave2, RSave3, RSave4, RSave5, RSave6, RSave7 As rdoResultset
Private SSave, SSave2, SSave3, SSave4, SSave5, SSave6, SSave7 As String

Private RJual1, RJual2, RJual3, RJual4 As rdoResultset
Private SJual1, SJual2, SJual3, SJual4 As String

Private RKAS, RKAS2, RKAS3 As rdoResultset
Private SKAS, SKAS2, SKAS3 As String

Private RKREDIT, RKREDIT2, RKREDIT3 As rdoResultset
Private SKREDIT, SKREDIT2, SKREDIT3 As String

Private RLABA, RLABA2 As rdoResultset
Private SLABA, SLABA2 As String

Private CEKKODE, SGLPART, SGLJASA, SGLSEDIA

Private KKODES, NNAMAS, BBIAYAS
Private NoUrutTrans As String

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Combo1 = "" Then
        Text15.SetFocus
    Else
        Text10.SetFocus
    End If
End If
End Sub

Private Sub Combo1_LostFocus()
If Combo1 = "" Then Exit Sub

If Left(Combo1, 3) = "IPT" Then
    SCari = "Select * From S002 where KODE_S='" + Trim(Combo1) + "'"
    Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
        If RCari.RowCount <> 0 Then
            Label14 = RCari("NAMA_S")
            Text10 = Format(RCari("BIAYA_S"), "##,###.00")
        Else
            MsgBox "KODE BELUM TERDAFTAR", vbSystemModal, "KONFIRMASI"
            Combo1.SetFocus
        End If
    RCari.Close
    Set RCari = Nothing
Else
    SCari = "Select * From B003A where KODE_JNS='" + Trim(Combo1) + "'"
    Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
        If RCari.RowCount <> 0 Then
            Label14 = RCari("NAMA_JNS")
            Text10 = Format(RCari("HJUAL"), "##,###.00")
        Else
            MsgBox "KODE BELUM TERDAFTAR", vbSystemModal, "KONFIRMASI"
            Combo1.SetFocus
        End If
    RCari.Close
    Set RCari = Nothing
End If
Text15 = 0
End Sub

Private Sub Command1_Click()
Dim Tanya
Tanya = MsgBox("MASUKKAN DATA", vbOKCancel, "KONFIRMASI")
Grid.ZOrder

If Tanya = vbCancel Then
    Combo1.SetFocus
    Exit Sub
End If

Do While Text14 > 0
    Text14 = Text14 - 1
    Call Simpan_QTY
Loop
    
    Call IsiGrid
    Call IsiGridA
    'Combo1 = ""
    Label14 = ""
    Text10 = ""
    Text14 = 1
    Combo1.SetFocus
    
If Text13 > 0 Then
    Text16 = grid3.TextMatrix(grid3.Row, 2) - CCur(Text13)
Else
    Text16 = grid3.TextMatrix(grid3.Row, 2)
End If

Text12 = grid3.TextMatrix(grid3.Row, 2)

End Sub

Private Sub Simpan_QTY()
SSave = "Select * From S003A"
Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
RSave.AddNew
    RSave("KODE_S") = Combo1
        If Left(Combo1, 3) = "IPT" Then
            Text100 = CCur(Text100) + 1
        Else
            Text200 = CCur(Text200) + 1
        End If
    RSave("NAMA_S") = Label14
    RSave("BIAYA_S") = CCur(Text10)
    
    If Left(Combo1, 3) = "IPT" Then
        RSave("DISKON_BELI") = 0
        RSave("LABA") = 0
    Else
        SCari = "Select * From B004 where KODE_JNS='" + Trim(Combo1) + "' Order  by NO_URUT Desc"
        Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
            If RCari.RowCount <> 0 Then
                DISKONBELI = RCari("DISKON_BELI")
                LABALABA = CCur(Text10) * RCari("DISKON_BELI") / 100
                
                RSave("DISKON_BELI") = CCur(DISKONBELI)
                RSave("LABA") = CCur(LABALABA)
                RSave("STS_NOURUT") = RCari("NO_URUT")
            Else
                SCari2 = "Select * From B004A where KODE_JNS='" + Trim(Combo1) + "'"
                Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenDynamic, rdConcurRowVer)
                    If RCari2.RowCount <> 0 Then
                        DISKONBELI = RCari2("DISKON_BELI")
                        LABALABA = CCur(Text10) * RCari2("DISKON_BELI") / 100
                        
                        RSave("DISKON_BELI") = CCur(DISKONBELI)
                        RSave("LABA") = CCur(LABALABA)
                        RSave("STS_NOURUT") = "B004A"
                    Else
                        RSave("DISKON_BELI") = 2227 / 100
                        RSave("LABA") = CCur(Text10) * 27.27 / 100
                        RSave("STS_NOURUT") = "REGULER"
                    End If
                RCari2.Close
                Set RCari2 = Nothing
            End If
        RCari.Close
        Set RCari = Nothing
    End If
    
RSave.Update
RSave.Close
Set RSave = Nothing
End Sub

Private Sub IsiGridGrid()
gridgrid.ZOrder
End Sub

Private Sub IsiGrid()
Dim Brs As Integer
Brs = 1
SCari = "Select * from S003A order by Kode_S"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset, rdConcurReadOnly)
If RCari.RowCount <> 0 Then
RCari.MoveFirst
Do While Not RCari.EOF
    With Grid
        .Rows = Brs + 1
        .Row = Brs
        .Col = 0: .Text = RCari("Kode_S")
        .Col = 1: .Text = RCari("Nama_S")
        .Col = 2: .Text = Format(RCari("Biaya_S"), "##,###.00")
    End With
    Brs = Brs + 1
RCari.MoveNext
Loop
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub IsiGridA()
Dim Brs As Integer
SCari = "Select * from S003AA"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset, rdConcurReadOnly)
If RCari.RowCount <> 0 Then
    With grid3
        .Rows = 1
        .Row = 0
        .Col = 0: .Text = RCari("CountOfKode_S")
        .Col = 2: .Text = Format(RCari("SumOfBiaya_S"), "##,###.00")
    End With
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub IsiGridGA()
Dim Brs As Integer
SCari = "Select * from S003BB"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset, rdConcurReadOnly)
If RCari.RowCount <> 0 Then
    With grid4
        .Rows = 1
        .Row = 0
        .Col = 0: .Text = RCari("CountOfKode_JNS")
        .Col = 2: .Text = Format(RCari("SumOfSaldo"), "##,###.00")
    End With
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub IsiGrid2()
Dim Brs As Integer
Brs = 1
SCari = "Select * from S003B order by Kode_JNS"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset, rdConcurReadOnly)
If RCari.RowCount <> 0 Then
RCari.MoveFirst
Do While Not RCari.EOF
    With Grid
        .Rows = Brs + 1
        .Row = Brs
        .Col = 0: .Text = RCari("Kode_JNS"): .CellAlignment = 4
        .Col = 1: .Text = RCari("Nama_JNS")
        .Col = 2: .Text = Format(RCari("SALDO"), "##,###.00")
    End With
    Brs = Brs + 1
RCari.MoveNext
Loop
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub Command2_Click()
OYEN = "Select * from S003 Where NO_POL like '%" + Trim(Text19) + "'"
GridInfo.Clear
GridInfo.Refresh
Call IsiGridInfo
End Sub

Private Sub Command7_Click()
STS_KSG = 0

SCari = "Select * from B004"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    Do While Not RCari.EOF
        RCari.EDIT
            RCari("STS_S003A") = 0
        RCari.Update
        RCari.MoveNext
    Loop
End If
RCari.Close
Set RCari = Nothing

SDel = "Delete * From S003A"
Set RDel = RDCO.OpenResultset(SDel, rdOpenDynamic, rdConcurRowVer)
RDel.Close
Set RDel = Nothing


Unload Me
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=DEALER", rdDriverNoPrompt, False, CN)

ClearTextBoxes Me

Call NoBukti
Call Combo
Call SiapkanGrid

Call Del

'Combo1 = ""
'Combo2 = ""
Label16 = "TOTAL BAYAR"
Label10 = Tanggal
Label14 = ""
Label15 = ""

Text2 = Time

Text13 = 0
Text14 = 1
Text15 = 0
Text11 = 0

Text100 = 0
Text200 = 0
NoUrutTrans = 0

Grid.ZOrder

'OYEN = "Select Top 50 NAMA, ALAMAT, TELEPON, NO_POL, NO_RANGKA, NO_MESIN, TIPE from S003"
'Call IsiGridInfo
Call SiapkanGridInfo

'Frame3.Visible = True
'Frame3.ZOrder

If STS_KSG = 1 Then
    Label15 = "TRANSAKSI KSG"
    Label15.ForeColor = &HFF&
Else
    Label15 = "TRANSAKSI NON KSG"
    Label15.ForeColor = &H0&
End If

Text17 = "9006206"
Label17 = "BIAYA BENGKEL"

Text18 = "1001113"
Label18 = "KAS BENGKEL"
End Sub

Private Sub Del()
SCari = "Select * from B004"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    Do While Not RCari.EOF
        RCari.EDIT
            RCari("STS_S003A") = 0
        RCari.Update
        RCari.MoveNext
    Loop
End If
RCari.Close
Set RCari = Nothing

SDel = "Delete * From S003A"
Set RDel = RDCO.OpenResultset(SDel, rdOpenDynamic, rdConcurRowVer)
RDel.Close
Set RDel = Nothing

End Sub

Private Sub SiapkanGrid()
With Grid
    .Cols = 3
    .Row = 0
    .Col = 0: .ColWidth(0) = "1500": .Text = "KODE": .CellAlignment = 4
    .Col = 1: .ColWidth(1) = "3000": .Text = "NAMA": .CellAlignment = 4
    .Col = 2: .ColWidth(2) = "2000": .Text = "BIAYA": .CellAlignment = 4
End With

With gridgrid
    .Cols = 3
    .Row = 0
    .Col = 0: .ColWidth(0) = "1500": .Text = "KODE": .CellAlignment = 4
    .Col = 1: .ColWidth(1) = "3000": .Text = "NAMA": .CellAlignment = 4
    .Col = 2: .ColWidth(2) = "2000": .Text = "BIAYA": .CellAlignment = 4
End With

With grid3
    .Cols = 3
    .Row = 0
    .Col = 0: .ColWidth(0) = "1500": .CellAlignment = 4
    .Col = 1: .ColWidth(1) = "3000": .Text = "TOTAL": .CellAlignment = 4
    .Col = 2: .ColWidth(2) = "2000":
End With
End Sub

Private Sub NoBukti()
Dim No As Double
SqlNo = "Select * from C013 where nama = '" + Operator + "'"
Set RSLNO = RDCO.OpenResultset(SqlNo, rdOpenDynamic, rdConcurRowVer)

No = Val(RSLNO("NoPelayanan")) + 1
Label3 = Trim(Digit(5, No))
RSLNO.Close
Set RSLNO = Nothing
End Sub

Private Sub NoBukti2()
SSave7 = "Select * From C013 where nama = '" + Operator + "'"
Set RSave7 = RDCO.OpenResultset(SSave7, rdOpenKeyset, rdConcurRowVer)
    No = RSave7("NoPelayanan")
    RSave7.EDIT
        RSave7("NoPelayanan") = No + 1
RSave7.Update
RSave7.Close
Set RSave7 = Nothing
End Sub

Private Sub Combo()
SPLY = "Select * From S002 order by KODE_S"
Set RPLY = RDCO.OpenResultset(SPLY, rdOpenDynamic, rdOpenKeyset)
RPLY.MoveFirst
Do While Not RPLY.EOF
    Combo1.AddItem RPLY("KODE_S")
RPLY.MoveNext
Loop
RPLY.Close
Set RPLY = Nothing
Combo1.ListIndex = 0

NNAMAS = ""

SSPR = "Select * From B003A order by KODE_JNS"
Set RSPR = RDCO.OpenResultset(SSPR, rdOpenDynamic, rdOpenKeyset)
RSPR.MoveFirst
Do While Not RSPR.EOF
    NNAMAS = Left(RSPR("KODE_JNS"), 3)
    If NNAMAS <> "M00" Then
        Combo1.AddItem RSPR("KODE_JNS")
    End If
RSPR.MoveNext
Loop
RSPR.Close
Set RSPR = Nothing
End Sub

Private Sub grid_dblClick()
Dim Tanya, KODE

KODE = ""
KODE = Grid.TextMatrix(Grid.Row, 0)

CEKKODE = Left(KODE, 3)
If CEKKODE = "IPT" Then
    Text100 = CCur(Text100) - 1
Else
    Text200 = CCur(Text200) - 1
End If


Tanya = MsgBox("YAKIN AKAN HAPUS DATA BAHAN " + Trim(KODE), vbOKCancel, "WARNING")
If Tanya = vbCancel Then Exit Sub
    SDel = "Delete From S003A where Kode_S = '" + Trim(KODE) + "'"
    Set RDel = RDCO.OpenResultset(SDel, rdOpenDynamic, rdConcurRowVer)
    
    If Text100 = 0 And Text200 = 0 Then
        Call IsiGridGrid
        Call IsiGridA
    Else
        Call IsiGrid
        Call IsiGridA
    End If
    
Grid.Refresh
grid3.Refresh
Text12 = grid3.TextMatrix(grid3.Row, 2)

If Text13 > 0 Then
    Text16 = grid3.TextMatrix(grid3.Row, 2) - CCur(Text13)
Else
    Text16 = grid3.TextMatrix(grid3.Row, 2)
End If

End Sub

Private Sub GridInfo_dblClick()
Text3 = Trim(Format(GridInfo.TextMatrix(GridInfo.Row, 0), ">"))
Text4 = Trim(Format(GridInfo.TextMatrix(GridInfo.Row, 1), ">"))
Text5 = Trim(Format(GridInfo.TextMatrix(GridInfo.Row, 2), ">"))
Text6 = Trim(Format(GridInfo.TextMatrix(GridInfo.Row, 3), ">"))
Text7 = Trim(Format(GridInfo.TextMatrix(GridInfo.Row, 4), ">"))
Text8 = Trim(Format(GridInfo.TextMatrix(GridInfo.Row, 5), ">"))
Text9 = Trim(Format(GridInfo.TextMatrix(GridInfo.Row, 6), ">"))

Combo1.SetFocus

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Text1 = "" Then Text1 = "-"
    Text3.SetFocus
End If
End Sub

Private Sub Text1_LostFocus()
Text1 = Format(Text1, ">")
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
    If KeyAscii = 13 Then Text14.SetFocus
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
    If KeyAscii = 13 Then Command1.SetFocus
End Sub

Private Sub Text14_LostFocus()
If Text14 = "" Then
    Text14.SetFocus
    Exit Sub
Else
    If Text14 < 1 Then
        Text14.SetFocus
        Exit Sub
    End If
End If
End Sub

Private Sub Text10_LostFocus()
Text10 = Format(Text10, "##,###.00")
'Call HARGA
End Sub

Private Sub HARGA()
SCari = "Select * from B003A where KODE_JNS = '" + Trim(Combo1) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    If CCur(RCari("HJual")) <> CCur(Text10) Then
        RCari.EDIT
            RCari("HJual") = CCur(Text10)
            RCari("HBeli") = CCur(Text10)
        RCari.Update
    End If
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub Text11_GotFocus()
If CCur(Text11) = 0 Then Text11 = ""
End Sub

Private Sub Text15_GotFocus()
Text15 = ""
Text12 = Format(CCur(Text16) - CCur(Text13), "##,###.00")
Label16 = "TOTAL BAYAR"
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

If Text11 = "" Then Text11.SetFocus

    If Not IsNumeric(Text11) Then
        Text11.SetFocus
        Text11 = ""
        MsgBox "DISKON (%) MENGGUNAKAN ANGKA", vbCritical, "TYPE DATA SALAH"
        Exit Sub
    End If


Text13 = Format(CCur(Text16) * CCur(Text11) / 100, "##,###.00")
Text11 = Format(Text11, "##,###.00")
Text15.SetFocus

Text12 = Format(CCur(Text12) - CCur(Text13), "##,###.00")

End If

End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

If Text15 = "" Then Text15.SetFocus

If Not IsNumeric(Text15) Then
    Text15.SetFocus
    Text15 = ""
    MsgBox "NOMINAL PEMBAYARAN MENGGUNAKAN ANGKA", vbCritical, "TYPE DATA SALAH"
    Exit Sub
End If
    
    If CCur(Text15) < CCur(Text12) Then
        Label16 = "TOTAL BAYAR"
        Text12 = Format(CCur(Text16) - CCur(Text13), "##,###.00")
        Text15.SetFocus
        Text15 = ""
        MsgBox "NOMINAL PEMBAYARAN KURANG", vbCritical, "WARNING"
        Exit Sub
    End If

    Label16 = "KEMBALIAN"
    Text12 = Format(CCur(Text15) - CCur(Text12), "##,###.00")
    
Text15 = Format(Text15, "##,###.00")
SendKeys vbTab
End If
End Sub

Private Sub Text3_Change()
'OYEN = "Select * from S003 Where NAMA like '%" + Trim(Text3) + "%'"
'GridInfo.Clear
'GridInfo.Refresh
'Call IsiGridInfo
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Text3 = "" Then Text3 = "-"
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text3_LostFocus()
Text3 = Format(Text3, ">")
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Text4 = "" Then Text4 = "-"
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text4_LostFocus()
Text4 = Format(Text4, ">")
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Text5 = "" Then Text5 = "-"
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text5_LostFocus()
Text5 = Format(Text5, ">")
End Sub

Private Sub Text6_Change()
'If Text6 = "" Then Exit Sub

'OYEN = "Select * from S003 Where NOPOL like '%" + Trim(Text3) + "%'"
'GridInfo.Clear
'GridInfo.Refresh
'Call IsiGridInfo
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Text6 = "" Then Text6 = "-"
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text6_LostFocus()
Text6 = Format(Text6, ">")
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Text7 = "" Then Text7 = "-"
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text7_LostFocus()
Text7 = Format(Text7, ">")
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Text8 = "" Then Text8 = "-"
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text8_LostFocus()
Text8 = Format(Text8, ">")
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Text9 = "" Then Text9 = "-"
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text9_LostFocus()
Text9 = Format(Text9, ">")
End Sub

Private Sub TmbSave_Click()
If Text15 = "" Then
    MsgBox "NOMINAL BAYAR KOSONG", vbCritical, "DATA TIDAK BOLEH KOSONG"
    Text15.SetFocus
    Exit Sub
End If

If Text100 = 0 Then
    Text1 = "-"
    Text2 = "-"
    Text3 = "-"
    Text4 = "-"
    Text5 = "-"
    Text6 = "-"
    Text7 = "-"
    Text8 = "-"
    Text9 = "-"
End If

If Text1 = "" Or Text2 = "" Then
    MsgBox "DATA MEKANIK MASIH KOSONG", vbCritical, "DATA TIDAK BOLEH KOSONG"
    Text1.SetFocus
    Exit Sub
End If

If Text3 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Or Text7 = "" Or Text8 = "" Or Text9 = "" Then
    MsgBox "INFO KENDARAAN MASIH KOSONG", vbCritical, "DATA TIDAK BOLEH KOSONG"
    Text3.SetFocus
    Exit Sub
End If

Dim Tanya
Tanya = MsgBox("TRANSAKSI SELESAI", vbSystemModal, "KONFIRMASI")
    If Tanya = vbOK Then
        Call SimpanS003
        Call LabaBengkel
        Call PERSEDIAAN_SPAREPART
        Call JurnalKas
        Call JURNALKREDIT
        Call LabaRugi
        If Text13 > 0 Then
            Call DiscPart
        End If
        Call Del_HISB004
    Else
        MsgBox "CANCEL", vbSystemModal, "KONFIRMASI"
    End If

Call NoBukti2

NoUrut = ""
NoUrut = Label3


ClearTextBoxes Me

STS_KSG = 0
Unload Me
S003A.Show

End Sub

Private Sub Del_HISB004()
SDel3 = "Delete * From B004 Where STS_S003A='1'"
Set RDel3 = RDCO.OpenResultset(SDel3, rdOpenDynamic, rdConcurRowVer)
RDel3.Close
Set RDel3 = Nothing
End Sub

Private Sub DiscPart()
Dim Enak

SSimpan2 = "Select * From G003 where codesl = '" + Trim(Text17) + "'"
Set RSimpan2 = RDCO.OpenResultset(SSimpan2, rdOpenDynamic, rdConcurRowVer)
Enak = RSimpan2("Posisi")
    If RSimpan2("Posisi") = "D" Then
        RSimpan2.EDIT
        A = CCur(RSimpan2("mutasid")) + CCur(Text13)
        RSimpan2("mutasid") = A
        RSimpan2("saldo") = CCur(RSimpan2("SaldoAwal")) + CCur(RSimpan2("mutasid")) - CCur(RSimpan2("mutasic"))
    ElseIf RSimpan2("Posisi") = "C" Then
        RSimpan2.EDIT
        A = CCur(RSimpan2("mutasid")) + CCur(Text13)
        RSimpan2("mutasid") = A
        RSimpan2("saldo") = CCur(RSimpan2("SaldoAwal")) - CCur(RSimpan2("mutasid")) + CCur(RSimpan2("mutasic"))
    End If
    
'        SSave2 = "Select * From G004"
'        Set RSave2 = RDCO.OpenResultset(SSave2, rdOpenDynamic, rdConcurRowVer)
'        RSave2.AddNew
'            RSave2("CodeCab") = CodeCab
'            RSave2("Codesl") = Trim(Text17)
'            RSave2("NamaSL") = Trim(Label17)
'            RSave2("NoBukti") = Label3
'            RSave2("Keterangan") = "DISCPART." + Trim(label173)
'            RSave2("NominalD") = CCur(Text13)
'            RSave2("NominalC") = 0
'            RSave2("Tanggal") = Tanggal
'            RSave2("UserCode") = Operator
'            RSave2("Jam") = Time
'        RSave2.Update
'        RSave2.Close
'        Set RSave2 = Nothing
        
RSimpan2.Update

            SDebet2 = "Select * From G005"
            Set RDebet2 = RDCO.OpenResultset(SDebet2, rdOpenKeyset, rdConcurRowVer)
                RDebet2.AddNew
                RDebet2("codecab") = CodeCab
                RDebet2("codesl") = Trim(Text17)
                RDebet2("namasl") = Label17
                RDebet2("nobukti") = Label3
                RDebet2("keterangan") = "DISCPART." + Trim(label173)
                RDebet2("nominald") = CCur(Text13)
                RDebet2("nominalc") = 0
                RDebet2("saldo") = RSimpan2("SALDO")
                RDebet2("tanggal") = Tanggal
                RDebet2("jam") = Time
                RDebet2("usercode") = Operator
            RDebet2.Update
            RDebet2.Close
            Set RDebet2 = Nothing
            
RSimpan2.Close
Set RSimpan2 = Nothing

SSimpan3 = "Select * From G003 where codesl = '" + Trim(Text18) + "'"
Set RSimpan3 = RDCO.OpenResultset(SSimpan3, rdOpenDynamic, rdConcurRowVer)
    If RSimpan3("Posisi") = "D" Then
        RSimpan3.EDIT
        A = CCur(RSimpan3("mutasic")) + CCur(Text13)
        RSimpan3("mutasic") = A
        RSimpan3("saldo") = CCur(RSimpan3("SaldoAwal")) + CCur(RSimpan3("mutasid")) - CCur(RSimpan3("mutasic"))
    ElseIf RSimpan3("Posisi") = "C" Then
        RSimpan3.EDIT
        A = CCur(RSimpan3("mutasic")) + CCur(Text13)
        RSimpan3("mutasic") = A
        RSimpan3("saldo") = CCur(RSimpan3("SaldoAwal")) - CCur(RSimpan3("mutasid")) + CCur(RSimpan3("mutasic"))
    End If
       
'        SSave3 = "Select * From G004"
'        Set RSave3 = RDCO.OpenResultset(SSave3, rdOpenDynamic, rdConcurRowVer)
'        RSave3.AddNew
'            RSave3("CodeCab") = CodeCab
'            RSave3("Codesl") = Trim(Text18)
'            RSave3("Namasl") = Label17
'            RSave3("NoBukti") = Label3
'            RSave3("Keterangan") = "DISCPART." + Trim(label173)
'            RSave3("NominalD") = 0
'            RSave3("NominalC") = CCur(Text13)
'            RSave3("Tanggal") = Tanggal
'            RSave3("UserCode") = Operator
'            RSave3("Jam") = Time
'        RSave3.Update
'        RSave3.Close
'        Set RSave3 = Nothing
        
RSimpan3.Update

            SCredit2 = "Select * From G005"
            Set RCredit2 = RDCO.OpenResultset(SCredit2, rdOpenKeyset, rdConcurRowVer)
            RCredit2.AddNew
                RCredit2("codecab") = CodeCab
                RCredit2("codesl") = Trim(Text18)
                RCredit2("namasl") = Label18
                RCredit2("nobukti") = Label3
                RCredit2("keterangan") = "DISCPART." + Trim(label173)
                RCredit2("nominald") = 0
                RCredit2("nominalc") = CCur(Text13)
                RCredit2("saldo") = RSimpan3("SALDO")
                RCredit2("tanggal") = Tanggal
                RCredit2("jam") = Time
                RCredit2("usercode") = Operator
            RCredit2.Update
            RCredit2.Close
            Set RCredit2 = Nothing


RSimpan3.Close
Set RSimpan3 = Nothing

SSimpan5 = "Select * From LabaRugi"
Set RSimpan5 = RDCO.OpenResultset(SSimpan5, rdOpenDynamic, rdConcurRowVer)
    SaldoD = CCur(RSimpan5("sumofmutasid"))
    SaldoC = CCur(RSimpan5("sumofmutasic"))
    
    SSave5 = "Select * From G003 where Posisi = 'L'"
    Set RSave5 = RDCO.OpenResultset(SSave5, rdOpenKeyset, rdConcurRowVer)
    Saldo = RSave5("saldoawal")
    RSave5.EDIT
        RSave5("mutasid") = SaldoD
        RSave5("mutasic") = SaldoC
        RSave5("saldo") = CCur(RSave5("SaldoAwal")) - CCur(RSave5("mutasid")) + CCur(RSave5("mutasic"))
    RSave5.Update
    RSave5.Close
    Set RSave5 = Nothing

RSimpan5.Close
Set RSimpan5 = Nothing
End Sub

Private Sub LabaBengkel()
SSave = "Select * From B005A"
Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
RSave.AddNew
    RSave("NO_FAKTUR") = Label3
    RSave("TANGGAL") = Tanggal
    
    SCari = "Select * from S003AA"
    Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset, rdConcurReadOnly)
    If RCari.RowCount <> 0 Then
        RSave("LABA") = RCari("SumOfLABA")
    End If
    RCari.Close
    Set RCari = Nothing
RSave.Update
RSave.Close
Set RSave = Nothing
End Sub

Private Sub SimpanS003()
'UPDATE JURNAL PENJUALAN
SSave2 = "Select * From S003"
Set RSave2 = RDCO.OpenResultset(SSave2, rdOpenDynamic, rdConcurRowVer)
RSave2.AddNew
    RSave2("NO_TRANS") = Label3
    RSave2("TGL_TRANS") = Tanggal
    RSave2("N_MEKANIK") = Trim(Text1)
    RSave2("JAM") = Trim(Text2)
    RSave2("NAMA") = Trim(Text3)
    RSave2("ALAMAT") = Trim(Text4)
    RSave2("TELEPON") = Trim(Text5)
    RSave2("NO_POL") = Trim(Text6)
    RSave2("NO_RANGKA") = Trim(Text7)
    RSave2("NO_MESIN") = Trim(Text8)
    
    RSave2("TIPE") = Trim(Text9)
    
    RSave2("TOTAL") = CCur(Text16)
    RSave2("BAYAR") = CCur(Text16) - CCur(Text13)
    RSave2("CASH") = CCur(Text15)
    
    RSave2("DISKON") = CCur(Text11)
    RSave2("NOMDIS") = CCur(Text13)
    
    RSave2("TERBILANG") = Terbilang(CCur(Text16) - CCur(Text13))
    
        
    RSave2("KEMBALIAN") = CCur(Text12)
    
'    If STS_KSG = 1 Then
'        RSave2("STS_KSG") = "1"
'    Else
'        RSave2("STS_KSG") = "0"
'    End If
    
RSave2.Update
RSave2.Close
Set RSave2 = Nothing

'UPDATE JURNAL PELAYANAN
If grid3.TextMatrix(grid3.Row, 0) = "" Then
Else
    SSave3 = "Select * From S003A order by KODE_S"
    Set RSave3 = RDCO.OpenResultset(SSave3, rdOpenDynamic, rdConcurRowVer)
    RSave3.MoveFirst
    Do While Not RSave3.EOF
        KODES = RSave3("KODE_S")
        NAMAS = RSave3("NAMA_S")
        BIAYAS = RSave3("BIAYA_S")
        
            SSave4 = "Select * From S004"
            Set RSave4 = RDCO.OpenResultset(SSave4, rdOpenDynamic, rdConcurRowVer)
            RSave4.AddNew
                NoUrutTrans = NoUrutTrans + 1
                RSave4("NO_TRANS") = Label3
                RSave4("TGL_TRANS") = Tanggal
'                RSave4("NO_URUT") = NoUrutTrans
                RSave4("KODE_S") = KODES
                RSave4("NAMA_S") = NAMAS
                RSave4("BIAYA_S") = BIAYAS
            RSave4.Update
            RSave4.Close
            Set RSave4 = Nothing
    RSave3.MoveNext
    Loop
End If

End Sub

Private Sub PERSEDIAAN_SPAREPART()
SJual = "Select * From S003A ORDER BY NO_URUT"
Set RJual = RDCO.OpenResultset(SJual, rdOpenKeyset, rdConcurRowVer)
RJual.MoveFirst
Do While Not RJual.EOF
    KODES = RJual("KODE_S")
    NAMAS = RJual("NAMA_S")
    BIAYAS = RJual("BIAYA_S")


'CABUT COY... JIKA BUKAN SPAREPART'
CEKKODE = Left(KODES, 3)
If CEKKODE = "IPT" Then

Else

    SJual2 = "Select * From B003A where KODE_JNS = '" + Trim(KODES) + "'"
    Set RJual2 = RDCO.OpenResultset(SJual2, rdOpenKeyset, rdConcurRowVer)
    CRD = RJual2("MUTASIC") + CCur(BIAYAS)
    JAKHIR = RJual2("SALDO") - CCur(BIAYAS)
    LABA = RJual2("HJUAL") - RJual2("HBELI")

'HISTORY JUAL SPAREPART'

        SJual3 = "Select * From B003 where KODE_JNS = '" + Trim(KODES) + "'"
        Set RJual3 = RDCO.OpenResultset(SJual3, rdOpenKeyset, rdConcurRowVer)
        JMLCRD = RJual3("JML_CRD") + 1
        JMLAKHIR = RJual3("JML_AKHIR") - 1

            SJual4 = "Select * From B005"
            Set RJual4 = RDCO.OpenResultset(SJual4, rdOpenKeyset, rdConcurRowVer)
            RJual4.AddNew
            RJual4("KODE_TRANS") = "JL"
            RJual4("KODE_JNS") = KODES
            RJual4("NAMA_JNS") = NAMAS
            RJual4("NO_FAKTUR") = Label3
            RJual4("NO_BUKTI") = Label3
            RJual4("KETERANGAN") = "JL.SP.NO." + KODES
            RJual4("JML_DBT") = 0
            RJual4("JML_CRD") = JMLCRD
            RJual4("JML_AKHIR") = JMLAKHIR
            RJual4("MUTASI_DBT") = 0
            RJual4("MUTASI_CRT") = BIAYAS
            RJual4("SALDO_AKHIR") = JAKHIR
            RJual4("H_POKOK") = RJual2("HBELI")
            RJual4("KAS") = LABA
            RJual4("TANGGAL") = Tanggal
            RJual4("TGL_TRANS") = Label10
            RJual4("LABA") = RJual("LABA")
            RJual4.Update
            RJual4.Close
            Set RJual4 = Nothing
            
'EDIT JUMLAH SPAREPART'
        RJual3.EDIT
        RJual3("JML_CRD") = CCur(JMLCRD)
        RJual3("JML_AKHIR") = CCur(JMLAKHIR)
        
        RJual3.Update
        RJual3.Close
        Set RJual3 = Nothing
        
'EDIT NOMINAL SPAREPART'
    RJual2.EDIT
    RJual2("MUTASIC") = CCur(CRD)
    RJual2("SALDO") = CCur(JAKHIR)
    
    RJual2.Update
    RJual2.Close
    Set RJual2 = Nothing

End If
RJual.MoveNext
Loop
RJual.Close
Set RJual = Nothing
    
End Sub

Private Sub JurnalKas()
SKAS = "Select * From S001 where KETERANGAN = '" + Trim(Operator) + "'"
Set RKAS = RDCO.OpenResultset(SKAS, rdOpenKeyset, rdConcurRowVer)
    SGLKAS = RKAS("SGL_KAS")
    
    SKAS2 = "Select * From G003 where CODESL = '" + Trim(SGLKAS) + "'"
    Set RKAS2 = RDCO.OpenResultset(SKAS2, rdOpenKeyset, rdConcurRowVer)
    NNAMASL = RKAS2("NAMASL")
    MMUTASID = RKAS2("mutasid") + CCur(Text16)
    SSALDO = RKAS2("Saldo") + CCur(Text16)
    
    RKAS2.EDIT
    RKAS2("MutasiD") = CCur(MMUTASID)
    RKAS2("Saldo") = CCur(SSALDO)

        SKAS3 = "Select * From G005"
        Set RKAS3 = RDCO.OpenResultset(SKAS3, rdOpenKeyset, rdConcurRowVer)
        RKAS3.AddNew
        RKAS3("codecab") = CodeCab
        RKAS3("codesl") = SGLKAS
        RKAS3("namasl") = NNAMASL
        RKAS3("nobukti") = Label3
        RKAS3("keterangan") = "SERVICE." + Text6
        RKAS3("nominald") = CCur(Text16)
        RKAS3("nominalc") = 0
        RKAS3("saldo") = SSALDO
        RKAS3("tanggal") = Tanggal
        RKAS3("jam") = Date
        RKAS3("usercode") = Operator
        RKAS3.Update
        RKAS3.Close
        Set RKAS3 = Nothing
    
    RKAS2.Update
    RKAS2.Close
    Set RKAS2 = Nothing
    
RKAS.Close
Set RKAS = Nothing
End Sub

Private Sub JURNALKREDIT()
SKAS = "Select * From S001 where KETERANGAN = '" + Trim(Operator) + "'"
Set RKAS = RDCO.OpenResultset(SKAS, rdOpenKeyset, rdConcurRowVer)
    
    SGLPART = RKAS("SGL_PART")
    Call KREDIT
    
    SGLJASA = RKAS("SGL_JASA")
    Call KREDIT2
    
    SGLSEDIA = RKAS("SGL_SEDIA")
    Call KREDIT3
    
RKAS.Close
Set RKAS = Nothing
End Sub

Private Sub KREDIT()
'CABUT COY... JIKA SPAREPART = 0'
If CCur(Text200) = 0 Then Exit Sub

SKREDIT = "Select * From G003 where CODESL = '" + Trim(SGLPART) + "'"
Set RKREDIT = RDCO.OpenResultset(SKREDIT, rdOpenKeyset, rdConcurRowVer)
    
    SKREDIT2 = "Select * From LABAPART"
    Set RKREDIT2 = RDCO.OpenResultset(SKREDIT2, rdOpenKeyset, rdConcurRowVer)

    LLABA = CCur(RKREDIT2("SumOfBiaya_S")) - CCur(RKREDIT2("SumOfHBeli"))
   
MUTC = CCur(RKREDIT("mutasic")) + CCur(LLABA)
SALD = CCur(RKREDIT("saldo")) + CCur(LLABA)

RKREDIT.EDIT
RKREDIT("mutasic") = CCur(MUTC)
RKREDIT("saldo") = CCur(SALD)

        SKAS3 = "Select * From G005"
        Set RKAS3 = RDCO.OpenResultset(SKAS3, rdOpenKeyset, rdConcurRowVer)
        RKAS3.AddNew
            RKAS3("codecab") = CodeCab
            RKAS3("codesl") = SGLPART
            RKAS3("namasl") = RKREDIT("namasl")
            RKAS3("nobukti") = Label3
            RKAS3("keterangan") = "J.SP.SERVICE." + Text6
            RKAS3("nominald") = 0
            RKAS3("nominalc") = CCur(LLABA)
            RKAS3("saldo") = CCur(SALD)
            RKAS3("tanggal") = Tanggal
            RKAS3("jam") = Date
            RKAS3("usercode") = Operator
        RKAS3.Update
        RKAS3.Close
        Set RKAS3 = Nothing

    RKREDIT2.Close
    Set RKREDIT2 = Nothing

RKREDIT.Update
RKREDIT.Close
Set RKREDIT = Nothing
End Sub

Private Sub KREDIT2()
'CABUT COY... JIKA JASA = 0'
If CCur(Text100) = 0 Then Exit Sub

SKREDIT = "Select * From G003 where CODESL = '" + Trim(SGLJASA) + "'"
Set RKREDIT = RDCO.OpenResultset(SKREDIT, rdOpenKeyset, rdConcurRowVer)
    
    SKREDIT2 = "Select * From LABAJASA"
    Set RKREDIT2 = RDCO.OpenResultset(SKREDIT2, rdOpenKeyset, rdConcurRowVer)

MUTC = CCur(RKREDIT("mutasic")) + CCur(RKREDIT2("SumOfBiaya_S"))
SALD = CCur(RKREDIT("saldo")) + CCur(RKREDIT2("SumOfBiaya_S"))

RKREDIT.EDIT
RKREDIT("mutasic") = CCur(MUTC)
RKREDIT("saldo") = CCur(SALD)

        SKAS3 = "Select * From G005"
        Set RKAS3 = RDCO.OpenResultset(SKAS3, rdOpenKeyset, rdConcurRowVer)
        RKAS3.AddNew
            RKAS3("codecab") = CodeCab
            RKAS3("codesl") = SGLJASA
            RKAS3("namasl") = RKREDIT("namasl")
            RKAS3("nobukti") = Label3
            RKAS3("keterangan") = "PDPT.SERVICE.IPT." + Text6
            RKAS3("nominald") = 0
            RKAS3("nominalc") = CCur(RKREDIT2("SumOfBiaya_S"))
            RKAS3("saldo") = CCur(SALD)
            RKAS3("tanggal") = Tanggal
            RKAS3("jam") = Date
            RKAS3("usercode") = Operator
        RKAS3.Update
        RKAS3.Close
        Set RKAS3 = Nothing

    RKREDIT2.Close
    Set RKREDIT2 = Nothing

RKREDIT.Update
RKREDIT.Close
Set RKREDIT = Nothing
End Sub

Private Sub KREDIT3()
'CABUT COY... JIKA SPAREPART = 0'
If CCur(Text200) = 0 Then Exit Sub

SKREDIT = "Select * From G003 where CODESL = '" + Trim(SGLSEDIA) + "'"
Set RKREDIT = RDCO.OpenResultset(SKREDIT, rdOpenKeyset, rdConcurRowVer)
    
    SKREDIT2 = "Select * From LABAPART"
    Set RKREDIT2 = RDCO.OpenResultset(SKREDIT2, rdOpenKeyset, rdConcurRowVer)

MUTC = CCur(RKREDIT("mutasic")) + CCur(RKREDIT2("SumOfHBeli"))
SALD = CCur(RKREDIT("saldo")) - CCur(RKREDIT2("SumOfHBeli"))

RKREDIT.EDIT
RKREDIT("mutasic") = CCur(MUTC)
RKREDIT("saldo") = CCur(SALD)

        SKAS3 = "Select * From G005"
        Set RKAS3 = RDCO.OpenResultset(SKAS3, rdOpenKeyset, rdConcurRowVer)
        RKAS3.AddNew
            RKAS3("codecab") = CodeCab
            RKAS3("codesl") = SGLSEDIA
            RKAS3("namasl") = RKREDIT("namasl")
            RKAS3("nobukti") = Label3
            RKAS3("keterangan") = "J.SP.SERVICE." + Text6
            RKAS3("nominald") = 0
            RKAS3("nominalc") = CCur(RKREDIT2("SumOfHBeli"))
            RKAS3("saldo") = CCur(SALD)
            RKAS3("tanggal") = Tanggal
            RKAS3("jam") = Date
            RKAS3("usercode") = Operator
        RKAS3.Update
        RKAS3.Close
        Set RKAS3 = Nothing

    RKREDIT2.Close
    Set RKREDIT2 = Nothing

RKREDIT.Update
RKREDIT.Close
Set RKREDIT = Nothing
End Sub

Private Sub LabaRugi()
SLABA = "Select * From LabaRugi"
Set RLABA = RDCO.OpenResultset(SLABA, rdOpenDynamic, rdConcurRowVer)
    SaldoD = CCur(RLABA("sumofmutasid"))
    SaldoC = CCur(RLABA("sumofmutasic"))
    
    SLABA2 = "Select * From G003 where Posisi = 'L'"
    Set RLABA2 = RDCO.OpenResultset(SLABA2, rdOpenKeyset, rdConcurRowVer)
    Saldo = RLABA2("saldoawal")
    RLABA2.EDIT
        RLABA2("mutasid") = SaldoD
        RLABA2("mutasic") = SaldoC
        RLABA2("saldo") = CCur(RLABA2("SaldoAwal")) - CCur(RLABA2("mutasid")) + CCur(RLABA2("mutasic"))
    RLABA2.Update
    RLABA2.Close
    Set RLABA2 = Nothing

RLABA.Close
Set RLABA = Nothing
End Sub

Private Sub SiapkanGridInfo()
With GridInfo
    .Cols = 7
    .Row = 0
    .Col = 0: .ColWidth(0) = "2000": .Text = "NAMA": .CellAlignment = 4
    .Col = 1: .ColWidth(1) = "3500": .Text = "ALAMAT": .CellAlignment = 4
    .Col = 2: .ColWidth(2) = "1500": .Text = "TELEPON": .CellAlignment = 4
    .Col = 3: .ColWidth(3) = "1500": .Text = "POLISI": .CellAlignment = 4
    .Col = 4: .ColWidth(4) = "1500": .Text = "RANGKA": .CellAlignment = 4
    .Col = 5: .ColWidth(5) = "1500": .Text = "MESIN": .CellAlignment = 4
    .Col = 6: .ColWidth(6) = "2000": .Text = "TIPE/WARNA/TAHUN": .CellAlignment = 4
End With
End Sub

Private Sub IsiGridInfo()
Dim Brs As Integer

Call SiapkanGridInfo

Brs = 1
SCari = OYEN
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset, rdConcurReadOnly)
If RCari.RowCount <> 0 Then
RCari.MoveFirst
Do While Not RCari.EOF
    With GridInfo
        .Rows = Brs + 1
        .Row = Brs
        .Col = 0: .Text = Trim(RCari("NAMA"))
        .Col = 1: .Text = Trim(RCari("ALAMAT"))
        .Col = 2: .Text = RCari("TELEPON")
        .Col = 3: .Text = RCari("NO_POL")
        .Col = 4: .Text = RCari("NO_RANGKA")
        .Col = 5: .Text = RCari("NO_MESIN")
        .Col = 6: .Text = RCari("TIPE")
    End With
    Brs = Brs + 1
RCari.MoveNext
Loop
End If
RCari.Close
Set RCari = Nothing
End Sub
