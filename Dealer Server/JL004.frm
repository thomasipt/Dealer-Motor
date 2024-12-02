VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form JL004 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TRANSAKSI KENDARAAN"
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8895
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   8895
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
      Height          =   420
      Left            =   3982
      TabIndex        =   83
      Top             =   8337
      Width           =   960
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   6570
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   124
      Width           =   1965
   End
   Begin VB.Frame Frame4 
      Caption         =   "Info Pembeli"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2325
      Left            =   3795
      TabIndex        =   38
      Top             =   1842
      Width           =   4950
      Begin VB.TextBox Text23 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   1260
         MaxLength       =   255
         TabIndex        =   3
         Text            =   "Text23"
         Top             =   1785
         Width           =   3525
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   1260
         MaxLength       =   255
         TabIndex        =   1
         Text            =   "Text3"
         Top             =   315
         Width           =   3525
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00FFFFFF&
         Height          =   990
         Left            =   1260
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   2
         Text            =   "JL004.frx":0000
         Top             =   735
         Width           =   3525
      End
      Begin VB.Label Label4 
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
         Left            =   105
         TabIndex        =   40
         Top             =   1125
         Width           =   1140
      End
      Begin VB.Label Label3 
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
         Left            =   105
         TabIndex        =   39
         Top             =   390
         Width           =   1140
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Info Kendaraan"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3480
      Left            =   120
      TabIndex        =   27
      Top             =   1842
      Width           =   3480
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Left            =   1260
         TabIndex        =   37
         Top             =   2325
         Width           =   2085
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Left            =   1260
         TabIndex        =   36
         Top             =   435
         Width           =   1875
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Left            =   1260
         TabIndex        =   35
         Top             =   2955
         Width           =   2085
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Left            =   1260
         TabIndex        =   34
         Top             =   1695
         Width           =   825
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Left            =   1260
         TabIndex        =   33
         Top             =   1065
         Width           =   1875
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
         Left            =   105
         TabIndex        =   32
         Top             =   2940
         Width           =   1140
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
         Left            =   105
         TabIndex        =   31
         Top             =   2310
         Width           =   1140
      End
      Begin VB.Label Label7 
         Caption         =   "TAHUN"
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
         Left            =   105
         TabIndex        =   30
         Top             =   1680
         Width           =   1140
      End
      Begin VB.Label Label6 
         Caption         =   "WARNA"
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
         Left            =   105
         TabIndex        =   29
         Top             =   1050
         Width           =   1140
      End
      Begin VB.Label Label5 
         Caption         =   "TYPE"
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
         Left            =   105
         TabIndex        =   28
         Top             =   420
         Width           =   1140
      End
   End
   Begin VB.TextBox Text12 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5160
      TabIndex        =   4
      Text            =   "Text12"
      Top             =   4362
      Width           =   3000
   End
   Begin VB.TextBox Text13 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5160
      TabIndex        =   5
      Text            =   "Text13"
      Top             =   4797
      Width           =   3000
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   2745
      Left            =   120
      TabIndex        =   20
      Top             =   5412
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   4842
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "BIAYA LAINNYA"
      TabPicture(0)   =   "JL004.frx":0006
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame7"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "PEMBAYARAN"
      TabPicture(1)   =   "JL004.frx":0022
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "PIUTANG"
      TabPicture(2)   =   "JL004.frx":003E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame7 
         Height          =   2355
         Left            =   -74940
         TabIndex        =   68
         Top             =   315
         Width           =   8535
         Begin VB.TextBox Text11 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6045
            TabIndex        =   10
            Text            =   "Text11"
            Top             =   990
            Width           =   2055
         End
         Begin VB.TextBox Text10 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6045
            TabIndex        =   9
            Text            =   "Text10"
            Top             =   465
            Width           =   2055
         End
         Begin VB.TextBox Text9 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2160
            TabIndex        =   8
            Text            =   "Text9"
            Top             =   1515
            Width           =   2055
         End
         Begin VB.TextBox Text8 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2160
            TabIndex        =   7
            Text            =   "Text8"
            Top             =   990
            Width           =   2055
         End
         Begin VB.TextBox Text7 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2160
            TabIndex        =   6
            Text            =   "Text7"
            Top             =   465
            Width           =   2055
         End
         Begin VB.Label Label19 
            Caption         =   "DISKON"
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
            Left            =   4680
            TabIndex        =   73
            Top             =   1020
            Width           =   1140
         End
         Begin VB.Label Label18 
            Caption         =   "BROKER"
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
            Left            =   4680
            TabIndex        =   72
            Top             =   495
            Width           =   1140
         End
         Begin VB.Label Label17 
            Caption         =   "KACAB"
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
            Left            =   795
            TabIndex        =   71
            Top             =   1545
            Width           =   1140
         End
         Begin VB.Label Label16 
            Caption         =   "JAKET"
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
            Left            =   795
            TabIndex        =   70
            Top             =   1020
            Width           =   1140
         End
         Begin VB.Label Label15 
            Caption         =   "BENSIN"
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
            Left            =   795
            TabIndex        =   69
            Top             =   495
            Width           =   1140
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2355
         Left            =   60
         TabIndex        =   57
         Top             =   315
         Width           =   8535
         Begin MSFlexGridLib.MSFlexGrid GridGL 
            Height          =   1140
            Left            =   4365
            TabIndex        =   67
            Top             =   1050
            Width           =   3750
            _ExtentX        =   6615
            _ExtentY        =   2011
            _Version        =   393216
            Appearance      =   0
         End
         Begin VB.Frame Frame2 
            Caption         =   "Transaksi Pembayaran"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2010
            Left            =   60
            TabIndex        =   63
            Top             =   210
            Width           =   3900
            Begin VB.TextBox Text14 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Height          =   360
               Left            =   1980
               TabIndex        =   11
               Text            =   "Text14"
               Top             =   315
               Width           =   1860
            End
            Begin VB.TextBox Text15 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Height          =   360
               Left            =   1980
               TabIndex        =   12
               Text            =   "Text15"
               Top             =   735
               Width           =   1860
            End
            Begin VB.TextBox Text16 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Height          =   360
               Left            =   1995
               TabIndex        =   14
               Text            =   "Text16"
               Top             =   1155
               Width           =   1860
            End
            Begin VB.Label Label29 
               Caption         =   "> TUNAI"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   300
               Left            =   105
               TabIndex        =   66
               Top             =   345
               Width           =   1230
            End
            Begin VB.Label Label31 
               Caption         =   "> PIUTANG"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   300
               Left            =   105
               TabIndex        =   65
               Top             =   1185
               Width           =   1755
            End
            Begin VB.Label Label32 
               Caption         =   "> NON TUNAI"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   300
               Left            =   105
               TabIndex        =   64
               Top             =   765
               Width           =   1755
            End
         End
         Begin VB.Frame Frame6 
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
            Left            =   4155
            TabIndex        =   58
            Top             =   210
            Width           =   3960
            Begin VB.TextBox Text18 
               BackColor       =   &H00C0C0C0&
               Height          =   360
               Left            =   1320
               TabIndex        =   13
               Text            =   "Text18"
               Top             =   390
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
               Left            =   2385
               TabIndex        =   59
               Top             =   390
               Width           =   1515
            End
            Begin VB.Label Label34 
               Caption         =   "KODE SUB GL"
               Height          =   330
               Left            =   135
               TabIndex        =   62
               Top             =   405
               Width           =   1455
            End
            Begin VB.Label Label35 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Label35"
               Height          =   330
               Left            =   165
               TabIndex        =   61
               Top             =   810
               Width           =   3660
            End
            Begin VB.Label InfoGL2 
               Alignment       =   2  'Center
               BackColor       =   &H00FFC0C0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Info"
               Height          =   360
               Left            =   2595
               TabIndex        =   60
               Top             =   390
               Width           =   675
            End
         End
         Begin VB.CommandButton TblSave 
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
            Left            =   4155
            TabIndex        =   21
            Top             =   1800
            Width           =   3960
         End
      End
      Begin VB.Frame Frame5 
         Height          =   2355
         Left            =   -74940
         TabIndex        =   41
         Top             =   315
         Width           =   8535
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   5775
            MaxLength       =   12
            TabIndex        =   18
            Text            =   "Text2"
            Top             =   1507
            Width           =   2175
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
            Left            =   3705
            TabIndex        =   43
            Top             =   1080
            WhatsThisHelpID =   870
            Width           =   705
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
            Left            =   2970
            TabIndex        =   42
            Top             =   1080
            WhatsThisHelpID =   870
            Width           =   615
         End
         Begin VB.TextBox Text19 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1470
            MaxLength       =   12
            TabIndex        =   15
            Text            =   "Text19"
            Top             =   1080
            Width           =   1410
         End
         Begin VB.TextBox Text20 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1440
            TabIndex        =   19
            Text            =   "Text20"
            Top             =   1920
            Width           =   6510
         End
         Begin VB.TextBox Text21 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   5775
            TabIndex        =   17
            Text            =   "Text21"
            Top             =   615
            Width           =   1710
         End
         Begin VB.TextBox Text22 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   5775
            TabIndex        =   16
            Text            =   "Text22"
            Top             =   210
            Width           =   1710
         End
         Begin VB.Label Label2 
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
            Left            =   4725
            TabIndex        =   84
            Top             =   1530
            Width           =   1155
         End
         Begin VB.Label Label52 
            Caption         =   "(dd/mm/yyyy)"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7560
            TabIndex        =   80
            Top             =   450
            Width           =   915
         End
         Begin VB.Label Label36 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1470
            TabIndex        =   56
            Top             =   1530
            Width           =   2610
         End
         Begin VB.Label Label41 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2295
            TabIndex        =   55
            Top             =   240
            Width           =   2310
         End
         Begin VB.Label Label40 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1470
            TabIndex        =   54
            Top             =   240
            Width           =   750
         End
         Begin VB.Label Label39 
            Caption         =   "KODE PIUTANG"
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
            Left            =   105
            TabIndex        =   53
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label Label38 
            Caption         =   "SYARAT"
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
            Left            =   105
            TabIndex        =   52
            Top             =   1950
            Width           =   1020
         End
         Begin VB.Label Label43 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Label43"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5775
            TabIndex        =   51
            Top             =   1065
            Width           =   2475
         End
         Begin VB.Label Label44 
            Caption         =   "NOMINAL"
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
            Left            =   4725
            TabIndex        =   50
            Top             =   1065
            Width           =   1155
         End
         Begin VB.Label Label45 
            Caption         =   "TGL JATUH"
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
            Left            =   4725
            TabIndex        =   49
            Top             =   645
            Width           =   1245
         End
         Begin VB.Label Label46 
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
            Left            =   4725
            TabIndex        =   48
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label47 
            Caption         =   "Nama"
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
            Left            =   105
            TabIndex        =   47
            Top             =   1530
            Width           =   1260
         End
         Begin VB.Label Label48 
            Caption         =   "No Customer"
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
            Left            =   105
            TabIndex        =   46
            Top             =   1110
            Width           =   1125
         End
         Begin VB.Label Label37 
            Caption         =   "NO. PIUTANG"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   105
            TabIndex        =   45
            Top             =   630
            Width           =   1125
         End
         Begin VB.Label Label42 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1470
            TabIndex        =   44
            Top             =   645
            Width           =   1875
         End
      End
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "TGL PENJUALAN"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4395
      TabIndex        =   82
      Top             =   132
      Width           =   1860
   End
   Begin VB.Label Label30 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "Label30"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   195
      TabIndex        =   81
      Top             =   162
      Width           =   3135
   End
   Begin VB.Label Label51 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2010
      TabIndex        =   79
      Top             =   1317
      Width           =   2055
   End
   Begin VB.Label Label50 
      Caption         =   "HARGA OTR  Rp."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   285
      TabIndex        =   78
      Top             =   1287
      Width           =   1875
   End
   Begin VB.Label Label49 
      Alignment       =   1  'Right Justify
      Caption         =   "TOTAL BAYAR   Rp."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4380
      TabIndex        =   77
      Top             =   1287
      Width           =   1875
   End
   Begin VB.Label Label33 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6480
      TabIndex        =   76
      Top             =   1317
      Width           =   2055
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6480
      TabIndex        =   75
      Top             =   897
      Width           =   2055
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      Caption         =   "BIAYA   Rp."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4800
      TabIndex        =   74
      Top             =   867
      Width           =   1455
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   "OTR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4215
      TabIndex        =   26
      Top             =   4452
      Width           =   720
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "BBN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4215
      TabIndex        =   25
      Top             =   4872
      Width           =   720
   End
   Begin VB.Label Label23 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6480
      TabIndex        =   24
      Top             =   552
      Width           =   2055
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      Caption         =   "HARGA DEALER"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4275
      TabIndex        =   23
      Top             =   522
      Width           =   1980
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Label10"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   285
      TabIndex        =   22
      Top             =   492
      Width           =   8250
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C00000&
      Height          =   960
      Left            =   3795
      Shape           =   4  'Rounded Rectangle
      Top             =   4302
      Width           =   4950
   End
End
Attribute VB_Name = "JL004"
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
JL003.Show
End Sub

Private Sub EntriSPL2_Click()
C012.Show 1
Text19.SetFocus
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=DEALER", rdDriverNoPrompt, False, CN)

Call Kosong
Call CariData
Cari_SubGL

Cash = 0
Bank = 0
KHutang = 0

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

Private Sub CariData()
SToket = "Select * From M001 where RANGKA = '" + Trim(NO_RANGKA) + "'"
Set RToket = RDCO.OpenResultset(SToket, rdOpenDynamic, rdOpenKeyset)
If RToket.RowCount <> 0 Then
    Label27 = RToket("TYPE")
    Label24 = RToket("WARNA")
    Label25 = RToket("TAHUN")
    Label28 = RToket("RANGKA")
    Label26 = RToket("MESIN")
    
    Label23 = Format(RToket("H_Beli"), "##,###.00")
    Label51 = Format(RToket("H_OTR"), "##,###.00")
    Label10 = RToket("MTS_MOTOR")
    
    Text12 = Label51
    
    Label30 = RToket("NO_FAK")
    
End If

If Label10 = CodeCab Then
    Label10 = "POSISI -->> " + N_CCAB
End If

End Sub

Private Sub Kosong()
Label30 = ""
Text1 = Tanggal
Text4 = Tanggal
Text2 = 0
Text3 = ""
Text5 = ""
Text23 = ""
Text12 = 0
Text13 = 0
Text7 = 0
Text8 = 0
Text9 = 0
Text10 = 0
Text11 = 0
Text14 = 0
Text15 = 0
Text16 = 0
Text18 = ""
Label35 = ""
Text19 = ""
Text20 = ""
Text22 = Tanggal
Text21 = Tanggal
SSTab2.Tab = 0
Frame7.Enabled = False
Frame1.Enabled = False
Frame5.Enabled = False
Text7.BackColor = &HC0C0C0
Text8.BackColor = &HC0C0C0
Text9.BackColor = &HC0C0C0
Text10.BackColor = &HC0C0C0
Text11.BackColor = &HC0C0C0
Text14.BackColor = &HC0C0C0
Text15.BackColor = &HC0C0C0
Text16.BackColor = &HC0C0C0
Text19.BackColor = &HC0C0C0
Text20.BackColor = &HC0C0C0
Text21.BackColor = &HC0C0C0
Text22.BackColor = &HC0C0C0
Call InfoGL2_Click
Call FrameGLNonAktif
End Sub

Private Sub GridGL_DblClick()
Call InfoGL2_Click
Text18 = GridGL.TextMatrix(GridGL.Row, 0)
Label35 = GridGL.TextMatrix(GridGL.Row, 1)
Text18.SetFocus
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
No_Nas = ""
Nama_Nas = ""
C013.Show 1
Text19 = No_Nas
Label36 = Nama_Nas
Text22.SetFocus
End Sub

Private Sub TblSave_Click()
Dim Tanya, NoTran

If CCur(Text14) = 0 Then
    Cash = 0
Else
    Cash = 1
End If

If CCur(Text15) = 0 Then
    Bank = 0
Else
    Bank = 1
End If

If CCur(Text16) = 0 Then
    KHutang = 0
Else
    KHutang = 1
End If

If Text1 = "" Then
    MsgBox "DATA TANGGAL PENJUALAN MASIH KOSONG", vbCritical, "DATA TIDAK BOLEH KOSONG"
    Text1.SetFocus
    Exit Sub
End If


If Text3 = "" Or Text5 = "" Or Text23 = "" Then
    MsgBox "DATA PEMBELI MASIH KOSONG", vbCritical, "DATA TIDAK BOLEH KOSONG"
    Text3.SetFocus
    Exit Sub
End If

If Text12 = "" Or Text13 = "" Then
    MsgBox "HARGA OTR / BIAYA BBN MASIH KOSONG", vbCritical, "DATA TIDAK BOLEH KOSONG"
    Text12.SetFocus
    Exit Sub
End If

Call JL004

If Cash = 0 And Bank = 0 And KHutang = 1 Then   'Hutang
    Tanya = MsgBox("YAKIN PROSES TRANSAKSI PENJUALAN PIUTANG Rp. " + Text16 + " ?", vbOKCancel, "SIMPAN TRANSAKSI PENJUALAN")
    If Tanya = vbCancel Then Exit Sub
    Call SetJual
    Call JurnalBahan
    Call JurnalBiaya
    Call JurnalBiaya2
    Call JurnalKas2
    Call JurnalHutang
    Call JurnalPendapatan
    Call Hutang
ElseIf Cash = 0 And Bank = 1 And KHutang = 0 Then   'NON TUNAI
    Tanya = MsgBox("YAKIN PROSES TRANSAKSI PENJUALAN NON TUNAI Rp. " + Text15 + " ?", vbOKCancel, "SIMPAN TRANSAKSI PENJUALAN")
    If Tanya = vbCancel Then Exit Sub
    Call SetJual
    Call JurnalBahan
    Call JurnalBiaya
    Call JurnalBiaya2
    Call JurnalKas2
    Call JurnalKas3
    Call JurnalPendapatan
ElseIf Cash = 0 And Bank = 1 And KHutang = 1 Then   'NON TUNAI + HUTANG
    Tanya = MsgBox("YAKIN PROSES TRANSAKSI PENJUALAN NON TUNAI Rp. " + Text15 + " , PIUTANG Rp. " + Text16 + " ?", vbOKCancel, "SIMPAN TRANSAKSI PENJUALAN")
    If Tanya = vbCancel Then Exit Sub
    Call SetJual
    Call JurnalBahan
    Call JurnalBiaya
    Call JurnalBiaya2
    Call JurnalKas2
    Call JurnalKas3
    Call JurnalHutang
    Call JurnalPendapatan
    Call Hutang
ElseIf Cash = 1 And Bank = 0 And KHutang = 0 Then   'TUNAI + HUTANG
    Tanya = MsgBox("YAKIN PROSES TRANSAKSI PENJUALAN TUNAI Rp. " + Text14 + " ?", vbOKCancel, "SIMPAN TRANSAKSI PENJUALAN")
    If Tanya = vbCancel Then Exit Sub
    Call SetJual
    Call JurnalKas
    Call JurnalBahan
    Call JurnalBiaya
    Call JurnalBiaya2
    Call JurnalKas2
    Call JurnalPendapatan
ElseIf Cash = 1 And Bank = 0 And KHutang = 1 Then   'TUNAI + Hutang
    Tanya = MsgBox("YAKIN PROSES TRANSAKSI PENJUALAN TUNAI Rp. " + Text14 + " , PIUTANG Rp. " + Text16 + " ?", vbOKCancel, "SIMPAN TRANSAKSI PENJUALAN")
    If Tanya = vbCancel Then Exit Sub
    Call SetJual
    Call JurnalKas
    Call JurnalBahan
    Call JurnalBiaya
    Call JurnalBiaya2
    Call JurnalKas2
    Call JurnalHutang
    Call JurnalPendapatan
    Call Hutang
ElseIf Cash = 1 And Bank = 1 And KHutang = 0 Then   'TUNAI + NON TUNAI
    Tanya = MsgBox("YAKIN PROSES TRANSAKSI PENJUALAN TUNAI Rp. " + Text14 + " , NON TUNAI BANK Rp. " + Text15 + " ?", vbOKCancel, "SIMPAN TRANSAKSI PENJUALAN")
    If Tanya = vbCancel Then Exit Sub
    Call SetJual
    Call JurnalKas
    Call JurnalBahan
    Call JurnalBiaya
    Call JurnalBiaya2
    Call JurnalKas2
    Call JurnalKas3
    Call JurnalPendapatan
ElseIf Cash = 1 And Bank = 1 And KHutang = 1 Then   'TUNAI + NON TUNAI + Hutang
    Tanya = MsgBox("YAKIN PROSES TRANSAKSI PENJUALAN TUNAI Rp. " + Text14 + " , NON TUNAI BANK Rp. " + Text15 + " , PIUTANG Rp. " + Text16 + " ?", vbOKCancel, "SIMPAN TRANSAKSI PENJUALAN")
    If Tanya = vbCancel Then Exit Sub
    Call SetJual
    Call JurnalKas
    Call JurnalBahan
    Call JurnalBiaya
    Call JurnalBiaya2
    Call JurnalKas2
    Call JurnalKas3
    Call JurnalHutang
    Call JurnalPendapatan
    Call Hutang
End If

Call LabaRugi

Unload Me
JL003.Show 1
End Sub

Private Sub text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text1_LostFocus()
If Text1 = "" Then Exit Sub
If Not IsDate(Text1) Then
    Text1.SetFocus
    Text1 = Tanggal
    MsgBox "TANGGAL HARUS TYPE TANGGAL", vbCritical, "TYPE DATA SALAH"
End If
Text1 = Format(Text1, "DD/MM/YYYY")
End Sub

Private Sub Text12_GotFocus()
If CCur(Text12) = 0 Then Text12 = ""
End Sub

Private Sub TEXT12_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Text12 <> Label51 Then
        MsgBox "HARGA OTR SALAH", vbCritical, "WARNING"
        Text12 = Label51
        Exit Sub
    End If
    SendKeys vbTab
End If
End Sub

Private Sub Text12_LostFocus()
If Text12 = "" Then Text12 = 0
If Not IsNumeric(Text12) Then
    Text12.SetFocus
    MsgBox "NOMINAL HARUS ANGKA", vbCritical, "TYPE DATA SALAH"
    Exit Sub
End If
Text12 = Format(Text12, "##,###.00")

End Sub

Private Sub Text13_GotFocus()
If CCur(Text13) = 0 Then Text13 = ""
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Frame7.Enabled = True
    SendKeys vbTab
End If
End Sub

Private Sub Text13_LostFocus()
If Text13 = "" Then Text13 = 0
If Not IsNumeric(Text13) Then
    Text13.SetFocus
    MsgBox "NOMINAL PEMBAYARAN HARUS ANGKA", vbCritical, "TYPE DATA SALAH"
    Exit Sub
End If
Text13 = Format(Text13, "##,###.00")

SSTab2.Tab = 0
Text7.BackColor = &HFFFFFF
Text8.BackColor = &HFFFFFF
Text9.BackColor = &HFFFFFF
Text10.BackColor = &HFFFFFF
Text11.BackColor = &HFFFFFF
End Sub

Private Sub Text14_GotFocus()
If CCur(Text14) = 0 Then Text14 = ""
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

Label33 = Format(CCur(Text14) + CCur(Text15) + CCur(Text16), "##,###.00")

    If CCur(Text14) > CCur(Text12) Then
        Text14.SetFocus
        MsgBox "TOTAL PEMBAYARAN TIDAK BOLEH MELEBIHI HARGA ON THE ROAD", vbCritical, "TOTAL PEMBAYARAN SALAH"
    ElseIf CCur(Text14) = CCur(Text12) Then
        Text15 = 0
        Text16 = 0
        Cash = 1
        TblSave.SetFocus
    ElseIf CCur(Text14) < CCur(Text12) Then
        Cash = 1
        Text15.SetFocus
    ElseIf CCur(Text14) = 0 Then
        Cash = 0
        Text15.SetFocus
    End If
    
End Sub

Private Sub Text15_GotFocus()
If Text15 = 0 Then Text15 = ""
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text15_LostFocus()
If Text15 = "" Then Text15 = 0
If Not IsNumeric(Text15) Then
    Text15.SetFocus
    MsgBox "NOMINAL PEMBAYARAN HARUS ANGKA", vbCritical, "TYPE DATA SALAH"
    Exit Sub
End If
Text15 = Format(Text15, "##,###.00")

Label33 = Format(CCur(Text14) + CCur(Text15) + CCur(Text16), "##,###.00")

If (CCur(Text15) + CCur(Text14)) > CCur(Text12) Then
    Text15.SetFocus
    MsgBox "TOTAL PEMBAYARAN TIDAK BOLEH MELEBIHI TOTAL PEMBELIAN", vbCritical, "TOTAL PEMBAYARAN SALAH"
    Exit Sub
End If

If CCur(Text15) > 0 Then
    Call FrameGLAktif
    Text18.SetFocus
    Bank = 1
    Exit Sub
ElseIf CCur(Text15) = 0 Then
    Bank = 0
    Call FrameGLNonAktif
    Text16.Enabled = True
    Text16.SetFocus
    Exit Sub
End If

End Sub

Private Sub Text16_GotFocus()
If Text14 = 0 And Text15 = 0 Then
    Text16 = Format(CCur(Text12), "##,###.00")
ElseIf Text14 <> 0 And Text15 = 0 Then
    Text16 = Format(CCur(Text12) - CCur(Text14), "##,###.00")
ElseIf Text14 <> 0 And Text15 <> 0 Then
    Text16 = Format(CCur(Text12) - CCur(Text14) - CCur(Text15), "##,###.00")
ElseIf Text14 = 0 And Text15 <> 0 Then
    Text16 = Format(CCur(Text12) - CCur(Text15), "##,###.00")
End If
If CCur(Text16) = 0 Then Text16 = ""
End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text16_LostFocus()
If Text16 = "" Then Text16 = 0
If Not IsNumeric(Text16) Then
    Text16.SetFocus
    MsgBox "NOMINAL PEMBAYARAN HARUS ANGKA", vbCritical, "TYPE DATA SALAH"
    Exit Sub
End If
Text16 = Format(Text16, "##,###.00")

Label33 = Format(CCur(Text14) + CCur(Text15) + CCur(Text16), "##,###.00")

If CCur(Text16) + CCur(Text15) + CCur(Text14) > CCur(Text12) Then
    Text16.SetFocus
    MsgBox "TOTAL PEMBAYARAN TIDAK BOLEH MELEBIHI HARGA ON THE ROAD", vbCritical, "TOTAL PEMBAYARAN SALAH"
    Exit Sub
End If

If CCur(Text16) > 0 Then
    KHutang = 1
    Call Hutang_Aktif
    Text19.SetFocus
    Text22 = Tanggal
    Text21 = Tanggal
        SPin = "Select * From P001 order by Kode_Pin"
        Set RPin = RDCO.OpenResultset(SPin, rdOpenDynamic, rdConcurRowVer)
        If RPin.RowCount <> 0 Then
            Label40 = RPin("Kode_Pin")
            Label41 = RPin("Nama_Pin")
        End If
        RPin.Close
        Set RPin = Nothing
        Label42 = Trim(Label40) + "." + Trim(Label30)
ElseIf CCur(Text16) = 0 Then
    KHutang = 0
    Call Hutang_Pasif
    TblSave.SetFocus
End If
End Sub

Private Sub Hutang_Aktif()
    Frame5.Enabled = True
    SSTab2.Tab = 2
    Label40 = ""
    Label41 = ""
    Label42 = ""
    Label43 = Text16
    Text22 = Tanggal
    Text21 = Tanggal
    Text19.BackColor = &HFFFFC0
    Text22.BackColor = &HFFFFC0
    Text21.BackColor = &HFFFFC0
    Text20.BackColor = &HFFFFC0
    Text19.SetFocus
End Sub

Private Sub Hutang_Pasif()
    Frame5.Enabled = False
End Sub

Private Sub Text18_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text18_LostFocus()
If Text18 = "" Then Exit Sub
SSGL = "Select CodeSL, NamaSL From G003 where CodeSL = '" + Trim(Text18) + "'"
Set RSGL = RDCO.OpenResultset(SSGL, rdOpenDynamic, rdConcurRowVer)
If RSGL.RowCount <> 0 Then
    Label35 = RSGL("NamaSL")
Else
    Text18 = ""
    Text18.SetFocus
    MsgBox "KODE SUB GENERAL LEDGER BELUM TERDAFTAR", vbInformation, "KODE GL BELUM TERDAFTAR"
End If
RSGL.Close
Set RSGL = Nothing

    If CCur(Text14) + CCur(Text15) = CCur(Text12) Then
        Text16.Enabled = False
        TblSave.SetFocus
        Bank = 1
        KHutang = 0
        Exit Sub
    ElseIf CCur(Text14) + CCur(Text15) < CCur(Text12) Then
        Text16.Enabled = True
        Text16.SetFocus
        Bank = 1
        KHutang = 0
        Exit Sub
    End If
    
End Sub

Private Sub Text19_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text2_GotFocus()
If CCur(Text2) = 0 Then Text2 = ""
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys vbTab
End If
End Sub

Private Sub Text2_LostFocus()
If Text2 = "" Then Text2 = 0
If Not IsNumeric(Text2) Then
    Text2.SetFocus
    MsgBox "NOMINAL INTENSIF HARUS ANGKA", vbCritical, "TYPE DATA SALAH"
    Exit Sub
End If
Text2 = Format(Text2, "##,###.00")
End Sub

Private Sub Text20_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Text19 = "" Then
        Text19.SetFocus
        MsgBox "DATA CUSTOMER KOSONG", vbCritical, "WARNING"
        Exit Sub
    Else
        SSTab2.Tab = 1
        TblSave.SetFocus
        Exit Sub
    End If
End If
End Sub

Private Sub Text20_LostFocus()
If Text20 = "" Then Text20 = "-"
End Sub

Private Sub Text21_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim jatuh As String
    Dim firstdate As Date
    Dim IntervalType As String
    Dim Number As Integer
    IntervalType = "m"
    firstdate = Text22
    Date = Text21
    jatuh = DateDiff(IntervalType, firstdate, Date)
    JW = jatuh
    SendKeys vbTab
End If
End Sub

Private Sub Text21_LostFocus()
If Text21 = "" Then Exit Sub
If Not IsDate(Text21) Then
    Text21.SetFocus
    Text21 = Tanggal
    MsgBox "TANGGAL MULAI KREDIT HARUS TYPE TANGGAL", vbCritical, "TYPE DATA SALAH"
End If
    If Text21 = Text22 Then
        MsgBox "TANGGAL JATUH TIDAK BOLEH SAMA", vbCritical, "WARNING"
        Text21.SetFocus
    End If
Text21 = Format(Text21, "DD/MM/YYYY")
End Sub

Private Sub Text22_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text22_LostFocus()
If Text22 = "" Then Exit Sub
If Not IsDate(Text22) Then
    Text22.SetFocus
    Text22 = Tanggal
    MsgBox "TANGGAL MULAI KREDIT HARUS TYPE TANGGAL", vbCritical, "TYPE DATA SALAH"
End If
Text22 = Format(Text22, "DD/MM/YYYY")
End Sub

Private Sub Text23_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Text23 = "" Then
        Text23 = "------"
        SendKeys vbTab
    Else
        Text23 = Format(Text23, ">")
        SendKeys vbTab
    End If
End If
End Sub

Private Sub text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Text3 = "" Then
        Text3 = "------"
        SendKeys vbTab
    Else
        Text3 = Format(Text3, ">")
        SendKeys vbTab
    End If
End If
End Sub

Private Sub text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text4_LostFocus()
If Text4 = "" Then
    Text4 = "-"
    Exit Sub
End If

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
    If Text5 = "" Then
        Text5 = "------"
        SendKeys vbTab
    Else
        Text5 = Format(Text5, ">")
        SendKeys vbTab
    End If
End If
End Sub

Private Sub Text7_GotFocus()
If CCur(Text7) = 0 Then Text7 = ""
Text7.BackColor = &HFFFFFF
End Sub

Private Sub text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text7_LostFocus()
If Text7 = "" Then Text7 = 0
If Not IsNumeric(Text7) Then
    Text7.SetFocus
    MsgBox "NOMINAL PEMBAYARAN HARUS ANGKA", vbCritical, "TYPE DATA SALAH"
    Exit Sub
End If
Text7 = Format(Text7, "##,###.00")
Label21 = Format(CCur(Text7) + CCur(Text8) + CCur(Text9) + CCur(Text10) + CCur(Text11), "##,###.00")
End Sub

Private Sub Text8_GotFocus()
If CCur(Text8) = 0 Then Text8 = ""
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text8_LostFocus()
If Text8 = "" Then Text8 = 0
If Not IsNumeric(Text8) Then
    Text8.SetFocus
    MsgBox "NOMINAL PEMBAYARAN HARUS ANGKA", vbCritical, "TYPE DATA SALAH"
    Exit Sub
End If
Text8 = Format(Text8, "##,###.00")
Label21 = Format(CCur(Text7) + CCur(Text8) + CCur(Text9) + CCur(Text10) + CCur(Text11), "##,###.00")
End Sub

Private Sub Text9_GotFocus()
If CCur(Text9) = 0 Then Text9 = ""
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text9_LostFocus()
If Text9 = "" Then Text9 = 0
If Not IsNumeric(Text9) Then
    Text9.SetFocus
    MsgBox "NOMINAL PEMBAYARAN HARUS ANGKA", vbCritical, "TYPE DATA SALAH"
    Exit Sub
End If
Text9 = Format(Text9, "##,###.00")
Label21 = Format(CCur(Text7) + CCur(Text8) + CCur(Text9) + CCur(Text10) + CCur(Text11), "##,###.00")
End Sub

Private Sub Text10_GotFocus()
If CCur(Text10) = 0 Then Text10 = ""
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text10_LostFocus()
If Text10 = "" Then Text10 = 0
If Not IsNumeric(Text10) Then
    Text10.SetFocus
    MsgBox "NOMINAL PEMBAYARAN HARUS ANGKA", vbCritical, "TYPE DATA SALAH"
    Exit Sub
End If
Text10 = Format(Text10, "##,###.00")
Label21 = Format(CCur(Text7) + CCur(Text8) + CCur(Text9) + CCur(Text10) + CCur(Text11), "##,###.00")
End Sub

Private Sub Text11_GotFocus()
If CCur(Text11) = 0 Then Text11 = ""
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Frame1.Enabled = True
    SSTab2.Tab = 1
    SendKeys vbTab
End If
End Sub

Private Sub Text11_LostFocus()
If Text11 = "" Then Text11 = 0
If Not IsNumeric(Text11) Then
    Text11.SetFocus
    MsgBox "NOMINAL PEMBAYARAN HARUS ANGKA", vbCritical, "TYPE DATA SALAH"
    Exit Sub
End If
Text11 = Format(Text11, "##,###.00")
Label21 = Format(CCur(Text7) + CCur(Text8) + CCur(Text9) + CCur(Text10) + CCur(Text11), "##,###.00")

Text14.BackColor = &HFFFFFF
Text15.BackColor = &HFFFFFF
Text16.BackColor = &HFFFFFF

End Sub

Private Sub FrameGLAktif()
Frame6.Enabled = True
Text18.Enabled = True
Text18.BackColor = &HFFFFC0
InfoGL.Enabled = True
Label35.Enabled = True
End Sub

Private Sub FrameGLNonAktif()
Frame6.Enabled = False
Text18.Enabled = False
Text18.BackColor = &HC0C0C0
InfoGL.Enabled = False
Label35.Enabled = False
Text18 = ""
Label35 = ""
End Sub

Private Sub JL004()
SGele = "Select * From JL004"
Set RGele = RDCO.OpenResultset(SGele, rdOpenKeyset, rdConcurRowVer)
RGele.AddNew

    RGele("NAMA") = Trim(Text3)
    RGele("ALAMAT_1") = Trim(Text5)
    RGele("ALAMAT_2") = Trim(Text23)
    RGele("TYPE") = Label27
    RGele("WARNA") = Label24
    RGele("TAHUN") = Label25
    
    RGele("RANGKA") = Label28
    RGele("MESIN") = Label26
    RGele("POSISI") = Label10
    RGele("H_BELI") = CCur(Label23)
    RGele("OTR") = CCur(Text12)
    RGele("BBN") = CCur(Text13)
    RGele("BENSIN") = CCur(Text7)
    RGele("JAKET") = CCur(Text8)
    RGele("KACAB") = CCur(Text9)
    RGele("BROKER") = CCur(Text10)
    
    RGele("DISKON") = CCur(Text11)
    RGele("TUNAI") = CCur(Text14)
    RGele("NON_TUNAI") = CCur(Text15)
    RGele("PIUTANG") = CCur(Text16)
    RGele("INTENSIF") = CCur(Text2)
    RGele("S_DISKON") = 0
    RGele("S_GL") = Trim(Text18)
    RGele("KODE_P") = Label40
    RGele("NO_P") = Label42
    RGele("NO_C") = Trim(Text19)
    RGele("T_MULAI") = Text22
    RGele("T_JATUH") = Text21
    
    RGele("SYARAT") = Trim(Text20)
    
RGele.Update
RGele.Close
Set RGele = Nothing
End Sub

Private Sub SetJual()
SJual = "Select * From M001 where TYPE = '" + Trim(Label27) + "' and WARNA = '" + Trim(Label24) + "' and RANGKA = '" + Trim(Label28) + "' and MESIN = '" + Trim(Label26) + "'"
Set RJual = RDCO.OpenResultset(SJual, rdOpenKeyset, rdConcurRowVer)
RJual.EDIT
    RJual("STS_JUAL") = 1
    RJual("NAMA_PEMBELI") = Trim(Text3)
    RJual("ALAMAT_1") = Trim(Text5)
    RJual("ALAMAT_2") = Trim(Text23)
    RJual("TGL_JUAL") = Tanggal
    
    If KHutang = 1 Then
        RJual("NO_HUTANG") = Trim(Label42)
    Else
        RJual("NO_HUTANG") = 0
    End If
    
    RJual("BBN") = CCur(Text13)
    RJual("BENSIN") = CCur(Text7)
    RJual("JAKET") = CCur(Text8)
    RJual("KACAB") = CCur(Text9)
    RJual("BROKER") = CCur(Text10)
    RJual("DISKON") = CCur(Text11)
    RJual("TUNAI") = CCur(Text14)
    RJual("NON_TUNAI") = CCur(Text15)
    RJual("PIUTANG") = CCur(Text16)
    RJual("INTENSIF") = CCur(Text2)
    
RJual.Update
RJual.Close
Set RJual = Nothing
End Sub

Private Sub JurnalKas()
SDebet = "Select * From C013 where Nama = '" + Trim(Operator) + "'"
Set RDebet = RDCO.OpenResultset(SDebet, rdOpenKeyset, rdConcurRowVer)
    G_DEBET = RDebet("GDebet")
    SDebet2 = "Select * From G003 where CODESL='" + Trim(G_DEBET) + "'"
    Set RDebet2 = RDCO.OpenResultset(SDebet2, rdOpenKeyset, rdConcurRowVer)
    MMUTASID = RDebet2("mutasid") + CCur(Text14)
    SSALDO = RDebet2("saldo") + CCur(Text14)
    RDebet2.EDIT
        RDebet2("mutasid") = CCur(MMUTASID)
        RDebet2("saldo") = CCur(SSALDO)
            SDebet3 = "Select * From G005"
            Set RDebet3 = RDCO.OpenResultset(SDebet3, rdOpenKeyset, rdConcurRowVer)
            RDebet3.AddNew
                RDebet3("codecab") = CodeCab
                RDebet3("codesl") = G_DEBET
                RDebet3("namasl") = RDebet2("NamaSL")
                RDebet3("nobukti") = Trim(Label30)
                RDebet3("keterangan") = "JL.TUNAI." + Label30 + "."
                RDebet3("nominald") = CCur(Text14)
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
RDebet.Close
Set RDebet = Nothing
End Sub

Private Sub JurnalBahan()
SBahan = "Select * From C013 where Nama = '" + Trim(Operator) + "'"
Set RBahan = RDCO.OpenResultset(SBahan, rdOpenKeyset, rdConcurRowVer)
    F_CREDIT = RBahan("FCredit")
    SBahan2 = "Select * From G003 where CODESL='" + Trim(F_CREDIT) + "'"
    Set RBahan2 = RDCO.OpenResultset(SBahan2, rdOpenKeyset, rdConcurRowVer)
    MMUTASIC = RBahan2("mutasic") + CCur(Label23)
    SSALDO = RBahan2("saldo") - CCur(Label23)
    RBahan2.EDIT
        RBahan2("mutasic") = CCur(MMUTASIC)
        RBahan2("saldo") = CCur(SSALDO)
            SBahan3 = "Select * From G005"
            Set RBahan3 = RDCO.OpenResultset(SBahan3, rdOpenKeyset, rdConcurRowVer)
            RBahan3.AddNew
                RBahan3("codecab") = CodeCab
                RBahan3("codesl") = F_DEBET
                RBahan3("namasl") = RBahan2("NamaSL")
                RBahan3("nobukti") = Label30
                RBahan3("keterangan") = "JL.TUNAI." + Label30 + "."
                RBahan3("nominald") = 0
                RBahan3("nominalc") = CCur(Label23)
                RBahan3("saldo") = SSALDO
                RBahan3("tanggal") = Tanggal
                RBahan3("jam") = Time
                RBahan3("usercode") = Operator
            RBahan3.Update
            RBahan3.Close
            Set RBahan3 = Nothing
    RBahan2.Update
    RBahan2.Close
    Set RBahan2 = Nothing
RBahan.Close
Set RBahan = Nothing
End Sub

Private Sub JurnalBiaya()
SBahan = "Select * From B001 where KODE_IND = '" + Trim(151) + "' "
Set RBahan = RDCO.OpenResultset(SBahan, rdOpenKeyset, rdConcurRowVer)
    GL_BBN = RBahan("SGL_BBN")
    SBahan2 = "Select * From G003 where CODESL='" + Trim(GL_BBN) + "'"
    Set RBahan2 = RDCO.OpenResultset(SBahan2, rdOpenKeyset, rdConcurRowVer)
    MMUTASID = RBahan2("mutasid") + CCur(Text13)
    SSALDO = RBahan2("saldo") + CCur(Text13)
    RBahan2.EDIT
        RBahan2("mutasid") = CCur(MMUTASID)
        RBahan2("saldo") = CCur(SSALDO)
            SBahan3 = "Select * From G005"
            Set RBahan3 = RDCO.OpenResultset(SBahan3, rdOpenKeyset, rdConcurRowVer)
            RBahan3.AddNew
                RBahan3("codecab") = CodeCab
                RBahan3("codesl") = GL_BBN
                RBahan3("namasl") = RBahan2("NamaSL")
                RBahan3("nobukti") = Label30
                RBahan3("keterangan") = "JL.BBN." + Label30 + "."
                RBahan3("nominald") = CCur(Text13)
                RBahan3("nominalc") = 0
                RBahan3("saldo") = SSALDO
                RBahan3("tanggal") = Tanggal
                RBahan3("jam") = Time
                RBahan3("usercode") = Operator
            RBahan3.Update
            RBahan3.Close
            Set RBahan3 = Nothing
    RBahan2.Update
    RBahan2.Close
    Set RBahan2 = Nothing
RBahan.Close
Set RBahan = Nothing
End Sub

Private Sub JurnalBiaya2()
SBahannya = "Select * From B001 where KODE_IND = '" + Trim(151) + "' "
Set RBahannya = RDCO.OpenResultset(SBahannya, rdOpenKeyset, rdConcurRowVer)
    GL_BIAYA = RBahannya("SGL_BIAYA")
    SBahannya2 = "Select * From G003 where CODESL='" + Trim(GL_BIAYA) + "'"
    Set RBahannya2 = RDCO.OpenResultset(SBahannya2, rdOpenKeyset, rdConcurRowVer)
    MMUTASID = RBahannya2("mutasid") + CCur(Label21)
    SSALDO = RBahannya2("saldo") + CCur(Label21)
    RBahannya2.EDIT
        RBahannya2("mutasid") = CCur(MMUTASID)
        RBahannya2("saldo") = CCur(SSALDO)
            SBahannya3 = "Select * From G005"
            Set RBahannya3 = RDCO.OpenResultset(SBahannya3, rdOpenKeyset, rdConcurRowVer)
            RBahannya3.AddNew
                RBahannya3("codecab") = CodeCab
                RBahannya3("codesl") = GL_BIAYA
                RBahannya3("namasl") = RBahannya2("NamaSL")
                RBahannya3("nobukti") = Label30
                RBahannya3("keterangan") = "JL.BIAYA." + Label30 + "."
                RBahannya3("nominald") = CCur(Label21)
                RBahannya3("nominalc") = 0
                RBahannya3("saldo") = SSALDO
                RBahannya3("tanggal") = Tanggal
                RBahannya3("jam") = Time
                RBahannya3("usercode") = Operator
            RBahannya3.Update
            RBahannya3.Close
            Set RBahannya3 = Nothing
    RBahannya2.Update
    RBahannya2.Close
    Set RBahannya2 = Nothing
RBahannya.Close
Set RBahannya = Nothing
End Sub

Private Sub JurnalPendapatan()
SPDPT = "Select * From B001 where KODE_IND = '" + Trim(151) + "' "
Set RPDPT = RDCO.OpenResultset(SPDPT, rdOpenKeyset, rdConcurRowVer)
    GL_PDPT = RPDPT("SGL_PDPT")
    SPDPT2 = "Select * From G003 where CODESL='" + Trim(GL_PDPT) + "'"
    Set RPDPT2 = RDCO.OpenResultset(SPDPT2, rdOpenKeyset, rdConcurRowVer)
    MMUTASIC = RPDPT2("mutasic") + (CCur(Text12) - CCur(Label23) + CCur(Text2))
    SSALDO = RPDPT2("saldo") + (CCur(Text12) - CCur(Label23) + CCur(Text2))
    RPDPT2.EDIT
        RPDPT2("mutasic") = CCur(MMUTASIC)
        RPDPT2("saldo") = CCur(SSALDO)
            SPDPT3 = "Select * From G005"
            Set RPDPT3 = RDCO.OpenResultset(SPDPT3, rdOpenKeyset, rdConcurRowVer)
            RPDPT3.AddNew
                RPDPT3("codecab") = CodeCab
                RPDPT3("codesl") = GL_PDPT
                RPDPT3("namasl") = RPDPT2("NamaSL")
                RPDPT3("nobukti") = Label30
                RPDPT3("keterangan") = "JL.PDPT." + Label30 + "."
                RPDPT3("nominald") = 0
                RPDPT3("nominalc") = CCur(Text12) - CCur(Label23) + CCur(Text2)
                RPDPT3("saldo") = SSALDO
                RPDPT3("tanggal") = Tanggal
                RPDPT3("jam") = Time
                RPDPT3("usercode") = Operator
            RPDPT3.Update
            RPDPT3.Close
            Set RPDPT3 = Nothing
    RPDPT2.Update
    RPDPT2.Close
    Set RPDPT2 = Nothing
RPDPT.Close
Set RPDPT = Nothing
End Sub

Private Sub JurnalKas2()
SCBiaya = "Select * From C013 where Nama = '" + Trim(Operator) + "'"
Set RCBiaya = RDCO.OpenResultset(SCBiaya, rdOpenKeyset, rdConcurRowVer)
    G_CREDIT = RCBiaya("GCredit")
    SCBiaya2 = "Select * From G003 where CODESL='" + Trim(G_CREDIT) + "'"
    Set RCBiaya2 = RDCO.OpenResultset(SCBiaya2, rdOpenKeyset, rdConcurRowVer)
    MMUTASIC = RCBiaya2("mutasic") + (CCur(Label21) + CCur(Text13))
    SSALDO = RCBiaya2("saldo") - (CCur(Label21) + CCur(Text13))
    RCBiaya2.EDIT
        RCBiaya2("mutasic") = CCur(MMUTASIC)
        RCBiaya2("saldo") = CCur(SSALDO)
            SCBiaya3 = "Select * From G005"
            Set RCBiaya3 = RDCO.OpenResultset(SCBiaya3, rdOpenKeyset, rdConcurRowVer)
            RCBiaya3.AddNew
                RCBiaya3("codecab") = CodeCab
                RCBiaya3("codesl") = G_CREDIT
                RCBiaya3("namasl") = RCBiaya2("NamaSL")
                RCBiaya3("nobukti") = Label30
                RCBiaya3("keterangan") = "JL.BBN/BIAYA." + Label30 + "."
                RCBiaya3("nominald") = 0
                RCBiaya3("nominalc") = CCur(Label21) + CCur(Text13)
                RCBiaya3("saldo") = SSALDO
                RCBiaya3("tanggal") = Tanggal
                RCBiaya3("jam") = Time
                RCBiaya3("usercode") = Operator
            RCBiaya3.Update
            RCBiaya3.Close
            Set RCBiaya3 = Nothing
    RCBiaya2.Update
    RCBiaya2.Close
    Set RCBiaya2 = Nothing
RCBiaya.Close
Set RCBiaya = Nothing
End Sub

Private Sub JurnalKas3()
SDBank2 = "Select * From G003 where codesl = '" + Trim(Text18) + "'"
Set RDBank2 = RDCO.OpenResultset(SDBank2, rdOpenKeyset, rdConcurRowVer)
MMUTASID = RDBank2("mutasid") + CCur(Text15)
SSALDO = RDBank2("saldo") + CCur(Text15)
RDBank2.EDIT
    RDBank2("mutasid") = CCur(MMUTASID)
    RDBank2("saldo") = CCur(SSALDO)
        SDBank3 = "Select * From G005"
        Set RDBank3 = RDCO.OpenResultset(SDBank3, rdOpenKeyset, rdConcurRowVer)
        RDBank3.AddNew
            RDBank3("codecab") = CodeCab
            RDBank3("codesl") = Trim(Text18)
            RDBank3("namasl") = RDBank2("NamaSL")
            RDBank3("nobukti") = Label30
            RDBank3("keterangan") = "JL.NON TUNAI BANK." + Label30 + "."
            RDBank3("nominald") = CCur(Text15)
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
SDBank = "Select * From P001 where KODE_PIN = '" + Trim(Label40) + "'"
Set RDBank = RDCO.OpenResultset(SDBank, rdOpenKeyset, rdConcurRowVer)
    GPIN = RDBank("SGL_PIN")
    SDBank2 = "Select * From G003 where CODESL='" + Trim(GPIN) + "'"
    Set RDBank2 = RDCO.OpenResultset(SDBank2, rdOpenKeyset, rdConcurRowVer)
    MMUTASID = RDBank2("mutasid") + CCur(Text16)
    SSALDO = RDBank2("saldo") + CCur(Text16)
    RDBank2.EDIT
        RDBank2("mutasid") = CCur(MMUTASID)
        RDBank2("saldo") = CCur(SSALDO)
            SDBank3 = "Select * From G005"
            Set RDBank3 = RDCO.OpenResultset(SDBank3, rdOpenKeyset, rdConcurRowVer)
            RDBank3.AddNew
                RDBank3("codecab") = CodeCab
                RDBank3("codesl") = GPIN
                RDBank3("namasl") = RDBank2("NamaSL")
                RDBank3("nobukti") = Label30
                RDBank3("keterangan") = "JL." + Label42 + "." + Label30 + "."
                RDBank3("nominald") = CCur(Text16)
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
RDBank.Close
Set RDBank = Nothing

SDBank = "Select * From P001 where KODE_PIN = '" + Trim(Label40) + "'"
Set RDBank = RDCO.OpenResultset(SDBank, rdOpenKeyset, rdConcurRowVer)
    GPIN = RDBank("SGL_INTS")
    SDBank2 = "Select * From G003 where CODESL='" + Trim(GPIN) + "'"
    Set RDBank2 = RDCO.OpenResultset(SDBank2, rdOpenKeyset, rdConcurRowVer)
    MMUTASID = RDBank2("mutasid") + CCur(Text2)
    SSALDO = RDBank2("saldo") + CCur(Text2)
    RDBank2.EDIT
        RDBank2("mutasid") = CCur(MMUTASID)
        RDBank2("saldo") = CCur(SSALDO)
            SDBank3 = "Select * From G005"
            Set RDBank3 = RDCO.OpenResultset(SDBank3, rdOpenKeyset, rdConcurRowVer)
            RDBank3.AddNew
                RDBank3("codecab") = CodeCab
                RDBank3("codesl") = GPIN
                RDBank3("namasl") = RDBank2("NamaSL")
                RDBank3("nobukti") = Label30
                RDBank3("keterangan") = "INTS." + Label42 + "." + Label30 + "."
                RDBank3("nominald") = CCur(Text2)
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
RDBank.Close
Set RDBank = Nothing
End Sub

Private Sub Hutang()
Dim jatuh As String
Dim firstdate As Date
Dim IntervalType As String
Dim Number As Integer
IntervalType = "m"
firstdate = Text22
Date = Text21
jatuh = DateDiff(IntervalType, firstdate, Date)
JW = jatuh

SOyen = "Select * From P002"
Set ROyen = RDCO.OpenResultset(SOyen, rdOpenKeyset, rdConcurRowVer)
ROyen.AddNew
    ROyen("KODE_PIN") = Trim(Label40)
    ROyen("NOMOR_PIN") = Trim(Label42)
    ROyen("NOMOR_NAS") = Trim(Text19)
    ROyen("NAMA_NAS") = Trim(Label36)
    ROyen("PLAFON") = CCur(Label43)
    ROyen("BAKI_DEBET") = CCur(Label43)
    ROyen("INTENSIF") = CCur(Text2)
    ROyen("TGL_MULAI") = Text22
    ROyen("TGL_JATUH") = Text21
    ROyen("JWAKTU") = JW
    ROyen("SYARAT_BYR") = Trim(Text20)
    ROyen("STATUS") = 0
    ROyen("TANGGAL") = Tanggal
    ROyen("USER_CODE") = Operator
    
ROyen.Update
ROyen.Close
Set ROyen = Nothing

SNovi = "Select * From P003"
Set RNovi = RDCO.OpenResultset(SNovi, rdOpenKeyset, rdConcurRowVer)
RNovi.AddNew
    RNovi("NOMOR_PIN") = Trim(Label42)
    RNovi("NAMA_NAS") = Trim(Label36)
    RNovi("NO_BUKTI") = Label30
    RNovi("KETERANGAN") = "PENCAIRAN" + Label40 + "." + Label42
    RNovi("POKOK") = 0
    RNovi("BUNGA") = 0
    RNovi("DENDA") = 0
    RNovi("BAKI_DEBET") = CCur(Label43)
    RNovi("TANGGAL") = Tanggal
    RNovi("USER_CODE") = Operator
    
RNovi.Update
RNovi.Close
Set RNovi = Nothing

End Sub

Private Sub LabaRugi()
SUhAh = "Select * From LabaRugi"
Set RUhAh = RDCO.OpenResultset(SUhAh, rdOpenDynamic, rdConcurRowVer)
    SaldoD = CCur(RUhAh("sumofmutasid"))
    SaldoC = CCur(RUhAh("sumofmutasic"))
    
    SSave5 = "Select * From G003 where POSISI = 'L'"
    Set RSave5 = RDCO.OpenResultset(SSave5, rdOpenDynamic, rdConcurRowVer)
    Saldo = RSave5("saldoawal")
    RSave5.EDIT
        RSave5("mutasid") = SaldoD
        RSave5("mutasic") = SaldoC
        RSave5("saldo") = CCur(RSave5("SaldoAwal")) - CCur(RSave5("mutasid")) + CCur(RSave5("mutasic"))
    RSave5.Update
    RSave5.Close
    Set RSave5 = Nothing

RUhAh.Close
Set RUhAh = Nothing
End Sub

