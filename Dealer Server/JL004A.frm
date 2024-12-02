VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form JL004A 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EDIT NO STNK / BPKB"
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   8895
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000000FF&
      Height          =   360
      Left            =   8055
      ScaleHeight     =   300
      ScaleWidth      =   315
      TabIndex        =   79
      Top             =   4797
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0C0FF&
      Height          =   360
      Left            =   1395
      TabIndex        =   1
      Text            =   "Text4"
      Top             =   4365
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
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
      Left            =   120
      TabIndex        =   2
      Top             =   4860
      Width           =   3480
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0C0FF&
      Height          =   360
      Left            =   1395
      TabIndex        =   0
      Text            =   "Text2"
      Top             =   3960
      Width           =   2175
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
      Left            =   7815
      TabIndex        =   72
      Top             =   8325
      Width           =   960
   End
   Begin VB.TextBox Text13 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
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
      Left            =   5025
      TabIndex        =   22
      Text            =   "Text13"
      Top             =   4797
      Width           =   3000
   End
   Begin VB.TextBox Text12 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Enabled         =   0   'False
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
      Left            =   5025
      TabIndex        =   21
      Text            =   "Text12"
      Top             =   4362
      Width           =   3000
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
      Height          =   2040
      Left            =   120
      TabIndex        =   10
      Top             =   1842
      Width           =   3480
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
         TabIndex        =   20
         Top             =   330
         Width           =   1050
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
         TabIndex        =   19
         Top             =   705
         Width           =   1050
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
         TabIndex        =   18
         Top             =   1020
         Width           =   1050
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
         TabIndex        =   17
         Top             =   1335
         Width           =   1050
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
         TabIndex        =   16
         Top             =   1650
         Width           =   1050
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Left            =   1260
         TabIndex        =   15
         Top             =   660
         Width           =   1875
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Left            =   1260
         TabIndex        =   14
         Top             =   975
         Width           =   825
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Left            =   1260
         TabIndex        =   13
         Top             =   1605
         Width           =   2085
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Left            =   1260
         TabIndex        =   12
         Top             =   345
         Width           =   1875
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Left            =   1260
         TabIndex        =   11
         Top             =   1290
         Width           =   2085
      End
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
      TabIndex        =   7
      Top             =   1842
      Width           =   4950
      Begin VB.TextBox Text5 
         BackColor       =   &H00C0C0FF&
         Height          =   990
         Left            =   1260
         MultiLine       =   -1  'True
         TabIndex        =   4
         Text            =   "JL004A.frx":0000
         Top             =   735
         Width           =   3525
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0C0FF&
         Height          =   360
         Left            =   1260
         TabIndex        =   3
         Text            =   "Text3"
         Top             =   315
         Width           =   3525
      End
      Begin VB.TextBox Text23 
         BackColor       =   &H00C0C0FF&
         Height          =   360
         Left            =   1260
         TabIndex        =   5
         Text            =   "Text23"
         Top             =   1785
         Width           =   3525
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
         TabIndex        =   9
         Top             =   390
         Width           =   1140
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
         TabIndex        =   8
         Top             =   1125
         Width           =   1140
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   360
      Left            =   6570
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   124
      Width           =   1965
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   2745
      Left            =   120
      TabIndex        =   23
      Top             =   5412
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   4842
      _Version        =   393216
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
      TabPicture(0)   =   "JL004A.frx":0006
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "PEMBAYARAN"
      TabPicture(1)   =   "JL004A.frx":0022
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "PIUTANG"
      TabPicture(2)   =   "JL004A.frx":003E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame5 
         Height          =   2355
         Left            =   -74940
         TabIndex        =   43
         Top             =   315
         Width           =   8535
         Begin VB.TextBox Text22 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
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
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   5775
            TabIndex        =   47
            Text            =   "Text22"
            Top             =   210
            Width           =   1710
         End
         Begin VB.TextBox Text21 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
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
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   5775
            TabIndex        =   46
            Text            =   "Text21"
            Top             =   615
            Width           =   1710
         End
         Begin VB.TextBox Text20 
            BackColor       =   &H00FFFFC0&
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
            Height          =   360
            Left            =   1470
            TabIndex        =   45
            Text            =   "Text20"
            Top             =   1920
            Width           =   6510
         End
         Begin VB.TextBox Text19 
            BackColor       =   &H00FFFFC0&
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
            Height          =   360
            Left            =   1470
            MaxLength       =   12
            TabIndex        =   44
            Text            =   "Text19"
            Top             =   1080
            Width           =   1410
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
            TabIndex        =   58
            Top             =   645
            Width           =   1875
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
            TabIndex        =   57
            Top             =   630
            Width           =   1125
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
            TabIndex        =   56
            Top             =   1110
            Width           =   1125
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
            TabIndex        =   55
            Top             =   1530
            Width           =   1260
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
            TabIndex        =   54
            Top             =   240
            Width           =   1335
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
            TabIndex        =   53
            Top             =   645
            Width           =   1245
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
            TabIndex        =   52
            Top             =   1065
            Width           =   1155
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
            TabIndex        =   50
            Top             =   1950
            Width           =   1020
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
            TabIndex        =   49
            Top             =   1530
            Width           =   2610
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
            TabIndex        =   48
            Top             =   450
            Width           =   915
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2355
         Left            =   -74940
         TabIndex        =   35
         Top             =   315
         Width           =   8535
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
            TabIndex        =   36
            Top             =   210
            Width           =   3900
            Begin VB.TextBox Text16 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Enabled         =   0   'False
               Height          =   360
               Left            =   1995
               TabIndex        =   39
               Text            =   "Text16"
               Top             =   1155
               Width           =   1860
            End
            Begin VB.TextBox Text15 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Enabled         =   0   'False
               Height          =   360
               Left            =   1980
               TabIndex        =   38
               Text            =   "Text15"
               Top             =   735
               Width           =   1860
            End
            Begin VB.TextBox Text14 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Enabled         =   0   'False
               Height          =   360
               Left            =   1980
               TabIndex        =   37
               Text            =   "Text14"
               Top             =   315
               Width           =   1860
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
               TabIndex        =   42
               Top             =   765
               Width           =   1755
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
               TabIndex        =   41
               Top             =   1185
               Width           =   1755
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
               TabIndex        =   40
               Top             =   345
               Width           =   1230
            End
         End
      End
      Begin VB.Frame Frame7 
         Height          =   2355
         Left            =   60
         TabIndex        =   24
         Top             =   315
         Width           =   8535
         Begin VB.PictureBox Picture3 
            BackColor       =   &H000000FF&
            Height          =   285
            Left            =   7605
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   78
            Top             =   465
            Width           =   285
         End
         Begin VB.PictureBox Picture2 
            BackColor       =   &H000000FF&
            Height          =   285
            Left            =   3735
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   77
            Top             =   1515
            Width           =   285
         End
         Begin VB.TextBox Text7 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
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
            Left            =   1620
            TabIndex        =   29
            Text            =   "Text7"
            Top             =   465
            Width           =   2055
         End
         Begin VB.TextBox Text8 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
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
            Left            =   1620
            TabIndex        =   28
            Text            =   "Text8"
            Top             =   990
            Width           =   2055
         End
         Begin VB.TextBox Text9 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
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
            Left            =   1620
            TabIndex        =   27
            Text            =   "Text9"
            Top             =   1515
            Width           =   2055
         End
         Begin VB.TextBox Text10 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
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
            Left            =   5505
            TabIndex        =   26
            Text            =   "Text10"
            Top             =   465
            Width           =   2055
         End
         Begin VB.TextBox Text11 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
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
            Left            =   5505
            TabIndex        =   25
            Text            =   "Text11"
            Top             =   990
            Width           =   2055
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
            TabIndex        =   34
            Top             =   495
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
            TabIndex        =   33
            Top             =   1020
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
            TabIndex        =   32
            Top             =   1545
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
            TabIndex        =   31
            Top             =   495
            Width           =   1140
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
            TabIndex        =   30
            Top             =   1020
            Width           =   1140
         End
      End
   End
   Begin VB.Label Label34 
      Caption         =   "LABA   Rp."
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
      Left            =   120
      TabIndex        =   76
      Top             =   8370
      Width           =   1020
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "LABA"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1215
      TabIndex        =   75
      Top             =   8355
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "NO STNK"
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
      Left            =   225
      TabIndex        =   74
      Top             =   4035
      Width           =   1140
   End
   Begin VB.Label Label8 
      Caption         =   "NO BPKB"
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
      Left            =   225
      TabIndex        =   73
      Top             =   4440
      Width           =   1140
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
      TabIndex        =   71
      Top             =   495
      Width           =   3930
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
      TabIndex        =   70
      Top             =   522
      Width           =   1980
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
      TabIndex        =   69
      Top             =   552
      Width           =   2055
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
      TabIndex        =   68
      Top             =   4872
      Width           =   720
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
      TabIndex        =   67
      Top             =   4452
      Width           =   720
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
      TabIndex        =   66
      Top             =   867
      Width           =   1455
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
      TabIndex        =   65
      Top             =   897
      Width           =   2055
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
      TabIndex        =   64
      Top             =   1317
      Width           =   2055
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
      TabIndex        =   63
      Top             =   1287
      Width           =   1875
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
      TabIndex        =   61
      Top             =   1317
      Width           =   2055
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
      TabIndex        =   60
      Top             =   162
      Width           =   3135
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
      TabIndex        =   59
      Top             =   132
      Width           =   1860
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
      TabIndex        =   62
      Top             =   1287
      Width           =   1875
   End
End
Attribute VB_Name = "JL004A"
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

Private Sub Command1_Click()
If Text2 = "" Or Text4 = "" Or Text3 = "" Or Text5 = "" Or Text23 = "" Then
    MsgBox "DATA TIDAK BOLEH KOSONG", vbCritical, "WARNING"
    Text1.SetFocus
    Exit Sub
End If

        SSave = "Select * From M001 where No_Fak = '" + Trim(Label30) + "'"
        Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
        RSave.EDIT
            RSave("NO_STNK") = Trim(Text2)
            RSave("NO_BPKB") = Trim(Text4)
            RSave("NAMA_PEMBELI") = Trim(Text3)
            RSave("ALAMAT_1") = Trim(Text5)
            RSave("ALAMAT_2") = Trim(Text23)
        RSave.Update
        RSave.Close
        Set RSave = Nothing

Unload Me
JL003A.Show
End Sub

Private Sub Command2_Click()
Unload Me
JL003A.Show
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=DEALER", rdDriverNoPrompt, False, CN)

SSTab2.Tab = 0
Call CariData
Call CariData2

Label21 = Format(CCur(Text7) + CCur(Text8) + CCur(Text9) + CCur(Text10) + CCur(Text11), "##,###.00")
Label33 = Format(CCur(Label51) - CCur(Label21), "##,###.00")
Label9 = Format(CCur(Label51) - CCur(Label23) - CCur(Label21), "##,###.00")

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
    
    Text3 = Format(RToket("NAMA_PEMBELI"), ">")
    Text5 = Format(RToket("ALAMAT_1"), ">")
    Text23 = Format(RToket("ALAMAT_2"), ">")
    
    Text7 = Format(RToket("BENSIN"), "##,###.00")
    Text8 = Format(RToket("JAKET"), "##,###.00")
    Text9 = Format(RToket("KACAB"), "##,###.00")
    Text10 = Format(RToket("BROKER"), "##,###.00")
    Text11 = Format(RToket("DISKON"), "##,###.00")
    Text13 = Format(RToket("BBN"), "##,###.00")
    
    Text14 = Format(RToket("TUNAI"), "##,###.00")
    Text15 = Format(RToket("NON_TUNAI"), "##,###.00")
    Text16 = Format(RToket("PIUTANG"), "##,###.00")
    
    Label42 = Format(RToket("NO_HUTANG"), ">")
    Text1 = RToket("TGL_JUAL")
    
    Text2 = Format(RToket("NO_STNK"), ">")
    Text4 = Format(RToket("NO_BPKB"), ">")
    
End If

If Label10 = CodeCab Then
    Label10 = "POSISI -->> " + N_CCAB
End If

RToket.Close
Set RToket = Nothing

End Sub

Private Sub CariData2()
SCari = "Select * From P002 where NOMOR_PIN = '" + Trim(Label42) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdOpenKeyset)
If RCari.RowCount <> 0 Then
    Text19 = RCari("NOMOR_NAS")
    Label36 = Format(RCari("NAMA_NAS"), ">")
    Text20 = Format(RCari("SYARAT_BYR"), ">")
    Text22 = RCari("TGL_MULAI")
    Text21 = RCari("TGL_JATUH")
    Label43 = Format(RCari("PLAFON"), "##,###.00")
Else
    Label42 = ""
    Text19 = ""
    Label36 = ""
    Text20 = ""
    Text22 = ""
    Text21 = ""
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub Picture1_DblClick()
Dim Tanya

Tanya = MsgBox("YAKIN PROSES TRANSAKSI BBN ?", vbOKCancel, "TRANSAKSI BBN")
If Tanya = vbCancel Then Exit Sub

SCari = "Select * from B001 where KODE_IND='151'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    G_DEBET = RCari("SGL_BBN")
    G_CREDIT = RCari("SGL_KAS")
End If
RCari.Close
Set RCari = Nothing

STS_Nama = "INPUT BIAYA BBN"
Call Biaya

End Sub

Private Sub Picture2_DblClick()
Dim Tanya

Tanya = MsgBox("YAKIN PROSES TRANSAKSI KACAB ?", vbOKCancel, "TRANSAKSI KACAB")
If Tanya = vbCancel Then Exit Sub

SCari = "Select * from B001 where KODE_IND='151'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    G_DEBET = 9006204
    G_CREDIT = RCari("SGL_KAS")
End If
RCari.Close
Set RCari = Nothing

STS_Nama = "INPUT BIAYA KACAB"
Call Biaya

End Sub

Private Sub Picture3_DblClick()
Dim Tanya

Tanya = MsgBox("YAKIN PROSES TRANSAKSI BROKER ?", vbOKCancel, "TRANSAKSI BROKER")
If Tanya = vbCancel Then Exit Sub

SCari = "Select * from B001 where KODE_IND='151'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    G_DEBET = RCari("SGL_BIAYA")
    G_CREDIT = RCari("SGL_KAS")
End If
RCari.Close
Set RCari = Nothing

STS_Nama = "INPUT BIAYA BROKER"
Call Biaya

End Sub

Private Sub Biaya()
STS_Rangka = Label28
STS_Mesin = Label26
STS_Biaya = 1

Unload Me
G003.Show 1
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text2 = Format(Text2, ">")
    SendKeys vbTab
End If
End Sub

Private Sub text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text4 = Format(Text4, ">")
    SendKeys vbTab
End If
End Sub

Private Sub text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text3 = Format(Text3, ">")
SendKeys vbTab
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text5 = Format(Text5, ">")
    SendKeys vbTab
End If
End Sub

Private Sub Text23_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text23 = Format(Text23, ">")
    Command1.SetFocus
End If
End Sub
