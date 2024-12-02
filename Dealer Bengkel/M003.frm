VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form M003 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TABEL DATA KENDARAAN"
   ClientHeight    =   10050
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14220
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10050
   ScaleWidth      =   14220
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   6240
      Top             =   9480
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "M003.frx":0000
      Left            =   10740
      List            =   "M003.frx":0002
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   720
      Width           =   3135
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "M003.frx":0004
      Left            =   10650
      List            =   "M003.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   3255
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1845
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   2715
      Width           =   2520
   End
   Begin VB.CommandButton Command3 
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
      Left            =   12090
      TabIndex        =   35
      Top             =   9510
      Width           =   2010
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "M003.frx":0008
      Left            =   5535
      List            =   "M003.frx":000A
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   720
      Width           =   3135
   End
   Begin VB.TextBox Text12 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7905
      TabIndex        =   28
      Text            =   "Text12"
      Top             =   8820
      Width           =   6195
   End
   Begin VB.CommandButton Command2 
      Caption         =   "HAPUS"
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
      Left            =   105
      TabIndex        =   27
      Top             =   9510
      Width           =   2010
   End
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   10650
      TabIndex        =   10
      Text            =   "Text10"
      Top             =   2520
      Width           =   3135
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   10650
      TabIndex        =   9
      Text            =   "Text9"
      Top             =   1995
      Width           =   3135
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   10650
      TabIndex        =   8
      Text            =   "Text8"
      Top             =   1470
      Width           =   3135
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2265
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   1455
      Width           =   2850
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2265
      MaxLength       =   65
      TabIndex        =   5
      Text            =   "Text5"
      Top             =   1995
      Width           =   6210
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1440
      TabIndex        =   1
      Text            =   "Text4"
      Top             =   720
      Width           =   1380
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1440
      TabIndex        =   0
      Text            =   "Text6"
      Top             =   210
      Width           =   2460
   End
   Begin VB.CommandButton Command1 
      Caption         =   "EDIT"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6660
      TabIndex        =   13
      Top             =   3375
      Width           =   2010
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
      Height          =   420
      Left            =   225
      TabIndex        =   12
      Top             =   3375
      Width           =   2010
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   4875
      Left            =   105
      TabIndex        =   14
      Top             =   3885
      Width           =   13995
      _ExtentX        =   24686
      _ExtentY        =   8599
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   65280
      ForeColorFixed  =   0
      BackColorBkg    =   16777152
      GridColor       =   8421504
      FocusRect       =   0
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
   Begin VB.TextBox Text11 
      Height          =   405
      Left            =   2265
      TabIndex        =   26
      Text            =   "Text11"
      Top             =   4470
      Width           =   1125
   End
   Begin VB.Frame Frame1 
      Height          =   510
      Left            =   105
      TabIndex        =   29
      Top             =   8730
      Width           =   7590
      Begin VB.OptionButton Option1 
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
         TabIndex        =   34
         Top             =   195
         Width           =   1065
      End
      Begin VB.OptionButton Option2 
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
         Left            =   1665
         TabIndex        =   33
         Top             =   195
         Width           =   1065
      End
      Begin VB.OptionButton Option3 
         Caption         =   "RANGKA"
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
         Left            =   3240
         TabIndex        =   32
         Top             =   195
         Width           =   1065
      End
      Begin VB.OptionButton Option4 
         Caption         =   "MESIN"
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
         Left            =   4815
         TabIndex        =   31
         Top             =   195
         Width           =   1065
      End
      Begin VB.OptionButton Option5 
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
         Left            =   6390
         TabIndex        =   30
         Top             =   195
         Width           =   1065
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6045
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   2715
      Width           =   2520
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "(dd/mm/yyyy)"
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
      Left            =   2835
      TabIndex        =   39
      Top             =   825
      Width           =   1050
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "HARGA POKOK, TYPE KENDARAAN DAN NO FAKTUR TIDAK DAPAT DIEDIT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   300
      Left            =   2280
      TabIndex        =   38
      Top             =   9570
      Width           =   9690
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "MUTASI"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9195
      TabIndex        =   37
      Top             =   3345
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FF00&
      Caption         =   "OFF ROAD"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   300
      TabIndex        =   36
      Top             =   2790
      Width           =   1470
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFC0C0&
      Caption         =   "NO MESIN"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9195
      TabIndex        =   25
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFC0C0&
      Caption         =   "NO RANGKA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9195
      TabIndex        =   24
      Top             =   2115
      Width           =   1455
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFC0C0&
      Caption         =   "TAHUN"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9195
      TabIndex        =   23
      Top             =   1590
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFC0FF&
      Caption         =   "TYPE                                                               WARNA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4245
      TabIndex        =   22
      Top             =   750
      Width           =   9540
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000FF00&
      Caption         =   "OTR"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5220
      TabIndex        =   21
      Top             =   2790
      Width           =   735
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080C0FF&
      Caption         =   "HARGA POKOK  Rp."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   495
      TabIndex        =   20
      Top             =   1530
      Width           =   1890
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080C0FF&
      Caption         =   "KETERANGAN"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   495
      TabIndex        =   19
      Top             =   2070
      Width           =   1365
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Label10"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   9570
      TabIndex        =   18
      Top             =   105
      Width           =   2055
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFC0&
      Caption         =   "TGL BELI"
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
      Left            =   360
      TabIndex        =   17
      Top             =   818
      Width           =   1365
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "NO. FAKTUR"
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
      Left            =   360
      TabIndex        =   16
      Top             =   300
      Width           =   1140
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   11580
      TabIndex        =   15
      Top             =   105
      Width           =   2475
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      Height          =   1155
      Left            =   195
      Shape           =   4  'Rounded Rectangle
      Top             =   105
      Width           =   3855
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      Height          =   1185
      Left            =   195
      Top             =   1350
      Width           =   8475
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      Height          =   1695
      Left            =   8985
      Top             =   1365
      Width           =   5070
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      Height          =   645
      Left            =   8985
      Shape           =   4  'Rounded Rectangle
      Top             =   3150
      Width           =   5070
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      Height          =   660
      Left            =   210
      Top             =   2625
      Width           =   8475
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFC0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   645
      Left            =   4170
      Top             =   585
      Width           =   9870
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "KENDARAAN"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4170
      TabIndex        =   40
      Top             =   225
      Width           =   3300
   End
End
Attribute VB_Name = "M003"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private RSave2, RSave As rdoResultset
Private SSave2, SSave As String

Private RPSN, RMSG, RCari10, RType, RType2, RSusuSore, RCari, RCari2, RCari3, RCari4 As rdoResultset
Private SPSN, SMSG, SCari10, SType, SType2, SSusuSore, SCari, SCari2, SCari3, SCari4 As String

Private QW, EDIT, NoNo, JJJ, SGLSEDIA, SGLCREDIT, SGLMODAL, HBLAMA

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys vbTab
End If
End Sub

Private Sub combo2_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF1
        Text10.SetFocus
End Select
End Sub

Private Sub Command1_Click()
If Text1 = "" Or Text8 = "" Or Text9 = "" Or Text10 = "" Or Combo3 = "" Or Combo2 = "" Or Text2 = "" Then
    MsgBox "MASIH ADA DATA YANG KOSONG !", vbCritical, "WARNING"
    Combo2.SetFocus
    Exit Sub
End If

        SSave = "Select * From M001 where No_Fak = '" + Trim(Text6) + "'"
        Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
        RSave.EDIT
            RSave("TYPE") = Trim(Combo1)
            RSave("WARNA") = Trim(Combo3)
            RSave("TAHUN") = Trim(Text8)
            RSave("RANGKA") = Trim(Text9)
            RSave("MESIN") = Trim(Text10)
            RSave("H_OTR") = CCur(Text1)
            RSave("H_BELI") = CCur(Text3)
            RSave("H_KOSONG") = CCur(Text2)
            RSave("MTS_MOTOR") = Trim(Combo2)
            
            'If CCur(Text12) <> CCur(Text3) Then
            '    Call PESAN
            'End If
            
        RSave.Update
        RSave.Close
        Set RSave = Nothing

Unload Me
M003.Show
End Sub

Private Sub PESAN()
G003.Show 1
End Sub

Private Sub Command2_Click()
MsgBox "JALANKAN PROSES TRANSAKSI PROSES GL SETELAH INI !", vbCritical, "PERUBAHAN HARGA BELI"

SDel = "Delete From M001 where NO_FAK = '" + Trim(Text11) + "'"
Set RDel = RDCO.OpenResultset(SDel, rdOpenDynamic, rdConcurRowVer)

Unload Me
M003.Show

End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=DEALER", rdDriverNoPrompt, False, CN)

Call SiapkanGrid
Call IsiGrid
Call Kosong
Call IsiCombo
Call IsiCombo3

Call NoSistem
Label10 = Tanggal
Text4 = Tanggal
Text8 = Year(Tanggal)
Text6 = "FAK NO." + Trim(Label3)

EDIT = 0
NoNo = 0
QW = 0

Label11.Visible = False

End Sub

Private Sub IsiCombo()
SType = "Select Nama_JNS from B003 where Kode_IND = '151' "
Set RType = RDCO.OpenResultset(SType, rdOpenDynamic, rdConcurRowVer)
If RType.RowCount <> 0 Then
    RType.MoveFirst
    Do Until RType.EOF
        Combo1.AddItem RType("Nama_JNS")
    RType.MoveNext
    Loop
End If
RType.Close
Set RType = Nothing
Combo1.ListIndex = 0

SType2 = "Select WARNA from B003WARNA"
Set RType2 = RDCO.OpenResultset(SType2, rdOpenDynamic, rdConcurRowVer)
If RType2.RowCount <> 0 Then
    RType2.MoveFirst
    Do Until RType2.EOF
        Combo3.AddItem RType2("WARNA")
    RType2.MoveNext
    Loop
Else
    MsgBox "TIPE KENDARAAN BELUM TERDAFTAR", vbInformation, "WARNING"
    Combo1 = ""
    Combo1.SetFocus
    Exit Sub
End If
RType2.Close
Set RType2 = Nothing
Combo3.ListIndex = 0
End Sub

Private Sub IsiCombo3()
SType = "Select NAMA from C012 where TIPE = '300' order by NAMA"
Set RType = RDCO.OpenResultset(SType, rdOpenDynamic, rdConcurRowVer)
If RType.RowCount <> 0 Then
    RType.MoveFirst
    Do Until RType.EOF
        Combo2.AddItem RType("NAMA")
    RType.MoveNext
    Loop
End If
RType.Close
Set RType = Nothing
Combo2.ListIndex = 0
End Sub

Private Sub Kosong()
ClearTextBoxes Me
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False
Option5.Value = False
Frame1.Visible = False
'Combo1 = ""
'Combo2 = ""
'Combo3 = ""
End Sub

Private Sub NoSistem()
Dim Nomor As Double
Dim InfoNomor As Double

SCari = "Select Top 1 No_Urut From M001 order by No_Urut Desc"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    Nomor = Val(RCari("No_Urut")) + 1
    Label3 = Nomor
Else
    Label3 = "1"
End If
RCari.Close
Set RCari = Nothing

End Sub

Private Sub SiapkanGrid()
With grid
    .Row = 0
    .Cols = 11
    .Col = 0: .ColWidth(0) = 500: .Text = "NO": .CellAlignment = 4: .CellFontBold = True
    .Col = 1: .ColWidth(1) = 1250: .Text = "FAK": .CellAlignment = 4: .CellFontBold = True
    .Col = 2: .ColWidth(2) = 2000: .Text = "TYPE": .CellAlignment = 4: .CellFontBold = True
    .Col = 3: .ColWidth(3) = 1500: .Text = "WARNA": .CellAlignment = 4: .CellFontBold = True
    .Col = 4: .ColWidth(4) = 1000: .Text = "RANGKA": .CellAlignment = 4: .CellFontBold = True
    .Col = 5: .ColWidth(5) = 1000: .Text = "MESIN": .CellAlignment = 4: .CellFontBold = True
    .Col = 6: .ColWidth(6) = 1000: .Text = "TAHUN": .CellAlignment = 4: .CellFontBold = True
    .Col = 7: .ColWidth(7) = 1250: .Text = "POKOK": .CellAlignment = 4: .CellFontBold = True
    .Col = 8: .ColWidth(8) = 1250: .Text = "OFF ROAD": .CellAlignment = 4: .CellFontBold = True
    .Col = 9: .ColWidth(9) = 1250: .Text = "OTR": .CellAlignment = 4: .CellFontBold = True
    .Col = 10: .ColWidth(10) = 1500: .Text = "MTS": .CellAlignment = 4: .CellFontBold = True
End With

End Sub

Private Sub IsiGrid()
SGrid = "Select * From M001 where STS_JUAL = '0' order by NO_FAK desc"
Set RGrid = RDCO.OpenResultset(SGrid, rdOpenKeyset, rdConcurReadOnly)
If RGrid.RowCount <> 0 Then
   RGrid.MoveFirst
   B = 1
   NoNo = 1
   Do Until RGrid.EOF
      grid.Rows = B + 1
      grid.Row = B
         With grid
              .Col = 0: .Text = NoNo: .CellAlignment = 4
              .Col = 1: .Text = RGrid("NO_FAK"): .CellAlignment = 4
              .Col = 2: .Text = RGrid("TYPE")
              .Col = 3: .Text = RGrid("WARNA"): .CellAlignment = 4
              .Col = 4: .Text = RGrid("RANGKA"): .CellAlignment = 4
              .Col = 5: .Text = RGrid("MESIN"): .CellAlignment = 4
              .Col = 6: .Text = RGrid("TAHUN"): .CellAlignment = 4
              .Col = 7: .Text = Format(RGrid("H_BELI"), "##,###.00")
              .Col = 8: .Text = Format(RGrid("H_KOSONG"), "##,###.00")
              .Col = 9: .Text = Format(RGrid("H_OTR"), "##,###.00")
              .Col = 10: .Text = RGrid("MTS_MOTOR")
         End With
      B = B + 1
      NoNo = NoNo + 1
      RGrid.MoveNext
   Loop
End If
RGrid.Close
Set RGrid = Nothing
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 13
    QW = 1
    
    TmbSave.Visible = False
                  
    Text1 = grid.TextMatrix(grid.Row, 8)
    Combo1 = grid.TextMatrix(grid.Row, 2)
    Text3 = grid.TextMatrix(grid.Row, 7)
    Text7 = grid.TextMatrix(grid.Row, 3)
    Text8 = grid.TextMatrix(grid.Row, 6)
    Text9 = grid.TextMatrix(grid.Row, 4)
    Text10 = grid.TextMatrix(grid.Row, 5)
    
    Text11 = grid.TextMatrix(grid.Row, 0)
    
    Text4.Visible = False
    Text5.Visible = False
    Text6.Visible = False
    
    Frame1.Visible = False
    Text12 = ""
End Select
End Sub

Private Sub grid_dblClick()
QW = 1

TmbSave.Visible = False
              
Text1 = grid.TextMatrix(grid.Row, 9)
Combo1 = grid.TextMatrix(grid.Row, 2)
Text3 = grid.TextMatrix(grid.Row, 7)
Combo3 = grid.TextMatrix(grid.Row, 3)
Combo2 = grid.TextMatrix(grid.Row, 10)
Text8 = grid.TextMatrix(grid.Row, 6)
Text9 = grid.TextMatrix(grid.Row, 4)
Text10 = grid.TextMatrix(grid.Row, 5)
Text6 = grid.TextMatrix(grid.Row, 1)
Text2 = grid.TextMatrix(grid.Row, 8)

Text11 = grid.TextMatrix(grid.Row, 0)

Text4.Visible = False
Text5.Visible = False
Text6.Visible = True

Frame1.Visible = False

'Text3.Enabled = False
Text6.Enabled = False
Combo1.Enabled = False

Label11.Visible = True

End Sub

Private Sub Option1_Click()
Option2.Value = False
Option3.Value = False
Option4.Value = False
Option5.Value = False

JJJ = ""
JJJ = "Select * From M001 where TYPE like '%" + Trim(Text12) + "%' and STS_JUAL = '0'"
Call SiapkanGrid
Call IsiGrid2

Option1.Value = False

End Sub

Private Sub IsiGrid2()
SGrid = JJJ
Set RGrid = RDCO.OpenResultset(SGrid, rdOpenKeyset, rdConcurReadOnly)
If RGrid.RowCount <> 0 Then
   RGrid.MoveFirst
   B = 1
   NoNo = 1
   Do Until RGrid.EOF
      grid.Rows = B + 1
      grid.Row = B
         With grid
              .Col = 0: .Text = NoNo: .CellAlignment = 4
              .Col = 1: .Text = RGrid("NO_FAK"): .CellAlignment = 4
              .Col = 2: .Text = RGrid("TYPE")
              .Col = 3: .Text = RGrid("WARNA"): .CellAlignment = 4
              .Col = 4: .Text = RGrid("RANGKA"): .CellAlignment = 4
              .Col = 5: .Text = RGrid("MESIN"): .CellAlignment = 4
              .Col = 6: .Text = RGrid("TAHUN"): .CellAlignment = 4
              .Col = 7: .Text = Format(RGrid("H_BELI"), "##,###.00")
              .Col = 8: .Text = Format(RGrid("H_KOSONG"), "##,###.00")
              .Col = 9: .Text = Format(RGrid("H_OTR"), "##,###.00")
              .Col = 10: .Text = RGrid("MTS_MOTOR")
         End With
      B = B + 1
      NoNo = NoNo + 1
      RGrid.MoveNext
   Loop
End If
RGrid.Close
Set RGrid = Nothing
End Sub

Private Sub Option2_Click()
Option1.Value = False
Option3.Value = False
Option4.Value = False
Option5.Value = False

JJJ = ""
JJJ = "Select * From M001 where WARNA like '%" + Trim(Text12) + "%' and STS_JUAL = '0'"
Call SiapkanGrid
Call IsiGrid2

Option2.Value = False
End Sub

Private Sub Option3_Click()
Option1.Value = False
Option2.Value = False
Option4.Value = False
Option5.Value = False

JJJ = ""
JJJ = "Select * From M001 where RANGKA = '" + Trim(Text12) + "' and STS_JUAL = '0'"
Call SiapkanGrid
Call IsiGrid2

Option3.Value = False
End Sub

Private Sub Option4_Click()
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option5.Value = False

JJJ = ""
JJJ = "Select * From M001 where MESIN = '" + Trim(Text12) + "' and STS_JUAL = '0'"
Call SiapkanGrid
Call IsiGrid2

Option4.Value = False
End Sub

Private Sub Option5_Click()
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False

JJJ = ""
JJJ = "Select * From M001 where TAHUN = '" + Trim(Text12) + "' and STS_JUAL = '0'"
Call SiapkanGrid
Call IsiGrid2

Option5.Value = False
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text1_LostFocus()
If Text1 = "" Then Text1 = 0
Text1 = Format(Text1, "##,###.00")
End Sub

Private Sub text1_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF1
        Text2.SetFocus
End Select
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys vbTab
    Text10 = Format(Text10, ">")
End If
End Sub

Private Sub Text10_LostFocus()
If Text10 = "" Then Exit Sub
If QW = 0 Then
    SSusuSore = "Select MESIN from M001 where MESIN = '" + Trim(Text10) + "'"
    Set RSusuSore = RDCO.OpenResultset(SSusuSore, rdOpenDynamic, rdConcurRowVer)
    If RSusuSore.RowCount <> 0 Then
        Text10.SetFocus
        MsgBox "NOMOR MESIN TELAH DIGUNAKAN", vbInformation, "WARNING"
        Text10 = ""
        Exit Sub
    End If
    RSusuSore.Close
    Set RSusuSore = Nothing
End If
End Sub

Private Sub text10_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF1
        Text9.SetFocus
End Select
End Sub

Private Sub Text12_click()
Call Kosong
Frame1.Visible = True
End Sub

Private Sub TEXT12_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text12 = Format(Text12, ">")
End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    'Combo3.Clear
    SendKeys vbTab
End If
End Sub

'Private Sub Combo3_GotFocus()
'Call IsiCombo2
'End Sub

Private Sub combo3_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF1
        Combo1.SetFocus
End Select
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call HARGA
    Combo3 = Format(Combo3, ">")
    SendKeys vbTab
End If
End Sub

Private Sub HARGA()
SCari = "Select * from B003A where NAMA_JNS = '" + Trim(Combo1) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    Text3 = Format(RCari("HBELI"), "##,###.00")
    Text2 = Format(RCari("HKOSONG"), "##,###.00")
    Text1 = Format(RCari("HJUAL"), "##,###.00")
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text2_LostFocus()
If Text2 = "" Then Text2 = 0
Text2 = Format(Text2, "##,###.00")
End Sub

Private Sub text2_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF1
        Text5.SetFocus
End Select
End Sub

Private Sub text3_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
    If KeyAscii = 13 Then SendKeys vbTab
    
If EDIT = 1 Then
    Text5 = "KOREKSI"
Else
    Text5 = "PEMBELIAN " + Trim(Text6)
End If

End Sub

Private Sub Text3_LostFocus()
If Text3 = "" Then Text3 = 0
Text3 = Format(Text3, "##,###.00")
End Sub

Private Sub text3_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF1
        Combo3.SetFocus
End Select
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text4_LostFocus()
If Text4 = "" Then Exit Sub
If Not IsDate(Text4) Then
    Text4.SetFocus
    MsgBox "TYPE DATA MENGGUNAKAN TANGGAL", vbCritical, "TYPE DATA SALAH"
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

Private Sub text5_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF1
        Text3.SetFocus
End Select
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys vbTab
    Text6 = Format(Text6, ">")
End If
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub text8_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF1
        Text1.SetFocus
End Select
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys vbTab
    Text9 = Format(Text9, ">")
End If
End Sub

Private Sub Text9_LostFocus()
If Text9 = "" Then Exit Sub
If QW = 0 Then
    SSusuSore = "Select RANGKA from M001 where RANGKA = '" + Trim(Text9) + "'"
    Set RSusuSore = RDCO.OpenResultset(SSusuSore, rdOpenDynamic, rdConcurRowVer)
    If RSusuSore.RowCount <> 0 Then
        Text9.SetFocus
        MsgBox "NOMOR RANGKA TELAH DIGUNAKAN", vbInformation, "WARNING"
        Text9 = ""
        Exit Sub
    End If
    RSusuSore.Close
    Set RSusuSore = Nothing
End If
End Sub

Private Sub text9_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF1
        Text8.SetFocus
End Select
End Sub

Private Sub Timer1_Timer()
Label11.ForeColor = RGB(Rnd * 500, Rnd * 605, Rnd * 700)
End Sub

Private Sub TmbSave_Click()
Dim Tanya

If Text6 = "" Or Text4 = "" Then
    MsgBox "FAKTUR PEMBELIAN / TANGGAL PEMBELIAN MASIH KOSONG", vbCritical, "DATA TIDAK BOLEH KOSONG"
    Text6.SetFocus
    Exit Sub
End If

If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text5 = "" Then
    MsgBox "HARGA POKOK /HARGA OFF ROAD / HARGA OTR / KETERANGAN MASIH KOSONG", vbCritical, "DATA TIDAK BOLEH KOSONG"
    Text3.SetFocus
    Exit Sub
End If

If Combo1 = "" Or Combo3 = "" Or Text8 = "" Or Text9 = "" Or Text10 = "" Then
    MsgBox "TYPE / WARNA / TAHUN / NO.RANGKA / NO.MESIN MASIH KOSONG", vbCritical, "DATA TIDAK BOLEH KOSONG"
    Combo1.SetFocus
    Exit Sub
End If

If Combo2 = "" Then
    MsgBox "MUTASI KENDARAAN MASIH KOSONG", vbCritical, "DATA TIDAK BOLEH KOSONG"
    Combo2.SetFocus
    Exit Sub
End If

Tanya = MsgBox("ANDA YAKIN PROSES PEMBELIAN KENDARAAN TYPE = " + Trim(Combo1) + ", WARNA = " + Trim(Combo3) + " ?", vbOKCancel, "PROSES TRANSAKSI")
If Tanya = vbCancel Then Exit Sub

SCari = "Select * From B003A where NAMA_JNS = '" + Trim(Combo1) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset, rdConcurRowVer)
RCari.EDIT
    RCari("HBeli") = CCur(Text3)
    RCari("HJual") = CCur(Text1)
    RCari("HKosong") = CCur(Text2)
RCari.Update
RCari.Close
Set RCari = Nothing

SCari10 = "Select * From B001 where KODE_IND='" + Trim(151) + "'"
Set RCari10 = RDCO.OpenResultset(SCari10, rdOpenKeyset, rdConcurRowVer)
    SGLSEDIA = RCari10("SGL_SEDIA")
    SGLCREDIT = RCari10("SGL_CREDIT")
    SGLMODAL = RCari10("SGL_MODAL")
        Call JurnalBahan
        Call JurnalDEBET
        Call JurnalCredit
        Call LabaRugi
RCari10.Close
Set RCari10 = Nothing

Unload Me
M003.Show

End Sub

Private Sub JurnalBahan()
SSave = "Select * From M001"
Set RSave = RDCO.OpenResultset(SSave, rdOpenKeyset, rdConcurRowVer)
RSave.AddNew
    RSave("TYPE") = Trim(Combo1)
    RSave("WARNA") = Trim(Combo3)
    RSave("TAHUN") = Trim(Text8)
    RSave("RANGKA") = Trim(Text9)
    RSave("MESIN") = Trim(Text10)
    RSave("CCAB") = CodeCab
    RSave("STS_JUAL") = 0
    RSave("TGL_INPUT") = Trim(Text4)
    RSave("H_BELI") = CCur(Text3)
    RSave("H_KOSONG") = CCur(Text2)
    RSave("H_OTR") = CCur(Text1)
    RSave("MTS_MOTOR") = Trim(Combo2)
    RSave("NO_FAK") = Trim(Text6)
    RSave("Tanggal") = Tanggal
RSave.Update
RSave.Close
Set RSave = Nothing
End Sub

Private Sub JurnalDEBET()
SCari = "Select * From G003 where codesl='" + Trim(SGLSEDIA) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset, rdConcurRowVer)
    MMUTASID = RCari("mutasid") + CCur(Text3)
    SSALDO = RCari("saldo") + CCur(Text3)
    NNAMA = RCari("NamaSL")
    RCari.EDIT
        RCari("mutasid") = CCur(MMUTASID)
        RCari("saldo") = CCur(SSALDO)

    SCari2 = "Select * From G005"
    Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenKeyset, rdConcurRowVer)
    RCari2.AddNew
        RCari2("codecab") = CodeCab
        RCari2("codesl") = SGLSEDIA
        RCari2("namasl") = NNAMA
        RCari2("nobukti") = Trim(Text6)
        RCari2("keterangan") = Trim(Text5)
        RCari2("nominald") = CCur(Text3)
        RCari2("nominalc") = 0
        RCari2("saldo") = SSALDO
        RCari2("tanggal") = Tanggal
        RCari2("jam") = Date
        RCari2("usercode") = Operator
    RCari2.Update
    RCari2.Close
    Set RCari2 = Nothing

RCari.Update
RCari.Close
Set RCari = Nothing

SCari3 = "Select * From G003 where codesl='" + Trim(SGLCREDIT) + "'"
Set RCari3 = RDCO.OpenResultset(SCari3, rdOpenKeyset, rdConcurRowVer)
    MMUTASID = RCari3("mutasid") + CCur(Text3)
    SSALDO = RCari3("saldo") + CCur(Text3)
    NNAMA = RCari3("NamaSL")
    RCari3.EDIT
        RCari3("mutasid") = CCur(MMUTASID)
        RCari3("saldo") = CCur(SSALDO)

    SCari4 = "Select * From G005"
    Set RCari4 = RDCO.OpenResultset(SCari4, rdOpenKeyset, rdConcurRowVer)
    RCari4.AddNew
        RCari4("codecab") = CodeCab
        RCari4("codesl") = SGLCREDIT
        RCari4("namasl") = NNAMA
        RCari4("nobukti") = Trim(Text6)
        RCari4("keterangan") = Trim(Text5)
        RCari4("nominald") = CCur(Text3)
        RCari4("nominalc") = 0
        RCari4("saldo") = SSALDO
        RCari4("tanggal") = Tanggal
        RCari4("jam") = Date
        RCari4("usercode") = Operator
    RCari4.Update
    RCari4.Close
    Set RCari4 = Nothing

RCari3.Update
RCari3.Close
Set RCari3 = Nothing

End Sub

Private Sub JurnalCredit()
SCariCari = "Select * From G003 where codesl='" + Trim(SGLCREDIT) + "'"
Set RCariCari = RDCO.OpenResultset(SCariCari, rdOpenKeyset, rdConcurRowVer)
    MMUTASIC = RCariCari("mutasic") + CCur(Text3)
    SSALDO = RCariCari("saldo") - CCur(Text3)
    NNAMA = RCariCari("NamaSL")
    RCariCari.EDIT
        RCariCari("mutasic") = CCur(MMUTASIC)
        RCariCari("saldo") = CCur(SSALDO)

    SCariCari2 = "Select * From G005"
    Set RCariCari2 = RDCO.OpenResultset(SCariCari2, rdOpenKeyset, rdConcurRowVer)
    RCariCari2.AddNew
        RCariCari2("codecab") = CodeCab
        RCariCari2("codesl") = SGLCREDIT
        RCariCari2("namasl") = NNAMA
        RCariCari2("nobukti") = Trim(Text6)
        RCariCari2("keterangan") = Trim(Text5)
        RCariCari2("nominald") = 0
        RCariCari2("nominalc") = CCur(Text3)
        RCariCari2("saldo") = SSALDO
        RCariCari2("tanggal") = Tanggal
        RCariCari2("jam") = Date
        RCariCari2("usercode") = Operator
    RCariCari2.Update
    RCariCari2.Close
    Set RCariCari2 = Nothing

RCariCari.Update
RCariCari.Close
Set RCariCari = Nothing

SCariCari3 = "Select * From G003 where codesl='" + Trim(SGLMODAL) + "'"
Set RCariCari3 = RDCO.OpenResultset(SCariCari3, rdOpenKeyset, rdConcurRowVer)
    MMUTASIC = RCariCari3("mutasic") + CCur(Text3)
    SSALDO = RCariCari3("saldo") + CCur(Text3)
    NNAMA = RCariCari3("NamaSL")
    RCariCari3.EDIT
        RCariCari3("mutasic") = CCur(MMUTASIC)
        RCariCari3("saldo") = CCur(SSALDO)

    SCariCari4 = "Select * From G005"
    Set RCariCari4 = RDCO.OpenResultset(SCariCari4, rdOpenKeyset, rdConcurRowVer)
    RCariCari4.AddNew
        RCariCari4("codecab") = CodeCab
        RCariCari4("codesl") = SGLMODAL
        RCariCari4("namasl") = NNAMA
        RCariCari4("nobukti") = Trim(Text6)
        RCariCari4("keterangan") = Trim(Text5)
        RCariCari4("nominald") = 0
        RCariCari4("nominalc") = CCur(Text3)
        RCariCari4("saldo") = SSALDO
        RCariCari4("tanggal") = Tanggal
        RCariCari4("jam") = Date
        RCariCari4("usercode") = Operator
    RCariCari4.Update
    RCariCari4.Close
    Set RCariCari4 = Nothing

RCariCari3.Update
RCariCari3.Close
Set RCariCari3 = Nothing

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
