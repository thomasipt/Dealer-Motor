VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form B003SS 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "KODE BARANG"
   ClientHeight    =   8940
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   9120
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   596
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   608
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "CETAK TABEL SPAREPART"
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
      Left            =   109
      TabIndex        =   25
      Top             =   8400
      Width           =   2520
   End
   Begin VB.CommandButton Command1 
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
      Left            =   8044
      TabIndex        =   20
      Top             =   8400
      Width           =   960
   End
   Begin VB.TextBox Text2 
      Height          =   360
      Left            =   1865
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   1500
      Width           =   5505
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Left            =   1865
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1050
      Width           =   1860
   End
   Begin VB.CommandButton TmbSave 
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
      Left            =   476
      TabIndex        =   7
      Top             =   3090
      Width           =   1000
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "B003SS.frx":0000
      Left            =   1865
      List            =   "B003SS.frx":000A
      TabIndex        =   6
      Text            =   "Combo2"
      Top             =   2415
      Width           =   1410
   End
   Begin VB.CommandButton TmbDel 
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
      Left            =   7558
      TabIndex        =   8
      Top             =   3090
      Width           =   1000
   End
   Begin VB.TextBox Text3 
      Height          =   360
      Left            =   1865
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   1950
      Width           =   915
   End
   Begin VB.TextBox Text5 
      Height          =   360
      Left            =   3784
      TabIndex        =   5
      Text            =   "Text5"
      Top             =   1950
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1865
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   630
      Width           =   1680
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   1865
      TabIndex        =   0
      Text            =   "Combo3"
      Top             =   150
      Width           =   1005
   End
   Begin Crystal.CrystalReport CRPT 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      Caption         =   "INFO DAFTAR HARGA SPAREPART"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   4530
      Left            =   109
      TabIndex        =   19
      Top             =   3675
      Width           =   8895
      Begin VB.CommandButton Command4 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4275
         Left            =   30
         TabIndex        =   26
         Top             =   210
         Width           =   345
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3930
         TabIndex        =   21
         Text            =   "Text4"
         Top             =   1890
         Width           =   3435
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3930
         TabIndex        =   22
         Text            =   "Text6"
         Top             =   3045
         Width           =   3435
      End
      Begin VB.Label Label4 
         Caption         =   "HARGA BELI"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1845
         TabIndex        =   24
         Top             =   1935
         Width           =   1500
      End
      Begin VB.Label Label8 
         Caption         =   "HARGA JUAL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1845
         TabIndex        =   23
         Top             =   3090
         Width           =   1500
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         Height          =   3495
         Left            =   570
         Top             =   990
         Width           =   8250
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   4530
      Left            =   109
      TabIndex        =   27
      ToolTipText     =   "Klik untuk edit"
      Top             =   3675
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   7990
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "KODE GOLONGAN"
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
      Left            =   297
      TabIndex        =   18
      Top             =   630
      Width           =   1500
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080C0FF&
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
      Left            =   297
      TabIndex        =   17
      Top             =   1073
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080C0FF&
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
      Height          =   315
      Left            =   297
      TabIndex        =   16
      Top             =   1523
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080C0FF&
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3852
      TabIndex        =   15
      Top             =   1050
      Width           =   3480
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080C0FF&
      Caption         =   "SATUAN"
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
      Left            =   297
      TabIndex        =   14
      Top             =   2415
      Width           =   1455
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Label11"
      Height          =   285
      Left            =   3627
      TabIndex        =   13
      Top             =   645
      Width           =   3705
   End
   Begin VB.Label Label7 
      BackColor       =   &H0080C0FF&
      Caption         =   "STYLE"
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
      Left            =   297
      TabIndex        =   12
      Top             =   1973
      Width           =   1500
   End
   Begin VB.Label Label12 
      BackColor       =   &H0080C0FF&
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
      Height          =   285
      Left            =   2899
      TabIndex        =   11
      Top             =   1995
      Width           =   825
   End
   Begin VB.Line Line1 
      X1              =   8.8
      X2              =   596.8
      Y1              =   195
      Y2              =   195
   End
   Begin VB.Label Label17 
      BackColor       =   &H0080C0FF&
      Caption         =   "DISTRIBUTOR"
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
      Left            =   297
      TabIndex        =   10
      Top             =   142
      Width           =   1410
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Label16"
      Height          =   285
      Left            =   2966
      TabIndex        =   9
      Top             =   165
      Width           =   2985
   End
End
Attribute VB_Name = "B003SS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private RGol, RCari, RKode, RDel, RDelBar, RSim, RSave, RSaveP, RDist As rdoResultset
Private SDelBar, SDist, SGol, SCari, Metode, SKode, SDel, SSim, SSave, SSaveP As String
Private Brs, MetodLaba, Ganti, TOKET, WARNA, SPART

Private Sub Combo1_GotFocus()
    SendKeys "{F4}"
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Combo1_LostFocus()
If Combo1 = "" Then Exit Sub
SGol = "Select Keterangan from B001 where Kode_ind = '" + Trim(Combo1) + "'"
Set RGol = RDCO.OpenResultset(SGol, rdOpenDynamic, rdConcurRowVer)
If RGol.RowCount <> 0 Then
    Label11 = RGol("Keterangan")
End If
RGol.Close
Set RGol = Nothing

If Combo1 = 151 Then
    Call KodeJenis2
    SPART = 1
ElseIf Combo1 = 152 Then
    Call KodeJenis
    SPART = 0
ElseIf Combo1 = 153 Then
    'Call KodeJenis3
    SPART = 0
End If
    If SPART = 1 Then
        Frame1.Visible = True
        Frame1.ZOrder
    Else
        Frame1.Visible = False
    End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Combo1 = "153" Then
        Frame1.Visible = True
        Frame1.ZOrder
        Text4.SetFocus
    ElseIf SPART = 1 Then
        Text4.SetFocus
    End If
End If
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Combo3_LostFocus()
If Combo3 = "" Then Exit Sub
SDist = "Select Nama_Distb from C007 where Kode_DistB = '" + Combo3 + "'"
Set RDist = RDCO.OpenResultset(SDist, rdOpenDynamic, rdConcurRowVer)
If RDist.RowCount <> 0 Then
    Label16 = RDist("Nama_Distb")
Else
    MsgBox "KODE DISTRIBUTOR BELUM TERDAFTAR", vbInformation, "KODE BLM TERDAFTAR"
    Combo3.SetFocus
End If
RDist.Close
Set RDist = Nothing
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
crpt.ReportFileName = App.Path + "\ReportD\B003SS.rpt"
crpt.WindowState = crptMaximized
crpt.WindowMaxButton = True
crpt.WindowMinButton = True
crpt.Action = 1
End Sub

Private Sub Command3_Click()
crpt.ReportFileName = App.Path + "\ReportD\B003A.rpt"
crpt.WindowState = crptMaximized
crpt.WindowMaxButton = True
crpt.WindowMinButton = True
crpt.Action = 1
End Sub

Private Sub Command4_Click()
Call Kosong
Frame1.Visible = False
grid.ZOrder
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=DEALER", rdDriverNoPrompt, False, CN)

Call Kosong
Call IsiGol
Call Siap
Call IsiGrid
Combo2.ListIndex = 0
Text3 = ""
Text4 = ""
'Call KodeJenis
Call KodeDistributor

TOKET = 0
SPART = 0
Frame1.Visible = False

End Sub

Private Sub Kosong()
ClearTextBoxes Me
Label5 = ""
Label11 = ""
Label16 = ""
TmbDel.Enabled = False
Ganti = 0
End Sub

Private Sub IsiGol()
Dim KodeG
SGol = "Select Kode_IND From B001 where KODE_IND = '153' order by Kode_Ind"
Set RGol = RDCO.OpenResultset(SGol, rdOpenDynamic, rdConcurRowVer)
If RGol.RowCount <> 0 Then
    RGol.MoveFirst
    Do While Not RGol.EOF
        Combo1.AddItem RGol("Kode_Ind")
    RGol.MoveNext
    Loop
End If

RGol.Close
Set RGol = Nothing
Combo1.ListIndex = 0
End Sub

Private Sub KodeJenis()
Dim AutoNomor As Double
SNo = "Select Top 1 Kode_Jns from B003 order by Kode_Jns desc"
Set RNo = RDCO.OpenResultset(SNo, rdOpenDynamic, rdConcurRowVer)
If RNo.RowCount <> 0 Then
    AutoNomor = Mid(RNo("Kode_Jns"), 3, 11) + 1
    Text1 = "TR" + Trim(Digit(11, AutoNomor))
Else
    Text1 = "TR00000000001"
End If
RNo.Close
Set RNo = Nothing
Label5 = Text1
End Sub

Private Sub KodeJenis2()
Dim AutoNomor As Double
SNo = "Select Top 1 Kode_Jns from B003 where KODE_IND = '151' order by Kode_Jns desc"
Set RNo = RDCO.OpenResultset(SNo, rdOpenDynamic, rdConcurRowVer)
If RNo.RowCount <> 0 Then
    AutoNomor = Mid(RNo("Kode_Jns"), 3, 11) + 1
    Text1 = "M" + Trim(Digit(5, AutoNomor))
Else
    Text1 = "M00001"
End If
RNo.Close
Set RNo = Nothing
Label5 = Text1

End Sub

Private Sub KodeJenis3()
Dim AutoNomor As Double
SNo = "Select Top 1 Kode_Jns from B003 where KODE_IND = '153' order by Kode_Jns desc"
Set RNo = RDCO.OpenResultset(SNo, rdOpenDynamic, rdConcurRowVer)
If RNo.RowCount <> 0 Then
    AutoNomor = Mid(RNo("Kode_Jns"), 3, 11) + 1
    Text1 = "SP" + Trim(Digit(4, AutoNomor))
Else
    Text1 = "SP0001"
End If
RNo.Close
Set RNo = Nothing
Label5 = Text1
End Sub

Private Sub KodeDistributor()
SDist = "Select Kode_Distb from C007 order by Kode_Distb"
Set RDist = RDCO.OpenResultset(SDist, rdOpenDynamic, rdConcurRowVer)
If RDist.RowCount <> 0 Then
    RDist.MoveFirst
    Do Until RDist.EOF
        Combo3.AddItem RDist("Kode_Distb")
    RDist.MoveNext
    Loop
End If
RDist.Close
Set RDist = Nothing
Combo3.ListIndex = 0
End Sub

Private Sub grid_dblClick()
TOKET = 1
grid.Col = 1
KODE = ""
Clipboard.SetText (grid.Text)
KODE = grid.Text
Text1 = KODE

Call Ngantuk

End Sub

Private Sub Ngantuk()
SCari = "Select * from B003 where Kode_Jns = '" + Text1 + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    Info = MsgBox("KODE JENIS BARANG SUDAH TERDAFTAR, AKAN DILAKUKAN EDIT", vbOKCancel, "KODE JENIS BARANG SUDAH TERDAFTAR")
    If Info = vbOK Then
        Combo1 = RCari("Kode_Ind")
        If IsNull(RCari("Kode_Distb")) Then
            Combo3 = ""
            Label16 = ""
        End If
        Combo3 = RCari("Kode_Distb")
        Call Combo3_LostFocus
            SGol = "Select Keterangan from B001 where Kode_ind = '" + Trim(Combo1) + "'"
            Set RGol = RDCO.OpenResultset(SGol, rdOpenDynamic, rdConcurRowVer)
            If RGol.RowCount <> 0 Then
                Label11 = RGol("Keterangan")
            End If
            RGol.Close
            Set RGol = Nothing
        
        Text2 = RCari("Nama_Jns")
        Combo2 = RCari("satuan")
        
        If Mid(Text1, 1, 3) = "002" Then
            Check3.Value = 1
            MsgBox "OK"
        Else
        
        End If
        Text3 = RCari("STYLE")
        Text5 = RCari("WARNA")
        TmbDel.Enabled = True
    Else
        Combo1.SetFocus
        Call Kosong
    End If
Else
    Text2 = ""
    Text3 = ""
    Text5 = ""
End If
RCari.Close
Set RCari = Nothing
Text1 = Format(Text1, ">")
Label5 = Text1
    If Combo1 = "153" Then
        Frame1.Visible = True
        Frame1.ZOrder
        
        SCari = "Select * from B003A where KODE_JNS = '" + Trim(Text1) + "'"
        Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
        If RCari.RowCount <> 0 Then
            Text4 = Format(RCari("HBELI"), "##,###.00")
            Text6 = Format(RCari("HJUAL"), "##,###.00")
        End If
        RCari.Close
        Set RCari = Nothing
        SPART = 1
    Else
        Frame1.Visible = False
    End If
End Sub

Private Sub text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text1_LostFocus()
Dim Info, Awal, Akhir

'If Text1 = "" Then
'    Exit Sub
'Else
'    Text1 = "TR" + Trim(Digit(11, Text1))
'End If

Call Ngantuk

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text2_LostFocus()
If Ganti = 1 Then Exit Sub
SCari = "Select Nama_JNS from B003 where Nama_JNS = '" + Trim(Text2) + "' and KODE_IND = '" + Trim(Combo1) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    Text2.SetFocus
    MsgBox "NAMA JENIS BARANG SUDAH DIGUNAKAN", vbInformation, "NAMA JENIS BARANG SUDAH TERDAFTAR"
    Exit Sub
End If
RCari.Close
Set RCari = Nothing
Text2 = Format(Text2, ">")
End Sub

Private Sub text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text3_LostFocus()
Text3 = Format(Text3, ">")
End Sub

Private Sub text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text4_LostFocus()
If Text4 = "" Then Text4.SetFocus
If Not IsNumeric(Text4) Then
    Text4.SetFocus
    MsgBox "NOMINAL HARGA BELI HARUS ANGKA", vbCritical, "TYPE DATA SALAH"
    Text4 = 0
    Exit Sub
End If
Text4 = Format(Text4, "##,###.00")
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text5_LostFocus()
Text5 = Format(Text5, ">")
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SPART = 1
    TmbSave.SetFocus
End If
End Sub

Private Sub Text6_LostFocus()
If Text6 = "" Then Text6.SetFocus
If Not IsNumeric(Text6) Then
    Text6.SetFocus
    MsgBox "NOMINAL HARGA JUAL HARUS ANGKA", vbCritical, "TYPE DATA SALAH"
    Text6 = 0
    Exit Sub
End If
Text6 = Format(Text6, "##,###.00")
End Sub

Private Sub TmbDel_Click()
Dim Tanya
Tanya = MsgBox("YAKIN AKAN HAPUS DATA BAHAN " + Trim(Text1), vbOKCancel, "YAKIN HAPUS KODE BAHAN ?")
If Tanya = vbCancel Then Exit Sub
If Text1 = "" Then Exit Sub
SDel = "Delete From B003 where Kode_JNS = '" + Trim(Text1) + "'"
Set RDel = RDCO.OpenResultset(SDel, rdOpenDynamic, rdConcurRowVer)
MsgBox "DATA TELAH DI HAPUS", vbCritical, "WARNING"
Unload Me
B003.Show
End Sub

Private Sub TmbSave_Click()
If Combo1 = "" Or Combo3 = "" Or Text1 = "" Or Text2 = "" Or Combo2 = "" Then
    MsgBox "MASIH ADA DATA YANG KOSONG", vbCritical, "DATA TIDAK BOLEH KOSONG"
    Exit Sub
End If

If TOKET = 0 Then
    Call ENTRIBRG
    If SPART = 1 Then
        Call PART
    End If
Else
    Call ENTRIBRG2
    Call PART2
End If

Call Kosong
Combo3.SetFocus
Unload Me
B003SS.Show
End Sub

Private Sub PART()
SSaveP = "Select * From B003A"
Set RSaveP = RDCO.OpenResultset(SSaveP, rdOpenDynamic, rdConcurRowVer)
RSaveP.AddNew
    RSaveP("KODE_JNS") = Trim(Text1)
    RSaveP("NAMA_JNS") = Trim(Text2)
    RSaveP("SALDOAWAL") = 0
    RSaveP("MUTASID") = 0
    RSaveP("MUTASIC") = 0
    RSaveP("SALDO") = 0
    RSaveP("HBELI") = CCur(Text4)
    RSaveP("HJUAL") = CCur(Text6)
    RSaveP("HKOSONG") = 0
    RSaveP("TANGGAL") = Tanggal
RSaveP.Update
RSaveP.Close
Set RSaveP = Nothing
End Sub

Private Sub PART2()
SSaveP = "Select * from B003A where Kode_Jns = '" + Text1 + "'"
Set RSaveP = RDCO.OpenResultset(SSaveP, rdOpenDynamic, rdConcurRowVer)
RSaveP.EDIT
    RSaveP("KODE_JNS") = Trim(Text1)
    RSaveP("NAMA_JNS") = Trim(Text2)
    RSaveP("HBELI") = CCur(Text4)
    RSaveP("HJUAL") = CCur(Text6)
    RSaveP("HKOSONG") = 0
    RSaveP("TANGGAL") = Tanggal
RSaveP.Update
RSaveP.Close
Set RSaveP = Nothing
End Sub

Private Sub ENTRIBRG2()
SSave = "Select * From B003 where Kode_JNS = '" + Trim(Text1) + "'"
Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
RSave.EDIT
        RSave("NAMA_JNS") = Trim(Text2)
        RSave("STYLE") = Trim(Text3)
        RSave("WARNA") = Trim(Text5)
        RSave("SATUAN") = Trim(Combo2)
RSave.Update
RSave.Close
Set RSave = Nothing
End Sub

Private Sub ENTRIBRG()
SSave = "Select * From B003"
Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
RSave.AddNew
        RSave("KODE_DISTB") = Trim(Combo3)
        RSave("KODE_IND") = Trim(Combo1)
        RSave("KODE_JNS") = Trim(Text1)
        RSave("NAMA_JNS") = Trim(Text2)
        RSave("STYLE") = Trim(Text3)
        RSave("WARNA") = Trim(Text5)
        RSave("SATUAN") = Trim(Combo2)
        
        RSave("JML_AWAL") = 0
        RSave("JML_DBT") = 0
        RSave("JML_CRD") = 0
        RSave("JML_AKHIR") = 0

        RSave("TANGGAL") = Tanggal
RSave.Update
RSave.Close
Set RSave = Nothing
End Sub

Private Sub IsiGrid()
SCari = "Select * from B003 where KODE_IND = '153' order by Kode_Jns"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset, rdConcurReadOnly)
If RCari.RowCount <> 0 Then
RCari.MoveFirst
Brs = 1
Do Until RCari.EOF
       With grid
        .Rows = Brs + 1
        .Row = Brs
        .Col = 0: .Text = RCari("Kode_Ind"): .CellAlignment = 4
        .Col = 1: .Text = RCari("Kode_Jns"): .CellAlignment = 2
        .Col = 2: .Text = RCari("Nama_Jns"): .CellAlignment = 2
        .Col = 3: .Text = RCari("Style"): .CellAlignment = 4
        .Col = 4: .Text = RCari("Warna"): .CellAlignment = 4
        .Col = 5: .Text = RCari("Satuan"): .CellAlignment = 4
      End With
      RCari.MoveNext
      Brs = Brs + 1
Loop
End If
RCari.Close
Set RCari = Nothing
If Brs > 14 Then
    grid.TopRow = Brs - 14
End If
End Sub

Private Sub Siap()
With grid
     .Cols = 6
     .Row = 0
     .Col = 0: .ColWidth(0) = 750: .Text = "INDUK": .CellAlignment = 4
     .Col = 1: .ColWidth(1) = 1750: .Text = "KODE": .CellAlignment = 4
     .Col = 2: .ColWidth(2) = 3750: .Text = "NAMA": .CellAlignment = 4
     .Col = 3: .ColWidth(3) = 750: .Text = "STYLE": .CellAlignment = 4
     .Col = 4: .ColWidth(4) = 750: .Text = "WARNA": .CellAlignment = 4
     .Col = 5: .ColWidth(5) = 750: .Text = "SATUAN": .CellAlignment = 4
End With
End Sub
