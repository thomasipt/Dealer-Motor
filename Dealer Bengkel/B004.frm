VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form B004 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DEBET / CREDIT BARANG"
   ClientHeight    =   9315
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8250
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9315
   ScaleWidth      =   8250
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
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
      Left            =   3615
      TabIndex        =   20
      Top             =   8715
      Width           =   960
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   4995
      Left            =   90
      TabIndex        =   19
      Top             =   3465
      Width           =   8070
      _ExtentX        =   14235
      _ExtentY        =   8811
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   65280
      ForeColorFixed  =   0
      BackColorBkg    =   16777152
      GridColor       =   8421504
      FocusRect       =   0
      Appearance      =   0
   End
   Begin VB.Frame Frame2 
      Height          =   3375
      Left            =   90
      TabIndex        =   3
      Top             =   8
      Width           =   8070
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1620
         TabIndex        =   17
         Text            =   "Combo1"
         Top             =   233
         Width           =   1680
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1620
         TabIndex        =   16
         Text            =   "Combo2"
         Top             =   1133
         Width           =   2325
      End
      Begin VB.CommandButton Command3 
         Caption         =   "BATAL"
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
         Left            =   5198
         TabIndex        =   15
         Top             =   2850
         Width           =   960
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   360
         Left            =   1590
         TabIndex        =   4
         Text            =   "Text4"
         Top             =   1965
         Width           =   1845
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00FFFFC0&
         Height          =   360
         Left            =   1590
         MaxLength       =   35
         TabIndex        =   6
         Text            =   "Text5"
         Top             =   2415
         Width           =   3885
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
         Left            =   2093
         TabIndex        =   8
         Top             =   2850
         Width           =   960
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3675
         TabIndex        =   18
         Top             =   233
         Width           =   1680
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         Height          =   360
         Left            =   1620
         TabIndex        =   14
         Top             =   660
         Width           =   3885
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label2"
         Height          =   360
         Left            =   1620
         TabIndex        =   13
         Top             =   1560
         Width           =   3885
      End
      Begin VB.Label Label4 
         Caption         =   "DEBET"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   105
         TabIndex        =   12
         Top             =   210
         Width           =   1230
      End
      Begin VB.Label Label5 
         Caption         =   "NAMA BARANG"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   105
         TabIndex        =   11
         Top             =   660
         Width           =   1230
      End
      Begin VB.Label Label6 
         Caption         =   "CREDIT"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   105
         TabIndex        =   10
         Top             =   1110
         Width           =   1230
      End
      Begin VB.Label Label7 
         Caption         =   "NAMA BARANG"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   105
         TabIndex        =   9
         Top             =   1560
         Width           =   1230
      End
      Begin VB.Label Label8 
         Caption         =   "JUMLAH"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   105
         TabIndex        =   7
         Top             =   2010
         Width           =   1230
      End
      Begin VB.Label Label9 
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
         Height          =   360
         Left            =   105
         TabIndex        =   5
         Top             =   2460
         Width           =   1230
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   90
      TabIndex        =   0
      Top             =   8
      Width           =   8070
      Begin VB.CommandButton Command2 
         Caption         =   "CREDIT"
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
         Left            =   5160
         TabIndex        =   2
         Top             =   2775
         Width           =   2280
      End
      Begin VB.CommandButton Command1 
         Caption         =   "DEBET"
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
         Left            =   810
         TabIndex        =   1
         Top             =   2775
         Width           =   2280
      End
   End
   Begin VB.Line Line1 
      X1              =   1328
      X2              =   6908
      Y1              =   2678
      Y2              =   2678
   End
End
Attribute VB_Name = "B004"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private RSimpan, RSave, RGol As rdoResultset
Private SSimpan, SSave, SGol As String

Private AAA, BBB As String

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Combo1_LostFocus()
If Combo1 = "" Then Exit Sub
SGol = "Select * from B003 where Kode_JNS = '" + Trim(Combo1) + "'"
Set RGol = RDCO.OpenResultset(SGol, rdOpenDynamic, rdConcurRowVer)
If RGol.RowCount <> 0 Then
    Label1 = RGol("Nama_JNS")
End If
RGol.Close
Set RGol = Nothing
Text4.SetFocus
Posisi = 1
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Combo2_LostFocus()
If Combo2 = "" Then Exit Sub
SGol = "Select * from B003 where Kode_JNS = '" + Trim(Combo1) + "'"
Set RGol = RDCO.OpenResultset(SGol, rdOpenDynamic, rdConcurRowVer)
If RGol.RowCount <> 0 Then
    Label2 = RGol("Nama_JNS")
End If
RGol.Close
Set RGol = Nothing
Text4.SetFocus
End Sub

Private Sub Command1_Click()
Frame1.Visible = False
Frame2.Visible = True

Combo1.Visible = True
Label1.Visible = True
Label4.Visible = True
Label5.Visible = True

Combo2.Visible = False
Label2.Visible = False
Label6.Visible = False
Label7.Visible = False

Combo1.SetFocus
Label3 = "DEBET"

End Sub

Private Sub Command2_Click()
Frame1.Visible = False
Frame2.Visible = True

Combo1.Visible = False
Label1.Visible = False
Label4.Visible = False
Label5.Visible = False

Combo2.Visible = True
Label2.Visible = True
Label6.Visible = True
Label7.Visible = True

Combo2.SetFocus
Label3 = "CREDIT"
End Sub

Private Sub Command3_Click()
Frame1.Visible = True
Frame2.Visible = False
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=DEALER", rdDriverNoPrompt, False, CN)

Frame1.Visible = True
Frame2.Visible = False

Call Kosong
Call IsiGol

Call SiapkanGrid
Call Tampilkan

NoNo = 0

End Sub

Private Sub SiapkanGrid()
With grid
    .Cols = 7
    .Row = 0
    .Col = 0: .ColWidth(0) = "500": .Text = "NO": .CellAlignment = 4
    .Col = 1: .ColWidth(1) = "1500": .Text = "KODE": .CellAlignment = 4
    .Col = 2: .ColWidth(2) = "2500": .Text = "NAMA": .CellAlignment = 4
    .Col = 3: .ColWidth(3) = "750": .Text = "J AWAL": .CellAlignment = 4
    .Col = 4: .ColWidth(4) = "750": .Text = "D": .CellAlignment = 4
    .Col = 5: .ColWidth(5) = "750": .Text = "C": .CellAlignment = 4
    .Col = 6: .ColWidth(6) = "750": .Text = "J AKHIR": .CellAlignment = 4
End With
End Sub

Private Sub Tampilkan()
Dim Brs As Integer
Brs = 1
NoNo = 1
SCari = "Select * from B003 where Kode_Ind = '152' "
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset, rdConcurReadOnly)
If RCari.RowCount <> 0 Then
RCari.MoveFirst
Do While Not RCari.EOF
    With grid
        .Rows = Brs + 1
        .Row = Brs
        .Col = 0: .Text = NoNo: .CellAlignment = 4
        .Col = 1: .Text = RCari("Kode_JNS")
        .Col = 2: .Text = RCari("NAMA_JNS")
        .Col = 3: .Text = RCari("JML_AWAL"): .CellAlignment = 4
        .Col = 4: .Text = RCari("JML_DBT"): .CellAlignment = 4
        .Col = 5: .Text = RCari("JML_CRD"): .CellAlignment = 4
        .Col = 6: .Text = RCari("JML_AKHIR"): .CellAlignment = 4
    End With
    Brs = Brs + 1
    NoNo = NoNo + 1
RCari.MoveNext
Loop
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub Kosong()
Label1 = ""
Label2 = ""
Text4 = ""
Text5 = ""
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text4_LostFocus()
If Text4 = "" Then Exit Sub
If Not IsNumeric(Text4) Then
    Text4.SetFocus
    MsgBox "NOMINAL MENGGUNAKAN TYPE ANGKA", vbCritical, "TYPE DATA SALAH"
    Text4 = ""
    Exit Sub
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text5_LostFocus()
Text5 = Format(Text5, ">")
End Sub

Private Sub TmbClose_Click()
Unload Me
End Sub

Private Sub TmbSave_Click()
Dim Tanya
If Label3 = "DEBET" Then
    Tanya = MsgBox("YAKIN PROSES TRANSAKSI BARANG ?", vbOKCancel, "KONFIRMASI")
    If Tanya = vbCancel Then Exit Sub
        Call JurnalD
ElseIf Label3 = "CREDIT" Then
    Tanya = MsgBox("YAKIN PROSES TRANSAKSI BARANG ?", vbOKCancel, "KONFIRMASI")
    If Tanya = vbCancel Then Exit Sub
        Call JurnalC
End If
Call Kosong
Unload Me
B004.Show
End Sub

Private Sub JurnalD()
SSimpan = "Select * From B003 where KODE_JNS = '" + Trim(Combo1) + "'"
Set RSimpan = RDCO.OpenResultset(SSimpan, rdOpenDynamic, rdConcurRowVer)
    AAA = CCur(RSimpan("JML_DBT")) + CCur(Text4)
    BBB = CCur(RSimpan("JML_AKHIR")) + CCur(Text4)
    RSimpan.EDIT
    RSimpan("JML_DBT") = AAA
    RSimpan("JML_AKHIR") = BBB
    
        SSave = "Select * From B004"
        Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
        RSave.AddNew
            RSave("Tanggal") = Tanggal
            RSave("Kode_JNS") = Trim(Combo1)
            RSave("Nama_JNS") = Trim(Label1)
            RSave("Debet") = CCur(Text4)
            RSave("Credit") = 0
            RSave("Saldo") = CCur(RSimpan("JML_AKHIR"))
            RSave("Keterangan") = Trim(Text5)
        RSave.Update
        RSave.Close
        Set RSave = Nothing

RSimpan.Update
RSimpan.Close
Set RSimpan = Nothing
End Sub

Private Sub JurnalC()
SSimpan = "Select * From B003 where KODE_JNS = '" + Trim(Combo1) + "'"
Set RSimpan = RDCO.OpenResultset(SSimpan, rdOpenDynamic, rdConcurRowVer)
    AAA = CCur(RSimpan("JML_CRD")) + CCur(Text4)
    BBB = CCur(RSimpan("JML_AKHIR")) - CCur(Text4)
    RSimpan.EDIT
    RSimpan("JML_CRD") = AAA
    RSimpan("JML_AKHIR") = BBB
    
        SSave = "Select * From B004"
        Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
        RSave.AddNew
            RSave("Tanggal") = Tanggal
            RSave("Kode_JNS") = Trim(Combo1)
            RSave("Nama_JNS") = Trim(Label2)
            RSave("Debet") = 0
            RSave("Credit") = CCur(Text4)
            RSave("Saldo") = CCur(RSimpan("JML_AKHIR"))
            RSave("Keterangan") = Trim(Text5)
        RSave.Update
        RSave.Close
        Set RSave = Nothing

RSimpan.Update
RSimpan.Close
Set RSimpan = Nothing
End Sub

Private Sub IsiGol()
Dim KodeG
SGol = "Select Kode_JNS From B003 order by Kode_JNS"
Set RGol = RDCO.OpenResultset(SGol, rdOpenDynamic, rdConcurRowVer)
If RGol.RowCount <> 0 Then
    RGol.MoveFirst
    Do While Not RGol.EOF
        Combo1.AddItem RGol("Kode_JNS")
        Combo2.AddItem RGol("Kode_JNS")
    RGol.MoveNext
    Loop
End If

RGol.Close
Set RGol = Nothing
Combo1.ListIndex = 0
Combo2.ListIndex = 0
End Sub


