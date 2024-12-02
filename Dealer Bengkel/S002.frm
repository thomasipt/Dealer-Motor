VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form S002 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PELAYANAN SERVICE KENDARAAN"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6405
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4545
   ScaleMode       =   0  'User
   ScaleWidth      =   6405
   StartUpPosition =   2  'CenterScreen
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
      Left            =   5175
      TabIndex        =   9
      Top             =   1785
      Width           =   1050
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
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
      Left            =   4177
      TabIndex        =   1
      Text            =   "Text3"
      Top             =   1005
      Width           =   2025
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   1867
      TabIndex        =   0
      Text            =   "Text2"
      Top             =   555
      Width           =   4335
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   2115
      Left            =   165
      TabIndex        =   3
      ToolTipText     =   "Klik untuk edit"
      Top             =   2310
      Width           =   6060
      _ExtentX        =   10689
      _ExtentY        =   3731
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
   Begin VB.CommandButton TblEdit 
      Caption         =   "EDIT"
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
      Left            =   165
      TabIndex        =   8
      Top             =   1740
      Width           =   1050
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
      Left            =   165
      TabIndex        =   2
      Top             =   1740
      Width           =   1050
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080C0FF&
      Caption         =   "KODE"
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
      Left            =   1860
      TabIndex        =   7
      Top             =   135
      Width           =   2640
   End
   Begin VB.Line Line1 
      X1              =   150
      X2              =   6255
      Y1              =   1575
      Y2              =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "BIAYA JASA Rp"
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
      Left            =   2400
      TabIndex        =   6
      Top             =   1065
      Width           =   1590
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080C0FF&
      Caption         =   "TYPE SERVICE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   202
      TabIndex        =   5
      Top             =   585
      Width           =   1590
   End
   Begin VB.Label Label1 
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
      Height          =   300
      Left            =   202
      TabIndex        =   4
      Top             =   135
      Width           =   1590
   End
End
Attribute VB_Name = "S002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private RCari, RSave, RNo As rdoResultset
Private SCari, SSave, SNo As String

Private EDIT

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=DEALER", rdDriverNoPrompt, False, CN)

Call Kosong
Call SiapkanGrid
Call KODE
Call Tampilkan
TblSave.ZOrder

EDIT = 0
End Sub

Private Sub KODE()
Dim AutoNomor As Double
SNo = "Select Kode_S from S002 order by Kode_S DESC"
Set RNo = RDCO.OpenResultset(SNo, rdOpenDynamic, rdConcurRowVer)
If RNo.RowCount <> 0 Then
    AutoNomor = Mid(RNo("Kode_S"), 5, 8) + 1
    Label4 = "IPT." + Trim(Digit(4, AutoNomor))
Else
    Label4 = "IPT.0001"
End If
RNo.Close
Set RNo = Nothing
End Sub

Private Sub Kosong()
Text2 = ""
Text3 = ""
End Sub

Private Sub SiapkanGrid()
With grid
    .Cols = 3
    .Row = 0
    .Col = 0: .ColWidth(0) = "1000": .Text = "KODE": .CellAlignment = 4
    .Col = 1: .ColWidth(1) = "3000": .Text = "TYPE SERVICE": .CellAlignment = 4
    .Col = 2: .ColWidth(2) = "1500": .Text = "BIAYA JASA": .CellAlignment = 4
End With
End Sub

Private Sub Tampilkan()
Dim Brs As Integer
Brs = 1
SCari = "Select * from S002 order by Kode_S"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset, rdConcurReadOnly)
If RCari.RowCount <> 0 Then
RCari.MoveFirst
Do While Not RCari.EOF
    With grid
        .Rows = Brs + 1
        .Row = Brs
        .Col = 0: .Text = RCari("Kode_S"): .CellAlignment = 4
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

Private Sub grid_dblClick()
Dim Tanya2
Tanya2 = MsgBox("ANDA YAKIN MELAKUKAN EDIT DATA PELAYANAN", vbSystemModal, "KONFIRMASI")
If Tanya2 = vbOK Then
    Label4 = grid.TextMatrix(grid.Row, 0)
    Text2 = grid.TextMatrix(grid.Row, 1)
    Text3 = grid.TextMatrix(grid.Row, 2)
    TblEdit.ZOrder
    EDIT = 1
End If
End Sub

Private Sub TblEdit_Click()
Dim Tanya

If Text2 = "" Or Text3 = "" Then
    MsgBox "MASIH ADA DATA YANG KOSONG", vbCritical, "DATA TIDAK BOLEH KOSONG"
    Exit Sub
End If

Tanya = MsgBox("MASUKKAN DATA PELAYANAN", vbSystemModal, "KONFIRMASI")
If Tanya = vbOK Then
    SSave = "Select * From S002 where KODE_S='" + Trim(Label4) + "'"
    Set RSave = RDCO.OpenResultset(SSave, rdOpenKeyset, rdConcurRowVer)
    RSave.EDIT
        RSave("Nama_S") = Text2
        RSave("Biaya_S") = CCur(Text3)
    RSave.Update
    RSave.Close
    Set RSave = Nothing
Else
    Exit Sub
End If

Unload Me
S002.Show 1
End Sub

Private Sub TblSave_Click()
Dim Tanya

If Text2 = "" Or Text3 = "" Then
    MsgBox "MASIH ADA DATA YANG KOSONG", vbCritical, "DATA TIDAK BOLEH KOSONG"
    Text2.SetFocus
    Exit Sub
End If

Tanya = MsgBox("MASUKKAN DATA PELAYANAN", vbSystemModal, "KONFIRMASI")
If Tanya = vbOK Then
    SSave = "Select * From S002"
    Set RSave = RDCO.OpenResultset(SSave, rdOpenKeyset, rdConcurRowVer)
    RSave.AddNew
        RSave("Kode_S") = Label4
        RSave("Nama_S") = Text2
        RSave("Biaya_S") = CCur(Text3)
    RSave.Update
    RSave.Close
    Set RSave = Nothing
Else
    Exit Sub
End If

Unload Me
S002.Show 1

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text2_LostFocus()
Text2 = Format(Text2, ">")
End Sub

Private Sub text3_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If

If KeyAscii = 13 Then
    If EDIT = 0 Then
        TblSave.SetFocus
    ElseIf EDIT = 1 Then
        TblEdit.SetFocus
    End If
    Text3 = Format(Text3, "##,###.00")
End If

End Sub
