VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form C013 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INFORMASI  CUSTOMER / SUPPLIER / DEBITUR / KREDITUR"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   11625
   StartUpPosition =   2  'CenterScreen
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
      Left            =   10620
      TabIndex        =   5
      Top             =   5250
      Width           =   960
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   2070
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   5280
      Width           =   2040
   End
   Begin VB.CheckBox Check1 
      Caption         =   "CARI UNTUK NAMA"
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
      Left            =   45
      TabIndex        =   3
      Top             =   5325
      Width           =   1860
   End
   Begin VB.CheckBox Check2 
      Caption         =   "CARI UNTUK NOMOR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4410
      TabIndex        =   2
      Top             =   5280
      Width           =   2040
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   6585
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   5280
      Width           =   1860
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5100
      Left            =   45
      TabIndex        =   0
      Top             =   105
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   8996
      _Version        =   393216
   End
End
Attribute VB_Name = "C013"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private RCari1, RCari As rdoResultset
Private SCari1, SCari As String

Private Sub Check1_Click()
If Check1.Value = 1 Then
    Text1.Enabled = True
    Text1.BackColor = &HFFFFC0
    Text1.SetFocus
    Text2 = ""
    Text2.Enabled = False
    Check2.Value = 0
    Check2.Enabled = False
Else
    Check2.Enabled = True
    Text1 = ""
    Text2 = ""
    Text1.Enabled = False
    Text2.Enabled = False
    Text1.BackColor = &HC0C0C0
    Text2.BackColor = &HC0C0C0
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
    Text2.Enabled = True
    Text2.BackColor = &HFFFFC0
    Text2.SetFocus
    Text2 = ""
    Text1.Enabled = False
    Check1.Value = 0
    Check1.Enabled = False
Else
    Check1.Enabled = True
    Text1 = ""
    Text2 = ""
    Text1.Enabled = False
    Text2.Enabled = False
    Text1.BackColor = &HC0C0C0
    Text2.BackColor = &HC0C0C0
End If
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=DEALER", rdDriverNoPrompt, False, CN)

Call SiapGrid
Text1 = ""
Text2 = ""
Check1.Value = 0
Check2.Value = 0
Call Check1_Click
Call Check2_Click
Call Tampilkan
End Sub

Private Sub SiapGrid()
With grid
    .Row = 0
    .RowHeight(0) = 400
    .Cols = 5
    .Col = 0: .ColWidth(0) = 800: .Text = "NOMOR": .CellAlignment = 4: .CellFontBold = True
    .Col = 1: .ColWidth(1) = 2500: .Text = "N A M A": .CellAlignment = 4: .CellFontBold = True
    .Col = 2: .ColWidth(2) = 4000: .Text = "A L A M A T": .CellAlignment = 4: .CellFontBold = True
    .Col = 3: .ColWidth(3) = 1700: .Text = "K O T A": .CellAlignment = 4: .CellFontBold = True
    .Col = 4: .ColWidth(4) = 2000: .Text = "NO. TELEPON": .CellAlignment = 4: .CellFontBold = True
End With
End Sub

Private Sub Tampilkan()
Dim Brs, Nama
Brs = 1
SCari = "Select NoNas, Nama, Alamat1, Kota, Telpon from C012 order by Nama"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
RCari.MoveFirst
Do Until RCari.EOF
    With grid
        .Rows = Brs + 1
        .Row = Brs
        .Col = 0: .Text = RCari("NoNas"): .CellAlignment = 4
        .Col = 1: .Text = RCari("Nama")
        .Col = 2: .Text = RCari("Alamat1")
        .Col = 3: .Text = RCari("Kota")
        .Col = 4: .Text = RCari("Telpon"): .CellAlignment = 4
        Brs = Brs + 1
    End With
RCari.MoveNext
Loop
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub grid_dblClick()
No_Nas = grid.TextMatrix(grid.Row, 0)
Nama_Nas = grid.TextMatrix(grid.Row, 1)
Unload Me
End Sub

Private Sub text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text1_LostFocus()
If Text1 = "" Then Exit Sub
Dim Brs, Nama
Brs = 1
Nama = "%" + Text1 + "%"
SCari = "Select NoNas, Nama, Alamat1, Kota, Telpon from C012 where Nama like '" + Trim(Nama) + "' order by Nama"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
RCari.MoveFirst
Do Until RCari.EOF
    With grid
        .Rows = Brs + 1
        .Row = Brs
        .Col = 0: .Text = RCari("NoNas"): .CellAlignment = 4
        .Col = 1: .Text = RCari("Nama")
        .Col = 2: .Text = RCari("Alamat1")
        .Col = 3: .Text = RCari("Kota")
        .Col = 4: .Text = RCari("Telpon"): .CellAlignment = 4
        Brs = Brs + 1
    End With
RCari.MoveNext
Loop
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text2_LostFocus()
If Text2 = "" Then Exit Sub
Dim Brs, No
Brs = 1
No = "%" + Text2 + "%"
SCari = "Select NoNas, Nama, Alamat1, Kota, Telpon from C012 where NoNas like '" + Trim(No) + "' order by NoNas"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
RCari.MoveFirst
Do Until RCari.EOF
    With grid
        .Rows = Brs + 1
        .Row = Brs
        .Col = 0: .Text = RCari("NoNas"): .CellAlignment = 4
        .Col = 1: .Text = RCari("Nama")
        .Col = 2: .Text = RCari("Alamat1")
        .Col = 3: .Text = RCari("Kota")
        .Col = 4: .Text = RCari("Telpon"): .CellAlignment = 4
        Brs = Brs + 1
    End With
RCari.MoveNext
Loop
End If
RCari.Close
Set RCari = Nothing
End Sub

