VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form B001 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "KODE KATEGORI BARANG"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   8115
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   8115
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
      Left            =   6840
      TabIndex        =   18
      Top             =   2985
      Width           =   1050
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   2549
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   150
      Width           =   690
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   2549
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   600
      Width           =   4335
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2549
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   1050
      Width           =   960
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2549
      TabIndex        =   3
      Text            =   "Text4"
      Top             =   1500
      Width           =   960
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2549
      TabIndex        =   4
      Text            =   "Text5"
      Top             =   1950
      Width           =   960
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2549
      TabIndex        =   5
      Text            =   "Text6"
      Top             =   2400
      Width           =   960
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
      Left            =   210
      TabIndex        =   6
      Top             =   2985
      Width           =   1050
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   2115
      Left            =   75
      TabIndex        =   17
      ToolTipText     =   "Klik untuk edit"
      Top             =   3600
      Width           =   7950
      _ExtentX        =   14023
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
   Begin VB.Label Label1 
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
      Height          =   300
      Left            =   885
      TabIndex        =   16
      Top             =   180
      Width           =   1590
   End
   Begin VB.Label Label2 
      Caption         =   "NAMA GOLONGAN"
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
      Left            =   885
      TabIndex        =   15
      Top             =   630
      Width           =   1590
   End
   Begin VB.Label Label3 
      Caption         =   "SGL PERSEDIAAN"
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
      Left            =   885
      TabIndex        =   14
      Top             =   1080
      Width           =   1590
   End
   Begin VB.Label Label4 
      Caption         =   "SGL PENDAPATAN"
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
      Left            =   885
      TabIndex        =   13
      Top             =   1530
      Width           =   1590
   End
   Begin VB.Label Label5 
      Caption         =   "SGL DISKON JUAL"
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
      Left            =   885
      TabIndex        =   12
      Top             =   1980
      Width           =   1590
   End
   Begin VB.Label Label6 
      Caption         =   "SGL DISKON BELI"
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
      Left            =   885
      TabIndex        =   11
      Top             =   2430
      Width           =   1590
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3630
      TabIndex        =   10
      Top             =   1095
      Width           =   3300
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label8"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3630
      TabIndex        =   9
      Top             =   1545
      Width           =   3300
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3630
      TabIndex        =   8
      Top             =   1995
      Width           =   3300
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label10"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3630
      TabIndex        =   7
      Top             =   2445
      Width           =   3300
   End
End
Attribute VB_Name = "B001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private RCari, RSave, RSGL As rdoResultset
Private SCari, SSave, SSGL, KodeSGL As String

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=DEALER", rdDriverNoPrompt, False, CN)

Call Kosong
Call SiapkanGrid
Call Tampilkan
End Sub

Private Sub Kosong()
Label7 = ""
Label8 = ""
Label9 = ""
Label10 = ""
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""
End Sub

Private Sub SiapkanGrid()
With grid
    .Cols = 6
    .Row = 0
    .Col = 0: .ColWidth(0) = "1000": .Text = "KODE": .CellAlignment = 4
    .Col = 1: .ColWidth(1) = "2500": .Text = "NAMA": .CellAlignment = 4
    .Col = 2: .ColWidth(2) = "1000": .Text = "SGL SEDIA": .CellAlignment = 4
    .Col = 3: .ColWidth(3) = "1000": .Text = "SGL PDPT": .CellAlignment = 4
    .Col = 4: .ColWidth(4) = "1000": .Text = "DISC JUAL": .CellAlignment = 4
    .Col = 5: .ColWidth(5) = "1000": .Text = "DISC BELI": .CellAlignment = 4
End With
End Sub

Private Sub Tampilkan()
Dim Brs As Integer
Brs = 1
SCari = "Select * from B001 order by Kode_Ind"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset, rdConcurReadOnly)
If RCari.RowCount <> 0 Then
RCari.MoveFirst
Do While Not RCari.EOF
    With grid
        .Rows = Brs + 1
        .Row = Brs
        .Col = 0: .Text = RCari("Kode_Ind"): .CellAlignment = 4
        .Col = 1: .Text = RCari("Keterangan")
        .Col = 2: .Text = RCari("SGL_Sedia"): .CellAlignment = 4
        .Col = 3: .Text = RCari("SGL_PDPT"): .CellAlignment = 4
        .Col = 4: .Text = RCari("SGL_DisJual"): .CellAlignment = 4
        .Col = 5: .Text = RCari("SGL_DisBeli"): .CellAlignment = 4
    End With
    Brs = Brs + 1
RCari.MoveNext
Loop
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub TblSave_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Then
    MsgBox "MASIH ADA DATA YANG KOSONG", vbCritical, "DATA TIDAK BOLEH KOSONG"
    Exit Sub
End If

SSave = "Select * From B001"
Set RSave = RDCO.OpenResultset(SSave, rdOpenKeyset, rdConcurRowVer)
RSave.AddNew
    RSave("Kode_Ind") = Text1
    RSave("Keterangan") = Text2
    RSave("SGL_Sedia") = Text3
    RSave("SGL_PDPT") = Text4
    RSave("SGL_Disjual") = Text5
    RSave("SGL_DisBeli") = Text6
RSave.Update
RSave.Close
Set RSave = Nothing
Call Kosong
Call Tampilkan
Text1.SetFocus
End Sub

Private Sub text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text1_LostFocus()
If Text1 = "" Then Exit Sub
SCari = "Select * from B001 where Kode_ind = '" + Text1 + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset, rdConcurReadOnly)
If RCari.RowCount <> 0 Then
    Text1.SetFocus
    MsgBox "DATA SUDAH TERDAFTAR", vbInformation, "DATA SUDAH TERDAFTAR"
    Exit Sub
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text2_LostFocus()
Text2 = Format(Text2, ">")
End Sub

Private Sub text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text3_LostFocus()
If Text3 = "" Then Exit Sub
SSGL = "Select NamaSL from G003 where CodeSL = '" + Text3 + "'"
Set RSGL = RDCO.OpenResultset(SSGL, rdOpenKeyset, rdConcurReadOnly)
If RSGL.RowCount <> 0 Then
    Label7 = RSGL("NamaSL")
Else
    Text3.SetFocus
    MsgBox "KODE SUB GENERAL LEDGER BELUM TERDAFTAR", vbInformation, "KODE SGL BELUM TERDAFTAR"
End If
RSGL.Close
Set RSGL = Nothing
End Sub

Private Sub text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text4_LostFocus()
If Text4 = "" Then Exit Sub
SSGL = "Select NamaSL from G003 where CodeSL = '" + Text4 + "'"
Set RSGL = RDCO.OpenResultset(SSGL, rdOpenKeyset, rdConcurReadOnly)
If RSGL.RowCount <> 0 Then
    Label8 = RSGL("NamaSL")
Else
    Text4.SetFocus
    MsgBox "KODE SUB GENERAL LEDGER BELUM TERDAFTAR", vbInformation, "KODE SGL BELUM TERDAFTAR"
End If
RSGL.Close
Set RSGL = Nothing
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text5_LostFocus()
If Text5 = "" Then Exit Sub
SSGL = "Select NamaSL from G003 where CodeSL = '" + Text5 + "'"
Set RSGL = RDCO.OpenResultset(SSGL, rdOpenKeyset, rdConcurReadOnly)
If RSGL.RowCount <> 0 Then
    Label9 = RSGL("NamaSL")
Else
    Text5.SetFocus
    MsgBox "KODE SUB GENERAL LEDGER BELUM TERDAFTAR", vbInformation, "KODE SGL BELUM TERDAFTAR"
End If
RSGL.Close
Set RSGL = Nothing
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text6_LostFocus()
If Text6 = "" Then Exit Sub
SSGL = "Select NamaSL from G003 where CodeSL = '" + Text6 + "'"
Set RSGL = RDCO.OpenResultset(SSGL, rdOpenKeyset, rdConcurReadOnly)
If RSGL.RowCount <> 0 Then
    Label10 = RSGL("NamaSL")
Else
    Text6.SetFocus
    MsgBox "KODE SUB GENERAL LEDGER BELUM TERDAFTAR", vbInformation, "KODE SGL BELUM TERDAFTAR"
End If
RSGL.Close
Set RSGL = Nothing
End Sub

