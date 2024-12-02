VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form P001 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ENTRI KODE PIUTANG"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6645
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   6645
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
      Left            =   5565
      TabIndex        =   12
      Top             =   2153
      Width           =   960
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
      TabIndex        =   4
      Top             =   2153
      Width           =   960
   End
   Begin VB.TextBox Text4 
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
      Left            =   1680
      MaxLength       =   7
      TabIndex        =   3
      Text            =   "Text4"
      Top             =   1613
      Width           =   1005
   End
   Begin VB.TextBox Text3 
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
      Left            =   1680
      MaxLength       =   7
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   1118
      Width           =   1005
   End
   Begin VB.TextBox Text2 
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
      Left            =   1680
      MaxLength       =   35
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   623
      Width           =   4785
   End
   Begin VB.TextBox Text1 
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
      Left            =   1680
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   128
      Width           =   1005
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   2445
      Left            =   120
      TabIndex        =   11
      ToolTipText     =   "Klik untuk edit"
      Top             =   2730
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   4313
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
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Label6"
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
      Left            =   2730
      TabIndex        =   10
      Top             =   1658
      Width           =   3660
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Label5"
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
      Left            =   2730
      TabIndex        =   9
      Top             =   1163
      Width           =   3660
   End
   Begin VB.Label Label4 
      Caption         =   "SUB GL PDPT"
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
      Left            =   210
      TabIndex        =   8
      Top             =   1658
      Width           =   1320
   End
   Begin VB.Label Label3 
      Caption         =   "SUB GL PIUTANG"
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
      Left            =   210
      TabIndex        =   7
      Top             =   1163
      Width           =   1545
   End
   Begin VB.Label Label2 
      Caption         =   "NAMA PIUTANG"
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
      Left            =   210
      TabIndex        =   6
      Top             =   668
      Width           =   1365
   End
   Begin VB.Label Label1 
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
      Height          =   285
      Left            =   210
      TabIndex        =   5
      Top             =   173
      Width           =   1320
   End
End
Attribute VB_Name = "P001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private RKode, RGL, RSave As rdoResultset
Private SKode, SGL, SSave As String
Private Editing

Private Sub Command1_Click()
If Editing = 1 Then
    SSave = "Select * from P001"
    Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdOpenKeyset)
        RSave.AddNew
        RSave("Kode_Pin") = Text1
        RSave("Nama_Pin") = Text2
        RSave("SGL_Pin") = Text3
        RSave("SGL_Pdpt") = Text4
        RSave.Update
    RSave.Close
    Set RSave = Nothing
ElseIf Editing = 2 Then
    SSave = "Select * from P001 where Kode_Pin = '" + Text1 + "'"
    Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdOpenKeyset)
        RSave("Nama_Pin") = Text2
        RSave("SGL_Pin") = Text3
        RSave("SGL_PDPT") = Text4
        RSave.Update
    RSave.Close
    Set RSave = Nothing
End If
Call Tampilkan
Call Kosong
Text1.SetFocus
End Sub

Private Sub Command2_Click()
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
Editing = 1
Label5 = ""
Label6 = ""
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
End Sub

Private Sub SiapkanGrid()
With grid
    .Row = 0
    .Cols = 4
    .Col = 0: .ColWidth(0) = 1000: .Text = "KODE": .CellAlignment = 4: .CellFontBold = True
    .Col = 1: .ColWidth(1) = 2500: .Text = "NAMA PIUTANG": .CellAlignment = 4: .CellFontBold = True
    .Col = 2: .ColWidth(2) = 1250: .Text = "SGL PIUTANG": .CellAlignment = 4: .CellFontBold = True
    .Col = 3: .ColWidth(3) = 1250: .Text = "SGL PDPT": .CellAlignment = 4: .CellFontBold = True
End With
End Sub

Private Sub Tampilkan()
Dim Brs
Brs = 1
SKode = "Select * From P001 order by Kode_Pin"
Set RKode = RDCO.OpenResultset(SKode, rdOpenDynamic, rdOpenKeyset)
If RKode.RowCount <> 0 Then
RKode.MoveFirst
Do Until RKode.EOF
    With grid
        .Rows = Brs + 1
        .Row = Brs
        .Col = 0: .Text = RKode("Kode_Pin"): .CellAlignment = 4
        .Col = 1: .Text = RKode("Nama_Pin")
        .Col = 2: .Text = RKode("SGL_Pin"): .CellAlignment = 4
        .Col = 3: .Text = RKode("SGL_Pdpt"): .CellAlignment = 4
        Brs = Brs + 1
    End With
RKode.MoveNext
Loop
End If
RKode.Close
Set RKode = Nothing
End Sub

Private Sub text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text1_LostFocus()
Dim Tanya
If Text1 = "" Then Exit Sub
SKode = "Select * From P001 where Kode_Pin = '" + Text1 + "'"
Set RKode = RDCO.OpenResultset(SKode, rdOpenDynamic, rdOpenKeyset)
If RKode.RowCount <> 0 Then
    Tanya = MsgBox("KODE HUTANG TELAH TERDAFTAR", vbOKCancel, "ANDA AKAN MELAKUKAN EDIT")
    If Tanya = vbOK Then
        Text2 = RKode("Nama_Pin")
        Text3 = RKode("SGL_Pin")
        Text4 = RKode("SGL_Pdpt")
        Call Text3_LostFocus
        Call Text4_LostFocus
        Editing = 2
    Else
        Text1.SetFocus
        Call Kosong
    End If
End If
RKode.Close
Set RKode = Nothing
Text1 = Format(Text1, ">")
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
SGL = "Select NamaSL from G003 where CodeSL = '" + Text3 + "'"
Set RGL = RDCO.OpenResultset(SGL, rdOpenDynamic, rdOpenKeyset)
If RGL.RowCount <> 0 Then
    Label5 = RGL("NamaSL")
Else
    Text3.SetFocus
    MsgBox "KODE SGL BELUM TERDAFTAR", vbInformation, "KODE SL BELUM TERDAFTAR"
End If
RGL.Close
Set RGL = Nothing
End Sub

Private Sub text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text4_LostFocus()
If Text4 = "" Then Exit Sub
SGL = "Select NamaSL from G003 where CodeSL = '" + Text4 + "'"
Set RGL = RDCO.OpenResultset(SGL, rdOpenDynamic, rdOpenKeyset)
If RGL.RowCount <> 0 Then
    Label6 = RGL("NamaSL")
Else
    Text4.SetFocus
    MsgBox "KODE SGL BELUM TERDAFTAR", vbInformation, "KODE SL BELUM TERDAFTAR"
End If
RGL.Close
Set RGL = Nothing
End Sub


