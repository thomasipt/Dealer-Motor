VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form PO2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CETAK PURCHASE ORDER"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   7785
   StartUpPosition =   2  'CenterScreen
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
      Left            =   3412
      TabIndex        =   21
      Top             =   2625
      Width           =   960
   End
   Begin VB.Frame Frame1 
      Caption         =   "Edit Tabel"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   510
      TabIndex        =   17
      Top             =   5985
      Width           =   7155
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4725
         TabIndex        =   20
         Text            =   "Text7"
         Top             =   315
         Width           =   2295
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2415
         TabIndex        =   19
         Text            =   "Text6"
         Top             =   315
         Width           =   2295
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   105
         TabIndex        =   18
         Text            =   "Text5"
         Top             =   315
         Width           =   2295
      End
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2632
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   3255
      Width           =   2512
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   112
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   3255
      Width           =   2500
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   5172
      TabIndex        =   3
      Text            =   "Text4"
      Top             =   3255
      Width           =   2500
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1470
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1155
      Width           =   4740
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   3180
      Left            =   105
      TabIndex        =   8
      Top             =   3555
      Width           =   7560
      _ExtentX        =   13335
      _ExtentY        =   5609
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   65280
      ForeColorFixed  =   0
      BackColorBkg    =   16777152
      Appearance      =   0
   End
   Begin VB.CommandButton OK 
      Caption         =   "OK"
      Height          =   285
      Left            =   6615
      TabIndex        =   4
      Top             =   3885
      Width           =   435
   End
   Begin VB.Shape Shape1 
      Height          =   855
      Left            =   45
      Top             =   105
      Width           =   7680
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "JUMLAH UNIT"
      Height          =   285
      Left            =   5205
      TabIndex        =   16
      Top             =   210
      Width           =   2400
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "TGL BELI"
      Height          =   285
      Left            =   2685
      TabIndex        =   15
      Top             =   210
      Width           =   2400
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "NO FAK"
      Height          =   285
      Left            =   165
      TabIndex        =   14
      Top             =   210
      Width           =   2400
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
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
      Left            =   5205
      TabIndex        =   13
      Top             =   525
      Width           =   2400
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
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
      Left            =   2685
      TabIndex        =   12
      Top             =   525
      Width           =   2400
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
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
      Left            =   165
      TabIndex        =   11
      Top             =   525
      Width           =   2400
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Label4"
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
      Left            =   1470
      TabIndex        =   10
      Top             =   2100
      Width           =   6180
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
      Left            =   1470
      TabIndex        =   9
      Top             =   1680
      Width           =   3975
   End
   Begin VB.Label Label3 
      Caption         =   "ALAMAT"
      Height          =   300
      Left            =   105
      TabIndex        =   7
      Top             =   2100
      Width           =   1545
   End
   Begin VB.Label Label2 
      Caption         =   "DARI DEALER"
      Height          =   300
      Left            =   105
      TabIndex        =   6
      Top             =   1680
      Width           =   1365
   End
   Begin VB.Label Label1 
      Caption         =   "KEPADA"
      Height          =   285
      Left            =   105
      TabIndex        =   5
      Top             =   1230
      Width           =   1320
   End
End
Attribute VB_Name = "PO2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private REG, RPO, RGrid, RSave As rdoResultset
Private SEG, SPO, SGrid, SSave As String
Private No, TIPE, WARNA, UNIT

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=DEALER", rdDriverNoPrompt, False, CN)

Call CekPO2

Call Kosong
Call SiapkanGrid
Call Isi

No = 0
Frame1.Visible = False

End Sub

Private Sub CekPO2()
Dim Tanya
SPO = "Select NO_FAK from PO2 where NO_FAK = '" + Trim(FAKTUR) + "'"
Set RPO = RDCO.OpenResultset(SPO, rdOpenDynamic, rdConcurRowVer)
If RPO.RowCount <> 0 Then
    Tanya = MsgBox("FAK. " + Trim(FAKTUR) + " TELAH DIBUATKAN PO, LAKUKAN EDIT PO ?", vbOKCancel, "KONFIRMASI")
    If Tanya = vbOK Then
        Call Kosong
        Call SiapkanGrid
        Call IsiGrid
        Text2.Visible = True
        Text3.Visible = True
        Text4.Visible = True
    Else
        FAKTUR = ""
        TANGGAL_FAKTUR = ""
        JML_UNIT = ""
    End If
End If
RPO.Close
Set RPO = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
PO1.Show
End Sub

Private Sub Isi()
Label5 = N_CCAB
Label4 = N_ALAMAT
Label6 = FAKTUR
Label7 = TANGGAL_FAKTUR
Label8 = JML_UNIT
End Sub

Private Sub Kosong()
Text1 = ""
Label5 = ""
Label4 = ""
Text2 = ""
Text3 = ""
Text4 = ""
End Sub

Private Sub SiapkanGrid()
With grid
    .Row = 0
    .Cols = 3
    .Col = 0: .ColWidth(0) = 2500: .Text = "TYPE": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 12
    .Col = 1: .ColWidth(1) = 2512: .Text = "WARNA": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 12
    .Col = 2: .ColWidth(2) = 2500: .Text = "UNIT": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 12
End With
End Sub

Private Sub Grid_dblClick()
TIPE = ""
WARNA = ""
UNIT = ""

TIPE = grid.TextMatrix(grid.Row, 0)
WARNA = grid.TextMatrix(grid.Row, 1)
UNIT = grid.TextMatrix(grid.Row, 2)

Text5 = TIPE
Text6 = WARNA
Text7 = UNIT

Frame1.Visible = True
Frame1.ZOrder
End Sub

Private Sub OK_Click()
No = No + Trim(Text4)
SSave = "Select * from PO2"
Set RSave = RDCO.OpenResultset(SSave, rdOpenKeyset, rdConcurRowVer)
RSave.AddNew
    RSave("NO_FAK") = Trim(Label6)
    RSave("TGL_FAK") = Trim(Label7)
    RSave("JML") = Trim(Label8)
    RSave("NO_URUT") = No
    RSave("TYPE") = Trim(Text2)
    RSave("WARNA") = Trim(Text3)
    RSave("UNIT") = Trim(Text4)
    RSave("KPD") = Trim(Text1)
    RSave("DARI") = Trim(Label5)
    RSave("ALAMAT") = Trim(Label4)
RSave.Update
RSave.Close
Set RSave = Nothing

Call IsiGrid
Text2 = ""
Text3 = ""
Text4 = ""
Text2.SetFocus
    If No >= Label8 Then
        Text2.Visible = False
        Text3.Visible = False
        Text4.Visible = False
        Exit Sub
    End If
End Sub

Private Sub IsiGrid()
SGrid = "Select * From PO2 where NO_FAK= '" + Trim(FAKTUR) + "'"
Set RGrid = RDCO.OpenResultset(SGrid, rdOpenKeyset, rdConcurReadOnly)
If RGrid.RowCount <> 0 Then
   RGrid.MoveFirst
   B = 1
   Do Until RGrid.EOF
      grid.Rows = B + 1
      grid.Row = B
         With grid
              .Col = 0: .Text = RGrid("TYPE"): .CellAlignment = 4
              .Col = 1: .Text = RGrid("WARNA"): .CellAlignment = 4
              .Col = 2: .Text = RGrid("UNIT"): .CellAlignment = 4
         End With
      B = B + 1
      RGrid.MoveNext
   Loop
End If
RGrid.Close
Set RGrid = Nothing

End Sub

Private Sub text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text1 = Format(Text1, ">")
    SendKeys vbTab
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text2 = Format(Text2, ">")
    SendKeys vbTab
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text3 = Format(Text3, ">")
    SendKeys vbTab
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text4 = Format(Text4, ">")
    SendKeys vbTab
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text5 = Format(Text5, ">")
    SendKeys vbTab
End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text6 = Format(Text6, ">")
    SendKeys vbTab
End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text7 = Format(Text7, ">")
    SendKeys vbTab
    Frame1.Visible = False
    Call EditGrid
    Call IsiGrid
End If
End Sub

Private Sub EditGrid()
SEG = "Select * From PO2 where Type = '" + Trim(TIPE) + "'"
Set REG = RDCO.OpenResultset(SEG, rdOpenDynamic, rdConcurRowVer)
REG.EDIT
    REG("Type") = Text5
    REG("Warna") = Text6
    REG("Unit") = CCur(Text7)
REG.Update
REG.Close
Set REG = Nothing
No = No - Trim(Text7)

    If No < Label8 Then
        Text2.Visible = True
        Text3.Visible = True
        Text4.Visible = True
        Text2 = ""
        Text3 = ""
        Text4 = ""
        Text2.SetFocus
        Command1.SetFocus
        Exit Sub
    End If
End Sub
