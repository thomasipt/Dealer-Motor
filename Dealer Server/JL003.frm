VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form JL003 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TABEL STOCK KENDARAAN"
   ClientHeight    =   8475
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13665
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   13665
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
      Left            =   12630
      TabIndex        =   5
      Top             =   7957
      Width           =   960
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7680
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   8002
      Width           =   3795
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1815
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   8002
      Width           =   3795
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   7725
      Left            =   75
      TabIndex        =   0
      Top             =   97
      Width           =   13515
      _ExtentX        =   23839
      _ExtentY        =   13626
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   65280
      ForeColorFixed  =   0
      BackColorBkg    =   16777152
      GridColor       =   8421504
      FocusRect       =   0
      AllowUserResizing=   3
      Appearance      =   0
   End
   Begin VB.Label Label2 
      Caption         =   "No Mesin"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6120
      TabIndex        =   4
      Top             =   8070
      Width           =   1905
   End
   Begin VB.Label Label1 
      Caption         =   "No Rangka"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   8070
      Width           =   1905
   End
End
Attribute VB_Name = "JL003"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private RGrid As rdoResultset
Private SGrid As String

Private NoNo

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=DEALER", rdDriverNoPrompt, False, CN)

Call SiapkanGrid
Call IsiGrid

Text1 = ""
Text2 = ""
NoNo = 0

End Sub

Private Sub SiapkanGrid()
With grid
    .Row = 0
    .Cols = 8
    .Col = 0: .ColWidth(0) = 500: .Text = "NO": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 1: .ColWidth(1) = 2500: .Text = "TYPE": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 2: .ColWidth(2) = 1500: .Text = "WARNA": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 3: .ColWidth(3) = 1500: .Text = "TAHUN": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 4: .ColWidth(4) = 1500: .Text = "NO. RANGKA": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 5: .ColWidth(5) = 1500: .Text = "NO. MESIN": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 6: .ColWidth(6) = 1500: .Text = "TGL_INPUT": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 7: .ColWidth(7) = 2500: .Text = "HARGA": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
End With
End Sub
    
Private Sub IsiGrid()
SGrid = "Select * From M001 where STS_JUAL = '0' order by RANGKA Asc"
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
            .Col = 1: .Text = RGrid("TYPE")
            .Col = 2: .Text = RGrid("WARNA"): .CellAlignment = 4
            .Col = 3: .Text = RGrid("TAHUN"): .CellAlignment = 4
            .Col = 4: .Text = RGrid("RANGKA"): .CellAlignment = 4
            .Col = 5: .Text = RGrid("MESIN"): .CellAlignment = 4
            .Col = 6: .Text = RGrid("TGL_INPUT"): .CellAlignment = 4
            .Col = 7: .Text = Format(RGrid("H_BELI"), "##,###.00"): .CellFontBold = True
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
    NO_RANGKA = ""
    NO_RANGKA = grid.TextMatrix(grid.Row, 4)
    JL004.Show
    Unload Me
End Select
End Sub

Private Sub grid_dblClick()
NO_RANGKA = ""
NO_RANGKA = grid.TextMatrix(grid.Row, 4)
Unload Me
JL004.Show
End Sub

Private Sub text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text1_LostFocus()
If Text1 = "" Then Exit Sub
Dim Brs, IndekNama
Brs = 1
IndekNama = "%" + Text1 + "%"
SGrid = "Select * From M001 where Rangka like '" + IndekNama + "' and STS_JUAL = '0' "
Set RGrid = RDCO.OpenResultset(SGrid, rdOpenKeyset, rdConcurReadOnly)
If RGrid.RowCount <> 0 Then
   RGrid.MoveFirst
   B = 1
   Do Until RGrid.EOF
      grid.Rows = B + 1
      grid.Row = B
         With grid
            .Col = 0: .Text = RGrid("CCAB"): .CellAlignment = 4
            .Col = 1: .Text = RGrid("TYPE"): .CellAlignment = 4
            .Col = 2: .Text = RGrid("WARNA"): .CellAlignment = 4
            .Col = 3: .Text = RGrid("TAHUN"): .CellAlignment = 4
            .Col = 4: .Text = RGrid("RANGKA"): .CellAlignment = 4
            .Col = 5: .Text = RGrid("MESIN"): .CellAlignment = 4
            .Col = 6: .Text = RGrid("TGL_INPUT"): .CellAlignment = 4
            .Col = 7: .Text = Format(RGrid("H_BELI"), "##,###.00")
         End With
      B = B + 1
      RGrid.MoveNext
   Loop
Else
    grid.Clear
    Call SiapkanGrid
End If
RGrid.Close
Set RGrid = Nothing
Text1 = Format(Text1, ">")
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text2_LostFocus()
If Text2 = "" Then Exit Sub
Dim Brs, IndekNama
Brs = 1
IndekNama = "%" + Text2 + "%"
SGrid = "Select * From M001 where Mesin like '" + IndekNama + "' and STS_JUAL = '0' "
Set RGrid = RDCO.OpenResultset(SGrid, rdOpenKeyset, rdConcurReadOnly)
If RGrid.RowCount <> 0 Then
   RGrid.MoveFirst
   B = 1
   Do Until RGrid.EOF
      grid.Rows = B + 1
      grid.Row = B
         With grid
            .Col = 0: .Text = RGrid("CCAB"): .CellAlignment = 4
            .Col = 1: .Text = RGrid("TYPE"): .CellAlignment = 4
            .Col = 2: .Text = RGrid("WARNA"): .CellAlignment = 4
            .Col = 3: .Text = RGrid("TAHUN"): .CellAlignment = 4
            .Col = 4: .Text = RGrid("RANGKA"): .CellAlignment = 4
            .Col = 5: .Text = RGrid("MESIN"): .CellAlignment = 4
            .Col = 6: .Text = RGrid("TGL_INPUT"): .CellAlignment = 4
            .Col = 7: .Text = Format(RGrid("H_BELI"), "##,###.00")
         End With
      B = B + 1
      RGrid.MoveNext
   Loop
Else
    grid.Clear
    Call SiapkanGrid
End If
RGrid.Close
Set RGrid = Nothing
Text2 = Format(Text2, ">")
End Sub
