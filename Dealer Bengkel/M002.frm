VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form M002 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INFO KENDARAAN"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   12645
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   750
      Left            =   0
      TabIndex        =   1
      Top             =   2625
      Width           =   12825
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Height          =   360
         Left            =   105
         TabIndex        =   9
         Text            =   "Text2"
         Top             =   210
         Width           =   1500
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
         Height          =   360
         Left            =   11235
         TabIndex        =   8
         Top             =   210
         Width           =   960
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Height          =   360
         Left            =   7620
         TabIndex        =   7
         Text            =   "Text9"
         Top             =   210
         Width           =   1500
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Height          =   360
         Left            =   9120
         TabIndex        =   6
         Text            =   "Text7"
         Top             =   210
         Width           =   1500
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Height          =   360
         Left            =   6120
         TabIndex        =   5
         Text            =   "Text6"
         Top             =   210
         Width           =   1500
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Height          =   360
         Left            =   4620
         TabIndex        =   4
         Text            =   "Text5"
         Top             =   210
         Width           =   1500
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Height          =   360
         Left            =   3105
         TabIndex        =   3
         Text            =   "Text4"
         Top             =   210
         Width           =   1500
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Height          =   360
         Left            =   1605
         TabIndex        =   2
         Text            =   "Text3"
         Top             =   210
         Width           =   1500
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   2445
      Left            =   112
      TabIndex        =   0
      Top             =   105
      Width           =   12420
      _ExtentX        =   21908
      _ExtentY        =   4313
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   65280
      ForeColorFixed  =   0
      BackColorBkg    =   16777152
      GridColor       =   8421504
      Appearance      =   0
   End
End
Attribute VB_Name = "M002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private RGrid, RCari As rdoResultset
Private SGrid, SCari As String

Private Sub Command1_Click()
SCari = "Select * From M001 where NO_URUT = '" + Trim(Text1) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset, rdConcurRowVer)
RCari.EDIT
    RCari("TAHUN") = 2008
RCari.Update
RCari.Close
Set RCari = Nothing

Unload Me
M002.Show

End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=DEALER", rdDriverNoPrompt, False, CN)

Call SiapkanGrid
Call IsiGrid

Frame1.Visible = False

End Sub

Private Sub SiapkanGrid()
With grid
    .Row = 0
    .Cols = 9
    .Col = 0: .ColWidth(0) = 0: .Text = "NO": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 1: .ColWidth(1) = 1500: .Text = "TYPE": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 2: .ColWidth(2) = 1500: .Text = "WARNA": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 3: .ColWidth(3) = 1500: .Text = "RANGKA": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 4: .ColWidth(4) = 1500: .Text = "MESIN": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 5: .ColWidth(5) = 1500: .Text = "TAHUN": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 6: .ColWidth(6) = 1500: .Text = "DO/PSMUP": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 7: .ColWidth(7) = 1500: .Text = "SJ": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 8: .ColWidth(8) = 1500: .Text = "HARGA": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
End With
End Sub

Private Sub IsiGrid()
SGrid = "Select * From M001 order by NO Asc"
Set RGrid = RDCO.OpenResultset(SGrid, rdOpenKeyset, rdConcurReadOnly)
If RGrid.RowCount <> 0 Then
   RGrid.MoveFirst
   B = 1
   Do Until RGrid.EOF
      grid.Rows = B + 1
      grid.Row = B
         With grid
              .Col = 0: .Text = RGrid("NO_URUT"): .CellAlignment = 4
              .Col = 1: .Text = RGrid("TYPE"): .CellAlignment = 4
              .Col = 2: .Text = RGrid("WARNA"): .CellAlignment = 4
              .Col = 3: .Text = RGrid("RANGKA"): .CellAlignment = 4
              .Col = 4: .Text = RGrid("MESIN"): .CellAlignment = 4
              .Col = 5: .Text = RGrid("TAHUN"): .CellAlignment = 4
              .Col = 6: .Text = RGrid("DO"): .CellAlignment = 4
              .Col = 7: .Text = RGrid("SJ"): .CellAlignment = 4
              .Col = 8: .Text = Format(RGrid("H_BELI"), "##,###.00")
         End With
      B = B + 1
      RGrid.MoveNext
   Loop
End If
RGrid.Close
Set RGrid = Nothing
End Sub

Private Sub Grid_dblClick()
Frame1.Visible = True
NoUrut = ""
NoUrut = grid.TextMatrix(grid.Row, 0)
Text1 = NoUrut
Text2 = grid.TextMatrix(grid.Row, 1)
Text3 = grid.TextMatrix(grid.Row, 2)
Text4 = grid.TextMatrix(grid.Row, 3)
Text5 = grid.TextMatrix(grid.Row, 4)
Text6 = grid.TextMatrix(grid.Row, 5)
Text9 = grid.TextMatrix(grid.Row, 6)
Text7 = grid.TextMatrix(grid.Row, 7)
Text2.SetFocus
End Sub
