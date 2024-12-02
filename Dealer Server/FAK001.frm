VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form FAK001 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FAKTUR & NOTA"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14385
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   14385
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport crpt 
      Left            =   1890
      Top             =   6405
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      Caption         =   "CETAK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2850
      Left            =   3382
      TabIndex        =   2
      Top             =   1365
      Width           =   7620
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5190
         TabIndex        =   14
         Text            =   "Text5"
         Top             =   945
         Width           =   2325
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3885
         TabIndex        =   13
         Text            =   "Text4"
         Top             =   945
         Width           =   1485
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2415
         TabIndex        =   12
         Text            =   "Text3"
         Top             =   945
         Width           =   1485
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   105
         TabIndex        =   11
         Text            =   "Text2"
         Top             =   945
         Width           =   2325
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1515
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   390
         Width           =   6000
      End
      Begin VB.CommandButton Command2 
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
         Height          =   540
         Left            =   105
         TabIndex        =   3
         Top             =   2205
         Width           =   7410
      End
      Begin VB.OptionButton Option4 
         Caption         =   "KWITANSI 2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4515
         TabIndex        =   7
         Top             =   1785
         Width           =   2160
      End
      Begin VB.OptionButton Option3 
         Caption         =   "KWITANSI 1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4515
         TabIndex        =   6
         Top             =   1470
         Width           =   2160
      End
      Begin VB.OptionButton Option2 
         Caption         =   "FAKTUR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1050
         TabIndex        =   5
         Top             =   1785
         Width           =   2160
      End
      Begin VB.OptionButton Option1 
         Caption         =   "DELIVERY ORDER"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1050
         TabIndex        =   4
         Top             =   1470
         Width           =   2160
      End
      Begin VB.Label Label2 
         Caption         =   "TIPE / RANGKA / MESIN / TAHUN"
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
         TabIndex        =   9
         Top             =   735
         Width           =   3060
      End
      Begin VB.Label Label1 
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
         Height          =   225
         Left            =   105
         TabIndex        =   8
         Top             =   420
         Width           =   750
      End
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
      Left            =   6607
      TabIndex        =   0
      Top             =   6405
      Width           =   960
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   6120
      Left            =   30
      TabIndex        =   1
      Top             =   30
      Width           =   14325
      _ExtentX        =   25268
      _ExtentY        =   10795
      _Version        =   393216
      Rows            =   4
      FixedRows       =   2
      FixedCols       =   0
      BackColorFixed  =   14737632
      MergeCells      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "FAK001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Frame1.Visible = False
Grid.Visible = True
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=DEALER", rdDriverNoPrompt, False, CN)

Call SiapkanGrid
Call IsiGrid
Frame1.Visible = False

End Sub

Private Sub SiapkanGrid()
With Grid
    .Cols = 7
    .Rows = 3
    .Row = 0
    .Col = 0: .ColWidth(0) = 1500: .Text = "NO FAK": .CellAlignment = 4: .CellFontBold = True
    .Col = 1: .ColWidth(1) = 2500: .Text = "KENDARAAN": .CellAlignment = 4: .CellFontBold = True
    .Col = 2: .ColWidth(2) = 1500: .Text = "KENDARAAN": .CellAlignment = 4: .CellFontBold = True
    .Col = 3: .ColWidth(3) = 1500: .Text = "KENDARAAN": .CellAlignment = 4: .CellFontBold = True
    .Col = 4: .ColWidth(4) = 750: .Text = "KENDARAAN": .CellAlignment = 4: .CellFontBold = True
    .Col = 5: .ColWidth(5) = 1250: .Text = "PEMBELI": .CellAlignment = 4: .CellFontBold = True
    .Col = 6: .ColWidth(6) = 1250: .Text = "PEMBELI": .CellAlignment = 4: .CellFontBold = True
        
    .MergeCol(0) = True
    .MergeCol(1) = True
    .MergeCol(2) = True
    .MergeCol(3) = True
    .MergeCol(4) = True
    .MergeCol(5) = True
    .MergeCol(6) = True
    .MergeRow(0) = True
    .MergeRow(1) = True
    .Row = 1
    
    .Col = 0: .ColWidth(0) = 1500: .Text = "NO FAK": .CellAlignment = 4: .CellFontBold = True
    .Col = 1: .ColWidth(1) = 2500: .Text = "TIPE": .CellAlignment = 4: .CellBackColor = &HC0C0FF: .CellFontBold = True
    .Col = 2: .ColWidth(2) = 1500: .Text = "RANGKA": .CellAlignment = 4: .CellBackColor = &HC0C0FF: .CellFontBold = True
    .Col = 3: .ColWidth(3) = 1500: .Text = "MESIN": .CellAlignment = 4: .CellBackColor = &HC0C0FF: .CellFontBold = True
    .Col = 4: .ColWidth(4) = 750: .Text = "TAHUN": .CellAlignment = 4: .CellBackColor = &HC0C0FF: .CellFontBold = True
    .Col = 5: .ColWidth(5) = 2500: .Text = "NAMA": .CellAlignment = 4: .CellBackColor = &HC0FFC0: .CellFontBold = True
    .Col = 6: .ColWidth(6) = 3500: .Text = "ALAMAT": .CellAlignment = 4: .CellBackColor = &HC0FFC0: .CellFontBold = True
        
End With
End Sub

Private Sub IsiGrid()
SGrid = "Select * From M001 where STS_JUAL = '1' order by TGL_JUAL ASC"
Set RGrid = RDCO.OpenResultset(SGrid, rdOpenKeyset, rdConcurReadOnly)
If RGrid.RowCount <> 0 Then
   RGrid.MoveFirst
   B = 2
   Do Until RGrid.EOF
      Grid.Rows = B + 1
      Grid.Row = B
         With Grid
              .Col = 0: .Text = RGrid("NO_FAK"): .CellAlignment = 4
              .Col = 1: .Text = RGrid("TYPE")
              .Col = 2: .Text = RGrid("RANGKA"): .CellAlignment = 4
              .Col = 3: .Text = RGrid("MESIN"): .CellAlignment = 4
              .Col = 4: .Text = RGrid("TAHUN"): .CellAlignment = 4
              .Col = 5: .Text = RGrid("NAMA_PEMBELI")
              .Col = 6: .Text = RGrid("ALAMAT_2")
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
Grid.Visible = False
Text1 = Grid.TextMatrix(Grid.Row, 5)
Text2 = Grid.TextMatrix(Grid.Row, 1)
Text3 = Grid.TextMatrix(Grid.Row, 2)
Text4 = Grid.TextMatrix(Grid.Row, 3)
Text5 = Grid.TextMatrix(Grid.Row, 4)
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
    crpt.ReportFileName = "c:\Windows\ReportD\DO.rpt"
    crpt.SelectionFormula = "{M001.RANGKA} = '" + Trim(Text3) + "'"
    crpt.WindowState = crptMaximized
    crpt.WindowMaxButton = True
    crpt.WindowMinButton = True
    crpt.Action = 1
End If
    Option1.Value = False
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
    crpt.ReportFileName = "c:\Windows\ReportD\FAK.rpt"
    crpt.SelectionFormula = "{M001.RANGKA} = '" + Trim(Text3) + "'"
    crpt.WindowState = crptMaximized
    crpt.WindowMaxButton = True
    crpt.WindowMinButton = True
    crpt.Action = 1
End If
    Option2.Value = False
End Sub

Private Sub Option3_Click()
If Option3.Value = True Then
    crpt.ReportFileName = "c:\Windows\ReportD\KW1.rpt"
    crpt.SelectionFormula = "{M001.RANGKA} = '" + Trim(Text3) + "'"
    crpt.WindowState = crptMaximized
    crpt.WindowMaxButton = True
    crpt.WindowMinButton = True
    crpt.Action = 1
End If
    Option3.Value = False
End Sub

Private Sub Option4_Click()
If Option4.Value = True Then
    crpt.ReportFileName = "c:\Windows\ReportD\KW2.rpt"
    crpt.SelectionFormula = "{M001.RANGKA} = '" + Trim(Text3) + "'"
    crpt.WindowState = crptMaximized
    crpt.WindowMaxButton = True
    crpt.WindowMinButton = True
    crpt.Action = 1
End If
    Option4.Value = False
End Sub
