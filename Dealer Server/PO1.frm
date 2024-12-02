VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form PO1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TABEL FAKTUR PEMBELIAN"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   10080
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport crpt 
      Left            =   2415
      Top             =   1260
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      Height          =   2700
      Left            =   105
      TabIndex        =   1
      Top             =   75
      Width           =   9915
      Begin VB.CommandButton Command3 
         Caption         =   "CANCEL"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   3195
         TabIndex        =   4
         Top             =   1717
         Width           =   3690
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ENTRI PO"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   308
         TabIndex        =   3
         Top             =   457
         Width           =   3690
      End
      Begin VB.CommandButton Command2 
         Caption         =   "CETAK PO"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   6083
         TabIndex        =   2
         Top             =   472
         Width           =   3690
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   2595
      Left            =   83
      TabIndex        =   0
      Top             =   105
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   4577
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   65280
      BackColorBkg    =   16777152
      GridColor       =   0
      TextStyle       =   3
      TextStyleFixed  =   3
   End
End
Attribute VB_Name = "PO1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private RSLUser, RCari, RSave, RSave2, REdit, RHapus As rdoResultset
Private SQLUser, SCari, SSave, SSave2, SHapus, SEdit As String

Private Sub Command3_Click()
Frame1.Visible = False
grid.ZOrder
FAKTUR = ""
TANGGAL_FAKTUR = ""
JML_UNIT = ""
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=DEALER", rdDriverNoPrompt, False, CN)
Frame1.Visible = False
grid.ZOrder
Call SiapkanGrid
Call IsiGrid

If grid.TextMatrix(1, 0) = "" Then
    MsgBox "TIDAK ADA TRANSAKSI PEMBELIAN, LAKUKAN ENTRI FAKTUR DAHULU", vbSystemModal, "KONFIRMASI"
    grid.Visible = False
    Exit Sub
End If

End Sub

Private Sub SiapkanGrid()
With grid
    .Cols = 6
    .Row = 0
    .Col = 0: .ColWidth(0) = 1000: .Text = "NO FAK": .CellAlignment = 4
    .Col = 1: .ColWidth(1) = 1500: .Text = "TGL BELI": .CellAlignment = 4
    .Col = 2: .ColWidth(2) = 1000: .Text = "UNIT": .CellAlignment = 4
    .Col = 3: .ColWidth(3) = 2000: .Text = "KAS": .CellAlignment = 4
    .Col = 4: .ColWidth(4) = 2000: .Text = "NON TUNAI": .CellAlignment = 4
    .Col = 5: .ColWidth(5) = 2000: .Text = "JML TOTAL": .CellAlignment = 4
End With
End Sub

Private Sub IsiGrid()
sqlcs3 = "Select Top 10 NO_Fak, Tgl_Beli, Jumlah, Kas, Non_Kas, H_Jumlah From F001 where STS_FAK = 0 order by No_System Desc"
Set rscs3 = RDCO.OpenResultset(sqlcs3, rdOpenKeyset, rdConcurReadOnly)
If rscs3.RowCount <> 0 Then
   Call SiapkanGrid
   rscs3.MoveFirst
   B = 1
   Do Until rscs3.EOF
      grid.Rows = B + 1
      grid.Row = B
         With grid
              .Col = 0: .Text = rscs3("NO_FAK"): .CellAlignment = 4
              .Col = 1: .Text = rscs3("TGL_BELI"): .CellAlignment = 4
              .Col = 2: .Text = rscs3("JUMLAH"): .CellAlignment = 4
              .Col = 3: .Text = Format(rscs3("KAS"), "##,###.00")
              .Col = 4: .Text = Format(rscs3("NON_KAS"), "##,###.00")
              .Col = 5: .Text = Format(rscs3("H_JUMLAH"), "##,###.00")
         End With
      B = B + 1
      rscs3.MoveNext
   Loop
End If
rscs3.Close
Set rscs3 = Nothing
End Sub

Private Sub Grid_DblClick()
FAKTUR = ""
TANGGAL_FAKTUR = ""
JML_UNIT = ""
FAKTUR = grid.TextMatrix(grid.Row, 0)
TANGGAL_FAKTUR = grid.TextMatrix(grid.Row, 1)
JML_UNIT = grid.TextMatrix(grid.Row, 2)
Frame1.Visible = True
Frame1.ZOrder
End Sub

Private Sub Command2_Click()
CRPT.ReportFileName = "c:\windows\ReportD\CetakPO.rpt"
CRPT.SelectionFormula = "{PO2.NO_FAK}= '" + Trim(FAKTUR) + "'"
CRPT.WindowState = crptMaximized
CRPT.WindowMaxButton = False
CRPT.WindowMinButton = False
CRPT.Action = 1
Frame1.Visible = False
grid.ZOrder
End Sub

Private Sub Command1_Click()
Unload Me
PO2.Show
End Sub
