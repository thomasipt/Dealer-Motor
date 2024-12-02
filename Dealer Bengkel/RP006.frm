VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form RP006 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CUSTOMER SUPPORT"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13260
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   13260
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "CETAK"
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
      Left            =   83
      TabIndex        =   0
      Top             =   5130
      Width           =   1110
   End
   Begin VB.TextBox Text1 
      Height          =   420
      Left            =   2970
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   8055
      Width           =   1545
   End
   Begin VB.CommandButton cMDeXIT 
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
      Left            =   12068
      TabIndex        =   1
      Top             =   5130
      Width           =   1110
   End
   Begin Crystal.CrystalReport CRPT 
      Left            =   900
      Top             =   9360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   4875
      Left            =   83
      TabIndex        =   2
      Top             =   90
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   8599
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
End
Attribute VB_Name = "RP006"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private RGrid As rdoResultset
Private SGrid As String

Private Sub cMDeXIT_Click()
Unload Me
End Sub

Private Sub Command1_Click()
crpt.ReportFileName = App.Path + "\ReportD\RP006.rpt"
crpt.WindowState = crptMaximized
crpt.WindowMaxButton = False
crpt.WindowMinButton = False
crpt.Action = 1
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=DEALER", rdDriverNoPrompt, False, CN)
Text1 = Month(Tanggal)

Call SiapkanGrid
Call IsiGrid
End Sub

Private Sub SiapkanGrid()
With Grid
    .Row = 0
    .Cols = 6
    .Col = 0: .ColWidth(0) = 3500: .Text = "NAMA": .CellAlignment = 4: .CellFontBold = True
    .Col = 1: .ColWidth(1) = 3500: .Text = "ALAMAT": .CellAlignment = 4: .CellFontBold = True
    .Col = 2: .ColWidth(2) = 1250: .Text = "KOTA": .CellAlignment = 4: .CellFontBold = True
    .Col = 3: .ColWidth(3) = 1500: .Text = "IDENTITAS": .CellAlignment = 4: .CellFontBold = True
    .Col = 4: .ColWidth(4) = 1500: .Text = "TELP / HP": .CellAlignment = 4: .CellFontBold = True
    .Col = 5: .ColWidth(5) = 1250: .Text = "TGL LAHIR": .CellAlignment = 4: .CellFontBold = True
End With

End Sub

Private Sub IsiGrid()
SGrid = "SELECT * from RP006 where BULAN = '" + Trim(Text1) + "' order by NAMA_PEMBELI Asc"
Set RGrid = RDCO.OpenResultset(SGrid, rdOpenKeyset, rdConcurReadOnly)
If RGrid.RowCount <> 0 Then
   RGrid.MoveFirst
   B = 1
   Do Until RGrid.EOF
      Grid.Rows = B + 1
      Grid.Row = B
         With Grid
              .Col = 0: .Text = RGrid("NAMA_PEMBELI")
              .Col = 1: .Text = RGrid("ALAMAT_1")
              .Col = 2: .Text = RGrid("ALAMAT_2")
              .Col = 3: .Text = RGrid("NO_KTP"): .CellAlignment = 4
              .Col = 4: .Text = RGrid("TELP"): .CellAlignment = 4
              .Col = 5: .Text = RGrid("TGL_LAHIR"): .CellAlignment = 4
         End With
      B = B + 1
      NoNo = NoNo + 1
      RGrid.MoveNext
   Loop
End If
RGrid.Close
Set RGrid = Nothing
End Sub
