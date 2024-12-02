VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form S004 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TABEL SERVICE & SPAREPART"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   18480
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   18480
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option5 
      Caption         =   "JENIS"
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
      Left            =   5235
      TabIndex        =   8
      Top             =   4388
      Width           =   1170
   End
   Begin VB.OptionButton Option4 
      Caption         =   "MEKANIK"
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
      Left            =   3975
      TabIndex        =   7
      Top             =   4388
      Width           =   1170
   End
   Begin VB.OptionButton Option3 
      Caption         =   "JAM"
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
      Left            =   2715
      TabIndex        =   6
      Top             =   4388
      Width           =   1170
   End
   Begin VB.OptionButton Option2 
      Caption         =   "TGL"
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
      Left            =   1455
      TabIndex        =   5
      Top             =   4388
      Width           =   1170
   End
   Begin VB.OptionButton Option1 
      Caption         =   "NO"
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
      TabIndex        =   4
      Top             =   4388
      Width           =   1170
   End
   Begin VB.TextBox Text12 
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
      Left            =   6510
      TabIndex        =   3
      Text            =   "Text12"
      Top             =   4290
      Width           =   2625
   End
   Begin VB.CommandButton TmbSave 
      Caption         =   "CETAK"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   105
      TabIndex        =   1
      Top             =   5250
      Width           =   2475
   End
   Begin VB.CommandButton Command7 
      Caption         =   "KELUAR"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   15630
      TabIndex        =   0
      Top             =   5250
      Width           =   2475
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   4020
      Left            =   105
      TabIndex        =   2
      Top             =   105
      Width           =   18000
      _ExtentX        =   31750
      _ExtentY        =   7091
      _Version        =   393216
      Rows            =   4
      FixedRows       =   2
      FixedCols       =   0
      BackColorFixed  =   12632064
      BackColorBkg    =   16777152
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
Attribute VB_Name = "S004"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private Sub Command7_Click()
Unload Me
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=DEALER", rdDriverNoPrompt, False, CN)

Call SiapkanGrid
'Call IsiGrid
End Sub

Private Sub SiapkanGrid()
With Grid
    .Cols = 12
    .Rows = 3
    .Row = 0
    .Col = 0: .ColWidth(0) = 100: .Text = "NOMOR": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 1: .ColWidth(1) = 1000: .Text = "TANGGAL": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 2: .ColWidth(2) = 1000: .Text = "JAM": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 3: .ColWidth(3) = 1600: .Text = "MEKANIK": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 4: .ColWidth(4) = 1500: .Text = "BIAYA": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 5: .ColWidth(5) = 1250: .Text = "KENDARAAN": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 6: .ColWidth(6) = 1250: .Text = "KENDARAAN": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 7: .ColWidth(7) = 1250: .Text = "KENDARAAN": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 8: .ColWidth(8) = 1250: .Text = "KENDARAAN": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 9: .ColWidth(9) = 1250: .Text = "PEMILIK": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 10: .ColWidth(10) = 1250: .Text = "PEMILIK": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 11: .ColWidth(11) = 1250: .Text = "PEMILIK": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    
    
    .MergeCol(0) = True
    .MergeCol(1) = True
    .MergeCol(2) = True
    .MergeCol(3) = True
    .MergeCol(4) = True
    .MergeCol(5) = True
    .MergeCol(6) = True
    .MergeCol(7) = True
    .MergeCol(8) = True
    .MergeCol(9) = True
    .MergeCol(10) = True
    .MergeCol(11) = True
    .MergeRow(0) = True
    .MergeRow(1) = True
    
    .Row = 1
    .Col = 0: .ColWidth(0) = 1000: .Text = "NOMOR": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 1: .ColWidth(1) = 1000: .Text = "TANGGAL": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 2: .ColWidth(2) = 1000: .Text = "JAM": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 3: .ColWidth(3) = 1600: .Text = "MEKANIK": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 4: .ColWidth(4) = 1500: .Text = "BIAYA": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 5: .ColWidth(5) = 1250: .Text = "NO POLISI": .CellAlignment = 4: .CellBackColor = &HFFFF00
    .Col = 6: .ColWidth(6) = 1250: .Text = "NO RANGKA": .CellAlignment = 4: .CellBackColor = &HFFFF00
    .Col = 7: .ColWidth(7) = 1250: .Text = "NO MESIN": .CellAlignment = 4: .CellBackColor = &HFFFF00
    .Col = 8: .ColWidth(8) = 2500: .Text = "JENIS": .CellAlignment = 4: .CellBackColor = &HFFFF00
    .Col = 9: .ColWidth(9) = 2500: .Text = "NAMA": .CellAlignment = 4: .CellBackColor = &HFFFF00
    .Col = 10: .ColWidth(10) = 1500: .Text = "TELEPON": .CellAlignment = 4: .CellBackColor = &HFFFF00
    .Col = 11: .ColWidth(11) = 5000: .Text = "ALAMAT": .CellAlignment = 4: .CellBackColor = &HFFFF00
    
End With
End Sub
