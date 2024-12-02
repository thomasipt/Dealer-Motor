VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form RP011A 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "STATEMENT GENERAL LEDGER"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10695
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   10695
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   7200
      Width           =   1335
   End
   Begin VB.CommandButton CmdExit 
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
      Height          =   400
      Left            =   9645
      TabIndex        =   0
      Top             =   4980
      Width           =   990
   End
   Begin VB.OptionButton Option1 
      Caption         =   "STATEMENT SEMUA SUB GL"
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
      Left            =   45
      TabIndex        =   1
      Top             =   5025
      Width           =   3015
   End
   Begin Crystal.CrystalReport CRPT 
      Left            =   675
      Top             =   5640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   4875
      Left            =   52
      TabIndex        =   2
      ToolTipText     =   "Klik untuk cetak"
      Top             =   0
      Width           =   10590
      _ExtentX        =   18680
      _ExtentY        =   8599
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
End
Attribute VB_Name = "RP011A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private RCari As rdoResultset
Private SCari As String

Private D, M, M1, Y, Hari
Private Tahun, TglAng

Private Sub cMDeXIT_Click()
Unload Me
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=DEALER", rdDriverNoPrompt, False, CN)

Call Siap
Call IsiGrid

Call Montok
End Sub

Private Sub Siap()
With grid
     .Cols = 6
     .Row = 0
     .Col = 0: .ColWidth(0) = 1000: .Text = "CODE GL": .CellAlignment = 4
     .Col = 1: .ColWidth(1) = 3000: .Text = "NAMA": .CellAlignment = 4
     .Col = 2: .ColWidth(2) = 1500: .Text = "SALDO AWAL": .CellAlignment = 4
     .Col = 3: .ColWidth(3) = 1500: .Text = "DEBET": .CellAlignment = 4
     .Col = 4: .ColWidth(4) = 1500: .Text = "CREDIT": .CellAlignment = 4
     .Col = 5: .ColWidth(5) = 1500: .Text = "SALDO": .CellAlignment = 4
End With
End Sub

Private Sub IsiGrid()
SCari = "Select * from G003 where MutasiD not like 0 or MutasiC not like 0 order by CodeSL ASC"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset, rdConcurReadOnly)
If RCari.RowCount <> 0 Then
RCari.MoveFirst
Brs = 1
Do Until RCari.EOF
       With grid
        .Rows = Brs + 1
        .Row = Brs
        .Col = 0: .Text = RCari("CodeSL"): .CellAlignment = 4
        .Col = 1: .Text = RCari("NamaSL"):
        .Col = 2: .Text = Format(RCari("SaldoAwal"), "##,###.00")
        .Col = 3: .Text = Format(RCari("MutasiD"), "##,###.00")
        .Col = 4: .Text = Format(RCari("MutasiC"), "##,###.00")
        .Col = 5: .Text = Format(RCari("Saldo"), "##,###.00")
      End With
      RCari.MoveNext
      Brs = Brs + 1
Loop
End If
RCari.Close
Set RCari = Nothing
If Brs > 14 Then
    grid.TopRow = Brs - 14
End If
End Sub

Private Sub grid_dblClick()
Text1 = grid.TextMatrix(grid.Row, 0)
Call QWERTY
End Sub

Private Sub Option1_Click()
Dim A As Integer
A = Month(Tanggal)
If Option1.Value = True Then
    crpt.ReportFileName = App.Path + "\ReportD\HisGl.rpt"
    crpt.SelectionFormula = "{G005.tanggal} in date (" & Y & "," & M & "," & D & ") to date (" & Y & "," & M & "," & D & ")  "
    crpt.WindowState = crptMaximized
    crpt.WindowMaxButton = False
    crpt.WindowMinButton = False
    crpt.Action = 1
End If
Option1.Value = False
    crpt.Reset
End Sub

Private Sub Montok()
D = Day(Tanggal)
M = Month(Tanggal)
M1 = M + 1
Y = Year(Tanggal)
Tahun = Format(DateSerial(Y, M1, D), "DD/MM/YYYY")
TglAng = Format(DateSerial(Y, M, D), "DD/MM/YYYY")
Hari = DateDiff("d", TglAng, Tahun)
End Sub

Private Sub QWERTY()
Dim A As Integer
A = Month(Tanggal)
If Text1 = "" Then Exit Sub
SCari = "Select NamaSl From G005 where CodeSL= '" + Trim(Text1) + "'  order by NoUrut"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    Label3 = RCari("Namasl")
    crpt.ReportFileName = App.Path + "\ReportD\HisGL.rpt"
    crpt.SelectionFormula = "{G005.Codesl} = '" + Trim(Text1) + "' and {G005.tanggal} in date (" & Y & "," & M & "," & D & ") to date (" & Y & "," & M & "," & Hari & ")  "
    crpt.WindowState = crptMaximized
    crpt.WindowMaxButton = False
    crpt.WindowMinButton = False
    crpt.Action = 1
    crpt.Reset
Else
    MsgBox "TIDAK ADA MUTASI", vbCritical, "CODE SUB GL"
    Text1.SetFocus
End If
RCari.Close
Set RCari = Nothing
End Sub
