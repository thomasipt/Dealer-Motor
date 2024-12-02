VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form SH001 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HISTORY PELAYANAN"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   12930
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
      Caption         =   "NOPOL"
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
      Left            =   6750
      TabIndex        =   2
      Top             =   75
      Width           =   1500
   End
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
      Left            =   105
      TabIndex        =   8
      Top             =   7470
      Width           =   1500
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
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
      Left            =   98
      TabIndex        =   0
      Text            =   "1"
      Top             =   75
      Width           =   4560
   End
   Begin VB.CommandButton Command2 
      Caption         =   "NAMA"
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
      Left            =   5220
      TabIndex        =   1
      Top             =   75
      Width           =   1500
   End
   Begin VB.CommandButton Command3 
      Caption         =   "RANGKA"
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
      Left            =   8280
      TabIndex        =   3
      Top             =   75
      Width           =   1500
   End
   Begin VB.CommandButton Command4 
      Caption         =   "MESIN"
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
      Left            =   9810
      TabIndex        =   4
      Top             =   75
      Width           =   1500
   End
   Begin VB.CommandButton Command5 
      Caption         =   "MEKANIK"
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
      Left            =   11340
      TabIndex        =   5
      Top             =   75
      Width           =   1500
   End
   Begin VB.CommandButton Command6 
      Caption         =   "SEMUA"
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
      Left            =   11333
      TabIndex        =   6
      Top             =   7470
      Width           =   1500
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   6810
      Left            =   105
      TabIndex        =   7
      ToolTipText     =   "Klik untuk edit"
      Top             =   585
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   12012
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
      AllowUserResizing=   3
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
   Begin Crystal.CrystalReport CRPT 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3765
      TabIndex        =   9
      Top             =   7515
      Width           =   5415
   End
End
Attribute VB_Name = "SH001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private RGol, RCari, RKode, RDel, RDelBar, RSim, RSave, RSaveP, RDist As rdoResultset
Private SDelBar, SDist, SGol, SCari, Metode, SKode, SDel, SSim, SSave, SSaveP As String
Private Brs, MetodLaba, Ganti, TOKET, WARNA, SPART, Montok

Private Sub Command1_Click()
    crpt.ReportFileName = App.Path + "\ReportD\NOTAALL.rpt"
    crpt.SelectionFormula = "{S003.STS_HIS}= '1'"
    crpt.WindowState = crptMaximized
    crpt.WindowMaxButton = False
    crpt.WindowMinButton = False
    crpt.Action = 1
    crpt.Reset
End Sub

Private Sub Command2_Click()
Dim IndekNama

Call Kosong

IndekNama = "%" + Text1 + "%"
Montok = "Select * from S003 where NAMA like '" + IndekNama + "' order by NO_TRANS Asc"
Call IsiGrid2
End Sub

Private Sub Command3_Click()
Dim IndekNama

Call Kosong

IndekNama = "%" + Text1 + "%"
Montok = "Select * from S003 where NO_RANGKA like '" + IndekNama + "' order by NO_TRANS Asc"
Call IsiGrid2
End Sub

Private Sub Command4_Click()
Dim IndekNama

Call Kosong

IndekNama = "%" + Text1 + "%"
Montok = "Select * from S003 where NO_MESIN like '" + IndekNama + "' order by NO_TRANS Asc"
Call IsiGrid2
End Sub

Private Sub Command5_Click()
Dim IndekNama

Call Kosong

IndekNama = "%" + Text1 + "%"
Montok = "Select * from S003 where N_MEKANIK like '" + IndekNama + "' order by NO_TRANS Asc"
Call IsiGrid2
End Sub

Private Sub Command6_Click()
Unload Me
SH001.Show 1
End Sub

Private Sub Command7_Click()
Dim IndekNama

Call Kosong

IndekNama = "%" + Text1 + "%"
Montok = "Select * from S003 where NO_POL like '" + IndekNama + "' order by NO_TRANS Asc"
Call IsiGrid2
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=DEALER", rdDriverNoPrompt, False, CN)
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""

Call Siap
Call IsiGrid

Montok = ""
Label1.Visible = False
Command1.Visible = False

Call Kosong

End Sub

Private Sub Kosong()
SCari = "Select * from S003"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
RCari.MoveFirst
Do Until RCari.EOF
    RCari.EDIT
        RCari("STS_HIS") = 0
    RCari.Update
    RCari.MoveNext
Loop
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub Siap()
With grid
     .Cols = 9
     .Row = 0
     .Col = 0: .ColWidth(0) = 1000: .Text = "NO TRANS": .CellAlignment = 4
     .Col = 1: .ColWidth(1) = 1000: .Text = "TGL": .CellAlignment = 4
     .Col = 2: .ColWidth(2) = 3500: .Text = "NAMA": .CellAlignment = 4
     .Col = 3: .ColWidth(3) = 5000: .Text = "ALAMAT": .CellAlignment = 4
     .Col = 4: .ColWidth(4) = 1000: .Text = "NOPOL": .CellAlignment = 4
     .Col = 5: .ColWidth(5) = 1000: .Text = "RANGKA": .CellAlignment = 4
     .Col = 6: .ColWidth(6) = 1000: .Text = "MESIN": .CellAlignment = 4
     .Col = 7: .ColWidth(7) = 3000: .Text = "TIPE": .CellAlignment = 4
     .Col = 8: .ColWidth(8) = 1000: .Text = "MEKANIK": .CellAlignment = 4
End With
End Sub

Private Sub IsiGrid()
SCari = "Select top 100 no_trans, tgl_trans, nama, alamat, no_pol, no_rangka, no_mesin, tipe, n_mekanik from S003 order by NO_TRANS Desc"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset, rdConcurReadOnly)
If RCari.RowCount <> 0 Then
RCari.MoveFirst
Brs = 1
Do Until RCari.EOF
       With grid
        .Rows = Brs + 1
        .Row = Brs
        .Col = 0: .Text = RCari("no_trans"): .CellAlignment = 4
        .Col = 1: .Text = RCari("tgl_trans"): .CellAlignment = 4
        .Col = 2: .Text = RCari("nama")
        .Col = 3: .Text = RCari("alamat")
        .Col = 4: .Text = RCari("no_pol"): .CellAlignment = 4
        .Col = 5: .Text = RCari("no_rangka"): .CellAlignment = 4
        .Col = 6: .Text = RCari("no_mesin"): .CellAlignment = 4
        .Col = 7: .Text = RCari("tipe")
        .Col = 8: .Text = RCari("n_mekanik"): .CellAlignment = 4
      End With
      RCari.MoveNext
      Brs = Brs + 1
Loop
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub IsiGrid2()
grid.Clear
grid.Refresh

Call Siap

SCari = Montok
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
RCari.MoveFirst
Brs = 1
Do Until RCari.EOF
       With grid
        .Rows = Brs + 1
        .Row = Brs
        .Col = 0: .Text = RCari("no_trans"): .CellAlignment = 4
        .Col = 1: .Text = RCari("tgl_trans"): .CellAlignment = 4
        .Col = 2: .Text = RCari("nama")
        .Col = 3: .Text = RCari("alamat")
        .Col = 4: .Text = RCari("no_pol"): .CellAlignment = 4
        .Col = 5: .Text = RCari("no_rangka"): .CellAlignment = 4
        .Col = 6: .Text = RCari("no_mesin"): .CellAlignment = 4
        .Col = 7: .Text = RCari("tipe")
        .Col = 8: .Text = RCari("n_mekanik"): .CellAlignment = 4
      End With
      
      RCari.EDIT
        RCari("STS_HIS") = 1
      RCari.Update
      
      RCari.MoveNext
      Brs = Brs + 1
      
Loop
End If
RCari.Close
Set RCari = Nothing

Label1.Visible = True
Label1 = Brs - 1

Command1.Visible = True

End Sub
