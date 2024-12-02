VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form B003A 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PENCARIAN SPAREPART"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   12810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   12810
   StartUpPosition =   2  'CenterScreen
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
      Left            =   11280
      TabIndex        =   8
      Top             =   7245
      Width           =   1500
   End
   Begin VB.CommandButton Command5 
      Caption         =   "STOCK"
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
      Left            =   8595
      TabIndex        =   4
      Top             =   7245
      Width           =   1500
   End
   Begin VB.CommandButton Command4 
      Caption         =   "JUAL"
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
      Left            =   7095
      TabIndex        =   3
      Top             =   7245
      Width           =   1500
   End
   Begin VB.CommandButton Command3 
      Caption         =   "BELI"
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
      Left            =   5610
      TabIndex        =   2
      Top             =   7245
      Width           =   1500
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
      Left            =   2115
      TabIndex        =   1
      Top             =   7245
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "KODE"
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
      Left            =   90
      TabIndex        =   0
      Top             =   7245
      Width           =   2040
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2070
      TabIndex        =   6
      Text            =   "2"
      Top             =   60
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   90
      TabIndex        =   5
      Text            =   "1"
      Top             =   60
      Width           =   1995
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   6810
      Left            =   45
      TabIndex        =   7
      ToolTipText     =   "Klik untuk edit"
      Top             =   360
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
Attribute VB_Name = "B003A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private RGol, RCari, RKode, RDel, RDelBar, RSim, RSave, RSaveP, RDist As rdoResultset
Private SDelBar, SDist, SGol, SCari, Metode, SKode, SDel, SSim, SSave, SSaveP As String
Private Brs, MetodLaba, Ganti, TOKET, WARNA, SPART, MONTOK

Private Sub Command1_Click()
MONTOK = "Select * from B003A_CARI order by KODE_JNS"
Call IsiGrid2
End Sub

Private Sub Command2_Click()
MONTOK = "Select * from B003A_CARI order by NAMA_JNS"
Call IsiGrid2
End Sub

Private Sub Command3_Click()
MONTOK = "Select * from B003A_CARI order by HBELI"
Call IsiGrid2
End Sub

Private Sub Command4_Click()
MONTOK = "Select * from B003A_CARI order by HJUAL"
Call IsiGrid2
End Sub

Private Sub Command5_Click()
MONTOK = "Select * from B003A_CARI order by JML_AKHIR"
Call IsiGrid2
End Sub

Private Sub Command6_Click()
Unload Me
B003A.Show
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
MONTOK = ""

End Sub

Private Sub IsiGrid()
SCari = "Select * from B003A_CARI order by KODE_JNS"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset, rdConcurReadOnly)
If RCari.RowCount <> 0 Then
RCari.MoveFirst
Brs = 1
Do Until RCari.EOF
       With grid
        .Rows = Brs + 1
        .Row = Brs
        .Col = 0: .Text = RCari("Kode_Jns"): .CellAlignment = 2
        .Col = 1: .Text = RCari("Nama_Jns"): .CellAlignment = 2
        .Col = 2: .Text = Format(RCari("HBeli"), "##,###.00"): .CellAlignment = 4
        .Col = 3: .Text = Format(RCari("HJual"), "##,###.00"): .CellAlignment = 4
        .Col = 4: .Text = Format(RCari("JML_AKHIR"), "##,###"): .CellAlignment = 4
        .Col = 5: .Text = RCari("SATUAN"): .CellAlignment = 4
        .Col = 6: .Text = RCari("STYLE"): .CellAlignment = 4
        .Col = 7: .Text = RCari("WARNA"): .CellAlignment = 4
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
Call Siap

SCari = MONTOK
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset, rdConcurReadOnly)
If RCari.RowCount <> 0 Then
RCari.MoveFirst
Brs = 1
Do Until RCari.EOF
       With grid
        .Rows = Brs + 1
        .Row = Brs
        .Col = 0: .Text = RCari("Kode_Jns"): .CellAlignment = 2
        .Col = 1: .Text = RCari("Nama_Jns"): .CellAlignment = 2
        .Col = 2: .Text = Format(RCari("HBeli"), "##,###.00"): .CellAlignment = 4
        .Col = 3: .Text = Format(RCari("HJual"), "##,###.00"): .CellAlignment = 4
        .Col = 4: .Text = RCari("JML_AKHIR"): .CellAlignment = 4
        .Col = 5: .Text = RCari("SATUAN"): .CellAlignment = 4
        .Col = 6: .Text = RCari("STYLE"): .CellAlignment = 4
        .Col = 7: .Text = RCari("WARNA"): .CellAlignment = 4
      End With
      RCari.MoveNext
      Brs = Brs + 1
Loop
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub Siap()
With grid
     .Cols = 8
     .Row = 0
     .Col = 0: .ColWidth(0) = 2000: .Text = "KODE": .CellAlignment = 4
     .Col = 1: .ColWidth(1) = 3500: .Text = "NAMA": .CellAlignment = 4
     .Col = 2: .ColWidth(2) = 1500: .Text = "BELI": .CellAlignment = 4
     .Col = 3: .ColWidth(3) = 1500: .Text = "JUAL": .CellAlignment = 4
     .Col = 4: .ColWidth(4) = 1500: .Text = "STOCK": .CellAlignment = 4
     .Col = 5: .ColWidth(5) = 750: .Text = "SATUAN": .CellAlignment = 4
     .Col = 6: .ColWidth(6) = 750: .Text = "STYLE": .CellAlignment = 4
     .Col = 7: .ColWidth(7) = 750: .Text = "WARNA": .CellAlignment = 4
End With
End Sub

Private Sub text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text1_LostFocus()
If Text1 = "" Then
    Exit Sub
Else
    Text1 = Format(Text1, ">")
    MONTOK = "Select * from B003A_CARI where KODE_JNS like '%" + Trim(Text1) + "%'"
    Call IsiGrid2
    Text1 = ""
    Command1.SetFocus
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text2_LostFocus()
If Text2 = "" Then
    Exit Sub
Else
    Text2 = Format(Text2, ">")
    MONTOK = "Select * from B003A_CARI where NAMA_JNS like '%" + Trim(Text2) + "%'"
    Call IsiGrid2
    Text2 = ""
    Command1.SetFocus
End If
End Sub

