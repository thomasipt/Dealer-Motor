VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form P002 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INFORMASI PIUTANG"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12765
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   12765
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
      Left            =   11730
      TabIndex        =   1
      Top             =   6465
      Width           =   960
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   5820
      Left            =   75
      TabIndex        =   0
      Top             =   495
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   10266
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   65280
      ForeColorFixed  =   0
      BackColorBkg    =   16777152
      GridColor       =   8421504
      FocusRect       =   0
      Appearance      =   0
   End
   Begin VB.Label Label3 
      Caption         =   ">> Double klik untuk melakukan transaksi pelunasan piutang"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   75
      TabIndex        =   2
      Top             =   90
      Width           =   9150
   End
End
Attribute VB_Name = "P002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private RCari As rdoResultset
Private SCari As String

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=DEALER", rdDriverNoPrompt, False, CN)

Call SiapkanGrid
Call IsiGrid

NoNo = 0

End Sub

Private Sub SiapkanGrid()
With grid
    .Cols = 9
    .Row = 0
    .Col = 0: .ColWidth(0) = 500: .Text = "NO": .CellAlignment = 4: .CellFontBold = True
    .Col = 1: .ColWidth(1) = 2000: .Text = "NO PIUTANG": .CellAlignment = 4: .CellFontBold = True
    .Col = 2: .ColWidth(2) = 2500: .Text = "NAMA CUSTOMER": .CellAlignment = 4: .CellFontBold = True
    .Col = 3: .ColWidth(3) = 1600: .Text = "PLAFON": .CellAlignment = 4: .CellFontBold = True
    .Col = 4: .ColWidth(4) = 1600: .Text = "OUT STANDING": .CellAlignment = 4: .CellFontBold = True
    .Col = 5: .ColWidth(5) = 1600: .Text = "INTENSIF": .CellAlignment = 4: .CellFontBold = True
    .Col = 6: .ColWidth(6) = 1500: .Text = "TGL MULAI": .CellAlignment = 4: .CellFontBold = True
    .Col = 7: .ColWidth(7) = 1500: .Text = "JTH TEMPO": .CellAlignment = 4: .CellFontBold = True
    .Col = 8: .ColWidth(8) = 1000: .Text = "SYRT": .CellAlignment = 4: .CellFontBold = True
End With
End Sub

Private Sub IsiGrid()
Dim Brs
Brs = 1
NoNo = 1
SCari = "Select * from P002 where Status = '" + Trim("0") + "' order by Nama_Nas"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    RCari.MoveFirst
    Do While Not RCari.EOF
        With grid
        .Rows = Brs + 1
        .Row = Brs
        .Col = 0: .Text = NoNo: .CellAlignment = 4
        .Col = 1: .Text = RCari("Nomor_Pin")
        
        IPT = RCari("Nomor_Pin")
        
        SCari2 = "Select * from M001 where NO_HUTANG = '" + Trim(IPT) + "'"
        Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenDynamic, rdConcurRowVer)
        If RCari2.RowCount <> 0 Then
            .Col = 2: .Text = RCari2("NAMA_PEMBELI")
        End If
        RCari2.Close
        Set RCari2 = Nothing
        
        .Col = 3: .Text = Format(RCari("Plafon"), "##,###.00")
        .Col = 4: .Text = Format(RCari("Baki_Debet"), "##,###.00")
        .Col = 5: .Text = Format(RCari("Intensif"), "##,###.00")
        .Col = 6: .Text = RCari("Tgl_Mulai"): .CellAlignment = 4
        .Col = 7: .Text = RCari("Tgl_Jatuh"): .CellAlignment = 4
        .Col = 8: .Text = RCari("Syarat_Byr"): .CellAlignment = 4
        End With
    RCari.MoveNext
    Brs = Brs + 1
    NoNo = NoNo + 1
    Loop
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 13
    NoPinjaman = ""
    NoPinjaman = grid.TextMatrix(grid.Row, 1)
    
    If NoPinjaman = "" Then
        MsgBox "DATA KOSONG", vbCritical, "WARNING"
    Else
        P006.Show
    End If
    Unload Me
End Select
End Sub

Private Sub grid_dblClick()
NoPinjaman = ""
NoPinjaman = grid.TextMatrix(grid.Row, 1)

If NoPinjaman = "" Then
    MsgBox "DATA KOSONG", vbCritical, "WARNING"
Else
    P006.Show
End If
Unload Me
End Sub

