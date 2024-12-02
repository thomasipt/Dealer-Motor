VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form JL003A 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DAFTAR PENJUALAN"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13665
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   13665
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
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
      Left            =   11055
      TabIndex        =   0
      Text            =   "Text4"
      Top             =   22
      Width           =   2535
   End
   Begin VB.TextBox Text3 
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
      Left            =   8355
      TabIndex        =   9
      Text            =   "Text3"
      Top             =   8175
      Width           =   2535
   End
   Begin Crystal.CrystalReport crpt 
      Left            =   1920
      Top             =   5115
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
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
      Left            =   11520
      TabIndex        =   3
      Top             =   8130
      Width           =   960
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
      Left            =   1110
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   8175
      Width           =   2535
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
      Left            =   4755
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   8175
      Width           =   2535
   End
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
      TabIndex        =   4
      Top             =   8130
      Width           =   960
   End
   Begin VB.Frame Frame1 
      Caption         =   "KONFIRMASI ....."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9090
      Left            =   -60
      TabIndex        =   12
      Top             =   -240
      Width           =   13785
      Begin VB.OptionButton Option5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "CETAK FAKTUR PENGIRIMAN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   275
         Left            =   5115
         TabIndex        =   21
         Top             =   5295
         Width           =   3435
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "CETAK KWITANSI PENJUALAN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   275
         Left            =   5115
         TabIndex        =   18
         Top             =   4965
         Width           =   3435
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "CETAK FAKTUR (SPF)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   275
         Left            =   5115
         TabIndex        =   17
         Top             =   4635
         Width           =   3435
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "EDIT KENDARAAN TERJUAL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   275
         Left            =   5115
         TabIndex        =   16
         Top             =   3690
         Width           =   3435
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "CETAK KWITANSI"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   275
         Left            =   5115
         TabIndex        =   15
         Top             =   4320
         Width           =   3435
      End
      Begin VB.TextBox Text11 
         Height          =   375
         Left            =   -1095
         TabIndex        =   14
         Text            =   "Text11"
         Top             =   -285
         Width           =   1635
      End
      Begin VB.CommandButton Command3 
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
         Height          =   420
         Left            =   7867
         TabIndex        =   13
         Top             =   5910
         Width           =   960
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0C0C0&
         Height          =   3045
         Left            =   4837
         ScaleHeight     =   2985
         ScaleWidth      =   3930
         TabIndex        =   20
         Top             =   2783
         Width           =   3990
      End
      Begin VB.Frame Frame2 
         Caption         =   "Faktur Sistem"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   3990
         Left            =   4702
         TabIndex        =   19
         Top             =   2460
         Width           =   4260
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   7725
      Left            =   75
      TabIndex        =   5
      Top             =   360
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
   Begin VB.Label Label5 
      Caption         =   "No Fak"
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
      Left            =   10110
      TabIndex        =   11
      Top             =   90
      Width           =   1905
   End
   Begin VB.Label Label4 
      Caption         =   "Nama"
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
      Left            =   7425
      TabIndex        =   10
      Top             =   8243
      Width           =   1905
   End
   Begin VB.Label Label3 
      Caption         =   ">> Double klik untuk melakukan edit kendaraan terjual / cetak faktur ....."
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
      TabIndex        =   8
      Top             =   45
      Width           =   7170
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
      Left            =   75
      TabIndex        =   7
      Top             =   8250
      Width           =   1905
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
      Left            =   3825
      TabIndex        =   6
      Top             =   8250
      Width           =   1905
   End
End
Attribute VB_Name = "JL003A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private RGrid As rdoResultset
Private SGrid As String

Private NoNo

Private Sub Command1_Click()
    crpt.ReportFileName = App.Path + "\ReportD\SALDOSEDIAMTR_LAKU.rpt"
    crpt.SelectionFormula = "{M001.STS_Jual} = '1'"
    crpt.WindowState = crptMaximized
    crpt.WindowMaxButton = True
    crpt.WindowMinButton = True
    crpt.Action = 1
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Unload Me
JL003A.Show
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=DEALER", rdDriverNoPrompt, False, CN)

Call SiapkanGrid
Call IsiGrid

Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""

NoNo = 0

Frame1.Visible = False

End Sub

Private Sub SiapkanGrid()
With Grid
    .Row = 0
    .Cols = 9
    .Col = 0: .ColWidth(0) = 1500: .Text = "NO": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 1: .ColWidth(1) = 2500: .Text = "TYPE": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 2: .ColWidth(2) = 1500: .Text = "WARNA": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 3: .ColWidth(3) = 1500: .Text = "TAHUN": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 4: .ColWidth(4) = 1500: .Text = "NO. RANGKA": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 5: .ColWidth(5) = 1500: .Text = "NO. MESIN": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 6: .ColWidth(6) = 1500: .Text = "TGL_INPUT": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 7: .ColWidth(7) = 2500: .Text = "HARGA": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 8: .ColWidth(8) = 2500: .Text = "NAMA": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
End With
End Sub
    
Private Sub IsiGrid()
SGrid = "Select * From M001 where STS_JUAL = '1' order by NO_FAK Desc"
Set RGrid = RDCO.OpenResultset(SGrid, rdOpenKeyset, rdConcurReadOnly)
If RGrid.RowCount <> 0 Then
   RGrid.MoveFirst
   B = 1
   NoNo = 1
   Do Until RGrid.EOF
      Grid.Rows = B + 1
      Grid.Row = B
         With Grid
            .Col = 0: .Text = RGrid("NO_FAK")
            .Col = 1: .Text = RGrid("TYPE")
            .Col = 2: .Text = RGrid("WARNA"): .CellAlignment = 4
            .Col = 3: .Text = RGrid("TAHUN"): .CellAlignment = 4
            .Col = 4: .Text = RGrid("RANGKA"): .CellAlignment = 4
            .Col = 5: .Text = RGrid("MESIN"): .CellAlignment = 4
            .Col = 6: .Text = RGrid("TGL_INPUT"): .CellAlignment = 4
            .Col = 7: .Text = Format(RGrid("H_BELI"), "##,###.00"): .CellFontBold = True
            .Col = 8: .Text = RGrid("NAMA_PEMBELI"): .CellFontBold = True
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
    NO_RANGKA = Grid.TextMatrix(Grid.Row, 4)
    JL004.Show
    Unload Me
End Select
End Sub

Private Sub grid_dblClick()
NO_RANGKA = ""
NO_RANGKA = Grid.TextMatrix(Grid.Row, 4)
NO_MESIN = ""
NO_MESIN = Grid.TextMatrix(Grid.Row, 5)
NO_FAKTUR = ""
NO_FAKTUR = Grid.TextMatrix(Grid.Row, 0)

Text11 = NO_FAKTUR
Frame1.Visible = True
Frame1.ZOrder

'Unload Me
'JL004A.Show

End Sub

Private Sub Option1_Click()
Unload Me
JL004A.Show 1
End Sub

Private Sub Option2_Click()
Unload Me
JL004KWT.Show 1
End Sub

Private Sub Option3_Click()
Unload Me
JL004SPF.Show 1
End Sub

Private Sub Option4_Click()
Unload Me
JL004KWT_JUAL.Show 1
End Sub

Private Sub Option5_Click()
Unload Me
JL004KWT_KIRIM.Show 1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text1_LostFocus()
If Text1 = "" Then Exit Sub
Dim Brs, IndekNama
Brs = 1
IndekNama = "%" + Text1 + "%"
SGrid = "Select * From M001 where Rangka like '" + IndekNama + "' and STS_JUAL = '1' "
Set RGrid = RDCO.OpenResultset(SGrid, rdOpenKeyset, rdConcurReadOnly)
If RGrid.RowCount <> 0 Then
   RGrid.MoveFirst
   B = 1
   Do Until RGrid.EOF
      Grid.Rows = B + 1
      Grid.Row = B
         With Grid
            .Col = 0: .Text = RGrid("NO_FAK")
            .Col = 1: .Text = RGrid("TYPE")
            .Col = 2: .Text = RGrid("WARNA"): .CellAlignment = 4
            .Col = 3: .Text = RGrid("TAHUN"): .CellAlignment = 4
            .Col = 4: .Text = RGrid("RANGKA"): .CellAlignment = 4
            .Col = 5: .Text = RGrid("MESIN"): .CellAlignment = 4
            .Col = 6: .Text = RGrid("TGL_INPUT"): .CellAlignment = 4
            .Col = 7: .Text = Format(RGrid("H_BELI"), "##,###.00"): .CellFontBold = True
            .Col = 8: .Text = RGrid("NAMA_PEMBELI"): .CellFontBold = True
         End With
      B = B + 1
      RGrid.MoveNext
   Loop
Else
    Grid.Clear
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
SGrid = "Select * From M001 where Mesin like '" + IndekNama + "' and STS_JUAL = '1' "
Set RGrid = RDCO.OpenResultset(SGrid, rdOpenKeyset, rdConcurReadOnly)
If RGrid.RowCount <> 0 Then
   RGrid.MoveFirst
   B = 1
   Do Until RGrid.EOF
      Grid.Rows = B + 1
      Grid.Row = B
         With Grid
            .Col = 0: .Text = RGrid("NO_FAK")
            .Col = 1: .Text = RGrid("TYPE")
            .Col = 2: .Text = RGrid("WARNA"): .CellAlignment = 4
            .Col = 3: .Text = RGrid("TAHUN"): .CellAlignment = 4
            .Col = 4: .Text = RGrid("RANGKA"): .CellAlignment = 4
            .Col = 5: .Text = RGrid("MESIN"): .CellAlignment = 4
            .Col = 6: .Text = RGrid("TGL_INPUT"): .CellAlignment = 4
            .Col = 7: .Text = Format(RGrid("H_BELI"), "##,###.00"): .CellFontBold = True
            .Col = 8: .Text = RGrid("NAMA_PEMBELI"): .CellFontBold = True
         End With
      B = B + 1
      RGrid.MoveNext
   Loop
Else
    Grid.Clear
    Call SiapkanGrid
End If
RGrid.Close
Set RGrid = Nothing
Text2 = Format(Text2, ">")
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text3_LostFocus()
If Text3 = "" Then Exit Sub
Dim Brs, IndekNama
Brs = 1
IndekNama = "%" + Text3 + "%"
SGrid = "Select * From M001 where NAMA_PEMBELI like '" + IndekNama + "' and STS_JUAL = '1' "
Set RGrid = RDCO.OpenResultset(SGrid, rdOpenKeyset, rdConcurReadOnly)
If RGrid.RowCount <> 0 Then
   RGrid.MoveFirst
   B = 1
   Do Until RGrid.EOF
      Grid.Rows = B + 1
      Grid.Row = B
         With Grid
            .Col = 0: .Text = RGrid("NO_FAK")
            .Col = 1: .Text = RGrid("TYPE")
            .Col = 2: .Text = RGrid("WARNA"): .CellAlignment = 4
            .Col = 3: .Text = RGrid("TAHUN"): .CellAlignment = 4
            .Col = 4: .Text = RGrid("RANGKA"): .CellAlignment = 4
            .Col = 5: .Text = RGrid("MESIN"): .CellAlignment = 4
            .Col = 6: .Text = RGrid("TGL_INPUT"): .CellAlignment = 4
            .Col = 7: .Text = Format(RGrid("H_BELI"), "##,###.00"): .CellFontBold = True
            .Col = 8: .Text = RGrid("NAMA_PEMBELI"): .CellFontBold = True
         End With
      B = B + 1
      RGrid.MoveNext
   Loop
Else
    Grid.Clear
    Call SiapkanGrid
End If
RGrid.Close
Set RGrid = Nothing
Text3 = Format(Text3, ">")
End Sub

Private Sub Text4_Change()
Grid.Clear
Grid.Refresh
Call SiapkanGrid

SGrid = "Select * From M001 where NO_FAK like '%" + Trim(Text4) + "%' and STS_JUAL = '1' order by NO_FAK Desc "
Set RGrid = RDCO.OpenResultset(SGrid, rdOpenKeyset, rdConcurReadOnly)
If RGrid.RowCount <> 0 Then
   RGrid.MoveFirst
   B = 1
   NoNo = 1
   Do Until RGrid.EOF
      Grid.Rows = B + 1
      Grid.Row = B
         With Grid
            .Col = 0: .Text = RGrid("NO_FAK")
            .Col = 1: .Text = RGrid("TYPE")
            .Col = 2: .Text = RGrid("WARNA"): .CellAlignment = 4
            .Col = 3: .Text = RGrid("TAHUN"): .CellAlignment = 4
            .Col = 4: .Text = RGrid("RANGKA"): .CellAlignment = 4
            .Col = 5: .Text = RGrid("MESIN"): .CellAlignment = 4
            .Col = 6: .Text = RGrid("TGL_INPUT"): .CellAlignment = 4
            .Col = 7: .Text = Format(RGrid("H_BELI"), "##,###.00"): .CellFontBold = True
            .Col = 8: .Text = RGrid("NAMA_PEMBELI"): .CellFontBold = True
         End With
      B = B + 1
      NoNo = NoNo + 1
      RGrid.MoveNext
   Loop
End If
RGrid.Close
Set RGrid = Nothing
End Sub

Private Sub Text1_GotFocus()
ClearTextBoxes Me
End Sub

Private Sub Text2_GotFocus()
ClearTextBoxes Me
End Sub

Private Sub Text3_GotFocus()
ClearTextBoxes Me
End Sub

Private Sub Text4_GotFocus()
ClearTextBoxes Me
End Sub
