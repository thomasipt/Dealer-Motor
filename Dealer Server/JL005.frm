VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form JL005 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MUTASI KENDARAAN ANTAR CABANG"
   ClientHeight    =   8745
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14715
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   14715
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
      Left            =   6877
      TabIndex        =   22
      Top             =   8085
      Width           =   960
   End
   Begin VB.Frame Frame3 
      Caption         =   "MUTASI"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7830
      Left            =   97
      TabIndex        =   0
      Top             =   75
      Width           =   14520
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4185
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   1155
         Width           =   3270
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4185
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "Text3"
         Top             =   1995
         Width           =   3270
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4185
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "Text4"
         Top             =   2415
         Width           =   3270
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4185
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Text5"
         Top             =   2835
         Width           =   3270
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4185
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "Text2"
         Top             =   1575
         Width           =   3270
      End
      Begin VB.CommandButton Command1 
         Caption         =   "BATAL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   9700
         TabIndex        =   8
         Top             =   7035
         Width           =   1000
      End
      Begin VB.CommandButton TblSave 
         Caption         =   "SIMPAN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   4015
         TabIndex        =   7
         Top             =   7035
         Width           =   1000
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Left            =   9817
         TabIndex        =   5
         Text            =   "Text8"
         Top             =   405
         Width           =   1140
      End
      Begin VB.Frame Frame1 
         Caption         =   "Tujuan Barang"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   9510
         TabIndex        =   4
         Top             =   2520
         Width           =   3315
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   105
            TabIndex        =   20
            Text            =   "Combo1"
            Top             =   210
            Width           =   3150
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Asal Barang"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   9510
         TabIndex        =   1
         Top             =   1050
         Width           =   3315
         Begin VB.TextBox Text6 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   300
            Left            =   105
            TabIndex        =   19
            Text            =   "Text6"
            Top             =   210
            Width           =   3135
         End
      End
      Begin VB.Label Label5 
         Caption         =   "No. Mesin"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1890
         TabIndex        =   18
         Top             =   2880
         Width           =   2355
      End
      Begin VB.Label Label4 
         Caption         =   "No. Rangka"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1890
         TabIndex        =   17
         Top             =   2460
         Width           =   2355
      End
      Begin VB.Label Label3 
         Caption         =   "Tahun"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1890
         TabIndex        =   16
         Top             =   2040
         Width           =   2355
      End
      Begin VB.Label Label2 
         Caption         =   "Warna"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1890
         TabIndex        =   15
         Top             =   1620
         Width           =   2355
      End
      Begin VB.Label Label1 
         Caption         =   "Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1890
         TabIndex        =   14
         Top             =   1200
         Width           =   2355
      End
      Begin VB.Label Label31 
         Caption         =   "Tanggal Mutasi"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   7747
         TabIndex        =   6
         Top             =   353
         Width           =   2355
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label10"
         Height          =   285
         Left            =   5587
         TabIndex        =   3
         Top             =   420
         Width           =   1140
      End
      Begin VB.Label Label9 
         Caption         =   "Tanggal System"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3757
         TabIndex        =   2
         Top             =   360
         Width           =   2355
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   7695
      Left            =   97
      TabIndex        =   21
      Top             =   143
      Width           =   14520
      _ExtentX        =   25612
      _ExtentY        =   13573
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
End
Attribute VB_Name = "JL005"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private RCari4, RCari3, RCari2, RCari, RGrid As rdoResultset
Private SCari4, SCari3, SCari2, SCari, SGrid As String

Private NoNo

Private Sub Command1_Click()
Call MutasiPasif
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=DEALER", rdDriverNoPrompt, False, CN)

Call SiapkanGrid
Call IsiGrid

Call MutasiPasif

Label10 = Tanggal
Text8 = Tanggal
NoNo = 0

End Sub

Private Sub MutasiPasif()
Frame3.Visible = False
grid.Enabled = True
grid.ZOrder
End Sub

Private Sub MutasiAktif()
Frame3.Visible = True
Frame3.ZOrder
grid.Enabled = False
End Sub

Private Sub SiapkanGrid()
With grid
    .Row = 0
    .Cols = 9
    .Col = 0: .ColWidth(0) = 500: .Text = "NO": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 1: .ColWidth(1) = 2500: .Text = "TYPE": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 2: .ColWidth(2) = 1500: .Text = "WARNA": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 3: .ColWidth(3) = 1500: .Text = "TAHUN": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 4: .ColWidth(4) = 1500: .Text = "NO. RANGKA": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 5: .ColWidth(5) = 1500: .Text = "NO. MESIN": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 6: .ColWidth(6) = 1500: .Text = "TGL_INPUT": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 7: .ColWidth(7) = 1500: .Text = "HARGA": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 8: .ColWidth(8) = 2000: .Text = "MTS_MOTOR": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
End With
End Sub
    
Private Sub IsiGrid()
SGrid = "Select * From M001 where STS_JUAL = '0' order by RANGKA ASC"
Set RGrid = RDCO.OpenResultset(SGrid, rdOpenKeyset, rdConcurReadOnly)
If RGrid.RowCount <> 0 Then
   RGrid.MoveFirst
   B = 1
   NoNo = 1
   Do Until RGrid.EOF
      grid.Rows = B + 1
      grid.Row = B
         With grid
            .Col = 0: .Text = NoNo: .CellAlignment = 4
            .Col = 1: .Text = RGrid("TYPE")
            .Col = 2: .Text = RGrid("WARNA"): .CellAlignment = 4
            .Col = 3: .Text = RGrid("TAHUN"): .CellAlignment = 4
            .Col = 4: .Text = RGrid("RANGKA"): .CellAlignment = 4
            .Col = 5: .Text = RGrid("MESIN"): .CellAlignment = 4
            .Col = 6: .Text = RGrid("TGL_INPUT"): .CellAlignment = 4
            .Col = 7: .Text = Format(RGrid("H_BELI"), "##,###.00")
            .Col = 8: .Text = RGrid("MTS_MOTOR"): .CellAlignment = 4
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
    Call MutasiAktif
    Frame1.Visible = True
    
    Text1 = grid.TextMatrix(grid.Row, 1)
    Text2 = grid.TextMatrix(grid.Row, 2)
    Text3 = grid.TextMatrix(grid.Row, 3)
    Text4 = grid.TextMatrix(grid.Row, 4)
    Text5 = grid.TextMatrix(grid.Row, 5)
    Text6 = grid.TextMatrix(grid.Row, 8)
    
    Call IsiCombo
    
    SCari = "Select * From C012 where NoNas= '" + Trim(Text6) + "'"
    Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
    If RCari.RowCount <> 0 Then
        Text6 = RCari("Nama")
    End If
    
    RCari.Close
    Set RCari = Nothing
End Select
End Sub

Private Sub grid_dblClick()
Call MutasiAktif
Frame1.Visible = True

Text1 = grid.TextMatrix(grid.Row, 1)
Text2 = grid.TextMatrix(grid.Row, 2)
Text3 = grid.TextMatrix(grid.Row, 3)
Text4 = grid.TextMatrix(grid.Row, 4)
Text5 = grid.TextMatrix(grid.Row, 5)
Text6 = grid.TextMatrix(grid.Row, 8)

Call IsiCombo

SCari = "Select * From C012 where NoNas= '" + Trim(Text6) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    Text6 = RCari("Nama")
End If

RCari.Close
Set RCari = Nothing

End Sub

Private Sub IsiCombo()
Dim KodeG
SCari = "Select * From C012 order by NoNas Desc"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    RCari.MoveFirst
    Do While Not RCari.EOF
        Combo1.AddItem RCari("Nama")
    RCari.MoveNext
    Loop
End If

RCari.Close
Set RCari = Nothing
Combo1.ListIndex = 0
End Sub

Private Sub TblSave_Click()
SCari2 = "Select * From C012 where Nama = '" + Trim(Combo1) + "'"
Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenDynamic, rdConcurRowVer)
If RCari2.RowCount <> 0 Then
    IPT = RCari2("NAMA")
    
    SCari3 = "Select * From M001 where Rangka = '" + Trim(Text4) + "'"
    Set RCari3 = RDCO.OpenResultset(SCari3, rdOpenKeyset, rdConcurRowVer)
    RCari3.EDIT
        RCari3("MTS_MOTOR") = IPT
    RCari3.Update
    RCari3.Close
    Set RCari3 = Nothing
    
        SCari4 = "Select * From JL005"
        Set RCari4 = RDCO.OpenResultset(SCari4, rdOpenKeyset, rdConcurRowVer)
        RCari4.AddNew
        RCari4("TGL") = Text8
        RCari4("ccab") = Combo1
        RCari4("Type") = Trim(Text1)
        RCari4("Rangka") = Trim(Text4)
        RCari4("Mesin") = Trim(Text5)
        RCari4("Tanggal") = Tanggal
        RCari4("User_Code") = Operator
        RCari4.Update
        RCari4.Close
        Set RCari4 = Nothing
        
End If
RCari2.Close
Set RCari2 = Nothing

Unload Me
JL005.Show
End Sub
