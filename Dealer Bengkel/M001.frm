VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form M001 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ENTRI DATA MOTOR"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   12510
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   7560
      TabIndex        =   6
      Text            =   "Text9"
      Top             =   3930
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SIMPAN"
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
      Left            =   5775
      TabIndex        =   20
      Top             =   4515
      Width           =   960
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   10575
      TabIndex        =   8
      Text            =   "Text8"
      Top             =   3930
      Width           =   1500
   End
   Begin MSFlexGridLib.MSFlexGrid GridFak 
      Height          =   2085
      Left            =   2775
      TabIndex        =   19
      Top             =   960
      Width           =   3960
      _ExtentX        =   6985
      _ExtentY        =   3678
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   1
      Appearance      =   0
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   9060
      TabIndex        =   7
      Text            =   "Text7"
      Top             =   3930
      Width           =   1500
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   6060
      TabIndex        =   5
      Text            =   "Text6"
      Top             =   3930
      Width           =   1500
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   4560
      TabIndex        =   4
      Text            =   "Text5"
      Top             =   3930
      Width           =   1500
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   3045
      TabIndex        =   3
      Text            =   "Text4"
      Top             =   3930
      Width           =   1500
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   1545
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   3930
      Width           =   1500
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   45
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   3930
      Width           =   1500
   End
   Begin VB.Frame Frame1 
      Caption         =   "Info Faktur"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   1365
      TabIndex        =   10
      Top             =   105
      Width           =   9780
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   1395
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   465
         Width           =   2370
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "Label6"
         Height          =   300
         Left            =   8295
         TabIndex        =   17
         Top             =   315
         Width           =   1350
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "Label4"
         Height          =   300
         Left            =   6090
         TabIndex        =   16
         Top             =   690
         Width           =   3555
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "Label5"
         Height          =   300
         Left            =   5670
         TabIndex        =   15
         Top             =   315
         Width           =   1350
      End
      Begin VB.Label Label3 
         Caption         =   "HARGA BELI          Rp."
         Height          =   225
         Left            =   4305
         TabIndex        =   14
         Top             =   735
         Width           =   2520
      End
      Begin VB.Label Label2 
         Caption         =   "JUMLAH UNIT"
         Height          =   225
         Left            =   7140
         TabIndex        =   13
         Top             =   360
         Width           =   1365
      End
      Begin VB.Label Label8 
         Caption         =   "TGL PEMBELIAN"
         Height          =   225
         Left            =   4305
         TabIndex        =   12
         Top             =   360
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "NO. FAKTUR"
         Height          =   225
         Left            =   105
         TabIndex        =   11
         Top             =   540
         Width           =   1140
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   2445
      Left            =   45
      TabIndex        =   18
      Top             =   1455
      Width           =   12420
      _ExtentX        =   21908
      _ExtentY        =   4313
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   65280
      ForeColorFixed  =   0
      BackColorBkg    =   16777152
      Appearance      =   0
   End
   Begin VB.CommandButton OK 
      Caption         =   "!"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8100
      TabIndex        =   9
      Top             =   2730
      Width           =   555
   End
End
Attribute VB_Name = "M001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private RSSTSM, RSimpanM001, RSimpanM001B, RDel, RGrid2, RFak, RGridFak As rdoResultset
Private SSTSM, SSimpanM001, SSimpanM001B, SDel, SGrid2, SFak, SGridFak As String
Private J_UNIT

Private Sub Command1_Click()
Call SimpanM001
Call STSM001
Unload Me
M001.Show
End Sub

Private Sub STSM001()
SSTSM = "Select * From F001 where NO_FAK = '" + Trim(Text1) + "'"
Set RSTSM = RDCO.OpenResultset(SSTSM, rdOpenDynamic, rdConcurRowVer)
RSTSM.EDIT
    RSTSM("STS_M001") = 1
RSTSM.Update
RSTSM.Close
Set RSTSM = Nothing
End Sub

Private Sub SimpanM001()
SSimpanM001 = "Select * From M002"
Set RSimpanM001 = RDCO.OpenResultset(SSimpanM001, rdOpenKeyset, rdConcurRowVer)
RSimpanM001.MoveFirst
Do While Not RSimpanM001.EOF
    
    SSimpanM001B = "Select * From M001"
    Set RSimpanM001B = RDCO.OpenResultset(SSimpanM001B, rdOpenKeyset, rdConcurRowVer)
    RSimpanM001B.AddNew
        RSimpanM001B("TYPE") = RSimpanM001("TYPE")
        RSimpanM001B("WARNA") = RSimpanM001("WARNA")
        RSimpanM001B("TAHUN") = RSimpanM001("TAHUN")
        RSimpanM001B("RANGKA") = RSimpanM001("RANGKA")
        RSimpanM001B("MESIN") = RSimpanM001("MESIN")
        RSimpanM001B("CCAB") = CodeCab
        RSimpanM001B("DO") = RSimpanM001("DO")
        RSimpanM001B("SJ") = RSimpanM001("SJ")
        RSimpanM001B("STS_JUAL") = 0
        RSimpanM001B("TGL_INPUT") = Tanggal
        RSimpanM001B("H_BELI") = CCur(RSimpanM001("H_BELI"))
        RSimpanM001B("MTS_MOTOR") = CodeCab
    RSimpanM001B.Update
    RSimpanM001B.Close
    Set RSimpanM001B = Nothing

RSimpanM001.MoveNext
Loop
RSimpanM001.Close
Set RSimpanM001 = Nothing
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=DEALER", rdDriverNoPrompt, False, CN)

Call IsiGridFak
Call Kosong
Call SiapkanGrid2
Call DelM002

GridFak.Visible = False

J_UNIT = 0
End Sub

Private Sub DelM002()
SDel = "Delete * From M002"
Set RDel = RDCO.OpenResultset(SDel, rdOpenDynamic, rdConcurRowVer)
RDel.Close
Set RDel = Nothing
End Sub

Private Sub IsiGridFak()
With GridFak
    .Cols = 2
    .Row = 0
    .Col = 0: .ColWidth(0) = 1500: .Text = "NO FAK": .CellAlignment = 4
    .Col = 1: .ColWidth(1) = 2000: .Text = "TGL BELI": .CellAlignment = 4
End With
    
    SGridFak = "Select Top 10 No_System,No_Fak, Tgl_Beli From F001 where STS_M001=0 order by NO_SYSTEM Desc"
    Set RGridFak = RDCO.OpenResultset(SGridFak, rdOpenKeyset, rdConcurReadOnly)
    If RGridFak.RowCount <> 0 Then
       RGridFak.MoveFirst
       B = 1
       Do Until RGridFak.EOF
          GridFak.Rows = B + 1
          GridFak.Row = B
             With GridFak
                  .Col = 0: .Text = RGridFak("NO_FAK"): .CellAlignment = 4
                  .Col = 1: .Text = RGridFak("TGL_BELI"): .CellAlignment = 4
             End With
          B = B + 1
          RGridFak.MoveNext
       Loop
    End If
    RGridFak.Close
    Set RGridFak = Nothing
    
End Sub

Private Sub Kosong()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Text9 = ""
Text7 = ""
Text8 = ""
Label4 = ""
Label5 = ""
Label6 = ""
End Sub

Private Sub SiapkanGrid2()
With Grid2
    .Row = 0
    .Cols = 8
    .Col = 0: .ColWidth(0) = 1500: .Text = "TYPE": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 1: .ColWidth(1) = 1500: .Text = "WARNA": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 2: .ColWidth(2) = 1500: .Text = "RANGKA": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 3: .ColWidth(3) = 1500: .Text = "MESIN": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 4: .ColWidth(4) = 1500: .Text = "TAHUN": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 5: .ColWidth(5) = 1500: .Text = "DO/PSMUP": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 6: .ColWidth(6) = 1500: .Text = "SJ": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
    .Col = 7: .ColWidth(7) = 1500: .Text = "HARGA": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 10
End With
End Sub

Private Sub GridFak_dblClick()
FAKTUR = ""
FAKTUR = GridFak.TextMatrix(GridFak.Row, 0)
Text1 = GridFak.TextMatrix(GridFak.Row, 0)
GridFak.Visible = False
Text1.SetFocus
End Sub

Private Sub OK_Click()
If J_UNIT = Label6 Or J_HARGA = Label4 Then
    MsgBox "JUMLAH ENTRI KENDARAAN TIDAK BOLEH MELEBIHI JUMLAH UNIT", vbCritical, "WARNING"
    Text2 = ""
    Text3 = ""
    Text4 = ""
    Text5 = ""
    Text6 = ""
    Text7 = ""
    Text8 = ""
    Text9 = ""
    Text2.SetFocus
    Exit Sub
Else
    J_UNIT = J_UNIT + 1
    Call SimpanM002
    Call IsiGrid2
End If
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Text7 = ""
Text8 = ""
Text9 = ""
Text2.SetFocus
End Sub

Private Sub IsiGrid2()
SGrid2 = "Select * From M002 order by NO Asc"
Set RGrid2 = RDCO.OpenResultset(SGrid2, rdOpenKeyset, rdConcurReadOnly)
If RGrid2.RowCount <> 0 Then
   RGrid2.MoveFirst
   B = 1
   Do Until RGrid2.EOF
      Grid2.Rows = B + 1
      Grid2.Row = B
         With Grid2
              .Col = 0: .Text = RGrid2("TYPE"): .CellAlignment = 4
              .Col = 1: .Text = RGrid2("WARNA"): .CellAlignment = 4
              .Col = 2: .Text = RGrid2("RANGKA"): .CellAlignment = 4
              .Col = 3: .Text = RGrid2("MESIN"): .CellAlignment = 4
              .Col = 4: .Text = RGrid2("TAHUN"): .CellAlignment = 4
              .Col = 5: .Text = RGrid2("DO"): .CellAlignment = 4
              .Col = 6: .Text = RGrid2("SJ"): .CellAlignment = 4
              .Col = 7: .Text = Format(RGrid2("H_BELI"), "##,###.00")
         End With
      B = B + 1
      RGrid2.MoveNext
   Loop
End If
RGrid2.Close
Set RGrid2 = Nothing
End Sub

Private Sub SimpanM002()
SSave = "Select * from M002"
Set RSave = RDCO.OpenResultset(SSave, rdOpenKeyset, rdConcurRowVer)
RSave.AddNew
    RSave("TYPE") = Trim(Text2)
    RSave("WARNA") = Trim(Text3)
    RSave("RANGKA") = Trim(Text4)
    RSave("MESIN") = Trim(Text5)
    RSave("TAHUN") = Trim(Text6)
    RSave("DO") = Trim(Text9)
    RSave("SJ") = Trim(Text7)
    RSave("H_BELI") = CCur(Text8)
RSave.Update
RSave.Close
Set RSave = Nothing
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF1
        GridFak.Visible = True
        GridFak.ZOrder
    Case vbKeyEscape
        GridFak.Visible = False
End Select
End Sub

Private Sub text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text1_LostFocus()
If Text1 = "" Then Exit Sub
SFak = "Select * From F001 where NO_FAK = '" + Trim(Text1) + "'"
Set RFak = RDCO.OpenResultset(SFak, rdOpenDynamic, rdOpenKeyset)
If RFak.RowCount <> 0 Then
        Label5 = RFak("TGL_BELI")
        Label6 = RFak("JUMLAH")
        Label4 = Format(RFak("H_JUMLAH"), "##,###.00")
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text2 = Format(Text2, ">")
    SendKeys vbTab
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text3 = Format(Text3, ">")
    SendKeys vbTab
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text4 = Format(Text4, ">")
    SendKeys vbTab
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text5 = Format(Text5, ">")
    SendKeys vbTab
End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text9 = Format(Text9, ">")
    SendKeys vbTab
End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text7 = Format(Text7, ">")
    SendKeys vbTab
End If
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text8 = Format(Text8, "##,###.00")
    SendKeys vbTab
End If
End Sub

