VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form C012 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ENTRI DATA CUSTUMER / SUPPLIER / CABANG"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8670
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   8670
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
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
      Left            =   7545
      TabIndex        =   13
      Top             =   2580
      Width           =   960
   End
   Begin VB.CommandButton Command2 
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
      Left            =   3855
      TabIndex        =   12
      Top             =   5670
      Width           =   960
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   1523
      MaxLength       =   5
      TabIndex        =   0
      Text            =   "Text5"
      Top             =   90
      Width           =   1050
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   1523
      MaxLength       =   35
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   555
      Width           =   3930
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   1523
      MaxLength       =   50
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   1005
      Width           =   6690
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   1523
      MaxLength       =   35
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   1455
      Width           =   2715
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   1523
      MaxLength       =   35
      TabIndex        =   4
      Text            =   "Text4"
      Top             =   1905
      Width           =   2715
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
      Left            =   105
      TabIndex        =   5
      Top             =   2580
      Width           =   960
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   2445
      Left            =   105
      TabIndex        =   11
      ToolTipText     =   "Klik untuk edit"
      Top             =   3150
      Width           =   8400
      _ExtentX        =   14817
      _ExtentY        =   4313
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
   Begin Crystal.CrystalReport CRPT 
      Left            =   6090
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label2 
      Caption         =   "KODE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   465
      TabIndex        =   10
      Top             =   105
      Width           =   825
   End
   Begin VB.Label Label3 
      Caption         =   "N A M A"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   465
      TabIndex        =   9
      Top             =   600
      Width           =   1050
   End
   Begin VB.Label Label4 
      Caption         =   "ALAMAT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   465
      TabIndex        =   8
      Top             =   1050
      Width           =   1050
   End
   Begin VB.Label Label5 
      Caption         =   "KOTA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   465
      TabIndex        =   7
      Top             =   1500
      Width           =   915
   End
   Begin VB.Label Label6 
      Caption         =   "TELEPON"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   465
      TabIndex        =   6
      Top             =   1950
      Width           =   960
   End
   Begin VB.Line Line1 
      X1              =   150
      X2              =   8520
      Y1              =   2445
      Y2              =   2445
   End
End
Attribute VB_Name = "C012"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private RCari As rdoResultset
Private SCari As String

Private Sub Check1_Click()
Check2.Value = 0
Check3.Value = 0
Command4.Height = 540
End Sub

Private Sub Check2_Click()
Check1.Value = 0
Check3.Value = 0
Command4.Height = 540
End Sub

Private Sub Check3_Click()
Check1.Value = 0
Check2.Value = 0
Command4.Height = 540
End Sub

Private Sub Command1_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Then
    MsgBox "MASIH ADA DATA YANG KOSONG", vbCritical, "DATA TIDAK BOLEH KOSONG"
    Exit Sub
End If

SSimpan = "Select * From C012"
Set RSimpan = RDCO.OpenResultset(SSimpan, rdOpenKeyset, rdConcurRowVer)
RSimpan.AddNew
    RSimpan("Nonas") = Text5
    RSimpan("Nama") = Text1
    RSimpan("Alamat1") = Text2
    RSimpan("Kota") = Text3
    RSimpan("Telpon") = Text4
RSimpan.Update
RSimpan.Close
Set RSimpan = Nothing

PESAN = True
Call CariNomor
Call Kosong
Text1.SetFocus
Unload Me
C012.Show
End Sub

Private Sub Command2_Click()
crpt.ReportFileName = App.Path + "\ReportD\C012.rpt"
crpt.WindowState = crptMaximized
crpt.WindowMaxButton = True
crpt.WindowMinButton = True
crpt.Action = 1
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=DEALER", rdDriverNoPrompt, False, CN)

'Call CariNomor

Call Kosong
PESAN = False

Call SiapkanGrid
Call Tampilkan

End Sub

Private Sub SiapkanGrid()
With grid
    .Row = 0
    .Cols = 5
    .Col = 0: .ColWidth(0) = 1000: .Text = "NO": .CellAlignment = 4: .CellFontBold = True
    .Col = 1: .ColWidth(1) = 2500: .Text = "NAMA": .CellAlignment = 4: .CellFontBold = True
    .Col = 2: .ColWidth(2) = 3000: .Text = "ALAMAT": .CellAlignment = 4: .CellFontBold = True
    .Col = 3: .ColWidth(3) = 1250: .Text = "KOTA": .CellAlignment = 4: .CellFontBold = True
    .Col = 4: .ColWidth(4) = 1250: .Text = "TELEPON": .CellAlignment = 4: .CellFontBold = True
End With
End Sub

Private Sub Tampilkan()
Dim Brs
Brs = 1
SKode = "Select * From C012 order by NoNas Asc"
Set RKode = RDCO.OpenResultset(SKode, rdOpenDynamic, rdOpenKeyset)
If RKode.RowCount <> 0 Then
RKode.MoveFirst
Do Until RKode.EOF
    With grid
        .Rows = Brs + 1
        .Row = Brs
        .Col = 0: .Text = RKode("NoNas"): .CellAlignment = 4
        .Col = 1: .Text = RKode("Nama")
        .Col = 2: .Text = RKode("Alamat1"): .CellAlignment = 4
        .Col = 3: .Text = RKode("Kota"): .CellAlignment = 4
        .Col = 4: .Text = RKode("Telpon"): .CellAlignment = 4
        Brs = Brs + 1
    End With
RKode.MoveNext
Loop
End If
RKode.Close
Set RKode = Nothing
End Sub

Private Sub Kosong()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
End Sub

Private Sub CariNomor()
Dim Nomor As Double
Dim InfoNomor As Double

SCari = "Select Top 1 NoNas From C012 order by NoNas Desc"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    Nomor = Val(RCari("NoNas")) + 1
    InfoNomor = Digit(5, Val(RCari("Nonas")))
    If PESAN = True Then
        MsgBox "NOMOR TERSIMPAN " + Trim(Label6), vbOKOnly, "DATA TERSIMPAN"
    End If
    Label1 = Nomor
Else
    Label1 = "00001"
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys vbTab
    Text1 = Format(Text1, ">")
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys vbTab
    Text2 = Format(Text2, ">")
End If
End Sub

Private Sub text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys vbTab
    Text3 = Format(Text3, ">")
End If
End Sub

Private Sub text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys vbTab
    Text4 = Format(Text4, ">")
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys vbTab
    Text5 = Format(Text5, ">")
End If
End Sub

