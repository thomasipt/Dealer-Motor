VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form RP011 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "STATEMENT GENERAL LEDGER"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5805
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   5805
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "TAMPILKAN"
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
      Left            =   2115
      TabIndex        =   1
      Top             =   1935
      Width           =   1575
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
      Left            =   2407
      TabIndex        =   2
      Top             =   3150
      Width           =   990
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   1140
      MaxLength       =   7
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1050
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
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
      Left            =   1395
      TabIndex        =   3
      Top             =   2610
      Width           =   3015
   End
   Begin Crystal.CrystalReport CRPT 
      Left            =   3555
      Top             =   4275
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   330
      Left            =   1057
      TabIndex        =   7
      Top             =   255
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   582
      _Version        =   393216
      Format          =   54591489
      CurrentDate     =   39531
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   330
      Left            =   3937
      TabIndex        =   8
      Top             =   255
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   582
      _Version        =   393216
      Format          =   54591489
      CurrentDate     =   39531
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "TGL AWAL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   67
      TabIndex        =   10
      Top             =   315
      Width           =   990
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "TGL AKHIR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   2902
      TabIndex        =   9
      Top             =   315
      Width           =   990
   End
   Begin VB.Label Label1 
      Caption         =   "CODE SL"
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
      Left            =   60
      TabIndex        =   6
      Top             =   1050
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "NAMA SL"
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
      Left            =   60
      TabIndex        =   4
      Top             =   1410
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label3"
      Height          =   300
      Left            =   1140
      TabIndex        =   5
      Top             =   1410
      Width           =   3615
   End
   Begin VB.Line Line1 
      X1              =   -165
      X2              =   5940
      Y1              =   2535
      Y2              =   2535
   End
   Begin VB.Line Line2 
      X1              =   -165
      X2              =   5940
      Y1              =   2970
      Y2              =   2970
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      BorderWidth     =   3
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   915
      Left            =   0
      Top             =   0
      Width           =   6225
   End
End
Attribute VB_Name = "RP011"
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

Private Sub Command1_Click()
Dim A As Integer
A = Month(Tanggal)

D1 = Day(DTPicker1)
M1 = Month(DTPicker1)
Y1 = Year(DTPicker1)

D2 = Day(DTPicker2)
M2 = Month(DTPicker2)
Y2 = Year(DTPicker2)

If Text1 = "" Then Exit Sub
    
    crpt.ReportFileName = App.Path + "\ReportD\HisGL.rpt"
    crpt.SelectionFormula = "{G005.Codesl} = '" + Trim(Text1) + "' and {G005.tanggal} in date (" & Y1 & "," & M1 & "," & D1 & ") to date (" & Y2 & "," & M2 & "," & D2 & ")"
    crpt.WindowState = crptMaximized
    crpt.WindowMaxButton = False
    crpt.WindowMinButton = False
    crpt.Action = 1
    crpt.Reset
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=DEALER", rdDriverNoPrompt, False, CN)

Text1 = ""
Label3 = ""

Text2 = ""
Label6 = ""

Text3 = ""
Text4 = ""
Text5 = ""

Call Montok
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

Private Sub text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Montok()
D = 1
M = Month(Tanggal)
Y = Year(Tanggal)

DTPicker1 = Format(DateSerial(Y, M, D), "DD/MM/YYYY")
DTPicker2 = Tanggal

End Sub

Private Sub Text1_LostFocus()

If Text1 = "" Then Exit Sub

SCari = "Select NamaSl From G005 where CodeSL= '" + Trim(Text1) + "'  order by NoUrut"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    Label3 = RCari("Namasl")
Else
    MsgBox "BELUM ADA MUTASI (GL BELUM TERDAFTAR)", vbCritical, "CODE SUB GL"
    Text1.SetFocus
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Text2_LostFocus()
Dim A As Integer
A = Month(Tanggal)
If Text2 = "" Then Exit Sub
SCari = "Select NamaSl From G005 where CodeSL= '" + Trim(Text2) + "'  order by NoUrut"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    Label6 = RCari("Namasl")
    crpt.ReportFileName = App.Path + "\ReportD\HisGL2.rpt"
    crpt.SelectionFormula = "{STATGLT.BULAN} >= " & Text3 & " and {STATGLT.BULAN} <= " & Text5 & " and {STATGLT.THN} = " & Text4 & " and {STATGLT.Codesl} = '" + Trim(Text2) + "'"
    crpt.WindowState = crptMaximized
    crpt.WindowMaxButton = False
    crpt.WindowMinButton = False
    crpt.Action = 1
    crpt.Reset
Else
    MsgBox "BELUM ADA MUTASI (GL BELUM TERDAFTAR)", vbCritical, "CODE SUB GL"
    Text2.SetFocus
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub text3_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub text4_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub
