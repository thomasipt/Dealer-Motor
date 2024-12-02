VERSION 5.00
Begin VB.Form B003AA 
   Caption         =   "EDIT JUMLAH"
   ClientHeight    =   3480
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7890
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3480
   ScaleWidth      =   7890
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2250
      TabIndex        =   3
      Text            =   "Text4"
      Top             =   2160
      Width           =   1860
   End
   Begin VB.CommandButton TmbSave 
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
      Left            =   608
      TabIndex        =   4
      Top             =   2940
      Width           =   1000
   End
   Begin VB.CommandButton Command1 
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
      Left            =   6323
      TabIndex        =   5
      Top             =   2940
      Width           =   960
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2250
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   1575
      Width           =   1860
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2250
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   585
      Width           =   5505
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2250
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1080
      Width           =   5505
   End
   Begin VB.PictureBox Picture1 
      Height          =   1320
      Left            =   -90
      ScaleHeight     =   1260
      ScaleWidth      =   9090
      TabIndex        =   9
      Top             =   2745
      Width           =   9150
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "WARNING"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   180
      TabIndex        =   11
      Top             =   90
      Width           =   7575
   End
   Begin VB.Label Label1 
      Caption         =   "HARGA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   135
      TabIndex        =   10
      Top             =   2160
      Width           =   1830
   End
   Begin VB.Label Label2 
      Caption         =   "KODE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   135
      TabIndex        =   8
      Top             =   585
      Width           =   1830
   End
   Begin VB.Label Label3 
      Caption         =   "NAMA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   135
      TabIndex        =   7
      Top             =   1080
      Width           =   1830
   End
   Begin VB.Label Label7 
      Caption         =   "JUMLAH"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   135
      TabIndex        =   6
      Top             =   1575
      Width           =   1830
   End
End
Attribute VB_Name = "B003AA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private RGol, RCari, RKode, RDel, RDelBar, RSim, RSave, RSaveP, RSaveP2, RDist As rdoResultset
Private SDelBar, SDist, SGol, SCari, Metode, SKode, SDel, SSim, SSave, SSaveP, SSaveP2 As String
Private Brs, MetodLaba, Ganti, TOKET, WARNA, SPART, Montok

Private Sub Command1_Click()
Unload Me
B003A.Show 1
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=DEALER", rdDriverNoPrompt, False, CN)
Text1 = Kode
Text2 = NamaBar
Text3 = JumlahBar
Text4 = HargaBar
Call CekStatus

End Sub

Private Sub CekStatus()
SCari = "Select * From B003A where KODE_JNS = '" + Trim(Kode) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset)
If RCari.RowCount <> 0 Then
    Label4 = "POSISI EDIT = " + Trim(RCari("HKosong"))
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text1_LostFocus()
Text1 = Format(Text1, ">")
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text2_LostFocus()
Text2 = Format(Text2, ">")
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub text4_LostFocus()
Text4 = Format(Text4, "##,###.00")
End Sub

Private Sub TmbSave_Click()
If Text2 = "" Or Text3 = "" Or Text4 = "" Then
    MsgBox "MASIH ADA DATA YANG KOSONG", vbCritical, "WARNING"
    Exit Sub
End If

SSaveP = "Select * From B003 where KODE_JNS = '" + Trim(Kode) + "'"
Set RSaveP = RDCO.OpenResultset(SSaveP, rdOpenDynamic, rdConcurRowVer)
RSaveP.EDIT
    RSaveP("KODE_JNS") = Trim(Text1)
    RSaveP("NAMA_JNS") = Trim(Text2)
    RSaveP("JML_AWAL") = CCur(Text3)
    RSaveP("JML_AKHIR") = CCur(Text3)
    RSaveP("TANGGAL") = Tanggal
RSaveP.Update
RSaveP.Close
Set RSaveP = Nothing

SSaveP2 = "Select * From B003A where KODE_JNS = '" + Trim(Kode) + "'"
Set RSaveP2 = RDCO.OpenResultset(SSaveP2, rdOpenDynamic, rdConcurRowVer)
RSaveP2.EDIT
    RSaveP2("KODE_JNS") = Trim(Text1)
    RSaveP2("NAMA_JNS") = Trim(Text2)
    RSaveP2("SALDOAWAL") = CCur(Text3) * CCur(Text4)
    RSaveP2("SALDO") = CCur(Text3) * CCur(Text4)
    RSaveP2("HBeli") = CCur(Text4)
    RSaveP2("HJual") = CCur(Text4)
    
    RSaveP2("HKosong") = RSaveP2("HKosong") + 1
    
    RSaveP2("TANGGAL") = Tanggal
RSaveP2.Update
RSaveP2.Close
Set RSaveP2 = Nothing

Unload Me
B003A.Show 1

End Sub
