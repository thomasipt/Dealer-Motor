VERSION 5.00
Begin VB.Form LOGIN 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LOGIN"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4245
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   77.174
   ScaleMode       =   0  'User
   ScaleWidth      =   283
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
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
      Left            =   1725
      TabIndex        =   0
      Text            =   "Text1"
      ToolTipText     =   "Nama user"
      Top             =   180
      Width           =   2370
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
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
      IMEMode         =   3  'DISABLE
      Left            =   1725
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "Text1"
      ToolTipText     =   "Password user"
      Top             =   795
      Width           =   2370
   End
   Begin VB.CommandButton cmdCLOSE 
      Caption         =   "KELUAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2205
      TabIndex        =   3
      ToolTipText     =   "Klik untuk keluar"
      Top             =   1380
      Width           =   1890
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "MASUK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   150
      TabIndex        =   2
      ToolTipText     =   "Klik untuk masuk ke sistem"
      Top             =   1380
      Width           =   1890
   End
   Begin VB.Image Image1 
      Height          =   1035
      Left            =   150
      Picture         =   "LOGIN.frx":0000
      Stretch         =   -1  'True
      Top             =   180
      Width           =   1230
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000080&
      BorderWidth     =   3
      FillColor       =   &H008080FF&
      FillStyle       =   0  'Solid
      Height          =   2010
      Left            =   67
      Top             =   60
      Width           =   4110
   End
End
Attribute VB_Name = "LOGIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private SqlPass As String
Private tUser As rdoResultset
Private tMasuk As rdoResultset

Private RTgl, RHapus, RDel, RSave2, RSave3, RSave4, RCari, RCari2, RSLNO, rscs3 As rdoResultset
Private STgl, SHapus, SDel, SSave2, SSave3, SSave4, SCari, SCari2, SqlNo, sqlcs3, KODE As String

Private Sub cmdCLOSE_Click()
End
End Sub

Private Sub Masuk2()
SCari = "Select * From C013 where UserCode = '" + Text1 + "' and Password = '" + Text2 + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset)
If RCari.RowCount <> 0 Then
    Call Masuk
    Unload Me
Else
    LOGIN.Hide
    MsgBox "ANDA TIDAK BERHAK LOG IN KE SYSTEM", vbCritical, "KONFIRMASI"
    LOGIN.Show
    Text1 = ""
    Text2 = ""
    Text1.SetFocus
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub Masuk()
SqlPass = "Select * from C013 where UserCode =  '" + Trim(Text1) + "' "
Set tMasuk = RDCO.OpenResultset(SqlPass, rdOpenDynamic, rdConcurRowVer)
If tMasuk.RowCount <> 0 Then
    If tMasuk("MAIN") = "01" Then
        Operator = Trim(tMasuk("Nama"))
        CodeCab = tMasuk("CodeCab")
        Status = tMasuk("Main")
        'Operator = tMasuk("UserCode")
        NoUser = tMasuk("NoUrut")
        G_DEBET = tMasuk("GDebet")
        G_CREDIT = tMasuk("GCredit")
        F_DEBET = tMasuk("FDebet")
        F_CREDIT = tMasuk("FCredit")
        MAINSALE.Show
    ElseIf tMasuk("MAIN") = "03" Then
        Operator = Trim(tMasuk("Nama"))
        CodeCab = tMasuk("CodeCab")
        Status = tMasuk("Main")
        'Operator = tMasuk("UserCode")
        NoUser = tMasuk("NoUrut")
        G_DEBET = tMasuk("GDebet")
        G_CREDIT = tMasuk("GCredit")
        F_DEBET = tMasuk("FDebet")
        F_CREDIT = tMasuk("FCredit")
        MAINSERVICE.Show
    ElseIf tMasuk("MAIN") = "02" Then
        Operator = Trim(tMasuk("Nama"))
        CodeCab = tMasuk("CodeCab")
        Status = tMasuk("Main")
        'Operator = tMasuk("UserCode")
        NoUser = tMasuk("NoUrut")
        G_DEBET = tMasuk("GDebet")
        G_CREDIT = tMasuk("GCredit")
        F_DEBET = tMasuk("FDebet")
        F_CREDIT = tMasuk("FCredit")
        MAINSUPER.Show
    End If
End If
tMasuk.Close
Set tMasuk = Nothing
End Sub

Private Sub CmdOK_Click()
Call Masuk2
End Sub

Private Sub TGL()
STgl = "Select * from A001"
Set RTgl = RDCO.OpenResultset(STgl, rdOpenDynamic, rdConcurRowVer)
If RTgl.RowCount <> 0 Then
    Tanggal = RTgl("Tanggal")
    N_CCAB = RTgl("Nama")
    N_ALAMAT = RTgl("Alamat")
Else
End If
RTgl.Close
Set RTgl = Nothing
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=DEALER", rdDriverNoPrompt, False, CN)
Text1 = ""
Text2 = ""
Call TGL
End Sub

Private Sub text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Text1_LostFocus()
Text1 = Format(Text1, ">")
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Text2_LostFocus()
Text2 = Format(Text2, ">")
End Sub



