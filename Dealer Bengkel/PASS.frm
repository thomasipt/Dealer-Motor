VERSION 5.00
Begin VB.Form PASS 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PERUBAHAN PASSWORD"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4725
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   4725
   StartUpPosition =   2  'CenterScreen
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
      IMEMode         =   3  'DISABLE
      Left            =   2137
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   2452
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
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
      IMEMode         =   3  'DISABLE
      Left            =   2145
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   457
      Width           =   2175
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "SIMPAN"
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
      Left            =   202
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3457
      Width           =   1890
   End
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
      IMEMode         =   3  'DISABLE
      Left            =   2137
      MaxLength       =   8
      TabIndex        =   1
      Text            =   "Text4"
      Top             =   1297
      Width           =   2175
   End
   Begin VB.TextBox Text5 
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
      IMEMode         =   3  'DISABLE
      Left            =   2137
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   2872
      Width           =   2175
   End
   Begin VB.TextBox Text6 
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
      IMEMode         =   3  'DISABLE
      Left            =   2137
      MaxLength       =   8
      TabIndex        =   2
      Text            =   "Text6"
      Top             =   1717
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
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
      Left            =   2670
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3457
      Width           =   1890
   End
   Begin VB.Image Image1 
      Height          =   795
      Left            =   712
      Picture         =   "PASS.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   870
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "KONFIRMASI"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   345
      TabIndex        =   11
      Top             =   2947
      Width           =   1275
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "PASSWORD LAMA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2137
      TabIndex        =   10
      Top             =   142
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H008080FF&
      Caption         =   "USER BARU"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   345
      TabIndex        =   9
      Top             =   1297
      Width           =   1485
   End
   Begin VB.Label Label4 
      BackColor       =   &H008080FF&
      Caption         =   "PASSWORD BARU"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   345
      TabIndex        =   8
      Top             =   2474
      Width           =   2820
   End
   Begin VB.Label Label5 
      BackColor       =   &H008080FF&
      Caption         =   "KONFIRMASI"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   345
      TabIndex        =   7
      Top             =   1792
      Width           =   1275
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000080&
      BorderWidth     =   3
      FillColor       =   &H008080FF&
      FillStyle       =   0  'Solid
      Height          =   3060
      Left            =   45
      Top             =   1087
      Width           =   4635
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000080&
      BorderWidth     =   3
      FillColor       =   &H008080FF&
      FillStyle       =   0  'Solid
      Height          =   960
      Left            =   45
      Top             =   37
      Width           =   4635
   End
End
Attribute VB_Name = "PASS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private RCari, RSave As rdoResultset
Private SCari, SSave As String


Private Sub CmdOK_Click()
If Text1 = "" Or Text2 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Then
    MsgBox "DATA MASIH KOSONG", vbCritical, "KONFIRMASI"
    Text1.SetFocus
    Exit Sub
Else
    SSave = "Select * From C013 where Nama = '" + Trim(Operator) + "'"
    Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
    RSave.EDIT
        RSave("Password") = Trim(Text5)
        RSave("UserCode") = Trim(Text6)
    RSave.Update
    RSave.Close
    Set RSave = Nothing
    End
End If
End Sub

Private Sub Command1_Click()
Unload Me
LOGIN.Show
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=DEALER", rdDriverNoPrompt, False, CN)

ClearTextBoxes Me
Text3 = Operator
End Sub

Private Sub text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text1 = Format(Text1, ">")
        SCari = "Select * From C013 where Nama = '" + Trim(Operator) + "'"
        Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset)
        If RCari("Password") = Text1 Then
            SendKeys "{TAB}"
        Else
            MsgBox "PASSWORD SALAH, SISTEM AKAN TERTUTUP", vbCritical, "KONFIRMASI"
            End
        End If
        RCari.Close
        Set RCari = Nothing
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text2 = Format(Text2, ">")
    SendKeys "{TAB}"
End If
End Sub

Private Sub text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text4 = Format(Text4, ">")
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text5 = Format(Text5, ">")
    SendKeys "{TAB}"
        If Text5 <> Text2 Then
            MsgBox "USER BARU TIDAK SESUAI", vbCritical, "KONFIRMASI"
            Text5 = ""
            Text5.SetFocus
        End If
End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text6 = Format(Text6, ">")
    SendKeys "{TAB}"
        If Text6 <> Text4 Then
            MsgBox "USER BARU TIDAK SESUAI", vbCritical, "KONFIRMASI"
            Text6 = ""
            Text6.SetFocus
        End If
End If
End Sub
