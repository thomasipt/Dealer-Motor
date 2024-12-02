VERSION 5.00
Begin VB.Form BACKUP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BACKUP & RESTORE"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4230
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   4230
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdBackup 
      Caption         =   "BACKUP"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   120
      TabIndex        =   3
      Top             =   105
      Width           =   1110
   End
   Begin VB.CommandButton CmdRestore 
      Caption         =   "RESTORE"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1110
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CLOSE"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   120
      TabIndex        =   1
      Top             =   2535
      Width           =   1110
   End
   Begin VB.FileListBox File1 
      Height          =   3015
      Left            =   1395
      Pattern         =   "*.IPT"
      TabIndex        =   0
      Top             =   105
      Width           =   2715
   End
End
Attribute VB_Name = "BACKUP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Bowook
Private D, M, Y As String

Private Sub CmdBackup_Click()
Dim B
    D = Day(Tanggal)
    M = Month(Tanggal)
    Y = Year(Tanggal)
    B = "DEALER_" + Trim(D) + "_" + Trim(M) + "_" + Trim(Y)
    ChDrive "D:\"
    FileCopy "D:\DATABASE\DEALER.mdb", ("D:\DATABASE\BACKUP\" + Trim(B) + ".IPT")
    MsgBox "PROSES BACKUP DATA TELAH SELESAI", vbInformation, "BACKUP DATA"
    File1.Refresh
End Sub

Private Sub CmdRestore_Click()
Dim B, A
D = Day(Tanggal)
M = Month(Tanggal)
Y = Year(Tanggal)
A = MsgBox("ANDA AKAN MENJALANKAN RESTORE DATA", vbOKCancel, "RESTORE DATA")
If File1 = "" Then
    MsgBox "PILIH DATA YANG AKAN DIRESTORE", vbInformation, "DATA KOSONG"
ElseIf A = vbOK Then
    FileCopy ("D:\DATABASE\BACKUP\" + File1), ("D:\DATABASE\DEALER.mdb")
    MsgBox "PROSES RESTORE DATA TELAH SELESAI", vbInformation, "RESTORE DATA"
End If
File1.Refresh
End Sub

Private Sub Command1_Click()
End
End Sub

Private Sub Form_Load()
File1.Path = "D:\DATABASE\BACKUP\"
End Sub

