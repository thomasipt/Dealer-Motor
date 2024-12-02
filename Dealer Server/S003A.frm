VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form S003A 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CETAK NOTA SERVICE & SPAREPART"
   ClientHeight    =   1200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5610
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   5610
   StartUpPosition =   2  'CenterScreen
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
      Height          =   390
      Left            =   1853
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   105
      Width           =   1905
   End
   Begin VB.CommandButton Command7 
      Caption         =   "KELUAR"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2940
      TabIndex        =   1
      Top             =   630
      Width           =   2475
   End
   Begin VB.CommandButton TmbSave 
      Caption         =   "CETAK"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   150
      TabIndex        =   0
      Top             =   630
      Width           =   2475
   End
   Begin Crystal.CrystalReport crpt 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "S003A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command7_Click()
Unload Me
End Sub

Private Sub Form_Load()
Text1 = NoUrut
End Sub

Private Sub TmbSave_Click()
crpt.ReportFileName = App.Path + "\ReportD\NOTA.rpt"
crpt.SelectionFormula = "{S003.NO_TRANS}= '" + Trim(Text1) + "'"
crpt.WindowState = crptMaximized
crpt.WindowMaxButton = True
crpt.WindowMinButton = True
crpt.Action = 1
End Sub
