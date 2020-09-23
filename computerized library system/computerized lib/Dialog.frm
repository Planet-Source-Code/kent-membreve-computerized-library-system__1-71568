VERSION 5.00
Begin VB.Form Dialog 
   BackColor       =   &H00808000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Info"
   ClientHeight    =   3195
   ClientLeft      =   6570
   ClientTop       =   4545
   ClientWidth     =   6030
   Icon            =   "Dialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFC0&
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   2760
      Width           =   6015
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFC0&
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   1320
      Width           =   6015
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Height          =   615
      Left            =   0
      TabIndex        =   5
      Top             =   1920
      Width           =   6015
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "USER Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFC0&
      Height          =   975
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   6015
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
Text1.Text = (frmLogin.Adodc1.Recordset.Fields(1))
Text2.Text = (frmLogin.Adodc1.Recordset.Fields(0))
End Sub

Private Sub OKButton_Click()
frmLogin.Show
Unload Me
End Sub
