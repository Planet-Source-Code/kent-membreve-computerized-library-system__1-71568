VERSION 5.00
Begin VB.Form about1 
   BackColor       =   &H00808000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About"
   ClientHeight    =   5700
   ClientLeft      =   5535
   ClientTop       =   3660
   ClientWidth     =   8535
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Show Pics"
      Height          =   495
      Left            =   7080
      TabIndex        =   11
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Credits"
      Height          =   615
      Left            =   4680
      TabIndex        =   5
      Top             =   4320
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   615
      Left            =   1560
      TabIndex        =   1
      Top             =   4320
      Width           =   2415
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFC0&
      Height          =   135
      Left            =   0
      TabIndex        =   10
      Top             =   720
      Width           =   8775
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFC0&
      Height          =   735
      Left            =   0
      TabIndex        =   9
      Top             =   4200
      Width           =   8775
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "Version 1.0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   5040
      Width           =   8775
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFC0&
      Height          =   135
      Left            =   0
      TabIndex        =   7
      Top             =   5520
      Width           =   8775
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "All Rights Reserve: STI Collegre Tagbilaran"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   2760
      Width           =   4335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00800000&
      Caption         =   "Email : kentoymem2141616@yahoo.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   1680
      Width           =   4815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00800000&
      Caption         =   "Copyright : 2008"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackColor       =   &H00800000&
      Caption         =   "Creator : Kent A. Membreve And Angel Bondal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   960
      Width           =   4695
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   240
      Picture         =   "about1.frx":0000
      Top             =   960
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "Computerized Library System Beta Version 1.5"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8535
   End
End
Attribute VB_Name = "about1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Show
Unload Me
End Sub

Private Sub Command2_Click()
Form14.Show
Unload Me
End Sub

Private Sub Command3_Click()
Dialog5.Show
Unload Me
End Sub

