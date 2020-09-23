VERSION 5.00
Begin VB.Form Form14 
   BackColor       =   &H00808000&
   Caption         =   "Credits"
   ClientHeight    =   5520
   ClientLeft      =   4470
   ClientTop       =   3015
   ClientWidth     =   10680
   Icon            =   "Form14.frx":0000
   LinkTopic       =   "Form14"
   ScaleHeight     =   5520
   ScaleWidth      =   10680
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8640
      TabIndex        =   9
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   800
      Left            =   7920
      Top             =   4440
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   800
      Left            =   7920
      Top             =   3720
   End
   Begin VB.Timer Timer1 
      Interval        =   8000
      Left            =   7920
      Top             =   3120
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   5280
      Width           =   10695
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "If Not For You We Would Never Be Here!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   1320
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Thank You For Your Support And For The Gift of Enlightenment!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   7815
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "6.Bohol Northeastern Colleges"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3840
      TabIndex        =   6
      Top             =   4440
      Width           =   3855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "5.Mr. Alberto G. Arbasto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   5
      Top             =   3840
      Width           =   3135
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "4.Mr. Philip Naparan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   4
      Top             =   3120
      Width           =   2775
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "1.Our Parents"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "3.Ms. Jill B. Estrella"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "2.Mr. Hipolito A. Merto jr."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      Top             =   2040
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form14.frx":0442
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   10095
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0FFC0&
      Height          =   855
      Left            =   0
      TabIndex        =   13
      Top             =   4200
      Width           =   10695
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0FFC0&
      Height          =   135
      Left            =   0
      TabIndex        =   12
      Top             =   3600
      Width           =   10695
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0FFC0&
      Height          =   975
      Left            =   0
      TabIndex        =   11
      Top             =   1560
      Width           =   10695
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0FFC0&
      Height          =   855
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   10695
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
Form1.Show
End Sub

Private Sub Timer1_Timer()
Timer3.Enabled = False
Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label7.Visible = False
Timer2.Enabled = True
End Sub

Private Sub Timer2_Timer()
Timer1.Enabled = False
Label8.Visible = True
Label9.Visible = True
Timer3.Enabled = True

End Sub

Private Sub Timer3_Timer()
Timer2.Enabled = False
Timer1.Enabled = True

End Sub
