VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3975
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   10050
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   Icon            =   "Ô.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   10050
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2880
      Top             =   2760
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2400
      Top             =   2760
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1800
      Top             =   2760
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   1200
      Top             =   2760
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   10095
      TabIndex        =   0
      Top             =   3240
      Width           =   10095
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Created by:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Kent A. Membreve And Angel Bondal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   255
         Left            =   1320
         TabIndex        =   2
         Top             =   120
         Width           =   3375
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Copyrights 2008"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   8000
      Left            =   4800
      Top             =   2760
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C0FFC0&
      Height          =   615
      Left            =   6960
      TabIndex        =   15
      Top             =   0
      Width           =   3135
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   4080
      TabIndex        =   14
      Top             =   0
      Width           =   2895
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   4080
      TabIndex        =   13
      Top             =   1800
      Width           =   2895
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0FFC0&
      Height          =   615
      Left            =   6960
      TabIndex        =   12
      Top             =   1440
      Width           =   3135
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0FFC0&
      Height          =   615
      Left            =   0
      TabIndex        =   11
      Top             =   1440
      Width           =   4095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Height          =   375
      Left            =   6960
      TabIndex        =   10
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   840
      Width           =   4095
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Computerized Library System Version 1.0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   2160
      Width           =   3615
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0FFC0&
      Height          =   615
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   4095
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1815
      Left            =   6000
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1815
      Left            =   5040
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1815
      Left            =   4080
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Me.MousePointer = vbHourglass
frmLogin.Adodc1.Recordset.Filter = "USER_NAME ='" & frmLogin.txtUserName.Text & "'"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.MousePointer = vbDefault
End Sub

Private Sub Timer1_Timer()
Form1.Refresh
Unload Me
Form1.Show
Form1.Label16.Caption = (frmLogin.txtUserName.Text)
Form1.Label17.Caption = (frmLogin.Text3.Text)
End Sub

Private Sub Timer2_Timer()
Label5.Visible = True
Timer3.Enabled = True
Timer2.Enabled = False
End Sub

Private Sub Timer3_Timer()
Label6.Visible = True
Timer4.Enabled = True
Timer3.Enabled = False
End Sub

Private Sub Timer4_Timer()
Label8.Visible = True
Timer5.Enabled = True
Timer4.Enabled = False
End Sub

Private Sub Timer5_Timer()
Label9.Visible = Not Label9.Visible
End Sub
