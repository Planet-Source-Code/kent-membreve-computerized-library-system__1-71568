VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808000&
   Caption         =   "Computerized Library System"
   ClientHeight    =   8925
   ClientLeft      =   225
   ClientTop       =   525
   ClientWidth     =   14415
   FillColor       =   &H00FFFFFF&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8925
   ScaleWidth      =   14415
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame11 
      Caption         =   "Search"
      Height          =   2055
      Left            =   5640
      TabIndex        =   205
      Top             =   5640
      Visible         =   0   'False
      Width           =   4215
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1800
         TabIndex        =   210
         Text            =   "ENTER SEARCH HERE"
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton Command40 
         Caption         =   "CANCEL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   209
         ToolTipText     =   "Click To Hide This Menu"
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton Command41 
         Caption         =   "BY Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   208
         ToolTipText     =   "Click To Begin Searching By Title"
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton Command43 
         Caption         =   "By ID Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   207
         ToolTipText     =   "Click To Begin Search By ID Number"
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command44 
         Caption         =   "By Course"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   206
         ToolTipText     =   "Click to activate searching by ISBN"
         Top             =   960
         Width           =   1335
      End
   End
   Begin VB.TextBox Txt19 
      Height          =   405
      Left            =   0
      TabIndex        =   204
      Text            =   "Text19"
      Top             =   6000
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text11 
      Height          =   375
      Left            =   0
      TabIndex        =   198
      Text            =   "Text11"
      Top             =   5400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame13 
      Caption         =   "Search"
      Height          =   2055
      Left            =   7320
      TabIndex        =   191
      Top             =   5520
      Visible         =   0   'False
      Width           =   4695
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1800
         TabIndex        =   197
         Text            =   "ENTER SEARCH HERE"
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton Command50 
         Caption         =   "CANCEL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   196
         ToolTipText     =   "Click To Hide This Menu"
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton Command51 
         Caption         =   "By Title"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   195
         ToolTipText     =   "Click To Begin Searching By Title"
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton Command52 
         Caption         =   "By Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   194
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton Command53 
         Caption         =   "By ID Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   193
         ToolTipText     =   "Click To Begin Search By ID Number"
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton Command54 
         Caption         =   "By ISBN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   192
         ToolTipText     =   "Click to activate searching by ISBN"
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Close"
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
      Left            =   12480
      TabIndex        =   139
      Top             =   8280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8175
      Left            =   840
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   14420
      _Version        =   393216
      Tabs            =   8
      Tab             =   6
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   8454016
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Lost Books Record"
      TabPicture(0)   =   "Form1.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Command46"
      Tab(0).Control(1)=   "Command17"
      Tab(0).Control(2)=   "Picture3"
      Tab(0).Control(3)=   "Command18"
      Tab(0).Control(4)=   "Command13"
      Tab(0).Control(5)=   "Command2"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Damage Books Record"
      TabPicture(1)   =   "Form1.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command14"
      Tab(1).Control(1)=   "Command3"
      Tab(1).Control(2)=   "Command15"
      Tab(1).Control(3)=   "Picture4"
      Tab(1).Control(4)=   "Command1"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Book Records"
      TabPicture(2)   =   "Form1.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame12"
      Tab(2).Control(1)=   "Cmd20"
      Tab(2).Control(2)=   "Cmd1"
      Tab(2).Control(3)=   "Command19"
      Tab(2).Control(4)=   "Command12"
      Tab(2).Control(5)=   "Command10"
      Tab(2).Control(6)=   "Frame4"
      Tab(2).Control(7)=   "Picture5"
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "Print"
      TabPicture(3)   =   "Form1.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame6"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Returned Books"
      TabPicture(4)   =   "Form1.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Command49"
      Tab(4).Control(1)=   "Frame7"
      Tab(4).Control(2)=   "Picture8"
      Tab(4).Control(3)=   "Command28"
      Tab(4).Control(4)=   "Command22"
      Tab(4).ControlCount=   5
      TabCaption(5)   =   "Borrower's Records"
      TabPicture(5)   =   "Form1.frx":04CE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Command30"
      Tab(5).Control(1)=   "Frame1"
      Tab(5).Control(2)=   "Command16"
      Tab(5).Control(3)=   "Command8"
      Tab(5).Control(4)=   "Command6"
      Tab(5).Control(5)=   "Command5"
      Tab(5).Control(6)=   "Command4"
      Tab(5).Control(7)=   "Frame2"
      Tab(5).Control(8)=   "Picture1"
      Tab(5).Control(9)=   "Label61"
      Tab(5).ControlCount=   10
      TabCaption(6)   =   "Borrowed Books"
      TabPicture(6)   =   "Form1.frx":04EA
      Tab(6).ControlEnabled=   -1  'True
      Tab(6).Control(0)=   "Picture6"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "Frame3"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "Command23"
      Tab(6).Control(2).Enabled=   0   'False
      Tab(6).Control(3)=   "Command24"
      Tab(6).Control(3).Enabled=   0   'False
      Tab(6).Control(4)=   "Command25"
      Tab(6).Control(4).Enabled=   0   'False
      Tab(6).Control(5)=   "Command26"
      Tab(6).Control(5).Enabled=   0   'False
      Tab(6).Control(6)=   "Command27"
      Tab(6).Control(6).Enabled=   0   'False
      Tab(6).Control(7)=   "Frame5"
      Tab(6).Control(7).Enabled=   0   'False
      Tab(6).Control(8)=   "Frame10"
      Tab(6).Control(8).Enabled=   0   'False
      Tab(6).ControlCount=   9
      TabCaption(7)   =   "Blacklisted Borrowers"
      TabPicture(7)   =   "Form1.frx":0506
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Command55"
      Tab(7).Control(1)=   "Cmdchanstat"
      Tab(7).Control(2)=   "Cmdrefresh"
      Tab(7).Control(3)=   "Picture2"
      Tab(7).Control(4)=   "Frame9"
      Tab(7).ControlCount=   5
      Begin VB.CommandButton Command55 
         Caption         =   "SEARCH"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -66960
         TabIndex        =   202
         Top             =   7320
         Width           =   1095
      End
      Begin VB.CommandButton Command49 
         Caption         =   "Pay Fines"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -64800
         TabIndex        =   190
         Top             =   7200
         Width           =   1335
      End
      Begin VB.CommandButton Command46 
         Caption         =   "Pay lost Book"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -68760
         TabIndex        =   189
         Top             =   7320
         Width           =   1695
      End
      Begin VB.Frame Frame12 
         Caption         =   "Search"
         Height          =   2055
         Left            =   -72840
         TabIndex        =   183
         Top             =   4920
         Visible         =   0   'False
         Width           =   4695
         Begin VB.CommandButton Command48 
            Caption         =   "By ISBN"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   188
            ToolTipText     =   "Click to activate searching by ISBN"
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton Command47 
            Caption         =   "By AUTHOR"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   187
            ToolTipText     =   "Click To Begin Search By ID Number"
            Top             =   960
            Width           =   1335
         End
         Begin VB.CommandButton Command45 
            Caption         =   "By Title"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   186
            ToolTipText     =   "Click To Begin Searching By Title"
            Top             =   600
            Width           =   1335
         End
         Begin VB.CommandButton Command42 
            Caption         =   "CANCEL"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1320
            TabIndex        =   185
            ToolTipText     =   "Click To Hide This Menu"
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   1800
            TabIndex        =   184
            Text            =   "Enter Search Here................."
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Search"
         Height          =   2055
         Left            =   7200
         TabIndex        =   174
         Top             =   5160
         Visible         =   0   'False
         Width           =   4695
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   1800
            TabIndex        =   181
            Text            =   "ENTER SEARCH HERE"
            Top             =   240
            Width           =   2175
         End
         Begin VB.CommandButton Command38 
            Caption         =   "CANCEL"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1320
            TabIndex        =   179
            ToolTipText     =   "Click To Hide This Menu"
            Top             =   1680
            Width           =   975
         End
         Begin VB.CommandButton Command37 
            Caption         =   "By Title"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   178
            ToolTipText     =   "Click To Begin Searching By Title"
            Top             =   600
            Width           =   1335
         End
         Begin VB.CommandButton Command36 
            Caption         =   "By Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   177
            Top             =   1320
            Width           =   1335
         End
         Begin VB.CommandButton Command35 
            Caption         =   "By ID Number"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   176
            ToolTipText     =   "Click To Begin Search By ID Number"
            Top             =   960
            Width           =   1335
         End
         Begin VB.CommandButton Command34 
            Caption         =   "By ISBN"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   175
            ToolTipText     =   "Click to activate searching by ISBN"
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Select "
         Height          =   2055
         Left            =   3960
         TabIndex        =   170
         Top             =   5400
         Visible         =   0   'False
         Width           =   2775
         Begin VB.CommandButton Command39 
            Caption         =   "CANCEL"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1560
            TabIndex        =   180
            ToolTipText     =   "Click To Hide This Menu"
            Top             =   1680
            Width           =   1095
         End
         Begin VB.CommandButton Command33 
            Caption         =   "Damage Book"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   173
            ToolTipText     =   "Select this to add a book on the damage books record"
            Top             =   1200
            Width           =   2295
         End
         Begin VB.CommandButton Command32 
            Caption         =   "Lost Book"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   172
            ToolTipText     =   "Select to add a borrowed book to lost books records"
            Top             =   720
            Width           =   2295
         End
         Begin VB.CommandButton Command31 
            Caption         =   "Borrow Transaction"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   171
            ToolTipText     =   "Select to add new borrow transaction"
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.CommandButton Command30 
         Caption         =   "Refresh"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -64080
         TabIndex        =   169
         Top             =   7440
         Width           =   975
      End
      Begin VB.Frame Frame1 
         Caption         =   "Please Select"
         Height          =   1095
         Left            =   -67200
         TabIndex        =   165
         Top             =   6120
         Visible         =   0   'False
         Width           =   2175
         Begin VB.CommandButton Command11 
            Caption         =   "Blacklisted"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   168
            Top             =   600
            Width           =   1215
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Safe"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   167
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.CommandButton Cmdchanstat 
         Caption         =   "Change Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -65880
         TabIndex        =   141
         Top             =   7320
         Width           =   1455
      End
      Begin VB.CommandButton Cmdrefresh 
         Caption         =   "&Refresh"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -64440
         TabIndex        =   140
         Top             =   7320
         Width           =   1215
      End
      Begin VB.CommandButton Command17 
         Caption         =   "&Refresh"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -64920
         TabIndex        =   138
         Top             =   7320
         Width           =   1095
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Show Options"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -66720
         TabIndex        =   137
         Top             =   7680
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Search"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -68160
         TabIndex        =   136
         Top             =   7560
         Width           =   1300
      End
      Begin VB.CommandButton Cmd20 
         Caption         =   "&Edit"
         Height          =   225
         Left            =   -66240
         TabIndex        =   135
         Top             =   7200
         Width           =   1215
      End
      Begin VB.CommandButton Cmd1 
         Caption         =   "&Save"
         Height          =   225
         Left            =   -67200
         TabIndex        =   134
         Top             =   7200
         Width           =   975
      End
      Begin VB.CommandButton Command19 
         Caption         =   "&ADD"
         Height          =   225
         Left            =   -68280
         TabIndex        =   131
         Top             =   7200
         Width           =   1095
      End
      Begin VB.CommandButton Command12 
         Caption         =   "&Refresh"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -65040
         TabIndex        =   130
         Top             =   7200
         Width           =   1300
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00C0FFC0&
         Height          =   3495
         Left            =   -74760
         ScaleHeight     =   3435
         ScaleWidth      =   12195
         TabIndex        =   128
         Top             =   840
         Width           =   12255
         Begin MSAdodcLib.Adodc Adodc8 
            Height          =   375
            Left            =   10440
            Top             =   3120
            Visible         =   0   'False
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   8
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   ""
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   ""
            Caption         =   "Adodc6"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin MSDataGridLib.DataGrid DataGrid8 
            Height          =   3135
            Left            =   0
            TabIndex        =   129
            Top             =   0
            Width           =   12255
            _ExtentX        =   21616
            _ExtentY        =   5530
            _Version        =   393216
            BackColor       =   12648384
            HeadLines       =   1
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00C0FFC0&
         Height          =   2415
         Left            =   -74760
         ScaleHeight     =   2355
         ScaleWidth      =   12075
         TabIndex        =   126
         Top             =   720
         Width           =   12135
         Begin VB.TextBox Txtcount 
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            Height          =   250
            Left            =   1440
            TabIndex        =   133
            Text            =   "Text1"
            Top             =   2160
            Width           =   975
         End
         Begin MSAdodcLib.Adodc Adodc6 
            Height          =   375
            Left            =   10320
            Top             =   2040
            Visible         =   0   'False
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   8
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   ""
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   ""
            Caption         =   "Adodc6"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin MSDataGridLib.DataGrid DataGrid6 
            Height          =   2055
            Left            =   0
            TabIndex        =   127
            Top             =   0
            Width           =   12135
            _ExtentX        =   21405
            _ExtentY        =   3625
            _Version        =   393216
            BackColor       =   12648384
            HeadLines       =   1
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Number of Records:"
            Height          =   255
            Left            =   0
            TabIndex        =   132
            Top             =   2160
            Width           =   1575
         End
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&ADD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -69120
         TabIndex        =   125
         Top             =   7440
         Width           =   1095
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00C0FFC0&
         Height          =   3015
         Left            =   -74640
         ScaleHeight     =   2955
         ScaleWidth      =   12075
         TabIndex        =   118
         Top             =   4080
         Width           =   12135
         Begin MSAdodcLib.Adodc Adodc7 
            Height          =   375
            Left            =   10080
            Top             =   2640
            Visible         =   0   'False
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   8
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   ""
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   ""
            Caption         =   "Adodc5"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin MSDataGridLib.DataGrid DataGrid7 
            Height          =   2655
            Left            =   0
            TabIndex        =   119
            Top             =   0
            Width           =   12075
            _ExtentX        =   21299
            _ExtentY        =   4683
            _Version        =   393216
            BackColor       =   12648384
            HeadLines       =   1
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   4
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Number of Records:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   0
            TabIndex        =   121
            Top             =   2640
            Width           =   1575
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "_"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1440
            TabIndex        =   120
            Top             =   2640
            Width           =   2175
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Blacklisted Barrowers"
         Height          =   3015
         Left            =   -74760
         TabIndex        =   117
         Top             =   960
         Width           =   12375
         Begin VB.TextBox Txtstatus 
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   7080
            TabIndex        =   154
            Text            =   "Text1"
            Top             =   1200
            Width           =   1695
         End
         Begin VB.TextBox Txtborid 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   147
            Text            =   "Text1"
            Top             =   480
            Width           =   1860
         End
         Begin VB.TextBox Txtbornam 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   146
            Text            =   "Text1"
            Top             =   840
            Width           =   1860
         End
         Begin VB.TextBox Txtcors 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   145
            Text            =   "Text1"
            Top             =   1200
            Width           =   1740
         End
         Begin VB.TextBox Txtcy 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   144
            Text            =   "Text1"
            Top             =   1560
            Width           =   1740
         End
         Begin VB.TextBox Txtader 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   6720
            TabIndex        =   143
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox Txtcontact 
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   6720
            TabIndex        =   142
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Borrower Status"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5640
            TabIndex        =   155
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label Label30 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Borrower's ID : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   153
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label Lbl18 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Course : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   152
            Top             =   1200
            Width           =   1935
         End
         Begin VB.Label Lbl17 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Current Year : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   151
            Top             =   1560
            Width           =   1935
         End
         Begin VB.Label Lbl16 
            BackStyle       =   0  'Transparent
            Caption         =   "Borrower's Name:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   150
            Top             =   840
            Width           =   2655
         End
         Begin VB.Label Lbl15 
            BackStyle       =   0  'Transparent
            Caption         =   "Address:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5640
            TabIndex        =   149
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Lbl14 
            BackStyle       =   0  'Transparent
            Caption         =   "Contact#:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5640
            TabIndex        =   148
            Top             =   840
            Width           =   975
         End
      End
      Begin VB.CommandButton Command18 
         Caption         =   "&Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -66000
         TabIndex        =   116
         Top             =   7320
         Width           =   1095
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Refresh"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -65160
         TabIndex        =   111
         Top             =   7560
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Paid"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -65880
         TabIndex        =   110
         Top             =   7560
         Width           =   735
      End
      Begin VB.CommandButton Command14 
         Caption         =   "&ADD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -66840
         TabIndex        =   109
         Top             =   7560
         Width           =   975
      End
      Begin VB.CommandButton Command13 
         Caption         =   "&ADD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -69840
         TabIndex        =   108
         Top             =   7320
         Width           =   1095
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70200
         TabIndex        =   92
         Top             =   7440
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Save"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -68040
         TabIndex        =   94
         Top             =   7440
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -65160
         TabIndex        =   104
         Top             =   7440
         Width           =   1095
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Barrower Information"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   -74880
         TabIndex        =   96
         Top             =   1080
         Width           =   12495
         Begin VB.TextBox Cnum2 
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   7320
            TabIndex        =   113
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox ader2 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   7320
            TabIndex        =   112
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox Text34 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   100
            Text            =   "Text1"
            Top             =   1200
            Width           =   1740
         End
         Begin VB.TextBox Text23 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   99
            Text            =   "Text1"
            Top             =   960
            Width           =   1740
         End
         Begin VB.TextBox Text22 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   98
            Text            =   "Text1"
            Top             =   600
            Width           =   1860
         End
         Begin VB.TextBox Text21 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   97
            Text            =   "Text1"
            Top             =   240
            Width           =   1860
         End
         Begin VB.Label Label68 
            BackStyle       =   0  'Transparent
            Caption         =   "Contact#:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6240
            TabIndex        =   115
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label67 
            BackStyle       =   0  'Transparent
            Caption         =   "Address:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6240
            TabIndex        =   114
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Lbl30 
            BackStyle       =   0  'Transparent
            Caption         =   "Borrower's Name:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   107
            Top             =   600
            Width           =   2655
         End
         Begin VB.Label Label52 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Current Year : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   103
            Top             =   1200
            Width           =   1935
         End
         Begin VB.Label Label51 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Course : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   102
            Top             =   960
            Width           =   1935
         End
         Begin VB.Label Label26 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Borrower's ID : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   101
            Top             =   240
            Width           =   1935
         End
         Begin VB.Image Image2 
            Height          =   1080
            Left            =   10320
            Picture         =   "Form1.frx":0522
            Stretch         =   -1  'True
            Top             =   360
            Width           =   1080
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0FFC0&
         Height          =   4215
         Left            =   -74880
         ScaleHeight     =   4155
         ScaleWidth      =   12435
         TabIndex        =   93
         Top             =   2760
         Width           =   12495
         Begin MSAdodcLib.Adodc Adodc1 
            Height          =   375
            Left            =   10560
            Top             =   3840
            Visible         =   0   'False
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   8
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   ""
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   ""
            Caption         =   "Adodc1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   3855
            Left            =   0
            TabIndex        =   95
            Top             =   0
            Width           =   12495
            _ExtentX        =   22040
            _ExtentY        =   6800
            _Version        =   393216
            BackColor       =   12648384
            HeadLines       =   1
            RowHeight       =   15
            AllowAddNew     =   -1  'True
            AllowDelete     =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Label Label62 
            BackStyle       =   0  'Transparent
            Caption         =   "_"
            Height          =   255
            Left            =   1440
            TabIndex        =   106
            Top             =   3840
            Width           =   615
         End
         Begin VB.Label Label53 
            BackStyle       =   0  'Transparent
            Caption         =   "Number of Records:"
            Height          =   255
            Left            =   0
            TabIndex        =   105
            Top             =   3840
            Width           =   1575
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Found"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -67080
         TabIndex        =   91
         Top             =   7320
         Width           =   1095
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -69600
         TabIndex        =   75
         Top             =   7200
         Width           =   1300
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Book Information"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   -74880
         TabIndex        =   62
         Top             =   960
         Width           =   12495
         Begin VB.TextBox Text10 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   2400
            TabIndex        =   81
            Text            =   "Text1"
            Top             =   1560
            Width           =   2100
         End
         Begin VB.TextBox Text9 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   2400
            TabIndex        =   80
            Text            =   "Text1"
            Top             =   1320
            Width           =   2100
         End
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   2400
            TabIndex        =   79
            Text            =   "Text1"
            Top             =   1080
            Width           =   2625
         End
         Begin VB.TextBox Text7 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   2400
            TabIndex        =   78
            Text            =   "Text1"
            Top             =   840
            Width           =   4305
         End
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   2400
            TabIndex        =   77
            Text            =   "Text1"
            Top             =   600
            Width           =   2100
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   2400
            TabIndex        =   76
            Text            =   "Text1"
            Top             =   360
            Width           =   2100
         End
         Begin VB.TextBox Text37 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   5520
            TabIndex        =   67
            Text            =   "Text1"
            Top             =   360
            Width           =   2580
         End
         Begin VB.Frame Frame8 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Book Details"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   7080
            TabIndex        =   63
            Top             =   720
            Width           =   3255
            Begin VB.TextBox Txtrem 
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1200
               TabIndex        =   158
               Text            =   "Text3"
               Top             =   720
               Width           =   1335
            End
            Begin VB.TextBox Txtbor 
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1200
               TabIndex        =   157
               Text            =   "Text2"
               Top             =   480
               Width           =   1335
            End
            Begin VB.TextBox Txtquan 
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1200
               TabIndex        =   156
               Text            =   "Text1"
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label Label50 
               BackStyle       =   0  'Transparent
               Caption         =   "Quantity :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   66
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label31 
               BackStyle       =   0  'Transparent
               Caption         =   "Borrowed :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   65
               Top             =   480
               Width           =   1095
            End
            Begin VB.Label Label19 
               BackStyle       =   0  'Transparent
               Caption         =   "Remaining :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   64
               Top             =   720
               Width           =   975
            End
         End
         Begin VB.Label Label60 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Category : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   74
            Top             =   1320
            Width           =   1935
         End
         Begin VB.Label Label59 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Year Published : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   73
            Top             =   1560
            Width           =   1935
         End
         Begin VB.Label Label58 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Price : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4920
            TabIndex        =   72
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label57 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Book NO : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   71
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label Label56 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "ISBN : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   70
            Top             =   600
            Width           =   1935
         End
         Begin VB.Label Label55 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Book Title : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   69
            Top             =   840
            Width           =   1935
         End
         Begin VB.Label Label54 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Author : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   68
            Top             =   1080
            Width           =   1935
         End
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00C0FFC0&
         Height          =   3855
         Left            =   -74880
         ScaleHeight     =   3795
         ScaleWidth      =   12435
         TabIndex        =   57
         Top             =   3000
         Width           =   12495
         Begin MSAdodcLib.Adodc Adodc2 
            Height          =   345
            Left            =   10080
            Top             =   3480
            Visible         =   0   'False
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   609
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   8
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   ""
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   ""
            Caption         =   "Adodc1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin MSDataGridLib.DataGrid DataGrid2 
            Height          =   3495
            Left            =   0
            TabIndex        =   58
            Top             =   0
            Width           =   12435
            _ExtentX        =   21934
            _ExtentY        =   6165
            _Version        =   393216
            BackColor       =   12648384
            HeadLines       =   1
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   4
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Label Label29 
            BackStyle       =   0  'Transparent
            Caption         =   "_"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1560
            TabIndex        =   61
            Top             =   3480
            Width           =   1695
         End
         Begin VB.Label Label27 
            BackStyle       =   0  'Transparent
            Caption         =   "  Number of Records:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   0
            TabIndex        =   60
            Top             =   3480
            Width           =   1575
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   " Record:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   15
            Index           =   4
            Left            =   0
            TabIndex        =   59
            Top             =   3795
            Width           =   735
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Book Information"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   -74760
         TabIndex        =   38
         Top             =   960
         Width           =   12135
         Begin VB.TextBox Text25 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   9480
            Locked          =   -1  'True
            TabIndex        =   84
            Text            =   "Text1"
            Top             =   1080
            Width           =   1860
         End
         Begin VB.TextBox Text24 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   9480
            Locked          =   -1  'True
            TabIndex        =   83
            Text            =   "Text1"
            Top             =   1440
            Width           =   1860
         End
         Begin VB.TextBox Text26 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   9480
            Locked          =   -1  'True
            TabIndex        =   46
            Text            =   "Text1"
            Top             =   600
            Width           =   1860
         End
         Begin VB.TextBox Text27 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   9480
            Locked          =   -1  'True
            TabIndex        =   45
            Text            =   "Text1"
            Top             =   360
            Width           =   1860
         End
         Begin VB.TextBox Text28 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   44
            Text            =   "Text1"
            Top             =   1440
            Width           =   2100
         End
         Begin VB.TextBox Text29 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   43
            Text            =   "Text1"
            Top             =   1200
            Width           =   2100
         End
         Begin VB.TextBox Text30 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   42
            Text            =   "Text1"
            Top             =   960
            Width           =   2625
         End
         Begin VB.TextBox Text31 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   41
            Text            =   "Text1"
            Top             =   720
            Width           =   4305
         End
         Begin VB.TextBox Text32 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   40
            Text            =   "Text1"
            Top             =   480
            Width           =   2100
         End
         Begin VB.TextBox Text33 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   39
            Text            =   "Text1"
            Top             =   240
            Width           =   2100
         End
         Begin VB.Label Label38 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Fines : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7320
            TabIndex        =   56
            Top             =   1440
            Width           =   1935
         End
         Begin VB.Label Label25 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Date Returned : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7320
            TabIndex        =   55
            Top             =   600
            Width           =   1935
         End
         Begin VB.Label Label41 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "No of Days After Due : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   7320
            TabIndex        =   54
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label42 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Barrower's Name : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   1200
            Width           =   1935
         End
         Begin VB.Label Label43 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Date Barrowed : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   1440
            Width           =   1935
         End
         Begin VB.Label Label44 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Date Due : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7320
            TabIndex        =   51
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label45 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Book NO : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label46 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "ISBN : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label Label47 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Book Title : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   720
            Width           =   1935
         End
         Begin VB.Label Label48 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Barrower's ID : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   960
            Width           =   1935
         End
      End
      Begin VB.PictureBox Picture8 
         BackColor       =   &H00C0FFC0&
         Height          =   2775
         Left            =   -74760
         ScaleHeight     =   2715
         ScaleWidth      =   12075
         TabIndex        =   36
         Top             =   3120
         Width           =   12135
         Begin MSDataGridLib.DataGrid DataGrid5 
            Height          =   2415
            Left            =   0
            TabIndex        =   123
            Top             =   0
            Width           =   12135
            _ExtentX        =   21405
            _ExtentY        =   4260
            _Version        =   393216
            BackColor       =   12648384
            HeadLines       =   1
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin MSAdodcLib.Adodc Adodc5 
            Height          =   330
            Left            =   9960
            Top             =   2400
            Visible         =   0   'False
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   582
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   8
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   ""
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   ""
            Caption         =   "Adodc5"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin VB.Label Label28 
            BackStyle       =   0  'Transparent
            Caption         =   "_"
            Height          =   255
            Left            =   1560
            TabIndex        =   124
            Top             =   2400
            Width           =   855
         End
         Begin VB.Label Label49 
            BackStyle       =   0  'Transparent
            Caption         =   "Number of  Records:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   37
            Top             =   2400
            Width           =   1935
         End
      End
      Begin VB.CommandButton Command28 
         Caption         =   "&Refresh"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   -66240
         TabIndex        =   35
         Top             =   7200
         Width           =   1300
      End
      Begin VB.CommandButton Command22 
         Caption         =   "&Search"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   -67680
         TabIndex        =   34
         Top             =   7200
         Width           =   1300
      End
      Begin VB.CommandButton Command27 
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5520
         TabIndex        =   32
         Top             =   7440
         Width           =   1305
      End
      Begin VB.CommandButton Command26 
         Caption         =   "&Search"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   8160
         TabIndex        =   31
         Top             =   7440
         Width           =   1300
      End
      Begin VB.CommandButton Command25 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   9480
         TabIndex        =   30
         Top             =   7440
         Width           =   1300
      End
      Begin VB.CommandButton Command24 
         Caption         =   "&Refresh"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   10800
         TabIndex        =   29
         Top             =   7440
         Width           =   1185
      End
      Begin VB.CommandButton Command23 
         Caption         =   "&Return"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6840
         TabIndex        =   28
         Top             =   7440
         Width           =   1305
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Book Information"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   3015
         Left            =   360
         TabIndex        =   20
         Top             =   840
         Width           =   12255
         Begin VB.TextBox Text15 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   2760
            Locked          =   -1  'True
            TabIndex        =   90
            Text            =   "Text1"
            Top             =   1560
            Width           =   2625
         End
         Begin VB.TextBox Text18 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   7080
            Locked          =   -1  'True
            TabIndex        =   89
            Text            =   "Text1"
            Top             =   360
            Width           =   1980
         End
         Begin VB.TextBox Text12 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   2760
            Locked          =   -1  'True
            TabIndex        =   88
            Text            =   "Text1"
            Top             =   360
            Width           =   2100
         End
         Begin VB.TextBox Text13 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   2760
            Locked          =   -1  'True
            TabIndex        =   87
            Text            =   "Text1"
            Top             =   720
            Width           =   2100
         End
         Begin VB.TextBox Text14 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   2760
            Locked          =   -1  'True
            TabIndex        =   86
            Text            =   "Text1"
            Top             =   1200
            Width           =   4425
         End
         Begin VB.TextBox Text16 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   2760
            Locked          =   -1  'True
            TabIndex        =   85
            Text            =   "Text1"
            Top             =   2040
            Width           =   2100
         End
         Begin VB.TextBox Text17 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   2760
            Locked          =   -1  'True
            TabIndex        =   21
            Text            =   "Text1"
            Top             =   2400
            Width           =   2100
         End
         Begin VB.Label Label2 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "ISBN : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   82
            Top             =   720
            Width           =   2295
         End
         Begin VB.Label Label40 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Barrower's ID : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   27
            Top             =   1560
            Width           =   2295
         End
         Begin VB.Label Label39 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Book Title : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   26
            Top             =   1200
            Width           =   2295
         End
         Begin VB.Label Label34 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Book NO : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   25
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label Label35 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Date Due : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6120
            TabIndex        =   24
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label36 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Date Barrowed : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   23
            Top             =   2400
            Width           =   2295
         End
         Begin VB.Label Label37 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Barrower's Name : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   22
            Top             =   2040
            Width           =   2295
         End
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H00C0FFC0&
         Height          =   2535
         Left            =   240
         ScaleHeight     =   2475
         ScaleWidth      =   12195
         TabIndex        =   16
         Top             =   4440
         Width           =   12255
         Begin MSAdodcLib.Adodc Adodc3 
            Height          =   375
            Left            =   10080
            Top             =   2160
            Visible         =   0   'False
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   8
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   ""
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   ""
            Caption         =   "Adodc3"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin MSDataGridLib.DataGrid DataGrid3 
            Height          =   2175
            Left            =   0
            TabIndex        =   17
            Top             =   0
            Width           =   12195
            _ExtentX        =   21511
            _ExtentY        =   3836
            _Version        =   393216
            BackColor       =   12648384
            HeadLines       =   1
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   4
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Label Label33 
            BackStyle       =   0  'Transparent
            Caption         =   "Number of  Records:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   0
            TabIndex        =   19
            Top             =   2160
            Width           =   1575
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "_"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   1560
            TabIndex        =   18
            Top             =   2160
            Width           =   1815
         End
      End
      Begin VB.Frame Frame6 
         Height          =   3615
         Left            =   -70680
         TabIndex        =   6
         Top             =   2040
         Width           =   4215
         Begin VB.CommandButton Command9 
            Caption         =   "&Print Detail"
            Height          =   375
            Left            =   600
            TabIndex        =   7
            Top             =   1920
            Width           =   3015
         End
         Begin VB.Label Label32 
            Caption         =   "PRINT INFORMATION"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   8
            Top             =   360
            Width           =   2775
         End
         Begin VB.Image Image1 
            Height          =   720
            Left            =   3000
            Picture         =   "Form1.frx":0AAC
            Stretch         =   -1  'True
            Top             =   240
            Width           =   840
         End
      End
      Begin VB.Label Label61 
         Caption         =   "Change Status v"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -66600
         TabIndex        =   182
         ToolTipText     =   "Click Show Options to Change Borrower Status"
         Top             =   7440
         Width           =   1455
      End
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   300
      Left            =   1080
      TabIndex        =   9
      Top             =   600
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   529
      _Version        =   393216
      Format          =   76939265
      CurrentDate     =   37798
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   0
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   39
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1036
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1730
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":200A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2704
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2FDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":38B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4192
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4A6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5346
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5C20
            Key             =   "btn11"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":64FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6DD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":736E
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7908
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":81E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9396
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9A90
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9BEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A184
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":AA5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":B158
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":B6F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":BC8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C566
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":CE40
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D3DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D974
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":E24E
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":EB28
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":F402
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":FCDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":105B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":10E90
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1176A
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":11D04
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":11E5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":123F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":12552
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Cmddet 
      Caption         =   "&Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   160
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton Command29 
      Caption         =   "Return"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   163
      Top             =   7080
      Width           =   975
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00808080&
      Height          =   6135
      Left            =   840
      ScaleHeight     =   6075
      ScaleWidth      =   11955
      TabIndex        =   10
      Top             =   960
      Width           =   12015
      Begin VB.CommandButton Command21 
         Caption         =   "Command21"
         Height          =   195
         Left            =   0
         TabIndex        =   159
         Top             =   6120
         Width           =   975
      End
      Begin MSAdodcLib.Adodc Adodc4 
         Height          =   345
         Left            =   600
         Top             =   5760
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   609
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid DataGrid4 
         Height          =   5775
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   11955
         _ExtentX        =   21087
         _ExtentY        =   10186
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   4
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "_"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3120
         TabIndex        =   14
         Top             =   5835
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "  of"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2760
         TabIndex        =   13
         Top             =   5835
         Width           =   255
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   " Record:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   12
         Top             =   5835
         Width           =   735
      End
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "By Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   200
      Top             =   600
      Width           =   1335
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "ALL DUE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   201
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label20 
      BackColor       =   &H00C0FFC0&
      Height          =   495
      Left            =   14400
      TabIndex        =   203
      Top             =   0
      Width           =   4575
   End
   Begin VB.Label Label63 
      BackStyle       =   0  'Transparent
      Caption         =   "Filter Opitons"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   199
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label18 
      Caption         =   "Label18"
      Height          =   495
      Left            =   0
      TabIndex        =   166
      Top             =   4440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label17 
      Caption         =   "Label17"
      Height          =   495
      Left            =   0
      TabIndex        =   164
      Top             =   3600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   """"""
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12960
      TabIndex        =   162
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Hello!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12360
      TabIndex        =   161
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFC0&
      Height          =   735
      Left            =   0
      TabIndex        =   33
      Top             =   7440
      Width           =   18975
   End
   Begin VB.Label Label5 
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Select a Date : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "WELCOME TO THE COMPUTERIZED LIBRARY SYSTEM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   14415
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   18975
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFC0&
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   1800
      Width           =   18975
   End
   Begin VB.Label Label8 
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
      TabIndex        =   1
      Top             =   4920
      Width           =   18975
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFC0&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   8400
      Width           =   18975
   End
   Begin VB.Label Label71 
      BackColor       =   &H00C0FFC0&
      Height          =   135
      Left            =   0
      TabIndex        =   122
      Top             =   7080
      Width           =   18975
   End
   Begin VB.Menu mnu 
      Caption         =   "&File"
      Begin VB.Menu adbok 
         Caption         =   "ADD Book"
      End
      Begin VB.Menu adbor 
         Caption         =   "ADD Borrower"
      End
      Begin VB.Menu newbor 
         Caption         =   "New Borrow Transaction"
      End
      Begin VB.Menu adbl 
         Caption         =   "ADD Blacklisted Borrower"
      End
      Begin VB.Menu userpro 
         Caption         =   "User Promotions"
      End
      Begin VB.Menu signout 
         Caption         =   "Sign Out"
      End
      Begin VB.Menu print 
         Caption         =   "Print"
         Shortcut        =   {F9}
      End
      Begin VB.Menu ext 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuview 
      Caption         =   "&View"
      Begin VB.Menu bokdet 
         Caption         =   "Book Records"
      End
      Begin VB.Menu bordet 
         Caption         =   "Borrowers Records"
      End
      Begin VB.Menu borrregs 
         Caption         =   "Damage Books"
         Shortcut        =   {F2}
      End
      Begin VB.Menu Borboks 
         Caption         =   "Borrowed Books"
         Shortcut        =   {F6}
      End
      Begin VB.Menu libtran 
         Caption         =   "Returned Books"
         Shortcut        =   {F7}
      End
      Begin VB.Menu bookreg 
         Caption         =   "Lost Books"
         Shortcut        =   {F3}
      End
      Begin VB.Menu blaklis 
         Caption         =   "Blacklisted Borrowers"
      End
      Begin VB.Menu Dueboks 
         Caption         =   "Due Books"
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu settings 
      Caption         =   "&Settings"
      Begin VB.Menu addcat 
         Caption         =   "ADD/Edit Category"
         Shortcut        =   ^E
      End
      Begin VB.Menu Accset 
         Caption         =   "Account Settings"
      End
      Begin VB.Menu chngfine 
         Caption         =   "Change Fine"
         Shortcut        =   ^F
      End
      Begin VB.Menu themes 
         Caption         =   "Themes"
         Begin VB.Menu pnp 
            Caption         =   "Ponkan na Ponkan"
         End
         Begin VB.Menu redapple 
            Caption         =   "Red Apple"
         End
         Begin VB.Menu dalandan 
            Caption         =   "Dalandan"
         End
         Begin VB.Menu pandan 
            Caption         =   "Pandan"
         End
         Begin VB.Menu banana 
            Caption         =   "Banana"
         End
         Begin VB.Menu choco 
            Caption         =   "Chocolate"
         End
         Begin VB.Menu Emo 
            Caption         =   "Emotional"
         End
      End
   End
   Begin VB.Menu help 
      Caption         =   "Help"
      Begin VB.Menu abt 
         Caption         =   "About"
         Shortcut        =   ^A
      End
      Begin VB.Menu crid 
         Caption         =   "Credits"
         Shortcut        =   ^C
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub abt_Click()
about1.Show
Form1.Hide
End Sub

Private Sub Accset_Click()
Dialog7.Show
End Sub

Private Sub adbl_Click()
SSTab1.Tab = 5
SSTab1.Visible = True
Command20.Visible = True
Command16.SetFocus
End Sub

Private Sub adbok_Click()
Form3.Show vbModal
End Sub

Private Sub adbor_Click()
Form2.Show vbModal
End Sub

Private Sub addcat_Click()
Dialog4.Show
End Sub

Private Sub Adodc1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
If Val(Adodc1.Recordset.RecordCount) <= 0 Then
    Adodc1.Caption = "0"
    Label3.Caption = "0"
    Text21.Text = ""
    Text22.Text = ""
    Text23.Text = ""
    Text24.Text = ""
    ader2.Text = ""
    Cnum2.Text = ""
Else
    Adodc1.Caption = (Adodc1.Recordset.AbsolutePosition)
    Label62.Caption = (Adodc1.Recordset.RecordCount)
    Text21.Text = (Adodc1.Recordset.Fields(0))
    Text22.Text = (Adodc1.Recordset.Fields(1))
    Text23.Text = (Adodc1.Recordset.Fields(2))
    Text34.Text = (Adodc1.Recordset.Fields(3))
    ader2.Text = (Adodc1.Recordset.Fields(4))
    Cnum2.Text = (Adodc1.Recordset.Fields(5))
End If
If Adodc1.Recordset.Fields(4) = "" Then
Adodc1.Recordset.Fields(4) = "N/A"

ElseIf Adodc1.Recordset.Fields(5) = "" Then
Adodc1.Recordset.Fields(5) = "N/A"
End If

End Sub

Private Sub Adodc2_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
If Val(Adodc2.Recordset.RecordCount) <= 0 Then
    Adodc2.Caption = "0"
    Label29.Caption = "0"
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    Text37.Text = ""
    Txtquan.Text = ""
    Txtbor.Text = ""
    Txtrem.Text = ""
Else
    Adodc2.Caption = (Adodc2.Recordset.AbsolutePosition)
    Label29.Caption = (Adodc2.Recordset.RecordCount)
    Text5.Text = (Adodc2.Recordset.Fields(0))
    Text6.Text = (Adodc2.Recordset.Fields(1))
    Text7.Text = (Adodc2.Recordset.Fields(2))
    Text8.Text = (Adodc2.Recordset.Fields(3))
    Text9.Text = (Adodc2.Recordset.Fields(4))
    Text10.Text = (Adodc2.Recordset.Fields(5))
    Text37.Text = (Adodc2.Recordset.Fields(6))
    Txtquan.Text = (Adodc2.Recordset.Fields(7))
    Txtbor.Text = (Adodc2.Recordset.Fields(8))
    Txtrem.Text = (Adodc2.Recordset.Fields(9))
    If Adodc2.Recordset.Fields(9) = 0 Then
    Adodc2.Recordset.Fields(10) = "Unavailable"
    Adodc2.Recordset.Update
    Else
    Adodc2.Recordset.Fields(10) = "available"
    Adodc2.Recordset.Update
    End If
End If
End Sub

Private Sub Adodc3_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
If Val(Adodc3.Recordset.RecordCount) <= 0 Then
    Adodc3.Caption = "0"
    Label30.Caption = "0"
    Text12.Text = ""
    Text13.Text = ""
    Text14.Text = ""
    Text15.Text = ""
    Text16.Text = ""
    Text17.Text = ""
    Text18.Text = ""
    Text19.Text = ""
    Text20.Text = ""
    Text21.Text = ""
    Text22.Text = ""
Else
    Adodc3.Caption = (Adodc3.Recordset.AbsolutePosition)
    Label6.Caption = (Adodc3.Recordset.RecordCount)
    Text12.Text = (Adodc3.Recordset.Fields(0))
    Text13.Text = (Adodc3.Recordset.Fields(1))
    Text14.Text = (Adodc3.Recordset.Fields(2))
    Text15.Text = (Adodc3.Recordset.Fields(3))
    Text16.Text = (Adodc3.Recordset.Fields(4))
    Text17.Text = (Adodc3.Recordset.Fields(5))
    Text18.Text = (Adodc3.Recordset.Fields(6))
    Text19.Text = (Adodc3.Recordset.Fields(7))
    Text20.Text = (Adodc3.Recordset.Fields(8))
    Text21.Text = (Adodc3.Recordset.Fields(9))
    Text22.Text = (Adodc3.Recordset.Fields(10))
End If

End Sub


Private Sub Adodc4_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
If Val(Adodc4.Recordset.RecordCount) <= 0 Then
    Adodc4.Caption = "0"
    Label21.Caption = "0"
Else
    Adodc4.Caption = (Adodc4.Recordset.AbsolutePosition)
    Label21.Caption = (Adodc4.Recordset.RecordCount)
    Txt19.Text = (Adodc4.Recordset.Fields(6))
End If

End Sub

Private Sub Adodc5_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
If Val(Adodc5.Recordset.RecordCount) <= 0 Then
    Adodc5.Caption = "0"
    Label28.Caption = "0"
    Text23.Text = ""
    Text24.Text = ""
    Text25.Text = ""
    Text26.Text = ""
    Text27.Text = ""
    Text28.Text = ""
    Text29.Text = ""
    Text30.Text = ""
    Text31.Text = ""
    Text32.Text = ""
    Text33.Text = ""
Else
    Adodc5.Caption = (Adodc5.Recordset.AbsolutePosition)
    Label28.Caption = (Adodc5.Recordset.RecordCount)
    Text33.Text = (Adodc5.Recordset.Fields(0))
    Text32.Text = (Adodc5.Recordset.Fields(1))
    Text31.Text = (Adodc5.Recordset.Fields(2))
    Text30.Text = (Adodc5.Recordset.Fields(3))
    Text29.Text = (Adodc5.Recordset.Fields(4))
    Text28.Text = (Adodc5.Recordset.Fields(5))
    Text27.Text = (Adodc5.Recordset.Fields(6))
    Text26.Text = (Adodc5.Recordset.Fields(7))
    Text25.Text = (Adodc5.Recordset.Fields(8))
    Text24.Text = (Adodc5.Recordset.Fields(9))
    Text23.Text = (Adodc5.Recordset.Fields(10))
End If

End Sub



Private Sub Adodc6_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
If Val(Adodc6.Recordset.RecordCount) <= 0 Then
Txtcount.Text = (Adodc6.Recordset.RecordCount)
Else
Txtcount.Text = (Adodc6.Recordset.RecordCount)
End If

End Sub

Private Sub Adodc7_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
If Val(Adodc1.Recordset.RecordCount) <= 0 Then
    Label11.Caption = "0"
    Txtborid.Text = ""
    Txtbornam.Text = ""
    Txtcors.Text = ""
    Txtcy.Text = ""
    Txtader.Text = ""
    Txtcontact.Text = ""
    Txtstatus.Text = ""
Else
    Label11.Caption = (Adodc7.Recordset.RecordCount)
    Txtborid.Text = (Adodc7.Recordset.Fields(0))
    Txtbornam.Text = (Adodc7.Recordset.Fields(1))
    Txtcors.Text = (Adodc7.Recordset.Fields(2))
    Txtcy.Text = (Adodc7.Recordset.Fields(3))
    Txtader.Text = (Adodc7.Recordset.Fields(4))
    Txtcontact.Text = (Adodc7.Recordset.Fields(5))
    Txtstatus.Text = (Adodc7.Recordset.Fields(6))
End If
If Adodc1.Recordset.Fields(4) = "" Then
Adodc1.Recordset.Fields(4) = "N/A"

ElseIf Adodc1.Recordset.Fields(5) = "" Then
Adodc1.Recordset.Fields(5) = "N/A"
End If

End Sub

Private Sub banana_Click()
Label4.BackColor = &HC0FFFF
Label8.BackColor = &HC0FFFF
Label7.BackColor = &HC0FFFF
Label9.BackColor = &HC0FFFF
Label10.BackColor = &HC0FFFF
Label12.BackColor = &HC0FFFF
pnp.Checked = False
redapple.Checked = False
dalandan.Checked = False
pandan.Checked = False
banana.Checked = True
choco.Checked = False
Emo.Checked = False
End Sub

Private Sub blaklis_Click()
SSTab1.Tab = 7
SSTab1.Visible = True
Command20.Visible = True
End Sub

Private Sub bokdet_Click()
SSTab1.Visible = True
SSTab1.Tab = 2
Command20.Visible = True
End Sub

Private Sub bookreg_Click()
SSTab1.Tab = 0
SSTab1.Visible = True
Command20.Visible = True
End Sub

Private Sub Borboks_Click()
SSTab1.Tab = 6
SSTab1.Visible = True
Command20.Visible = True
End Sub

Private Sub borname_Change()
If borname = "" Or borname = "--------------------------------------" Then
Else
Command3.Enabled = True
End If
End Sub

Private Sub bordet_Click()
SSTab1.Visible = True
SSTab1.Tab = 5
Command20.Visible = True

End Sub

Private Sub borrregs_Click()
SSTab1.Tab = 1
SSTab1.Visible = True
Command20.Visible = True
End Sub

Private Sub chngfine_Click()
Dialog3.Show
End Sub

Private Sub choco_Click()
Label4.BackColor = &H80&
Label8.BackColor = &H80&
Label7.BackColor = &H80&
Label9.BackColor = &H80&
Label10.BackColor = &H80&
Label12.BackColor = &H80&
pnp.Checked = False
redapple.Checked = False
dalandan.Checked = False
pandan.Checked = False
banana.Checked = False
choco.Checked = True
Emo.Checked = False
End Sub

Private Sub Combo1_Click()
'If Val(Adodc6.Recordset.RecordCount) <= 0 Then
'    Adodc6.Caption = "0"
 '   Combo1.Text = ""
'Else
   ' Adodc6.Caption = (Adodc6.Recordset.AbsolutePosition)
    'Combo1.Text = (Adodc6.Recordset.Fields(0))
'End If
End Sub

Private Sub Cmdchanstat_Click()
Command5.Enabled = True
Dim str1, str2, str3 As String
str1 = InputBox("1 = Safe." & vbCrLf & "2 = Blacklisted", "Search Option")
If str1 = 1 Then
    Adodc7.Recordset.Fields(6) = "Safe"
    Adodc7.Recordset.Update
    Adodc7.Refresh
ElseIf str1 = 2 Then
    Adodc7.Recordset.Fields(6) = "Blacklisted"
    Adodc7.Recordset.Update
    Adodc7.Refresh
Else
MsgBox "That is not one of the choices"
    End If
End Sub

Private Sub Cmddet_Click()
SSTab1.Visible = True
SSTab1.Tab = 6
Adodc3.Recordset.Filter = "DATE_DUE ='" & (Txt19.Text) & "'"
Command20.Visible = True
Command27.SetFocus
End Sub

Private Sub Cmdrefresh_Click()
Adodc7.Refresh

End Sub

Private Sub Command1_Click()
Frame13.Visible = True
Text11.Text = "dam"
End Sub

Private Sub Cmd1_Click()
    Adodc2.Recordset.Fields(0) = Text5.Text
    Adodc2.Recordset.Fields(1) = Text6.Text
    Adodc2.Recordset.Fields(2) = Text7.Text
    Adodc2.Recordset.Fields(3) = Text8.Text
    Adodc2.Recordset.Fields(4) = Text9.Text
    Adodc2.Recordset.Fields(5) = Text10.Text
    Adodc2.Recordset.Fields(6) = Text37.Text
    Adodc2.Recordset.Fields(7) = Txtquan.Text
    Adodc2.Recordset.Fields(8) = Txtbor.Text
    Adodc2.Recordset.Fields(9) = Txtrem.Text
    Adodc2.Recordset.Update
Frame4.Enabled = False
Frame8.Enabled = False
Cmd1.Enabled = False
Cmd20.Enabled = True
End Sub

Private Sub Cmd20_Click()
Cmd1.Enabled = True
Cmd20.Enabled = False
Frame8.Enabled = True
Frame4.Enabled = True
Adodc2.RecordSource = Adodc2.Recordset.EditMode
End Sub

Private Sub Command10_Click()
Frame12.Visible = True
End Sub


Private Sub Command11_Click()
Adodc1.Recordset.Fields(6) = "Blacklisted"
Adodc1.Recordset.Update
Adodc1.Refresh
Adodc7.Refresh
Label18.Caption = "Blacklisted"
Command5.Enabled = True
End Sub

Private Sub Command13_Click()
Form7.Show vbModal
End Sub

Private Sub Command14_Click()
SSTab1.Tab = 6
End Sub

Private Sub Command15_Click()
Adodc8.Refresh
End Sub

Private Sub Command16_Click()
If Command16.Caption = "Show Options" Then
Frame1.Visible = True
Command16.Caption = "Hide Options"
ElseIf Command16.Caption = "Hide Options" Then
Frame1.Visible = False
Command16.Caption = "Show Options"
End If
End Sub

Private Sub Command17_Click()
Adodc6.Refresh
End Sub

Private Sub Command18_Click()
Frame13.Visible = True
Text11.Text = "los"
Text4.Text = "ENTER SEARCH HERE"
End Sub

Private Sub Command19_Click()
Form3.Show vbModal
End Sub

Private Sub Command12_Click()
Adodc2.Refresh
End Sub

Private Sub Command2_Click()
Form8.Show vbModal
End Sub

Private Sub Command20_Click()
SSTab1.Visible = False
Command20.Visible = False
End Sub

Private Sub Command21_Click()
Adodc2.Refresh
End Sub

Private Sub Command22_Click()
Frame13.Visible = True
Text4.Text = "ENTER SEARCH HERE"
Text11.Text = "ret"
End Sub

Private Sub Command23_Click()
Form5.Show vbModal
End Sub

Private Sub Command24_Click()
Adodc3.Refresh
End Sub

Private Sub Command25_Click()
On Error Resume Next
Dim repp2 As Integer
If Val(Adodc3.Recordset.RecordCount) <= 0 Then
    MsgBox "No more Records to be deleted.", vbInformation, "Confirm"
    Exit Sub
End If
repp2 = MsgBox("You are about to delete 1 record." & vbCrLf & "If you click YES, you won't be able to undo this delete operation." & vbCrLf & "Are you sure you want to delete these record?", vbCritical + vbYesNo, "Confirm Delete")
If repp2 = vbYes Then
    Adodc3.Recordset.Delete
    Adodc3.Recordset.MoveNext
    If Adodc3.Recordset.EOF Then
        Adodc3.Recordset.MoveLast
    End If
    Call ReturnBookQty
    MsgBox "Record has been successfuly deleted.", vbInformation, "Confirm"
End If

End Sub

Private Sub Command26_Click()
Frame10.Visible = True
End Sub

Private Sub Command27_Click()
Frame5.Visible = True
End Sub

Private Sub Command28_Click()
Adodc5.Refresh
End Sub

Private Sub Command29_Click()
SSTab1.Visible = True
SSTab1.Tab = 6
Adodc3.Recordset.Filter = "DATE_DUE ='" & (Txt19.Text) & "'"
Command20.Visible = True
Command23.SetFocus
End Sub

Private Sub Command3_Click()
If Adodc8.Recordset.EOF Then
MsgBox "No more Books To Be Paid"
Else
Form12.Label4.Caption = "frmdam"
Form12.Show vbModal
End If
End Sub

Private Sub Command30_Click()
Adodc1.Refresh
End Sub

Private Sub Command31_Click()
Form4.Show vbModal
Frame5.Visible = False
End Sub

Private Sub Command32_Click()
Form7.Show vbModal
Frame5.Visible = False
End Sub

Private Sub Command33_Click()
Form9.Show vbModal
Frame5.Visible = False
End Sub

Private Sub Command34_Click()
Adodc3.Recordset.MoveFirst
Adodc3.Recordset.Find "ISBN like '" & (Text1.Text) & "'"
If Adodc3.Recordset.EOF Then
    MsgBox "ISBN Match Not Found. Make Sure This ISBN Exist"
ElseIf Text1 = "" Then
MsgBox "Please Do not Put A Null Value While Searching"
ElseIf Adodc3.Recordset.Fields(1) = Text1.Text Then
    Adodc3.Recordset.Filter = "ISBN='" & (Text1.Text) & "'"
    Frame10.Visible = False
    Text1.Text = "ENTER SEARCH HERE"
    Command24.SetFocus
    Else
End If
End Sub

Private Sub Command35_Click()
Adodc3.Recordset.Find "BORROWERS_ID like '" & (Text1.Text) & "'"
If Adodc3.Recordset.EOF Then
MsgBox "ID Number Match Not Found. Make Sure This ID Number Exist"
ElseIf Text1.Text = "" Then
MsgBox "Please Do not Put A Null Value While Searching"
ElseIf Adodc3.Recordset.Fields(3) = Text1.Text Then
        Adodc3.Recordset.Filter = "BORROWERS_ID ='" & (Text1.Text) & "'"
        Frame10.Visible = False
        Text1.Text = "ENTER SEARCH HERE"
        Command24.SetFocus
        Else
             End If
End Sub

Private Sub Command36_Click()
Adodc3.Recordset.Find "BORROWERS_NAME like '" & (Text1.Text) & "'"
If Adodc3.Recordset.EOF Then
         MsgBox "That Name Does not Match any of the records. Make Sure This Name Exist"
ElseIf Text1.Text = "" Then
MsgBox "Please Do not Put A Null Value While Searching"
ElseIf Adodc3.Recordset.Fields(4) = Text1.Text Then
    Adodc3.Recordset.Filter = "BORROWERS_NAME ='" & (Text1.Text) & "'"
    Frame10.Visible = False
    Text1.Text = "ENTER SEARCH HERE"
    Command24.SetFocus
    Else
    
    End If

End Sub

Private Sub Command37_Click()
Adodc3.Recordset.Find "BOOK_TITLE like '" & (Text1.Text) & "'"
If Adodc3.Recordset.EOF Then
         MsgBox " That Title Does not Match any of the records. Make Sure This Title Exist"
ElseIf Text1.Text = "" Then
MsgBox "Please Do not Put A Null Value While Searching"
ElseIf Adodc3.Recordset.Fields(2) = Text1.Text Then
Adodc3.Recordset.Filter = "BOOK_TITLE ='" & (Text1.Text) & "'"
    Frame10.Visible = False
    Text1.Text = "ENTER SEARCH HERE"
    Command24.SetFocus
    Else
         End If
End Sub

Private Sub Command38_Click()
Frame10.Visible = False
Text1.Text = "ENTER SEARCH HERE"
End Sub

Private Sub Command4_Click()
Adodc1.RecordSource = Adodc1.Recordset.EditMode
Text21.Locked = False
Text22.Locked = False
Text23.Locked = False
Text34.Locked = False
ader2.Locked = False
Cnum2.Locked = False
Command4.Enabled = False
Command5.Enabled = True
End Sub

Private Sub Command40_Click()
Frame11.Visible = False
Text2.Text = "ENTER SEARCH HERE"
End Sub

Private Sub Command41_Click()
If Text11.Text = "bor" Then
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Find "BORROWERS_NAME like '" & (Text2.Text) & "'"
If Adodc1.Recordset.EOF Then
    MsgBox "Name Match Not Found. Make Sure That This NAME Exist"
ElseIf Text2 = "" Then
MsgBox "Please Do not Put A Null Value While Searching"
ElseIf Adodc1.Recordset.Fields(1) = Text2.Text Then
    Adodc1.Recordset.Filter = "BORROWERS_NAME='" & (Text2.Text) & "'"
    Frame11.Visible = False
    Text2.Text = "ENTER SEARCH HERE"
    Command30.SetFocus
    Else
End If
ElseIf Text11.Text = "bl" Then
Adodc7.Recordset.MoveFirst
Adodc7.Recordset.Find "BORROWERS_NAME like '" & (Text2.Text) & "'"
If Adodc7.Recordset.EOF Then
    MsgBox "Name Match Not Found. Make Sure That This NAME Exist"
ElseIf Text2 = "" Then
MsgBox "Please Do not Put A Null Value While Searching"
ElseIf Adodc7.Recordset.Fields(1) = Text2.Text Then
    Adodc7.Recordset.Filter = "BORROWERS_NAME='" & (Text2.Text) & "'"
    Frame11.Visible = False
    Text2.Text = "ENTER SEARCH HERE"
    Cmdrefresh.SetFocus
    Else
End If
End If
End Sub

Private Sub Command42_Click()
Frame12.Visible = False
Text3.Text = "Enter Search Here.................."
End Sub

Private Sub Command43_Click()
If Text11.Text = "bor" Then
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Find "BORROWERS_ID like '" & (Text2.Text) & "'"
If Adodc1.Recordset.EOF Then
    MsgBox "ID Number Match Not Found. Make Sure This ID Number Exist"
ElseIf Text2 = "" Then
MsgBox "Please Do not Put A Null Value While Searching"
ElseIf Adodc1.Recordset.Fields(0) = Text2.Text Then
    Adodc1.Recordset.Filter = "BORROWERS_ID='" & (Text2.Text) & "'"
    Frame11.Visible = False
    Text2.Text = "ENTER SEARCH HERE"
    Command30.SetFocus
    Else
End If
ElseIf Text11.Text = "bl" Then
Adodc7.Recordset.MoveFirst
Adodc7.Recordset.Find "BORROWERS_ID like '" & (Text2.Text) & "'"
If Adodc7.Recordset.EOF Then
    MsgBox "ID Number Match Not Found. Make Sure This ID Number Exist"
ElseIf Text2 = "" Then
MsgBox "Please Do not Put A Null Value While Searching"
ElseIf Adodc7.Recordset.Fields(0) = Text2.Text Then
    Adodc7.Recordset.Filter = "BORROWERS_ID='" & (Text2.Text) & "'"
    Frame11.Visible = False
    Text2.Text = "ENTER SEARCH HERE"
    Cmdrefresh.SetFocus
    Else
End If

End If
End Sub

Private Sub Command44_Click()
If Text11.Text = "bor" Then
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Find "COURSE like '" & (Text2.Text) & "'"
If Adodc1.Recordset.EOF Then
    MsgBox "Course Match Not Found. Make Sure That This Course Exist"
ElseIf Text2 = "" Then
MsgBox "Please Do not Put A Null Value While Searching"
ElseIf Adodc1.Recordset.Fields(2) = Text2.Text Then
    Adodc1.Recordset.Filter = "COURSE='" & (Text2.Text) & "'"
    Frame11.Visible = False
    Text2.Text = "ENTER SEARCH HERE"
    Command30.SetFocus
    Else
End If
ElseIf Text11.Text = "bl" Then
Adodc7.Recordset.MoveFirst
Adodc7.Recordset.Find "COURSE like '" & (Text2.Text) & "'"
If Adodc7.Recordset.EOF Then
    MsgBox "Course Match Not Found. Make Sure That This Course Exist"
ElseIf Text2 = "" Then
MsgBox "Please Do not Put A Null Value While Searching"
ElseIf Adodc7.Recordset.Fields(2) = Text2.Text Then
    Adodc7.Recordset.Filter = "COURSE='" & (Text2.Text) & "'"
    Frame11.Visible = False
    Text2.Text = "ENTER SEARCH HERE"
    Cmdrefresh.SetFocus
    Else
End If
End If
End Sub

Private Sub Command45_Click()
Adodc2.Recordset.MoveFirst
Adodc2.Recordset.Find "BOOK_TITLE like '" & (Text3.Text) & "'"
If Adodc2.Recordset.EOF Then
    MsgBox "Book Title Match Not Found. Make Sure This Book Title Exist"
ElseIf Text3 = "" Then
MsgBox "Please Do not Put A Null Value While Searching"
ElseIf Adodc2.Recordset.Fields(2) = Text3.Text Then
    Adodc2.Recordset.Filter = "BOOK_TITLE='" & (Text3.Text) & "'"
    Frame12.Visible = False
    Text3.Text = "Enter Search Here.................."
    Command12.SetFocus
    Else
End If

End Sub

Private Sub Command46_Click()
If Adodc6.Recordset.EOF Then
MsgBox "No more Books To Be Paid"
Else
Form12.Label4.Caption = "frmlost"
Form12.Show vbModal
End If

End Sub

Private Sub Command47_Click()
Adodc2.Recordset.MoveFirst
Adodc2.Recordset.Find "AUTHOR like '" & (Text3.Text) & "'"
If Adodc2.Recordset.EOF Then
    MsgBox "AUTHOR Match Not Found. Make Sure This AUTHOR Exist"
ElseIf Text2 = "" Then
MsgBox "Please Do not Put A Null Value While Searching"
ElseIf Adodc2.Recordset.Fields(3) = Text3.Text Then
    Adodc2.Recordset.Filter = "AUTHOR='" & (Text3.Text) & "'"
    Frame12.Visible = False
    Text3.Text = "Enter Search Here.................."
    Command12.SetFocus
    Else
End If

End Sub

Private Sub Command48_Click()
Adodc2.Recordset.MoveFirst
Adodc2.Recordset.Find "ISBN like '" & (Text3.Text) & "'"
If Adodc2.Recordset.EOF Then
    MsgBox "ID Number Match Not Found. Make Sure This ID Number Exist"
ElseIf Text2 = "" Then
MsgBox "Please Do not Put A Null Value While Searching"
ElseIf Adodc2.Recordset.Fields(1) = Text3.Text Then
    Adodc2.Recordset.Filter = "ISBN='" & (Text3.Text) & "'"
    Frame12.Visible = False
    Text3.Text = "Enter Search Here.................."
    Command12.SetFocus
    Else
End If

End Sub

Private Sub Command49_Click()
If Adodc5.Recordset.EOF Then
MsgBox "No more Books To Be Paid"
Else
Form12.Label4.Caption = "fines"
Form12.Show vbModal
End If

End Sub

Private Sub Command5_Click()
On Error Resume Next
Command30_Click
If Cnum2.Text = "" And ader2.Text = "" Then
MsgBox "Please Fill All The Needed Requirement"
Command4.Enabled = True
Command5.Enabled = False
Text21.Locked = True
Text22.Locked = True
Text23.Locked = True
Text34.Locked = True
ader2.Locked = True
Cnum2.Locked = True
Else
Adodc1.Recordset.Fields(0) = (Text21.Text)
Adodc1.Recordset.Fields(1) = (Text22.Text)
Adodc1.Recordset.Fields(2) = (Text23.Text)
Adodc1.Recordset.Fields(3) = (Text34.Text)
Adodc1.Recordset.Fields(4) = (ader2.Text)
Adodc1.Recordset.Fields(5) = (Cnum2.Text)
'Adodc1.Recordset.Fields(6) = (Label18.Caption)
Adodc1.Recordset.Update
Adodc1.Refresh
Form1.Refresh
Form7.Refresh
Command4.Enabled = True
Command5.Enabled = False
Text21.Locked = True
Text22.Locked = True
Text23.Locked = True
Text34.Locked = True
ader2.Locked = True
Cnum2.Locked = True
End If
End Sub

Private Sub Command50_Click()
Frame13.Visible = False
Text4.Text = "ENTER SEARCH HERE"

End Sub

Private Sub Command51_Click()
If Text11.Text = "dam" Then
Adodc8.Recordset.Find "BOOK_TITLE like '" & (Text4.Text) & "'"
If Adodc8.Recordset.EOF Then
         MsgBox " That Title Does not Match any of the records. Make Sure This Title Exist"
ElseIf Text4.Text = "" Then
MsgBox "Please Do not Put A Null Value While Searching"
ElseIf Adodc8.Recordset.Fields(2) = Text4.Text Then
Adodc8.Recordset.Filter = "BOOK_TITLE ='" & (Text4.Text) & "'"
    Frame13.Visible = False
    Text4.Text = "ENTER SEARCH HERE"
    Command15.SetFocus
Else
End If
ElseIf Text11.Text = "los" Then
Adodc6.Recordset.Find "BOOK_TITLE like '" & (Text4.Text) & "'"
If Adodc6.Recordset.EOF Then
         MsgBox " That Title Does not Match any of the records. Make Sure This Title Exist"
ElseIf Text4.Text = "" Then
MsgBox "Please Do not Put A Null Value While Searching"
ElseIf Adodc6.Recordset.Fields(2) = Text4.Text Then
Adodc6.Recordset.Filter = "BOOK_TITLE ='" & (Text4.Text) & "'"
    Frame13.Visible = False
    Text4.Text = "ENTER SEARCH HERE"
    Command17.SetFocus
Else
End If
ElseIf Text11.Text = "ret" Then
Adodc5.Recordset.Find "BOOK_TITLE like '" & (Text4.Text) & "'"
If Adodc5.Recordset.EOF Then
         MsgBox " That Title Does not Match any of the records. Make Sure This Title Exist"
ElseIf Text4.Text = "" Then
MsgBox "Please Do not Put A Null Value While Searching"
ElseIf Adodc5.Recordset.Fields(2) = Text4.Text Then
Adodc5.Recordset.Filter = "BOOK_TITLE ='" & (Text4.Text) & "'"
    Frame13.Visible = False
    Text4.Text = "ENTER SEARCH HERE"
    Command28.SetFocus
Else
End If

End If
End Sub

Private Sub Command52_Click()
If Text11.Text = "dam" Then
Adodc8.Recordset.Find "BORROWERS_NAME like '" & (Text4.Text) & "'"
If Adodc8.Recordset.EOF Then
         MsgBox "That Name Does not Match any of the records. Make Sure This Name Exist"
ElseIf Text4.Text = "" Then
MsgBox "Please Do not Put A Null Value While Searching"
ElseIf Adodc8.Recordset.Fields(4) = Text4.Text Then
    Adodc8.Recordset.Filter = "BORROWERS_NAME ='" & (Text4.Text) & "'"
    Frame13.Visible = False
    Text4.Text = "ENTER SEARCH HERE"
    Command15.SetFocus
    Else
    End If
    ElseIf Text11.Text = "los" Then
    Adodc6.Recordset.Find "BORROWERS_NAME like '" & (Text4.Text) & "'"
If Adodc6.Recordset.EOF Then
         MsgBox "That Name Does not Match any of the records. Make Sure This Name Exist"
ElseIf Text4.Text = "" Then
MsgBox "Please Do not Put A Null Value While Searching"
ElseIf Adodc6.Recordset.Fields(4) = Text4.Text Then
    Adodc6.Recordset.Filter = "BORROWERS_NAME ='" & (Text4.Text) & "'"
    Frame13.Visible = False
    Text4.Text = "ENTER SEARCH HERE"
    Command17.SetFocus
    Else
    End If
    ElseIf Text11.Text = "ret" Then
    Adodc5.Recordset.Find "BORROWERS_NAME like '" & (Text4.Text) & "'"
If Adodc5.Recordset.EOF Then
         MsgBox "That Name Does not Match any of the records. Make Sure This Name Exist"
ElseIf Text4.Text = "" Then
MsgBox "Please Do not Put A Null Value While Searching"
ElseIf Adodc5.Recordset.Fields(4) = Text4.Text Then
    Adodc5.Recordset.Filter = "BORROWERS_NAME ='" & (Text4.Text) & "'"
    Frame13.Visible = False
    Text4.Text = "ENTER SEARCH HERE"
    Command28.SetFocus
    Else
    End If

End If

End Sub

Private Sub Command53_Click()
If Text11.Text = "dam" Then
Adodc8.Recordset.Find "BORROWERS_ID like '" & (Text4.Text) & "'"
If Adodc8.Recordset.EOF Then
MsgBox "ID Number Match Not Found. Make Sure This ID Number Exist"
ElseIf Text4.Text = "" Then
MsgBox "Please Do not Put A Null Value While Searching"
ElseIf Adodc8.Recordset.Fields(3) = Text4.Text Then
        Adodc8.Recordset.Filter = "BORROWERS_ID ='" & (Text4.Text) & "'"
        Frame13.Visible = False
        Text4.Text = "ENTER SEARCH HERE"
        Command15.SetFocus
Else
End If
ElseIf Text11.Text = "los" Then
Adodc6.Recordset.Find "BORROWERS_ID like '" & (Text4.Text) & "'"
If Adodc6.Recordset.EOF Then
MsgBox "ID Number Match Not Found. Make Sure This ID Number Exist"
ElseIf Text4.Text = "" Then
MsgBox "Please Do not Put A Null Value While Searching"
ElseIf Adodc6.Recordset.Fields(3) = Text4.Text Then
        Adodc6.Recordset.Filter = "BORROWERS_ID ='" & (Text4.Text) & "'"
        Frame13.Visible = False
        Text4.Text = "ENTER SEARCH HERE"
        Command17.SetFocus
Else
End If
ElseIf Text11.Text = "ret" Then
Adodc5.Recordset.Find "BORROWERS_ID like '" & (Text4.Text) & "'"
If Adodc5.Recordset.EOF Then
MsgBox "ID Number Match Not Found. Make Sure This ID Number Exist"
ElseIf Text4.Text = "" Then
MsgBox "Please Do not Put A Null Value While Searching"
ElseIf Adodc5.Recordset.Fields(3) = Text4.Text Then
        Adodc5.Recordset.Filter = "BORROWERS_ID ='" & (Text4.Text) & "'"
        Frame13.Visible = False
        Text4.Text = "ENTER SEARCH HERE"
        Command28.SetFocus
Else
End If

End If
End Sub

Private Sub Command54_Click()
If Text11.Text = "dam" Then
Adodc8.Recordset.MoveFirst
Adodc8.Recordset.Find "ISBN like '" & (Text4.Text) & "'"
If Adodc8.Recordset.EOF Then
    MsgBox "ISBN Match Not Found. Make Sure This ISBN Exist"
ElseIf Text1 = "" Then
MsgBox "Please Do not Put A Null Value While Searching"
ElseIf Adodc8.Recordset.Fields(1) = Text4.Text Then
    Adodc8.Recordset.Filter = "ISBN='" & (Text4.Text) & "'"
    Frame13.Visible = False
    Text4.Text = "ENTER SEARCH HERE"
    Command15.SetFocus
    Else
End If
ElseIf Text11.Text = "los" Then
Adodc6.Recordset.MoveFirst
Adodc6.Recordset.Find "ISBN like '" & (Text4.Text) & "'"
If Adodc6.Recordset.EOF Then
    MsgBox "ISBN Match Not Found. Make Sure This ISBN Exist"
ElseIf Text1 = "" Then
MsgBox "Please Do not Put A Null Value While Searching"
ElseIf Adodc6.Recordset.Fields(1) = Text4.Text Then
    Adodc6.Recordset.Filter = "ISBN='" & (Text4.Text) & "'"
    Frame13.Visible = False
    Text4.Text = "ENTER SEARCH HERE"
    Command17.SetFocus
    Else
End If
ElseIf Text11.Text = "ret" Then
Adodc5.Recordset.MoveFirst
Adodc5.Recordset.Find "ISBN like '" & (Text4.Text) & "'"
If Adodc5.Recordset.EOF Then
    MsgBox "ISBN Match Not Found. Make Sure This ISBN Exist"
ElseIf Text1 = "" Then
MsgBox "Please Do not Put A Null Value While Searching"
ElseIf Adodc5.Recordset.Fields(1) = Text4.Text Then
    Adodc5.Recordset.Filter = "ISBN='" & (Text4.Text) & "'"
    Frame13.Visible = False
    Text4.Text = "ENTER SEARCH HERE"
    Command28.SetFocus
    Else
End If

End If
End Sub

Private Sub Command55_Click()
Frame11.Visible = True
Text11.Text = "bl"
End Sub

Private Sub Command6_Click()
Frame11.Visible = True
Text11.Text = "bor"
End Sub



Private Sub Command7_Click()
Adodc1.Recordset.Fields(6) = "Safe"
Adodc1.Recordset.Update
Adodc1.Refresh
Adodc7.Refresh
Label18.Caption = "Safe"
Command5.Enabled = True
End Sub

Private Sub Command8_Click()
Form2.Show vbModal
End Sub

Private Sub Command9_Click()
Form11.Show
End Sub

Private Sub crid_Click()
Form14.Show
Me.Hide
End Sub

Private Sub dalandan_Click()
Label4.BackColor = &HFF00&
Label8.BackColor = &HFF00&
Label7.BackColor = &HFF00&
Label9.BackColor = &HFF00&
Label10.BackColor = &HFF00&
Label12.BackColor = &HFF00&
pnp.Checked = False
redapple.Checked = False
dalandan.Checked = True
pandan.Checked = False
banana.Checked = False
choco.Checked = False
Emo.Checked = False
End Sub

Private Sub DataCombo1_Click(Area As Integer)
Call BookCategory
End Sub

Private Sub DTPicker1_Change()
Adodc4.Refresh
Adodc4.Recordset.Filter = ""
Adodc4.Recordset.Filter = "DATE_DUE ='" & Format(DTPicker1.Value, "mm/dd/yy") & "'"
End Sub

Private Sub Dueboks_Click()
SSTab1.Visible = False
Command20.Visible = False
End Sub

Private Sub Emo_Click()
Label4.BackColor = &H0&
Label8.BackColor = &H0&
Label7.BackColor = &H0&
Label9.BackColor = &H0&
Label10.BackColor = &H0&
Label12.BackColor = &H0&
pnp.Checked = False
redapple.Checked = False
dalandan.Checked = False
pandan.Checked = False
banana.Checked = False
choco.Checked = False
Emo.Checked = True
End Sub

Private Sub ext_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.MousePointer = vbHourglass
DTPicker1.Value = Format(Date, "mm/dd/yy")
Option1.Value = True
Me.Top = 0
Me.Left = 0
Adodc1.ConnectionString = ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & (App.Path & "\DataBase.mdb") & ";Persist Security Info=False;Jet OLEDB:Database Password=rjbbc")
    Adodc1.RecordSource = "Select * From BORROWERS Order by BORROWERS_NAME"
        Set DataGrid1.DataSource = Adodc1
Adodc2.ConnectionString = ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & (App.Path & "\DataBase.mdb") & ";Persist Security Info=False;Jet OLEDB:Database Password=rjbbc")
    Adodc2.RecordSource = "Select * From BOOKS Order by BOOK_TITLE"
        Set DataGrid2.DataSource = Adodc2
   Adodc3.ConnectionString = ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & (App.Path & "\DataBase.mdb") & ";Persist Security Info=False;Jet OLEDB:Database Password=rjbbc")
    Adodc3.RecordSource = "Select * From BORROWED_BOOKS Where BOOK_STATUS ='Borrowed' Order by BOOK_TITLE"
        Set DataGrid3.DataSource = Adodc3
Adodc4.ConnectionString = ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & (App.Path & "\DataBase.mdb") & ";Persist Security Info=False;Jet OLEDB:Database Password=rjbbc")
    Adodc4.RecordSource = "Select * From BORROWED_BOOKS Where DATE_RETURNED ='NOT YET' Order by BOOK_TITLE"
        Set DataGrid4.DataSource = Adodc4
    Adodc4.Recordset.Filter = ""
    Adodc4.Recordset.Filter = "DATE_DUE ='" & Format(DTPicker1.Value, "mm/dd/yy") & "'"
Adodc5.ConnectionString = ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & (App.Path & "\DataBase.mdb") & ";Persist Security Info=False;Jet OLEDB:Database Password=rjbbc")
    Adodc5.RecordSource = "Select * From BORROWED_BOOKS Where BOOK_STATUS ='Returned' Order by BOOK_TITLE"
        Set DataGrid5.DataSource = Adodc5
Adodc6.ConnectionString = ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & (App.Path & "\DataBase.mdb") & ";Persist Security Info=False;Jet OLEDB:Database Password=rjbbc")
    Adodc6.RecordSource = "Select * From BORROWED_BOOKS Where BOOK_STATUS ='Lost' Order by Book_TITLE"
Set DataGrid6.DataSource = Adodc6
Adodc8.ConnectionString = ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & (App.Path & "\DataBase.mdb") & ";Persist Security Info=False;Jet OLEDB:Database Password=rjbbc")
    Adodc8.RecordSource = "Select * From BORROWED_BOOKS Where BOOK_STATUS ='Damaged' Order by Book_TITLE"
Set DataGrid8.DataSource = Adodc8
Adodc7.ConnectionString = ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & (App.Path & "\DataBase.mdb") & ";Persist Security Info=False;Jet OLEDB:Database Password=rjbbc")
Adodc7.RecordSource = "Select * From BORROWERS Where BORROWING_STATUS ='Blacklisted' Order by BORROWERS_NAME"
Set DataGrid7.DataSource = Adodc7
DataGrid1.AllowUpdate = False
DataGrid2.AllowUpdate = False
DataGrid3.AllowUpdate = False
DataGrid4.AllowUpdate = False
DataGrid5.AllowUpdate = False
DataGrid6.AllowUpdate = False
Me.MousePointer = vbDefault
Label16.Caption = (frmLogin.txtUserName.Text)
Label17.Caption = Dialog7.Label5.Caption
End Sub




Private Sub Home_Click()
frmSplash.Show
Unload Me
End Sub

Private Sub rec_Click()

End Sub

Private Sub Form_Terminate()
Dim reply As Integer
reply = MsgBox("Are you sure you want to SIGN OUT?", vbExclamation + vbYesNo, "Library System")
If reply = vbYes Then
    End
Else
    Cancel = 1
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim reply As Integer
reply = MsgBox("Are you sure you want to exit?", vbExclamation + vbYesNo, "Library System")
If reply = vbYes Then
    End
Else
    Cancel = 1
End If

End Sub


Private Sub Label17_Change()
If Label17.Caption = "Administrator" Then
Command8.Enabled = True
Command4.Enabled = True
Command16.Enabled = True
Command19.Enabled = True
Cmd20.Enabled = True
Cmdchanstat.Enabled = True
adbok.Enabled = True
adbor.Enabled = True
adbl.Enabled = True
userpro.Enabled = True
Else
Command8.Enabled = False
Command4.Enabled = False
Command16.Enabled = False
Command19.Enabled = False
Cmd20.Enabled = False
Cmdchanstat.Enabled = False
adbok.Enabled = False
adbor.Enabled = False
adbl.Enabled = False
userpro.Enabled = False
End If
End Sub

Private Sub libtran_Click()
SSTab1.Visible = True
SSTab1.Tab = 4
Command20.Visible = True
End Sub

Private Sub newbor_Click()
Form4.Show vbModal
End Sub

Private Sub Option1_Change()
If Option1.Value = True Then
DTPicker1.Visible = True
Else
DTPicker1.Visible = False
End If
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
DTPicker1.Visible = True
Label5.Visible = True
End If
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
DTPicker1.Visible = False
Label5.Visible = False
Adodc4.Refresh
End If
End Sub

Private Sub pandan_Click()
Label4.BackColor = &HC0FFC0
Label8.BackColor = &HC0FFC0
Label7.BackColor = &HC0FFC0
Label9.BackColor = &HC0FFC0
Label10.BackColor = &HC0FFC0
Label12.BackColor = &HC0FFC0
pnp.Checked = False
redapple.Checked = False
dalandan.Checked = False
pandan.Checked = True
banana.Checked = False
choco.Checked = False
Emo.Checked = False
End Sub

Private Sub pnp_Click()
Label4.BackColor = &H80FF&
Label8.BackColor = &H80FF&
Label7.BackColor = &H80FF&
Label9.BackColor = &H80FF&
Label10.BackColor = &H80FF&
Label12.BackColor = &H80FF&
pnp.Checked = True
redapple.Checked = False
dalandan.Checked = False
pandan.Checked = False
banana.Checked = False
choco.Checked = False
Emo.Checked = False
End Sub

Private Sub print_Click()
SSTab1.Visible = True
SSTab1.Tab = 3
Command20.Visible = True
End Sub

Private Sub redapple_Click()
Label4.BackColor = &H8080FF
Label8.BackColor = &H8080FF
Label7.BackColor = &H8080FF
Label9.BackColor = &H8080FF
Label10.BackColor = &H8080FF
Label12.BackColor = &H8080FF
pnp.Checked = False
redapple.Checked = True
dalandan.Checked = False
pandan.Checked = False
banana.Checked = False
choco.Checked = False
Emo.Checked = False
End Sub


Private Sub signout_Click()
frmLogin.txtUserName.Text = ""
frmLogin.txtPassword.Text = ""
frmLogin.Adodc1.Refresh
Me.Refresh
frmLogin.Show
Me.Hide
frmLogin.Refresh
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)

End Sub

'Private Sub Text1_Change()
'On Error Resume Next
'Adodc3.Recordset.MoveFirst
'Adodc3.Recordset.Find "ISBN like '" & (Text1.Text) & "'"
'If Adodc3.Recordset.EOF Then
'Exit Sub
'Else
'Adodc3.Recordset.Filter = "ISBN '" & (Text1.Text) & "'"
'End If
'End Sub
Private Sub Text1_Change()

End Sub

Private Sub Text1_Click()
Text1.Text = ""
End Sub

Private Sub Text2_Click()
Text2.Text = ""
End Sub

Private Sub Text3_Click()
Text3.Text = ""
End Sub

Private Sub Text4_Click()
Text4.Text = ""
End Sub

Private Sub userpro_Click()
Dialog8.Show
End Sub
