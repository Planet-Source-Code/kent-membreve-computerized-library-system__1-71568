VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Dialog8 
   BackColor       =   &H00808000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   5865
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   10320
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command4 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   6120
      TabIndex        =   9
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete User"
      Height          =   375
      Left            =   4800
      TabIndex        =   4
      Top             =   5160
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Demote Account"
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   5160
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Upgrade Account"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   5160
      Width           =   1935
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2055
      Left            =   600
      TabIndex        =   1
      Top             =   2040
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   3625
      _Version        =   393216
      AllowUpdate     =   0   'False
      Enabled         =   -1  'True
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   8280
      Top             =   4680
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
   Begin VB.CommandButton OKButton 
      Caption         =   "BACK"
      Height          =   375
      Left            =   6840
      TabIndex        =   0
      Top             =   5160
      Width           =   1935
   End
   Begin VB.Label Label4 
      Height          =   375
      Left            =   1800
      TabIndex        =   8
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Account Level"
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "Dialog8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Adodc1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Label2.Caption = (Adodc1.Recordset.Fields(1))
Label4.Caption = (Adodc1.Recordset.Fields(3))
End Sub

Private Sub Command1_Click()
Adodc1.Recordset.Fields(3) = "Administrator"
Adodc1.Recordset.Update
Adodc1.Refresh
frmLogin.Adodc1.Refresh
Command4_Click
End Sub

Private Sub Command2_Click()
Adodc1.Recordset.Fields(3) = "Attendant"
Adodc1.Recordset.Update
Adodc1.Refresh
frmLogin.Adodc1.Refresh
Command4_Click
End Sub

Private Sub Command3_Click()
On Error Resume Next
Dim repp2 As Integer
If Val(Adodc1.Recordset.RecordCount) <= 0 Then
    MsgBox "No more Records to be deleted.", vbInformation, "Confirm"
    Exit Sub
End If
repp2 = MsgBox("You are about to delete 1 record." & vbCrLf & "If you click YES, you won't be able to undo this delete operation." & vbCrLf & "Are you sure you want to delete these record?", vbCritical + vbYesNo, "Confirm Delete")
If repp2 = vbYes Then
    Adodc1.Recordset.Delete
    Adodc1.Recordset.MoveNext
    If Adodc1.Recordset.EOF Then
        Adodc1.Recordset.MoveLast
    End If
    Call ReturnBookQty
    MsgBox "Record has been successfuly deleted.", vbInformation, "Confirm"
End If

End Sub

Private Sub Command4_Click()
Adodc1.Refresh
End Sub

Private Sub Form_Load()
Adodc1.ConnectionString = ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & (App.Path & "\DataBase.mdb") & ";Persist Security Info=False;Jet OLEDB:Database Password=rjbbc")
    Adodc1.RecordSource = "Select * From SECURITY_PASSWORD Order by USER_NAME"
        Set DataGrid1.DataSource = Adodc1

End Sub

Private Sub OKButton_Click()
Unload Me
End Sub
