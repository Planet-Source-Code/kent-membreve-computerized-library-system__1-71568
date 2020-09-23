Attribute VB_Name = "Module1"
Dim Conn As New ADODB.Connection
Dim rsBookCategory As New ADODB.Recordset
Dim rsBorrowedkBokNo As New ADODB.Recordset
Dim rsBorrowedBorID As New ADODB.Recordset
Dim rsReturnBookQty As New ADODB.Recordset
Dim rsSubtractBookQty As New ADODB.Recordset
Global finesCharge As Double
Dim rsfinesCharge As New ADODB.Recordset
Dim rsFines As New ADODB.Recordset
Global theFines As Double
Dim rsPassword As New ADODB.Recordset
Dim rsPassword1 As New ADODB.Recordset
Global pWords As String
Dim rsUserPassword As New ADODB.Recordset
Global varUserPassword As String
Dim rsyearPrint As New ADODB.Recordset
Dim rsauthorPrint As New ADODB.Recordset
Dim rsstatusPrint As New ADODB.Recordset
Public catRS As ADODB.Recordset
Public Sub Connect()
On Error Resume Next
Conn.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & (App.Path & "\DataBase.mdb") & ";Persist Security Info=False;Jet OLEDB:Database Password=rjbbc")
End Sub
Public Sub BookCategory()
Call Connect
rsBookCategory.Open "Select * From CATEGORY Order by CATEGORY", Conn, adOpenStatic, adLockOptimistic
Set Form3.DataCombo1.RowSource = rsBookCategory
    Form3.DataCombo1.ListField = "CATEGORY"
Set rsBookCategory = Nothing
Set Conn = Nothing
End Sub
'Public Sub BorrowedkBokNo()
'Call Connect
'rsBorrowedkBokNo.Open "Select * From BOOKS Order by BOOK_NO", Conn, adOpenStatic, adLockOptimistic
'Set Form4.DataCombo1.RowSource = rsBorrowedkBokNo
  '  Form4.DataCombo1.ListField = "BOOK_NO"
'Set rsBarrowedkBokNo = Nothing
'Set Conn = Nothing
'End Sub
'Public Sub BorrowedBorID()
'Call Connect
'rsBorrowedBorID.Open "Select * From BORROWERS Order by BORROWERS_ID", Conn, adOpenStatic, adLockOptimistic
'Set Form4.DataCombo2.RowSource = rsBorrowedBorID
   ' Form4.DataCombo2.ListField = "BORROWERS_ID"
'Set rsBorrowedBorID = Nothing
'Set Conn = Nothing
'End Sub
Public Sub SubtractBookQty()
On Error Resume Next
Call Connect
rsSubtractBookQty.Open "Select * From BOOKS Where BOOK_NO ='" & (Form4.Label15.Caption) & "' Order by BOOK_NO", Conn, adOpenStatic, adLockOptimistic
    rsSubtractBookQty.Fields(9) = Val(rsSubtractBookQty.Fields(9)) - 1
    rsSubtractBookQty.Fields(8) = Val(rsSubtractBookQty.Fields(8)) + 1
    rsSubtractBookQty.Update
Set rsSubtractBookQty = Nothing
Set Conn = Nothing
End Sub
Public Sub ReturnBookQty()
Call Connect
rsReturnBookQty.Open "Select * From BOOKS Where BOOK_NO ='" & (Form1.Adodc3.Recordset.Fields(0)) & "' Order by BOOK_NO", Conn, adOpenStatic, adLockOptimistic
    rsReturnBookQty.Fields(8) = Val(rsReturnBookQty.Fields(8)) - 1
    rsReturnBookQty.Fields(9) = Val(rsReturnBookQty.Fields(9)) + 1
    rsReturnBookQty.Update
Set rsReturnBookQty = Nothing
Set Conn = Nothing
End Sub
Public Sub finesCharge_()
Call Connect
rsfinesCharge.Open "Select * From FINES", Conn, adOpenStatic, adLockOptimistic
    finesCharge = (rsfinesCharge.Fields(0))
Set rsfinesCharge = Nothing
Set Conn = Nothing
End Sub
Public Sub setFines()
Call Connect
rsFines.Open "Select * From Fines", Conn, adOpenStatic, adLockOptimistic
    theFines = (rsFines.Fields(0))
Set rsFines = Nothing
Set Conn = Nothing
End Sub
Public Sub updateFines()
Call Connect
rsFines.Open "Select * From Fines", Conn, adOpenStatic, adLockOptimistic
    rsFines.Fields(0) = (Form6.Text1.Text)
    rsFines.Update
Set rsFines = Nothing
Set Conn = Nothing
End Sub
Public Sub Pword()
Call Connect
rsPassword.Open "Select * From SECURITY_PASSWORD ", Conn, adOpenStatic, adLockOptimistic
    pWords = (rsPassword.Fields(0))
Set rsPassword = Nothing
Set Conn = Nothing
End Sub
'Public Sub updatePword()
'Call Connect
'rsPassword1.Open "Select * From SECURITY_PASSWORD ", Conn, adOpenStatic, adLockOptimistic
 '   rsPassword1.Fields(0) = (Form8.Text2.Text)
  '  rsPassword1.Update
   ' MsgBox "Changes has been successfully save.", vbInformation, "Library System"
   ' Unload Form8
'Set rsPassword1 = Nothing
'Set Conn = Nothing
'End Sub
Public Sub UserPassword()
Call Connect
rsUserPassword.Open "Select * From SECURITY_PASSWORD ", Conn, adOpenStatic, adLockOptimistic
    varUserPassword = (rsUserPassword.Fields(0))
Set rsUserPassword = Nothing
Set Conn = Nothing
End Sub
Public Sub yearPrint()
Call Connect
rsyearPrint.Open "Select * From BORROWERS Order by CURRENT_YEAR", Conn, adOpenStatic, adLockOptimistic
    Set Form11.DataCombo1.RowSource = rsyearPrint
        Form11.DataCombo1.ListField = "CURRENT_YEAR"
Set rsyearPrint = Nothing
Set Conn = Nothing
End Sub
Public Sub authorPrint()
Call Connect
rsauthorPrint.Open "Select * From BOOKS Order by AUTHOR", Conn, adOpenStatic, adLockOptimistic
    Set Form11.DataCombo2.RowSource = rsauthorPrint
        Form11.DataCombo2.ListField = "AUTHOR"
Set rsauthorPrint = Nothing
Set Conn = Nothing
End Sub
Public Sub Stat()
Call Connect
rsstatusPrint.Open "Select * From BORROWED_BOOKS Order by BOOK_STATUS", Conn, adOpenStatic, adLockOptimistic
Set Form11.DataCombo3.RowSource = rsstatusPrint
        Form11.DataCombo3.ListField = "BOOK_STATUS"
Set rsauthorPrint = Nothing
Set Conn = Nothing
End Sub


