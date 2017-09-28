Attribute VB_Name = "Module1"
 Public con As New ADODB.Connection
 Public rscst As New ADODB.Recordset
 Public rsstf As New ADODB.Recordset
 Public rsCat As New ADODB.Recordset
 Public rslang As New ADODB.Recordset
 Public rspub As New ADODB.Recordset
 Public rsatr As New ADODB.Recordset
 Public rssup As New ADODB.Recordset
 Public rssub As New ADODB.Recordset
 Public rsadmin As New ADODB.Recordset
 Public rsbook As New ADODB.Recordset
 Public rsbookdetails As New ADODB.Recordset
 Public rsLogin As New ADODB.Recordset
 Public rsChange As New ADODB.Recordset
 Public rsLoginReg As New ADODB.Recordset

  Public rsPurchaseHead As New ADODB.Recordset
  Public rsPurchaseDetails As New ADODB.Recordset
  
  Public rsSalesead As New ADODB.Recordset
  Public rsSalesDetails As New ADODB.Recordset
  
 Public rs As New ADODB.Recordset
 
Public loginid As Integer
Public logintype As String
Public loginname As String
 
 

Public Sub Main()
con.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=db_bookshop;Data Source=AVIN"
Load frmsplash
frmsplash.Show

End Sub

Public Function Fillcombo(table As String, cmb As ComboBox, textfield As String, valuefield As String)
If rs.State = 1 Then rs.Close
rs.Open "select * from " & table & "", con, adOpenKeyset, adLockOptimistic
If rs.EOF = False And rs.BOF = False Then
rs.MoveFirst
cmb.Clear
While Not rs.EOF
cmb.AddItem (rs.Fields(textfield))
cmb.ItemData(cmb.NewIndex) = rs.Fields(valuefield)
rs.MoveNext
Wend
cmb.Text = "--select--"
End If

End Function
Public Function FillComboWithID(table As String, cmb As ComboBox, textfield As String, valuefield As String, condition As String)
If rs.State = 1 Then rs.Close
rs.Open "select * from " & table & " where " & condition & "", con, adOpenKeyset, adLockOptimistic
If rs.EOF = False And rs.BOF = False Then
rs.MoveFirst
cmb.Clear
While Not rs.EOF
cmb.AddItem (rs.Fields(textfield))
cmb.ItemData(cmb.NewIndex) = rs.Fields(valuefield)
rs.MoveNext
Wend
cmb.Text = "--select--"
End If

End Function
Public Function CboData(cbo As ComboBox) As Variant
CboData = 0
If cbo.ListIndex <> -1 Then
CboData = cbo.ItemData(cbo.ListIndex)
End If
End Function
Public Function GetComboIndex(cbo As ComboBox, itData As Variant) As Long
Dim i As Integer
GetComboIndex = -1
For i = 0 To cbo.ListCount - 1
If cbo.ItemData(i) = itData Then
GetComboIndex = i
Exit For
End If
Next
End Function
Public Sub SetComboItem(cbo As ComboBox, itData As Variant)
cbo.ListIndex = GetComboIndex(cbo, itData)
End Sub

Public Sub FillIntelliSense(lstBox As ListBox, fieldName As String, showControl As TextBox)
    With lstBox
        .Visible = True
        .Top = showControl.Top + showControl.Height
        .Left = showControl.Left
        Dim rs As New ADODB.Recordset
            lstBox.Clear
            rs.Open "select " & fieldName & " from " & showControl.Tag & " where " & fieldName & " like '" & showControl.Text & "%'", con, adOpenKeyset, adLockOptimistic
            If rs.RecordCount > 0 Then rs.MoveFirst
            While Not rs.EOF
             lstBox.AddItem (rs(0))
            rs.MoveNext
            Wend
            ' .SetFocus
    End With
End Sub

Public Function selectcombo(field As String, cmb As ComboBox)
Dim s As String
Dim la As Integer
s = field
la = cmb.ListCount
Dim i As Integer
For i = 0 To la - 1
If cmb.ItemData(i) = Val(s) Then
cmb.ListIndex = i
End If
Next i
End Function



