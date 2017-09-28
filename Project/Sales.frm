VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Sales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SALES"
   ClientHeight    =   9795
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   10200
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9795
   ScaleWidth      =   10200
   Begin VB.Frame Frame1 
      Height          =   9735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10095
      Begin VB.Frame Frame2 
         Height          =   1575
         Left            =   120
         TabIndex        =   28
         Top             =   120
         Width           =   9735
         Begin VB.TextBox txtinvoiceno 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1440
            TabIndex        =   32
            Top             =   360
            Width           =   2535
         End
         Begin VB.ComboBox cmbtype 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5400
            TabIndex        =   30
            Text            =   "--Select--"
            Top             =   1080
            Width           =   2535
         End
         Begin VB.ComboBox cmbcustomer 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5400
            TabIndex        =   29
            Top             =   360
            Width           =   2535
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Left            =   1440
            TabIndex        =   31
            Top             =   1080
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   556
            _Version        =   393216
            Format          =   39583745
            CurrentDate     =   38760
         End
         Begin VB.Label Label1 
            Caption         =   "Bill No"
            Height          =   255
            Left            =   240
            TabIndex        =   36
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Customer"
            Height          =   255
            Left            =   4440
            TabIndex        =   35
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Date"
            Height          =   255
            Left            =   240
            TabIndex        =   34
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label4 
            Caption         =   "Pay Type"
            Height          =   255
            Left            =   4440
            TabIndex        =   33
            Top             =   1080
            Width           =   735
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1935
         Left            =   120
         TabIndex        =   17
         Top             =   1800
         Width           =   9735
         Begin VB.ComboBox cmbcat 
            Height          =   315
            Left            =   1200
            TabIndex        =   23
            Top             =   360
            Width           =   2535
         End
         Begin VB.ComboBox cmbatr 
            Height          =   315
            Left            =   5280
            TabIndex        =   22
            Top             =   360
            Width           =   2535
         End
         Begin VB.ComboBox cmbpub 
            Height          =   315
            Left            =   1200
            TabIndex        =   21
            Top             =   1080
            Width           =   2535
         End
         Begin VB.ComboBox cmbsub 
            Height          =   315
            Left            =   5280
            TabIndex        =   20
            Top             =   1080
            Width           =   2535
         End
         Begin VB.CommandButton cmdview 
            Caption         =   "View"
            Height          =   375
            Left            =   6720
            TabIndex        =   19
            Top             =   1440
            Width           =   975
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   7800
            TabIndex        =   18
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label5 
            Caption         =   "Category"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label6 
            Caption         =   "Author"
            Height          =   255
            Left            =   4200
            TabIndex        =   26
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "Publisher"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label8 
            Caption         =   "Subject"
            Height          =   255
            Left            =   4200
            TabIndex        =   24
            Top             =   1080
            Width           =   855
         End
      End
      Begin VB.Frame Frame4 
         Height          =   975
         Left            =   120
         TabIndex        =   11
         Top             =   3840
         Width           =   9735
         Begin VB.ComboBox cmbisbn 
            Height          =   315
            Left            =   4080
            TabIndex        =   37
            Top             =   360
            Width           =   2175
         End
         Begin VB.ComboBox cmbbook 
            Height          =   315
            Left            =   600
            TabIndex        =   13
            Top             =   360
            Width           =   2175
         End
         Begin VB.TextBox txtrate 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6960
            TabIndex        =   12
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label Label9 
            Caption         =   "Book"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label10 
            Caption         =   "ISBN Number"
            Height          =   255
            Left            =   3000
            TabIndex        =   15
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label11 
            Caption         =   "Rate"
            Height          =   255
            Left            =   6480
            TabIndex        =   14
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame Frame5 
         Height          =   855
         Left            =   120
         TabIndex        =   7
         Top             =   4920
         Width           =   9735
         Begin VB.CommandButton cmdadd 
            Caption         =   "Add"
            Height          =   495
            Left            =   3000
            TabIndex        =   10
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmddelete 
            Caption         =   "Delete"
            Height          =   495
            Left            =   4260
            TabIndex        =   9
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdcancel 
            Caption         =   "Cancel"
            Height          =   495
            Left            =   5520
            TabIndex        =   8
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame6 
         Height          =   2175
         Left            =   120
         TabIndex        =   5
         Top             =   5880
         Width           =   9735
         Begin MSFlexGridLib.MSFlexGrid grddetails 
            Height          =   1815
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   8535
            _ExtentX        =   15055
            _ExtentY        =   3201
            _Version        =   393216
         End
      End
      Begin VB.Frame Frame7 
         Height          =   1335
         Left            =   120
         TabIndex        =   1
         Top             =   8160
         Width           =   9735
         Begin VB.TextBox txtgrandtotal 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   6120
            TabIndex        =   3
            Top             =   360
            Width           =   2775
         End
         Begin VB.CommandButton cmdsave 
            Caption         =   "Save"
            Height          =   375
            Left            =   7920
            TabIndex        =   2
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label12 
            Caption         =   "Grand Total"
            Height          =   255
            Left            =   5040
            TabIndex        =   4
            Top             =   360
            Width           =   975
         End
      End
   End
End
Attribute VB_Name = "Sales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim temp As New ADODB.Recordset



Private Sub cmbbook_Click()
FillComboWithID "tbl_bookdetails", cmbisbn, "ISBNno", "bd_id", "book_id=" & CboData(cmbbook) & " and bd_status='False'"
If temp.State = 1 Then temp.Close
temp.Open "select * from tbl_bookmaster where book_id=" & CboData(cmbbook) & "", con, adOpenKeyset, adLockOptimistic
If (temp.RecordCount > 0) Then
 txtrate.Text = temp.Fields("book_price")
 End If
 temp.Close

End Sub

Private Sub cmdadd_Click()
If addvalidation Then


Dim i, j As Integer

i = grddetails.Rows

If j <> i - 1 Then
 j = i - 1
End If

grddetails.TextMatrix(j, 0) = CboData(cmbisbn)
grddetails.TextMatrix(j, 1) = cmbbook.Text
grddetails.TextMatrix(j, 2) = 1
grddetails.TextMatrix(j, 3) = txtrate.Text
grddetails.TextMatrix(j, 4) = 1 * Val(txtrate.Text)
grddetails.TextMatrix(j, 5) = cmbisbn.Text

grddetails.Rows = grddetails.Rows + 1



If rsbookdetails.State = 1 Then rsbookdetails.Close
 rsbookdetails.Open "select * from tbl_bookdetails where bd_id=" & CboData(cmbisbn) & "", con, adOpenKeyset, adLockOptimistic
  rsbookdetails.Fields("bd_status") = 1
  rsbookdetails.Update
  rsbookdetails.Close


cmbisbn.Text = "--select--"
cmbbook.Text = "--select--"
txtrate.Text = ""
cmbbook.SetFocus

txtgrandtotal.Text = Val(txtgrandtotal.Text) + Val(grddetails.TextMatrix(j, 4))

End If

End Sub

Private Sub cmdcancel_Click()
cmbbook.Text = "--select--"
cmbisbn.Text = "--select--"
txtrate.Text = ""
End Sub

Private Sub cmddelete_Click()

Dim i As Integer
i = grddetails.row

If rsbookdetails.State = 1 Then rsbookdetails.Close
 rsbookdetails.Open "select * from tbl_bookdetails where bd_id=" & Val(grddetails.TextMatrix(i, 0)) & "", con, adOpenKeyset, adLockOptimistic
  rsbookdetails.Fields("bd_status") = 0
  
  rsbookdetails.Update
  rsbookdetails.Close
  
txtgrandtotal.Text = Val(txtgrandtotal.Text) - Val(grddetails.TextMatrix(i, 4))
grddetails.RemoveItem i





  
End Sub

Private Sub cmdSave_Click()
If headervalidation Then
If rsSalesead.State = 1 Then rsSalesead.Close
rsSalesead.Open "select * from tbl_salesheader", con, adOpenKeyset, adLockOptimistic
rsSalesead.AddNew
   rsSalesead.Fields("sh_billno") = txtinvoiceno.Text
   rsSalesead.Fields("sh_date") = DTPicker1.Value
   rsSalesead.Fields("cst_id") = CboData(cmbcustomer)
   rsSalesead.Fields("sh_paymenttype") = cmbtype.Text
   rsSalesead.Fields("sh_grand_total") = txtgrandtotal.Text
rsSalesead.Update
rsSalesead.Close


If temp.State = 1 Then temp.Close
temp.Open "select max(sh_id) as sid from tbl_salesheader", con, adOpenKeyset, adLockOptimistic
Dim salid As Integer
salid = temp.Fields("sid")

Dim n As Integer
n = grddetails.Rows

For i = 1 To n - 2
 
 If rsSalesDetails.State = 1 Then rsSalesDetails.Close
 rsSalesDetails.Open "select * from tbl_salesdetails", con, adOpenKeyset, adLockOptimistic
 rsSalesDetails.AddNew
   rsSalesDetails.Fields("bd_id") = Val(grddetails.TextMatrix(i, 0))
   rsSalesDetails.Fields("b_qty") = Val(grddetails.TextMatrix(i, 2))
   rsSalesDetails.Fields("sd_total") = Val(grddetails.TextMatrix(i, 4))
   rsSalesDetails.Fields("sh_id") = salid
rsSalesDetails.Update
rsSalesDetails.Close


If rsbookdetails.State = 1 Then rsbookdetails.Close
 rsbookdetails.Open "select * from tbl_bookdetails where bd_id=" & Val(grddetails.TextMatrix(i, 0)) & "", con, adOpenKeyset, adLockOptimistic
  rsbookdetails.Fields("bd_status") = 1
  
  rsbookdetails.Update
  rsbookdetails.Close




Next i

 
End If


End Sub

Private Sub cmdview_Click()
If viewvalidation Then
cmbbook.Clear
FillComboWithID "tbl_bookmaster", cmbbook, "book_name", "book_id", "c_id=" & CboData(cmbcat) & " and atr_id=" & CboData(cmbatr) & " and pub_id=" & CboData(cmbpub) & " and sub_id=" & CboData(cmbsub) & ""
End If

End Sub

Private Sub Command1_Click()
cmbatr.Clear
cmbbook.Clear
cmbcat.Clear
cmbsub.Clear
cmbpub.Clear

Fillcombo "tbl_customerdetails", cmbcustomer, "cst_name", "cst_id"
Fillcombo "tbl_bookmaster", cmbbook, "book_name", "book_id"
Fillcombo "tbl_category", cmbcat, "c_name", "c_id"
Fillcombo "tbl_author", cmbatr, "atr_name", "atr_id"
Fillcombo "tbl_publisher", cmbpub, "pub_name", "pub_id"
Fillcombo "tbl_subject", cmbsub, "sub_name", "sub_id"
End Sub

Private Sub Form_Load()
fillgrid

fillPayType
txtgrandtotal.Enabled = False
txtrate.Enabled = False

Fillcombo "tbl_customerdetails", cmbcustomer, "cst_name", "cst_id"
Fillcombo "tbl_bookmaster", cmbbook, "book_name", "book_id"
Fillcombo "tbl_category", cmbcat, "c_name", "c_id"
Fillcombo "tbl_author", cmbatr, "atr_name", "atr_id"
Fillcombo "tbl_publisher", cmbpub, "pub_name", "pub_id"
Fillcombo "tbl_subject", cmbsub, "sub_name", "sub_id"
End Sub

Private Sub fillPayType()
cmbtype.AddItem ("Cash")
cmbtype.AddItem ("Cheque")
End Sub
Private Sub fillgrid()
grddetails.Rows = 2
grddetails.Cols = 6

grddetails.ColWidth(0) = 0
grddetails.ColWidth(1) = 1400
grddetails.ColWidth(2) = 1400
grddetails.ColWidth(3) = 1400
grddetails.ColWidth(4) = 1400
grddetails.ColWidth(5) = 1400

grddetails.TextMatrix(0, 1) = "Book"
grddetails.TextMatrix(0, 2) = "Quantity"
grddetails.TextMatrix(0, 3) = "Rate"
grddetails.TextMatrix(0, 4) = "Total"
grddetails.TextMatrix(0, 5) = "ISBN No"

End Sub

Public Function headervalidation() As Boolean

If Trim(txtinvoiceno.Text) = "" Then
   MsgBox "Enter Bill Number", vbInformation, App.Title
   txtinvoiceno.SetFocus
   headervalidation = False
   Exit Function
End If

If Trim(cmbcustomer.Text) = "--select--" Then
  MsgBox "select Customer", vbInformation, App.Title
  cmbcustomer.SetFocus
  headervalidation = False
  Exit Function
End If

If Trim(cmbtype.Text) = "--select--" Then
  MsgBox "select type", vbInformation, App.Title
  cmbtype.SetFocus
  headervalidation = False
  Exit Function
End If
 
   If Trim(txtgrandtotal.Text) = "" Then
   MsgBox "Grand Total Not Found", vbInformation, App.Title
   headervalidation = False
   Exit Function
End If

headervalidation = True

End Function

Public Function viewvalidation() As Boolean

If Trim(cmbcat.Text) = "--select--" Then
  MsgBox "select Catgeory", vbInformation, App.Title
  cmbcat.SetFocus
  viewvalidation = False
  Exit Function
End If

If Trim(cmbatr.Text) = "--select--" Then
  MsgBox "select Author", vbInformation, App.Title
  cmbatr.SetFocus
  viewvalidation = False
  Exit Function
End If


If Trim(cmbpub.Text) = "--select--" Then
  MsgBox "select publisher", vbInformation, App.Title
  cmbpub.SetFocus
  viewvalidation = False
  Exit Function
End If

If Trim(cmbsub.Text) = "--select--" Then
  MsgBox "select subject", vbInformation, App.Title
  cmbsub.SetFocus
  viewvalidation = False
  Exit Function
End If




viewvalidation = True

End Function


Public Function addvalidation() As Boolean

If Trim(cmbbook.Text) = "--select--" Then
  MsgBox "select Book", vbInformation, App.Title
  cmbbook.SetFocus
  addvalidation = False
  Exit Function
End If

If Trim(cmbisbn.Text) = "--select--" Then
  MsgBox "select ISBN Number", vbInformation, App.Title
  cmbisbn.SetFocus
  addvalidation = False
  Exit Function
End If




addvalidation = True

End Function


