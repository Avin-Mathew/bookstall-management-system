VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form purchase 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PURCHASE"
   ClientHeight    =   9825
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   10260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9825
   ScaleWidth      =   10260
   Begin VB.Frame Frame1 
      Height          =   9735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10095
      Begin VB.Frame Frame7 
         Height          =   1335
         Left            =   120
         TabIndex        =   33
         Top             =   8160
         Width           =   9735
         Begin VB.CommandButton cmdsave 
            Caption         =   "Save"
            Height          =   375
            Left            =   7920
            TabIndex        =   36
            Top             =   840
            Width           =   1095
         End
         Begin VB.TextBox txtgrandtotal 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   6120
            TabIndex        =   35
            Top             =   360
            Width           =   2775
         End
         Begin VB.Label Label12 
            Caption         =   "Grand Total"
            Height          =   255
            Left            =   5040
            TabIndex        =   34
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame Frame6 
         Height          =   2175
         Left            =   120
         TabIndex        =   30
         Top             =   5880
         Width           =   9735
         Begin MSFlexGridLib.MSFlexGrid grddetails 
            Height          =   1815
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Width           =   8535
            _ExtentX        =   15055
            _ExtentY        =   3201
            _Version        =   393216
         End
      End
      Begin VB.Frame Frame5 
         Height          =   855
         Left            =   120
         TabIndex        =   26
         Top             =   4920
         Width           =   9735
         Begin VB.CommandButton cmdcancel 
            Caption         =   "Cancel"
            Height          =   495
            Left            =   5520
            TabIndex        =   29
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmddelete 
            Caption         =   "Delete"
            Height          =   495
            Left            =   4260
            TabIndex        =   28
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdadd 
            Caption         =   "Add"
            Height          =   495
            Left            =   3000
            TabIndex        =   27
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame4 
         Height          =   975
         Left            =   120
         TabIndex        =   19
         Top             =   3840
         Width           =   9735
         Begin VB.TextBox txtrate 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6960
            TabIndex        =   25
            Top             =   360
            Width           =   2175
         End
         Begin VB.TextBox txtisbn 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4200
            TabIndex        =   23
            Top             =   360
            Width           =   2175
         End
         Begin VB.ComboBox cmbbook 
            Height          =   315
            Left            =   600
            TabIndex        =   21
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label Label11 
            Caption         =   "Rate"
            Height          =   255
            Left            =   6480
            TabIndex        =   24
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label10 
            Caption         =   "ISBN Number"
            Height          =   255
            Left            =   3000
            TabIndex        =   22
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label9 
            Caption         =   "Book"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1935
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   9735
         Begin VB.CommandButton Command1 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   7800
            TabIndex        =   37
            Top             =   1440
            Width           =   975
         End
         Begin VB.CommandButton cmdview 
            Caption         =   "View"
            Height          =   375
            Left            =   6720
            TabIndex        =   18
            Top             =   1440
            Width           =   975
         End
         Begin VB.ComboBox cmbsub 
            Height          =   315
            Left            =   5280
            TabIndex        =   17
            Top             =   1080
            Width           =   2535
         End
         Begin VB.ComboBox cmbpub 
            Height          =   315
            Left            =   1200
            TabIndex        =   15
            Top             =   1080
            Width           =   2535
         End
         Begin VB.ComboBox cmbatr 
            Height          =   315
            Left            =   5280
            TabIndex        =   13
            Top             =   360
            Width           =   2535
         End
         Begin VB.ComboBox cmbcat 
            Height          =   315
            Left            =   1200
            TabIndex        =   11
            Top             =   360
            Width           =   2535
         End
         Begin VB.Label Label8 
            Caption         =   "Subject"
            Height          =   255
            Left            =   4200
            TabIndex        =   16
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "Publisher"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label6 
            Caption         =   "Author"
            Height          =   255
            Left            =   4200
            TabIndex        =   12
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "Category"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1575
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   9735
         Begin VB.ComboBox cmbsup 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5400
            TabIndex        =   32
            Top             =   360
            Width           =   2535
         End
         Begin VB.ComboBox cmbtype 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5400
            TabIndex        =   8
            Text            =   "--select--"
            Top             =   1080
            Width           =   2535
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Left            =   1440
            TabIndex        =   6
            Top             =   1080
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   556
            _Version        =   393216
            Format          =   93716481
            CurrentDate     =   38760
         End
         Begin VB.TextBox txtinvoiceno 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1440
            TabIndex        =   3
            Top             =   360
            Width           =   2535
         End
         Begin VB.Label Label4 
            Caption         =   "Pay Type"
            Height          =   255
            Left            =   4440
            TabIndex        =   7
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "Date"
            Height          =   255
            Left            =   240
            TabIndex        =   5
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label2 
            Caption         =   "Supplier"
            Height          =   255
            Left            =   4440
            TabIndex        =   4
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Invoice No"
            Height          =   255
            Left            =   240
            TabIndex        =   2
            Top             =   360
            Width           =   975
         End
      End
   End
End
Attribute VB_Name = "purchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim temp As New ADODB.Recordset
Private Sub cmdadd_Click()
If addvalidation Then


Dim i, j As Integer

i = grddetails.Rows

If j <> i - 1 Then
 j = i - 1
End If

grddetails.TextMatrix(j, 0) = CboData(cmbbook)
grddetails.TextMatrix(j, 1) = cmbbook.Text
grddetails.TextMatrix(j, 2) = 1
grddetails.TextMatrix(j, 3) = txtrate.Text
grddetails.TextMatrix(j, 4) = 1 * Val(txtrate.Text)
grddetails.TextMatrix(j, 5) = txtisbn.Text

grddetails.Rows = grddetails.Rows + 1


txtisbn.Text = ""
txtisbn.SetFocus

txtgrandtotal.Text = Val(txtgrandtotal.Text) + Val(grddetails.TextMatrix(j, 4))

End If


End Sub

Private Sub cmdcancel_Click()
txtrate.Text = ""
txtisbn.Text = ""
cmbbook.Text = "--select--"
End Sub

Private Sub cmddelete_Click()

Dim i As Integer
i = grddetails.row
txtgrandtotal.Text = Val(txtgrandtotal.Text) - Val(grddetails.TextMatrix(i, 4))
grddetails.RemoveItem i

End Sub

Private Sub cmdSave_Click()

If headervalidation Then

If rsPurchaseHead.State = 1 Then rsPurchaseHead.Close
rsPurchaseHead.Open "select * from tbl_purchasehead", con, adOpenKeyset, adLockOptimistic
rsPurchaseHead.AddNew
   rsPurchaseHead.Fields("invoice_no") = txtinvoiceno.Text
   rsPurchaseHead.Fields("ph_date") = DTPicker1.Value
   rsPurchaseHead.Fields("sup_id") = CboData(cmbsup)
   rsPurchaseHead.Fields("ph_paytype") = cmbtype.Text
   rsPurchaseHead.Fields("ph_grand_total") = txtgrandtotal.Text
rsPurchaseHead.Update
rsPurchaseHead.Close






If temp.State = 1 Then temp.Close
temp.Open "select max(ph_id) as pid from tbl_purchasehead", con, adOpenKeyset, adLockOptimistic
Dim purid As Integer
purid = temp.Fields("pid")

Dim n As Integer
n = grddetails.Rows

For i = 1 To n - 2
 
 If rsPurchaseDetails.State = 1 Then rsPurchaseDetails.Close
 rsPurchaseDetails.Open "select * from tbl_purchasedetails", con, adOpenKeyset, adLockOptimistic
 rsPurchaseDetails.AddNew
   rsPurchaseDetails.Fields("book_id") = Val(grddetails.TextMatrix(i, 0))
   rsPurchaseDetails.Fields("book_qty") = Val(grddetails.TextMatrix(i, 2))
   rsPurchaseDetails.Fields("book_unitprice") = Val(grddetails.TextMatrix(i, 3))
   rsPurchaseDetails.Fields("book_amount") = Val(grddetails.TextMatrix(i, 4))
   rsPurchaseDetails.Fields("book_isbnno") = Val(grddetails.TextMatrix(i, 5))
   rsPurchaseDetails.Fields("ph_id") = purid
   
rsPurchaseDetails.Update
rsPurchaseDetails.Close


If rsbookdetails.State = 1 Then rsbookdetails.Close
 rsbookdetails.Open "select * from tbl_bookdetails", con, adOpenKeyset, adLockOptimistic
  rsbookdetails.AddNew
  rsbookdetails.Fields("book_id") = Val(grddetails.TextMatrix(i, 0))
  rsbookdetails.Fields("ISBNno") = Val(grddetails.TextMatrix(i, 5))
  rsbookdetails.Fields("bd_status") = 0
  rsbookdetails.Update
  rsbookdetails.Close




Next i


cmbatr.Text = "--select--"
cmbbook.Text = "--select--"
cmbcat.Text = "--select--"
cmbpub.Text = "--select--"
cmbsub.Text = "--select--"
cmbsup.Text = "--select--"
cmbtype.Text = "--select--"

txtgrandtotal.Text = ""
txtinvoiceno.Text = ""
txtisbn.Text = ""
txtrate.Text = ""



 
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

Fillcombo "tbl_supplier", cmbsup, "sup_name", "sup_id"
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

Fillcombo "tbl_supplier", cmbsup, "sup_name", "sup_id"
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
   MsgBox "Enter Invoice Number", vbInformation, App.Title
   txtinvoiceno.SetFocus
   headervalidation = False
   Exit Function
End If

If Trim(cmbsup.Text) = "--select--" Then
  MsgBox "select supplier", vbInformation, App.Title
  cmbsup.SetFocus
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

If Trim(txtisbn.Text) = "" Then
  MsgBox "Enter ISBN No", vbInformation, App.Title
  txtisbn.SetFocus
  addvalidation = False
  Exit Function
End If


If Trim(txtrate.Text) = "" Then
  MsgBox "Enter Rate", vbInformation, App.Title
  txtrate.SetFocus
  addvalidation = False
  Exit Function
End If

addvalidation = True

End Function

