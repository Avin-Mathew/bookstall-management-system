VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Bkdetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BOOK DETAILS"
   ClientHeight    =   6360
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   8850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   8850
   Begin VB.Frame Frame1 
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8775
      Begin VB.Frame Frame4 
         Height          =   855
         Left            =   120
         TabIndex        =   17
         Top             =   3240
         Width           =   8415
         Begin VB.CommandButton cmdcancel 
            Caption         =   "&CANCEL"
            Height          =   495
            Left            =   6480
            TabIndex        =   22
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmddelete 
            Caption         =   "&DELETE"
            Height          =   495
            Left            =   4920
            TabIndex        =   21
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdedit 
            Caption         =   "&EDIT"
            Height          =   495
            Left            =   3480
            TabIndex        =   20
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdupdate 
            Caption         =   "&UPDATE"
            Height          =   495
            Left            =   2040
            TabIndex        =   19
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdadd 
            Appearance      =   0  'Flat
            Caption         =   "&ADD"
            Height          =   495
            Left            =   360
            TabIndex        =   18
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2895
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   8415
         Begin VB.ComboBox cmbsub 
            Height          =   315
            Left            =   5640
            TabIndex        =   26
            Top             =   2400
            Width           =   2535
         End
         Begin VB.TextBox txtnoc 
            Appearance      =   0  'Flat
            Height          =   405
            Left            =   1560
            TabIndex        =   24
            Top             =   2160
            Width           =   2775
         End
         Begin VB.ComboBox cmblanguage 
            Height          =   315
            Left            =   5640
            TabIndex        =   16
            Top             =   1860
            Width           =   2535
         End
         Begin VB.ComboBox cmbauthor 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5640
            TabIndex        =   15
            Top             =   1320
            Width           =   2535
         End
         Begin VB.ComboBox cmbpublisher 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5640
            TabIndex        =   14
            Top             =   780
            Width           =   2535
         End
         Begin VB.ComboBox cmbcategory 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "Bkdetails.frx":0000
            Left            =   5640
            List            =   "Bkdetails.frx":0002
            TabIndex        =   13
            Top             =   240
            Width           =   2535
         End
         Begin VB.TextBox txtqty 
            Appearance      =   0  'Flat
            Height          =   405
            Left            =   1560
            TabIndex        =   12
            Top             =   1520
            Width           =   2775
         End
         Begin VB.TextBox txtprice 
            Appearance      =   0  'Flat
            Height          =   405
            Left            =   1560
            TabIndex        =   11
            Top             =   880
            Width           =   2775
         End
         Begin VB.TextBox txtname 
            Appearance      =   0  'Flat
            Height          =   405
            Left            =   1560
            TabIndex        =   10
            Top             =   240
            Width           =   2775
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Language"
            Height          =   195
            Left            =   4680
            TabIndex        =   27
            Top             =   1920
            Width           =   720
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Subject"
            Height          =   195
            Left            =   4680
            TabIndex        =   25
            Top             =   2400
            Width           =   540
         End
         Begin VB.Label Label1 
            Caption         =   "No of Copies"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Publisher"
            Height          =   195
            Left            =   4680
            TabIndex        =   9
            Top             =   840
            Width           =   645
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Opening Quantity"
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   1560
            Width           =   1230
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Category"
            Height          =   195
            Left            =   4680
            TabIndex        =   7
            Top             =   240
            Width           =   630
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Price"
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   960
            Width           =   360
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   4440
            TabIndex        =   5
            Top             =   2040
            Width           =   45
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Author"
            Height          =   195
            Left            =   4680
            TabIndex        =   4
            Top             =   1320
            Width           =   465
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   " Name"
            Height          =   195
            Left            =   120
            TabIndex        =   3
            Top             =   360
            Width           =   465
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grddetails 
         Height          =   1935
         Left            =   120
         TabIndex        =   1
         Top             =   4200
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   3413
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "Bkdetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub cmdcancel_Click()
txtname.Text = ""
txtprice.Text = ""
txtqty.Text = ""
txtnoc.Text = ""
cmbpublisher.Text = "--select--"
cmblanguage.Text = "--select--"
cmbauthor.Text = "--select--"
cmbsub.Text = "--select--"
cmbcategory.Text = "--select--"
Frame2.Enabled = False
txtname.Tag = ""
cmdadd.Enabled = True
cmdupdate.Enabled = False
cmdedit.Enabled = False
cmddelete.Enabled = False
cmdcancel.Enabled = True



End Sub

Private Sub cmddelete_Click()
If rsbook.State = 1 Then rsbook.Close
rsbook.Open "select * from tbl_bookmaster where book_id=" & txtname.Tag & "", con, adOpenKeyset, adLockOptimistic
rsbook.Delete
rsbook.Update
rsbook.Close

fillgrid

txtname.Text = ""
txtprice.Text = ""
txtqty.Text = ""
txtnoc.Text = ""
cmbpublisher.Text = "--select--"
cmblanguage.Text = "--select--"
cmbauthor.Text = "--select--"
cmbsub.Text = "--select--"
cmbcategory.Text = "--select--"
Frame2.Enabled = False
txtname.Tag = ""
cmdadd.Enabled = True
cmdupdate.Enabled = False
cmdedit.Enabled = False
cmddelete.Enabled = False
cmdcancel.Enabled = True
End Sub

Private Sub cmdedit_Click()
Frame2.Enabled = True
cmdupdate.Enabled = True
cmddelete.Enabled = False
cmdedit.Enabled = False
txtname.SetFocus
End Sub

Private Sub Form_Load()

Frame2.Enabled = False
cmdadd.Enabled = True
cmdupdate.Enabled = False
cmddelete.Enabled = False
cmdedit.Enabled = False
cmdcancel.Enabled = True

Fillcombo "tbl_category", cmbcategory, "c_name", "c_id"
Fillcombo "tbl_publisher", cmbpublisher, "pub_name", "pub_id"
Fillcombo "tbl_language", cmblanguage, "l_name", "l_id"
Fillcombo "tbl_author", cmbauthor, "atr_name", "atr_id"
Fillcombo "tbl_subject", cmbsub, "sub_name", "sub_id"

fillgrid



End Sub

Private Sub grddetails_Click()
If rsbook.State = 1 Then rsbook.Close
rsbook.Open "select * from tbl_bookmaster where book_id=" & grddetails.TextMatrix(grddetails.row, 0) & "", con, adOpenKeyset, adLockOptimistic
If rsbook.RecordCount > 0 Then
txtname.Text = rsbook.Fields("book_name")
txtprice.Text = rsbook.Fields("book_price")
txtnoc.Text = rsbook.Fields("book_nocopies")
txtqty.Text = rsbook.Fields("book_openingqty")
txtname.Tag = rsbook.Fields("book_id")
selectcombo rsbook.Fields("c_id"), cmbcategory
selectcombo rsbook.Fields("pub_id"), cmbpublisher
selectcombo rsbook.Fields("l_id"), cmblanguage
selectcombo rsbook.Fields("atr_id"), cmbauthor
selectcombo rsbook.Fields("sub_id"), cmbsub
End If
rsbook.Close
cmdedit.Enabled = True
cmddelete.Enabled = True
cmdadd.Enabled = False
End Sub



Private Sub cmdadd_Click()
cmdupdate.Enabled = True
cmdcancel.Enabled = True
Frame2.Enabled = True
txtname.SetFocus
cmdadd.Enabled = False
cmdedit.Enabled = False
cmddelete.Enabled = False
End Sub

Private Sub cmdupdate_Click()
If vallidation Then
If txtname.Tag = "" Then
If rsbook.State = 1 Then rsbook.Close
rsbook.Open "select * from tbl_bookmaster", con, adOpenKeyset, adLockOptimistic
rsbook.AddNew
rsbook.Fields("book_name") = txtname.Text
rsbook.Fields("book_price") = txtprice.Text
rsbook.Fields("book_nocopies") = txtnoc.Text
rsbook.Fields("book_openingqty") = txtqty.Text
rsbook.Fields("c_id") = CboData(cmbcategory)
rsbook.Fields("pub_id") = CboData(cmbpublisher)
rsbook.Fields("l_id") = CboData(cmblanguage)
rsbook.Fields("atr_id") = CboData(cmbauthor)
rsbook.Fields("sub_id") = CboData(cmbsub)
rsbook.Update
rsbook.Close

txtname.Text = ""
txtprice.Text = ""
txtqty.Text = ""
txtnoc.Text = ""
cmbpublisher.Text = "--select--"
cmblanguage.Text = "--select--"
cmbauthor.Text = "--select--"
cmbsub.Text = "--select--"
cmbcategory.Text = "--select--"
Frame2.Enabled = False

cmdadd.Enabled = True
cmdupdate.Enabled = False
cmdedit.Enabled = False
cmddelete.Enabled = False
cmdcancel.Enabled = True



Else
If rsbook.State = 1 Then rsbook.Close
rsbook.Open "select * from tbl_bookmaster where book_id=" & txtname.Tag & "", con, adOpenKeyset, adLockOptimistic
rsbook.Fields("book_name") = txtname.Text
rsbook.Fields("book_price") = txtprice.Text
rsbook.Fields("book_nocopies") = txtnoc.Text
rsbook.Fields("book_openingqty") = txtqty.Text
rsbook.Fields("c_id") = CboData(cmbcategory)
rsbook.Fields("pub_id") = CboData(cmbpublisher)
rsbook.Fields("l_id") = CboData(cmblanguage)
rsbook.Fields("atr_id") = CboData(cmbauthor)
rsbook.Fields("sub_id") = CboData(cmbsub)
rsbook.Update
rsbook.Close

txtname.Text = ""
txtprice.Text = ""
txtqty.Text = ""
txtnoc.Text = ""
cmbpublisher.Text = "--select--"
cmblanguage.Text = "--select--"
cmbauthor.Text = "--select--"
cmbsub.Text = "--select--"
cmbcategory.Text = "--select--"
Frame2.Enabled = False
txtname.Tag = ""
cmdadd.Enabled = True
cmdupdate.Enabled = False
cmdedit.Enabled = False
cmddelete.Enabled = False
cmdcancel.Enabled = True


End If
End If
fillgrid





End Sub

Public Function fillgrid()
grddetails.Rows = 1
grddetails.Cols = 10


grddetails.ColWidth(0) = 0
grddetails.ColWidth(1) = 1700
grddetails.ColWidth(2) = 1700
grddetails.ColWidth(3) = 1700
grddetails.ColWidth(4) = 1700
grddetails.ColWidth(5) = 1700
grddetails.ColWidth(6) = 1700
grddetails.ColWidth(7) = 1700
grddetails.ColWidth(8) = 1700
grddetails.ColWidth(9) = 1700



grddetails.TextMatrix(0, 1) = "Book Name"
grddetails.TextMatrix(0, 2) = "Price"
grddetails.TextMatrix(0, 3) = "Copies"
grddetails.TextMatrix(0, 4) = "Opening Stock"
grddetails.TextMatrix(0, 5) = "Catgeory"
grddetails.TextMatrix(0, 6) = "publisher"
grddetails.TextMatrix(0, 7) = "language"
grddetails.TextMatrix(0, 8) = "author"
grddetails.TextMatrix(0, 9) = "subject"






If rsbook.State = 1 Then rsbook.Close
rsbook.Open "select * from tbl_bookmaster b,tbl_category c,tbl_publisher p,tbl_author a,tbl_language l ,tbl_subject s where c.c_id=b.c_id and  p.pub_id=b.pub_id and a.atr_id=b.atr_id and l.l_id=b.l_id and s.sub_id=b.sub_id ", con, adOpenKeyset, adLockOptimistic
If rsbook.RecordCount > 0 Then
rsbook.MoveFirst
While Not rsbook.EOF
grddetails.Rows = grddetails.Rows + 1
grddetails.TextMatrix(grddetails.Rows - 1, 0) = rsbook.Fields("book_id")
grddetails.TextMatrix(grddetails.Rows - 1, 1) = rsbook.Fields("book_name")
grddetails.TextMatrix(grddetails.Rows - 1, 2) = rsbook.Fields("book_price")
grddetails.TextMatrix(grddetails.Rows - 1, 3) = rsbook.Fields("book_nocopies")
grddetails.TextMatrix(grddetails.Rows - 1, 4) = rsbook.Fields("book_openingqty")
grddetails.TextMatrix(grddetails.Rows - 1, 5) = rsbook.Fields("c_name")
grddetails.TextMatrix(grddetails.Rows - 1, 6) = rsbook.Fields("pub_name")
grddetails.TextMatrix(grddetails.Rows - 1, 7) = rsbook.Fields("l_name")
grddetails.TextMatrix(grddetails.Rows - 1, 8) = rsbook.Fields("atr_name")
grddetails.TextMatrix(grddetails.Rows - 1, 9) = rsbook.Fields("sub_name")


rsbook.MoveNext
Wend
End If
rsbook.Close
End Function


Public Function vallidation() As Boolean
If Trim(txtname.Text) = "" Then
   MsgBox "Enter a name", vbInformation, App.Title
   txtname.SetFocus
   vallidation = False
Exit Function
End If

If Trim(txtprice.Text) = "" Then
  MsgBox "Enter the price", vbInformation, App.Title
  txtprice.SetFocus
  vallidation = False
  Exit Function
End If

If Trim(txtqty.Text) = "" Then
   MsgBox "Enter the opening quantity", vbInformation, App.Title
   txtqty.SetFocus
   vallidation = False
   Exit Function
End If

If Trim(txtnoc.Text) = "" Then
   MsgBox "Enter the no of copies", vbInformation, App.Title
   txtnoc.SetFocus
   vallidation = False
   Exit Function
End If
If Trim(cmbcategory.Text) = "--select--" Then
   MsgBox "Select category ", vbInformation, App.Title
   cmbcategory.SetFocus
   vallidation = False
   Exit Function
End If
If Trim(cmbpublisher.Text) = "--select--" Then
   MsgBox "Select Publisher", vbInformation, App.Title
   cmbpublisher.SetFocus
   vallidation = False
   Exit Function
End If
If Trim(cmbauthor.Text) = "--select--" Then
   MsgBox "Select Author ", vbInformation, App.Title
   cmbauthor.SetFocus
   vallidation = False
   Exit Function
End If
If Trim(cmblanguage.Text) = "--select--" Then
   MsgBox "Select Language", vbInformation, App.Title
   cmblanguage.SetFocus
   vallidation = False
   Exit Function
End If
If Trim(cmbsub.Text) = "--select--" Then
   MsgBox "Select Subject", vbInformation, App.Title
   cmbsub.SetFocus
   vallidation = False
   Exit Function
End If

vallidation = True

End Function
