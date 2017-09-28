VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmbookdetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Book Details"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   8430
   Begin VB.Frame Frame1 
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8295
      Begin MSFlexGridLib.MSFlexGrid grddetails 
         Height          =   2415
         Left            =   120
         TabIndex        =   16
         Top             =   3360
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   4260
         _Version        =   393216
      End
      Begin VB.Frame Frame3 
         Height          =   855
         Left            =   120
         TabIndex        =   10
         Top             =   2400
         Width           =   8055
         Begin VB.CommandButton cmdcancel 
            Caption         =   "&CANCEL"
            Height          =   375
            Left            =   6360
            TabIndex        =   15
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmddelete 
            Caption         =   "&DELETE"
            Height          =   375
            Left            =   5160
            TabIndex        =   14
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdedit 
            Caption         =   "&EDIT"
            Height          =   375
            Left            =   3960
            TabIndex        =   13
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdupdate 
            Caption         =   "&UPDATE"
            Height          =   375
            Left            =   2640
            TabIndex        =   12
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdadd 
            Caption         =   "&ADD"
            Height          =   375
            Left            =   1200
            TabIndex        =   11
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2055
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   8055
         Begin VB.TextBox txtisbn 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   5280
            TabIndex        =   9
            Top             =   1560
            Width           =   2415
         End
         Begin VB.ComboBox cmbbook 
            Height          =   315
            Left            =   960
            TabIndex        =   3
            Top             =   360
            Width           =   2415
         End
         Begin VB.Label lblsub 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   5280
            TabIndex        =   20
            Top             =   960
            Width           =   2415
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Subject"
            Height          =   195
            Left            =   4440
            TabIndex        =   19
            Top             =   960
            Width           =   540
         End
         Begin VB.Label ISN 
            AutoSize        =   -1  'True
            Caption         =   "ISBN no"
            Height          =   195
            Left            =   4440
            TabIndex        =   18
            Top             =   1560
            Width           =   600
         End
         Begin VB.Label lblcopy 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   5280
            TabIndex        =   17
            Top             =   360
            Width           =   2415
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Copies"
            Height          =   195
            Left            =   4440
            TabIndex        =   8
            Top             =   360
            Width           =   480
         End
         Begin VB.Label lblpub 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1080
            TabIndex        =   7
            Top             =   1560
            Width           =   2295
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Publisher"
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   1560
            Width           =   645
         End
         Begin VB.Label lblauthor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1080
            TabIndex        =   5
            Top             =   960
            Width           =   2295
         End
         Begin VB.Label Author 
            AutoSize        =   -1  'True
            Caption         =   "Author"
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Top             =   960
            Width           =   465
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Book"
            Height          =   195
            Left            =   240
            TabIndex        =   2
            Top             =   360
            Width           =   375
         End
      End
   End
End
Attribute VB_Name = "frmbookdetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub cmbbook_Click()
If rsbookdetails.State = 1 Then rsbookdetails.Close
rsbookdetails.Open "select * from tbl_bookmaster b,tbl_publisher p,tbl_author a,tbl_subject sub where b.sub_id=sub.sub_id and  p.pub_id=b.pub_id and a.atr_id=b.atr_id and b.book_id=" & CboData(cmbbook) & " ", con, adOpenKeyset, adLockOptimistic
If rsbookdetails.RecordCount > 0 Then
 lblauthor.Caption = rsbookdetails.Fields("atr_name")
 lblcopy.Caption = rsbookdetails.Fields("book_nocopies")
 lblpub.Caption = rsbookdetails.Fields("pub_name")
  lblsub.Caption = rsbookdetails.Fields("sub_name")
 
End If

rsbookdetails.Close

fillgrid
End Sub

Private Sub cmdcancel_Click()
cmbbook.Text = "--select--"
lblauthor.Caption = ""
lblcopy.Caption = ""
lblpub.Caption = ""
lblsub.Caption = ""
txtisbn.Text = ""
Frame2.Enabled = False

cmdadd.Enabled = True
cmdupdate.Enabled = False
cmddelete.Enabled = False
cmdedit.Enabled = False
cmdcancel.Enabled = True
txtisbn.Tag = ""


End Sub

Private Sub cmddelete_Click()

If rsbookdetails.State = 1 Then rsbookdetails.Close
rsbookdetails.Open "select * from tbl_bookdetails where bd_id=" & txtisbn.Tag & "", con, adOpenKeyset, adLockOptimistic
rsbookdetails.Delete
rsbookdetails.Update
rsbookdetails.Close
txtisbn.Tag = ""
fillgrid
cmddelete.Enabled = False
cmdadd.Enabled = True
cmdedit.Enabled = False
cmdupdate.Enabled = False
cmdcancel.Enabled = True
      cmbbook.Text = "--select--"
      lblauthor.Caption = ""
      lblcopy.Caption = ""
      lblpub.Caption = ""
      lblsub.Caption = ""
      txtisbn.Text = ""
      Frame2.Enabled = False
End Sub

Private Sub cmdedit_Click()
Frame2.Enabled = True
cmdupdate.Enabled = True
cmddelete.Enabled = False
cmdedit.Enabled = False
cmbbook.SetFocus
End Sub

Private Sub Form_Load()

Frame2.Enabled = False
cmdadd.Enabled = True
cmdupdate.Enabled = False
cmddelete.Enabled = False
cmdedit.Enabled = False
cmdcancel.Enabled = True

Fillcombo "tbl_bookmaster", cmbbook, "book_name", "book_id"

fillgrid

End Sub

Private Sub grddetails_Click()

If rsbookdetails.State = 1 Then rsbookdetails.Close
rsbookdetails.Open "select * from tbl_bookdetails where bd_id=" & grddetails.TextMatrix(grddetails.row, 0) & "", con, adOpenKeyset, adLockOptimistic
If rsbookdetails.RecordCount > 0 Then
txtisbn.Text = rsbookdetails.Fields("ISBNno")
txtisbn.Tag = rsbookdetails.Fields("bd_id")

selectcombo rsbookdetails.Fields("book_id"), cmbbook

End If


cmdedit.Enabled = True
cmddelete.Enabled = True
cmdadd.Enabled = False
End Sub



Private Sub cmdadd_Click()
cmdupdate.Enabled = True
cmdcancel.Enabled = True
Frame2.Enabled = True
cmbbook.SetFocus
cmdadd.Enabled = False
End Sub

Private Sub cmdupdate_Click()


If validation Then
If txtisbn.Tag = "" Then
   If rsbookdetails.State = 1 Then rsbookdetails.Close
   rsbookdetails.Open "select * from tbl_bookdetails", con, adOpenKeyset, adLockOptimistic
   rsbookdetails.AddNew
   rsbookdetails.Fields("book_id") = CboData(cmbbook)
   rsbookdetails.Fields("ISBNno") = txtisbn.Text
   rsbookdetails.Fields("bd_status") = 0
   rsbookdetails.Update
   rsbookdetails.Close
   
      cmbbook.Text = "--select--"
      lblauthor.Caption = ""
      lblcopy.Caption = ""
      lblpub.Caption = ""
      lblsub.Caption = ""
      txtisbn.Text = ""
      Frame2.Enabled = False

      cmdadd.Enabled = True
      cmdupdate.Enabled = False
      cmddelete.Enabled = False
       cmdedit.Enabled = False
      cmdcancel.Enabled = True
      
Else
  If rsbookdetails.State = 1 Then rsbookdetails.Close
   rsbookdetails.Open "select * from tbl_bookdetails where bd_id=" & txtisbn.Tag & "", con, adOpenKeyset, adLockOptimistic
   rsbookdetails.Fields("book_id") = CboData(cmbbook)
   rsbookdetails.Fields("ISBNno") = txtisbn.Text
   rsbookdetails.Update
   rsbookdetails.Close
   
   
      cmbbook.Text = "--select--"
      lblauthor.Caption = ""
      lblcopy.Caption = ""
      lblpub.Caption = ""
      lblsub.Caption = ""
      txtisbn.Text = ""
      Frame2.Enabled = False

      cmdadd.Enabled = True
      cmdupdate.Enabled = False
      cmddelete.Enabled = False
       cmdedit.Enabled = False
      cmdcancel.Enabled = True
        txtisbn.Tag = ""
   
End If

End If

fillgrid




End Sub

Public Function fillgrid()
grddetails.Rows = 1
grddetails.Cols = 3


grddetails.ColWidth(0) = 0
grddetails.ColWidth(1) = 1700
grddetails.ColWidth(2) = 1700




grddetails.TextMatrix(0, 1) = "Book Name"
grddetails.TextMatrix(0, 2) = "ISBN No"




If rsbookdetails.State = 1 Then rsbookdetails.Close
rsbookdetails.Open "select * from tbl_bookmaster b,tbl_bookdetails bd where b.book_id=bd.book_id ", con, adOpenKeyset, adLockOptimistic
If rsbookdetails.RecordCount > 0 Then
rsbookdetails.MoveFirst
While Not rsbookdetails.EOF
grddetails.Rows = grddetails.Rows + 1
grddetails.TextMatrix(grddetails.Rows - 1, 0) = rsbookdetails.Fields("bd_id")
grddetails.TextMatrix(grddetails.Rows - 1, 1) = rsbookdetails.Fields("book_name")
grddetails.TextMatrix(grddetails.Rows - 1, 2) = rsbookdetails.Fields("ISBNno")


rsbookdetails.MoveNext
Wend
End If
rsbookdetails.Close
End Function



Public Function validation() As Boolean
If Trim(cmbbook.Text) = "--select--" Then
   MsgBox "Enter the book", vbInformation, App.Title
   cmbbook.SetFocus
   validation = False
Exit Function
End If

If Trim(txtisbn.Text) = "" Then
   MsgBox "Enter the ISBN no.", vbInformation, App.Title
   txtisbn.SetFocus
   validation = False
   Exit Function
End If

validation = True
End Function



Private Sub txtisbn_KeyPress(KeyAscii As Integer)
If txtisbn.Text <> "" Then
If IsNumeric(txtisbn.Text) = False Then
MsgBox "Numbers only allowed"
txtisbn.Text = ""
End If
End If
End Sub
