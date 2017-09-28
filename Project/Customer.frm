VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Customer 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CUSTOMER DETAILS"
   ClientHeight    =   5700
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   9915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   9915
   Begin VB.Frame Frame1 
      Height          =   5535
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   9855
      Begin VB.Frame Frame3 
         Height          =   975
         Left            =   120
         TabIndex        =   9
         Top             =   2160
         Width           =   9615
         Begin VB.CommandButton cmdcancel 
            Caption         =   "&CANCEL"
            Height          =   495
            Left            =   6840
            TabIndex        =   14
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmddelete 
            Caption         =   "&DELETE"
            Height          =   495
            Left            =   5520
            TabIndex        =   13
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdedit 
            Caption         =   "&EDIT"
            Height          =   495
            Left            =   4200
            TabIndex        =   12
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdupdate 
            Caption         =   "&UPDATE"
            Height          =   495
            Left            =   2880
            TabIndex        =   11
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdadd 
            Caption         =   "&ADD"
            Height          =   495
            Left            =   1560
            TabIndex        =   10
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2055
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   9615
         Begin VB.TextBox txtctno 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   5880
            TabIndex        =   7
            Top             =   360
            Width           =   2775
         End
         Begin VB.TextBox txtaddress 
            Appearance      =   0  'Flat
            Height          =   855
            Left            =   1080
            MultiLine       =   -1  'True
            TabIndex        =   5
            Top             =   960
            Width           =   2775
         End
         Begin VB.TextBox txtname 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1080
            TabIndex        =   4
            Top             =   360
            Width           =   2775
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Contact no"
            Height          =   195
            Left            =   4800
            TabIndex        =   6
            Top             =   360
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Address"
            Height          =   195
            Left            =   240
            TabIndex        =   3
            Top             =   1080
            Width           =   570
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Name"
            Height          =   195
            Left            =   240
            TabIndex        =   2
            Top             =   360
            Width           =   420
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grddetails 
         Height          =   1815
         Left            =   120
         TabIndex        =   8
         Top             =   3240
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   3201
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "Customer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdadd_Click()
cmdupdate.Enabled = True
cmdcancel.Enabled = True
cmdedit.Enabled = False
cmddelete.Enabled = False
cmdadd.Enabled = False
Frame2.Enabled = True
txtname.SetFocus
End Sub

Private Sub cmdcancel_Click()
txtname.Text = ""
txtctno.Text = ""
txtaddress.Text = ""
Frame2.Enabled = False
txtname.Tag = ""
cmdadd.Enabled = True
cmdupdate.Enabled = False
cmdedit.Enabled = False
cmddelete.Enabled = False
cmdcancel.Enabled = True
End Sub

Private Sub cmddelete_Click()

If rscst.State = 1 Then rscst.Close
rscst.Open "select * from tbl_customerdetails where cst_id=" & txtname.Tag & "", con, adOpenKeyset, adLockOptimistic
rscst.Delete
rscst.Update
rscst.Close

fillgrid

txtname.Text = ""
txtctno.Text = ""
txtaddress.Text = ""
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
cmdadd.Enabled = False
txtname.SetFocus

End Sub

Private Sub cmdupdate_Click()
If validation Then
If txtname.Tag = "" Then
If rscst.State = 1 Then rscst.Close
rscst.Open "select * from tbl_customerdetails", con, adOpenKeyset, adLockOptimistic
rscst.AddNew
rscst.Fields("cst_name") = txtname.Text
rscst.Fields("cst_address") = txtaddress.Text
rscst.Fields("cst_contactno") = txtctno.Text
rscst.Update
rscst.Close

txtname.Text = ""
txtctno.Text = ""
txtaddress.Text = ""
Frame2.Enabled = False
cmdadd.Enabled = True
cmdupdate.Enabled = False
cmdedit.Enabled = False
cmddelete.Enabled = False
cmdcancel.Enabled = True


Else
If rscst.State = 1 Then rscst.Close
rscst.Open "select * from tbl_customerdetails where cst_id=" & txtname.Tag & "", con, adOpenKeyset, adLockOptimistic
rscst.Fields("cst_name") = txtname.Text
rscst.Fields("cst_address") = txtaddress.Text
rscst.Fields("cst_contactno") = txtctno.Text
rscst.Update
rscst.Close


txtname.Text = ""
txtctno.Text = ""
txtaddress.Text = ""
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

Private Sub Form_Load()

cmdadd.Enabled = True
cmdupdate.Enabled = False
cmdedit.Enabled = False
cmddelete.Enabled = False
cmdcancel.Enabled = True
Frame2.Enabled = False
fillgrid
End Sub

Private Sub grddetails_Click()
If rscst.State = 1 Then rscst.Close
rscst.Open "select * from tbl_customerdetails where cst_id=" & grddetails.TextMatrix(grddetails.row, 0) & "", con, adOpenKeyset, adLockOptimistic
If rscst.RecordCount > 0 Then
txtname.Text = rscst.Fields("cst_name")
txtaddress.Text = rscst.Fields("cst_address")
txtctno.Text = rscst.Fields("cst_contactno")
txtname.Tag = rscst.Fields("cst_id")

End If
cmdedit.Enabled = True
cmddelete.Enabled = True
cmdadd.Enabled = False
End Sub

Public Function fillgrid()
grddetails.Rows = 1
grddetails.Cols = 4


grddetails.ColWidth(0) = 0
grddetails.ColWidth(1) = 1700
grddetails.ColWidth(2) = 1700
grddetails.ColWidth(3) = 1700

grddetails.TextMatrix(0, 1) = "Customer"
grddetails.TextMatrix(0, 2) = "Address"
grddetails.TextMatrix(0, 3) = "Contact"

If rscst.State = 1 Then rscst.Close
rscst.Open "select * from tbl_customerdetails", con, adOpenKeyset, adLockOptimistic
If rscst.RecordCount > 0 Then
rscst.MoveFirst
While Not rscst.EOF
grddetails.Rows = grddetails.Rows + 1
grddetails.TextMatrix(grddetails.Rows - 1, 0) = rscst.Fields("cst_id")
grddetails.TextMatrix(grddetails.Rows - 1, 1) = rscst.Fields("cst_name")
grddetails.TextMatrix(grddetails.Rows - 1, 2) = rscst.Fields("cst_address")
grddetails.TextMatrix(grddetails.Rows - 1, 3) = rscst.Fields("cst_contactno")
rscst.MoveNext
Wend
End If
End Function

Public Function validation() As Boolean

If Trim(txtname.Text) = "" Then
   MsgBox "Enter a name", vbInformation, App.Title
   txtname.SetFocus
   validation = False
Exit Function
End If

If Trim(txtaddress.Text) = "" Then
  MsgBox "Enter the address", vbInformation, App.Title
  txtaddress.SetFocus
  validation = False
  Exit Function
End If

If Trim(txtctno.Text) = "" Then
   MsgBox "Enter the contact no", vbInformation, App.Title
   txtctno.SetFocus
   validation = False
   Exit Function
End If

validation = True

End Function




Private Sub txtctno_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = False And Not KeyAscii = 8 Then
MsgBox "Only numbers are allowed"
KeyAscii = 0
End If
End Sub
