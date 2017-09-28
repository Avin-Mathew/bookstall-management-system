VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Suplier 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SUPPLIER DETAILS"
   ClientHeight    =   5745
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   10305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   10305
   Begin VB.Frame Frame1 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10095
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   120
         TabIndex        =   8
         Top             =   2280
         Width           =   9735
         Begin VB.CommandButton cmdadd 
            Caption         =   "&ADD"
            Height          =   375
            Left            =   1560
            TabIndex        =   13
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdupdate 
            Caption         =   "&UPDATE"
            Height          =   375
            Left            =   2760
            TabIndex        =   12
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdedit 
            Caption         =   "&EDIT"
            Height          =   375
            Left            =   3960
            TabIndex        =   11
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmddelete 
            Caption         =   "&DELETE"
            Height          =   375
            Left            =   5160
            TabIndex        =   10
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdcancel 
            Caption         =   "&CANCEL"
            Height          =   375
            Left            =   6480
            TabIndex        =   9
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2055
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   9735
         Begin VB.TextBox txtname 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1320
            TabIndex        =   4
            Top             =   240
            Width           =   2655
         End
         Begin VB.TextBox txtaddress 
            Appearance      =   0  'Flat
            Height          =   615
            Left            =   1320
            MultiLine       =   -1  'True
            TabIndex        =   3
            Top             =   960
            Width           =   2655
         End
         Begin VB.TextBox txtctno 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   5880
            TabIndex        =   2
            Top             =   240
            Width           =   2655
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Name"
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   420
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Address"
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   960
            Width           =   570
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Contact no"
            Height          =   195
            Left            =   4920
            TabIndex        =   5
            Top             =   240
            Width           =   780
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grddetails 
         Height          =   2295
         Left            =   120
         TabIndex        =   14
         Top             =   3240
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   4048
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "Suplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdadd_Click()
cmdupdate.Enabled = True
cmddelete.Enabled = False
cmdedit.Enabled = False
cmdcancel.Enabled = True
cmdadd.Enabled = False
Frame2.Enabled = True
txtname.SetFocus
End Sub

Private Sub cmdcancel_Click()


txtname.Tag = ""
txtname.Text = ""
txtaddress.Text = ""
txtctno.Text = ""
Frame2.Enabled = False
cmdadd.Enabled = True
cmdupdate.Enabled = False
cmddelete.Enabled = False
cmdedit.Enabled = False
cmdcancel.Enabled = True
End Sub

Private Sub cmddelete_Click()
If rssup.State = 1 Then rssup.Close
rssup.Open "select * from tbl_supplier where sup_id=" & txtname.Tag & "", con, adOpenKeyset, adLockOptimistic
rssup.Delete
rssup.Update
rssup.Close

fillgrid

txtname.Tag = ""
cmddelete.Enabled = False
cmdadd.Enabled = True
cmdedit.Enabled = False
cmdupdate.Enabled = False


txtname.Text = ""
txtaddress.Text = " "
txtctno.Text = " "
Frame2.Enabled = False

End Sub

Private Sub cmdedit_Click()
Frame2.Enabled = True
cmdupdate.Enabled = True
cmddelete.Enabled = False
cmdedit.Enabled = False
txtname.SetFocus
End Sub

Private Sub cmdupdate_Click()
If validation Then
If txtname.Tag = "" Then
If rssup.State = 1 Then rssup.Close
rssup.Open "select * from tbl_supplier", con, adOpenKeyset, adLockOptimistic
rssup.AddNew
rssup.Fields("sup_name") = txtname.Text
rssup.Fields("sup_address") = txtaddress.Text
rssup.Fields("sup_contactno") = txtctno.Text
rssup.Update
rssup.Close

txtname.Tag = ""
txtname.Text = ""
txtaddress.Text = ""
txtctno.Text = ""
Frame2.Enabled = False
cmdadd.Enabled = True
cmdupdate.Enabled = False
cmddelete.Enabled = False
cmdedit.Enabled = False
cmdcancel.Enabled = True

Else

If rssup.State = 1 Then rssup.Close
rssup.Open "select * from tbl_supplier where sup_id=" & txtname.Tag & "", con, adOpenKeyset, adLockOptimistic
rssup.Fields("sup_name") = txtname.Text
rssup.Fields("sup_address") = txtaddress.Text
rssup.Fields("sup_contactno") = txtctno.Text
rssup.Update
rssup.Close

txtname.Tag = ""
txtname.Text = ""
txtaddress.Text = ""
txtctno.Text = ""
Frame2.Enabled = False
cmdadd.Enabled = True
cmdupdate.Enabled = False
cmddelete.Enabled = False
cmdedit.Enabled = False
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
If rssup.State = 1 Then rssup.Close
rssup.Open "select * from tbl_supplier where sup_id=" & grddetails.TextMatrix(grddetails.row, 0) & "", con, adOpenKeyset, adLockOptimistic
If rssup.RecordCount > 0 Then
txtname.Text = rssup.Fields("sup_name")
txtaddress.Text = rssup.Fields("sup_address")
txtctno.Text = rssup.Fields("sup_contactno")
txtname.Tag = rssup.Fields("sup_id")
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

grddetails.TextMatrix(0, 1) = "Supplier"
grddetails.TextMatrix(0, 2) = "Address"
grddetails.TextMatrix(0, 3) = "Contact"

If rssup.State = 1 Then rssup.Close
rssup.Open "select * from tbl_supplier", con, adOpenKeyset, adLockOptimistic
If rssup.RecordCount > 0 Then
rssup.MoveFirst
While Not rssup.EOF
grddetails.Rows = grddetails.Rows + 1
grddetails.TextMatrix(grddetails.Rows - 1, 0) = rssup.Fields("sup_id")
grddetails.TextMatrix(grddetails.Rows - 1, 1) = rssup.Fields("sup_name")
grddetails.TextMatrix(grddetails.Rows - 1, 2) = rssup.Fields("sup_address")
grddetails.TextMatrix(grddetails.Rows - 1, 3) = rssup.Fields("sup_contactno")
rssup.MoveNext
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
If txtctno.Text <> "" Then
If IsNumeric(txtctno.Text) = False Then
MsgBox "Numbers only allowed"
txtctno.Text = ""
End If
End If
End Sub
