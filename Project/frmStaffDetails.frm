VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Staff 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Staff"
   ClientHeight    =   6045
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   9180
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   9180
   Begin VB.Frame Frame1 
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9015
      Begin VB.Frame Frame3 
         Height          =   855
         Left            =   120
         TabIndex        =   10
         Top             =   2400
         Width           =   8775
         Begin VB.CommandButton cmdadd 
            Caption         =   "&ADD"
            Height          =   495
            Left            =   1440
            TabIndex        =   15
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdupdate 
            Caption         =   "&UPDATE"
            Height          =   495
            Left            =   2730
            TabIndex        =   14
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdedit 
            Caption         =   "&EDIT"
            Height          =   495
            Left            =   4020
            TabIndex        =   13
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmddelete 
            Caption         =   "&DELETE"
            Height          =   495
            Left            =   5310
            TabIndex        =   12
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdcancel 
            Caption         =   "&CANCEL"
            Height          =   495
            Left            =   6600
            TabIndex        =   11
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2055
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   8775
         Begin VB.TextBox txtname 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   840
            TabIndex        =   5
            Top             =   360
            Width           =   2895
         End
         Begin VB.TextBox txtadress 
            Appearance      =   0  'Flat
            Height          =   615
            Left            =   840
            MultiLine       =   -1  'True
            TabIndex        =   4
            Top             =   1080
            Width           =   3135
         End
         Begin VB.TextBox txtemail 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   5640
            TabIndex        =   3
            Top             =   1080
            Width           =   2895
         End
         Begin VB.TextBox txtcno 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   5640
            TabIndex        =   2
            Top             =   360
            Width           =   2895
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Name"
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   480
            Width           =   420
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Address"
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   1080
            Width           =   570
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Email"
            Height          =   195
            Left            =   4680
            TabIndex        =   7
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Contact No"
            Height          =   195
            Left            =   4680
            TabIndex        =   6
            Top             =   480
            Width           =   810
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grddetails 
         Height          =   2295
         Left            =   120
         TabIndex        =   16
         Top             =   3360
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   4048
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "Staff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


 Private Sub txtadres_Change()

End Sub

Private Sub cmdadd_Click()
cmdupdate.Enabled = True
cmdcancel.Enabled = True
cmdadd.Enabled = False
cmdedit.Enabled = False
cmddelete.Enabled = False
Frame2.Enabled = True
txtname.SetFocus
End Sub

Private Sub cmdcancel_Click()
   txtname.Text = ""
   txtcno.Text = ""
   txtemail.Text = ""
   txtadress.Text = " "
   txtname.Tag = ""
  cmdupdate.Enabled = False
cmdadd.Enabled = True
cmdedit.Enabled = False
cmddelete.Enabled = False
cmdcancel.Enabled = True
Frame2.Enabled = False

End Sub

Private Sub cmddelete_Click()
If rsstf.State = 1 Then rsstf.Close
rsstf.Open "select * from tbl_staff where stf_id=" & txtname.Tag & "", con, adOpenKeyset, adLockOptimistic
rsstf.Delete
rsstf.Update
rsstf.Close
txtname.Tag = ""
fillgrid
cmddelete.Enabled = False
cmdadd.Enabled = True
cmdedit.Enabled = False
txtname.Text = ""
txtcno.Text = ""
txtemail.Text = ""
txtadress.Text = " "

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
  If rsstf.State = 1 Then rsstf.Close
  rsstf.Open "select * from tbl_staff", con, adOpenKeyset, adLockOptimistic
  rsstf.AddNew
  rsstf.Fields("stf_name") = txtname.Text
  rsstf.Fields("stf_address") = txtadress.Text
  rsstf.Fields("stf_contactno") = txtcno.Text
  rsstf.Fields("stf_email") = txtemail.Text

  rsstf.Update
  rsstf.Close
  
   txtname.Text = ""
   txtcno.Text = ""
   txtemail.Text = ""
  txtadress.Text = " "

cmdupdate.Enabled = False
cmdadd.Enabled = True
cmdedit.Enabled = False
cmddelete.Enabled = False
cmdcancel.Enabled = True
Frame2.Enabled = False
  
  
  
Else

 If rsstf.State = 1 Then rsstf.Close
   rsstf.Open "select * from tbl_staff where stf_id=" & txtname.Tag & "", con, adOpenKeyset, adLockOptimistic
   rsstf.Fields("stf_name") = txtname.Text
   rsstf.Fields("stf_address") = txtadress.Text
   rsstf.Fields("stf_contactno") = txtcno.Text
   rsstf.Fields("stf_email") = txtemail.Text

   rsstf.Update
   rsstf.Close
   
   txtname.Text = ""
   txtcno.Text = ""
   txtemail.Text = ""
txtadress.Text = " "
txtname.Tag = ""
cmdupdate.Enabled = False
cmdadd.Enabled = True
cmdedit.Enabled = False
cmddelete.Enabled = False
cmdcancel.Enabled = True
Frame2.Enabled = False
   
   
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
If rsstf.State = 1 Then rsstf.Close
rsstf.Open "select * from tbl_staff where stf_id=" & grddetails.TextMatrix(grddetails.row, 0) & " ", con, adOpenKeyset, adLockOptimistic
If rsstf.RecordCount > 0 Then
txtname.Text = rsstf.Fields("stf_name")
txtemail.Text = rsstf.Fields("stf_email")
txtcno.Text = rsstf.Fields("stf_contactno")
txtadress.Text = rsstf.Fields("stf_address")

txtname.Tag = rsstf.Fields("stf_id")



End If
cmdedit.Enabled = True
cmddelete.Enabled = True
cmdadd.Enabled = False
End Sub

Public Function fillgrid()
grddetails.Rows = 1
grddetails.Cols = 5


grddetails.ColWidth(0) = 0
grddetails.ColWidth(1) = 1700
grddetails.ColWidth(2) = 1700
grddetails.ColWidth(3) = 1700
grddetails.ColWidth(4) = 1700




grddetails.TextMatrix(0, 1) = " Name"
grddetails.TextMatrix(0, 2) = "Contact no"
grddetails.TextMatrix(0, 3) = "Email"
grddetails.TextMatrix(0, 4) = "Address"

If rsstf.State = 1 Then rsstf.Close
rsstf.Open "select * from tbl_staff", con, adOpenKeyset, adLockOptimistic
If rsstf.RecordCount > 0 Then
rsstf.MoveFirst
While Not rsstf.EOF
grddetails.Rows = grddetails.Rows + 1
grddetails.TextMatrix(grddetails.Rows - 1, 0) = rsstf.Fields("stf_id")
grddetails.TextMatrix(grddetails.Rows - 1, 1) = rsstf.Fields("stf_name")
grddetails.TextMatrix(grddetails.Rows - 1, 2) = rsstf.Fields("stf_contactno")
grddetails.TextMatrix(grddetails.Rows - 1, 3) = rsstf.Fields("stf_email")
grddetails.TextMatrix(grddetails.Rows - 1, 4) = rsstf.Fields("stf_address")


rsstf.MoveNext
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

If Trim(txtadress.Text) = "" Then
  MsgBox "Enter the address", vbInformation, App.Title
  txtadress.SetFocus
  validation = False
  Exit Function
End If

If Trim(txtcno.Text) = "" Then
   MsgBox "Enter the contact no", vbInformation, App.Title
   txtcno.SetFocus
   validation = False
   Exit Function
End If
If Trim(txtemail.Text) = "" Then
   MsgBox "Enter the mail id", vbInformation, App.Title
   txtemail.SetFocus
   validation = False
   Exit Function
End If

validation = True

End Function

Private Sub txtcno_KeyPress(KeyAscii As Integer)
If txtcno.Text <> "" Then
If IsNumeric(txtcno.Text) = False Then
MsgBox "Numbers only allowed"
txtcno.Text = ""
End If
End If
End Sub
