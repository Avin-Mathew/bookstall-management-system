VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Admin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ADMIN"
   ClientHeight    =   6060
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   7380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   7380
   Begin VB.Frame frame1 
      Height          =   6015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      Begin MSFlexGridLib.MSFlexGrid grddetails 
         Height          =   1815
         Left            =   120
         TabIndex        =   20
         Top             =   4200
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   3201
         _Version        =   393216
         Cols            =   7
      End
      Begin VB.Frame Frame3 
         Height          =   975
         Left            =   120
         TabIndex        =   14
         Top             =   3120
         Width           =   7095
         Begin VB.CommandButton cmdcancel 
            Caption         =   "&CANCEL"
            Height          =   495
            Left            =   5640
            TabIndex        =   19
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmddelete 
            Caption         =   "&DELETE"
            Height          =   495
            Left            =   4200
            TabIndex        =   18
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdedit 
            Caption         =   "&EDIT"
            Height          =   495
            Left            =   2760
            TabIndex        =   17
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdupdate 
            Caption         =   "&UPDATE"
            Height          =   495
            Left            =   1440
            TabIndex        =   16
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdadd 
            Appearance      =   0  'Flat
            Caption         =   "&ADD"
            Height          =   495
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3015
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   7095
         Begin VB.TextBox txtwebsite 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   840
            TabIndex        =   22
            Top             =   2400
            Width           =   2295
         End
         Begin VB.TextBox txtpassword 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   4440
            TabIndex        =   13
            Top             =   1680
            Width           =   2295
         End
         Begin VB.TextBox txtuname 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   4440
            TabIndex        =   12
            Top             =   840
            Width           =   2295
         End
         Begin VB.TextBox txtcno 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   4440
            TabIndex        =   11
            Top             =   240
            Width           =   2295
         End
         Begin VB.TextBox txtemail 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   840
            TabIndex        =   7
            Top             =   1680
            Width           =   2295
         End
         Begin VB.TextBox txtaddress 
            Appearance      =   0  'Flat
            Height          =   615
            Left            =   840
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   6
            Top             =   840
            Width           =   2535
         End
         Begin VB.TextBox txtname 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   840
            TabIndex        =   5
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label A 
            Caption         =   "Website"
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   2520
            Width           =   1020
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Password"
            Height          =   195
            Left            =   3480
            TabIndex        =   10
            Top             =   1800
            Width           =   690
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "User name"
            Height          =   195
            Left            =   3480
            TabIndex        =   9
            Top             =   960
            Width           =   765
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Contact No"
            Height          =   195
            Left            =   3480
            TabIndex        =   8
            Top             =   360
            Width           =   810
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Email"
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Top             =   1800
            Width           =   375
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Address"
            Height          =   195
            Left            =   120
            TabIndex        =   3
            Top             =   960
            Width           =   570
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Name"
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   420
         End
      End
   End
End
Attribute VB_Name = "Admin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdadd_Click()
Frame2.Enabled = True
cmdadd.Enabled = False
cmdupdate.Enabled = True
cmddelete.Enabled = False
cmdcancel.Enabled = True
txtname.SetFocus
End Sub

Private Sub cmdcancel_Click()
txtname.Text = " "
txtaddress.Text = " "
txtemail.Text = " "
txtcno.Text = " "
txtuname.Text = " "
txtpassword.Text = " "
txtwebsite.Text = " "
End Sub

Private Sub cmddelete_Click()
If rsadmin.State = 1 Then rsadmin.Close
rsadmin.Open " select * from tbl_admin where admin_id = " & txtname.Tag & " ", con, adOpenKeyset, adLockOptimistic
rsadmin.Delete
rsadmin.Update
rsadmin.Close
txtname.Tag = " "
fillgrid
End Sub

Private Sub cmdedit_Click()
Frame2.Enabled = True
cmddelete.Enabled = False
cmdupdate.Enabled = True
End Sub

Private Sub cmdupdate_Click()
If Validate Then

If txtname.Tag = " " Then
If rsadmin.State = 1 Then rsadmin.Close
rsadmin.Open " select * from tbl_admin ", con, adOpenKeyset, adLockOptimistic
rsadmin.AddNew
rsadmin.Fields("admin_name") = txtname.Text
rsadmin.Fields("admin_address") = txtaddress.Text
rsadmin.Fields("admin_contact_no") = txtcno.Text
rsadmin.Fields("admin_uname") = txtuname.Text
rsadmin.Fields("admin_Password ") = txtpassword.Text
rsadmin.Fields("admin_website") = txtwebsite.Text
rsadmin.Update
rsadmin.Close
Else
If rsadmin.State = 1 Then rsadmin.Close
rsadmin.Open " select * from tbl_admin ", con, adOpenKeyset, adLockOptimistic
rsadmin.AddNew
rsadmin.Fields("admin_name") = txtname.Text
rsadmin.Fields("admin_address") = txtaddress.Text
rsadmin.Fields("admin_contact_no") = txtcno.Text
rsadmin.Fields("admin_uname") = txtuname.Text
rsadmin.Fields("admin_password") = txtpassword.Text
rsadmin.Fields("admin_website") = txtwebsite.Text
rsadmin.Fields("admin_email") = txtemail.Text
rsadmin.Update
rsadmin.Close
txtname.Tag = " "
End If

End If

cmdupdate.Enabled = False
cmdadd.Enabled = True
Frame2.Enabled = False
fillgrid
End Sub

Private Sub Form_Load()

Frame2.Enabled = False
cmdadd.Enabled = True
cmdupdate.Enabled = False
cmdcancel.Enabled = True
cmddelete.Enabled = False
cmdedit.Enabled = False
fillgrid
End Sub

Public Sub fillgrid()
grddetails.Rows = 1
grddetails.Cols = 9
grddetails.ColWidth(0) = 0
grddetails.ColWidth(1) = 1400
grddetails.ColWidth(2) = 1400
grddetails.ColWidth(3) = 1400
grddetails.ColWidth(4) = 1400
grddetails.ColWidth(5) = 1400
grddetails.ColWidth(6) = 1400
grddetails.ColWidth(7) = 1400
grddetails.TextMatrix(0, 1) = "name"
grddetails.TextMatrix(0, 2) = "address"
grddetails.TextMatrix(0, 3) = "email"
grddetails.TextMatrix(0, 4) = "website"
grddetails.TextMatrix(0, 5) = "contact no"
grddetails.TextMatrix(0, 6) = "user name"
grddetails.TextMatrix(0, 7) = "password"
If rsadmin.State = 1 Then rsadmin.Close
rsadmin.Open " select * from tbl_admin ", con, adOpenKeyset, adLockOptimistic
If (rsadmin.RecordCount > 0) Then
rsadmin.MoveFirst
While Not rsadmin.EOF
grddetails.Rows = grddetails.Rows + 1
grddetails.TextMatrix(grddetails.Rows - 1, 0) = rsadmin.Fields("admin_id")
grddetails.TextMatrix(grddetails.Rows - 1, 1) = rsadmin.Fields("admin_name")
grddetails.TextMatrix(grddetails.Rows - 1, 2) = rsadmin.Fields("admin_address")
grddetails.TextMatrix(grddetails.Rows - 1, 3) = rsadmin.Fields("admin_email")
grddetails.TextMatrix(grddetails.Rows - 1, 4) = rsadmin.Fields("admin_website")
grddetails.TextMatrix(grddetails.Rows - 1, 5) = rsadmin.Fields("admin_contact_no")
grddetails.TextMatrix(grddetails.Rows - 1, 6) = rsadmin.Fields("admin_uname")
grddetails.TextMatrix(grddetails.Rows - 1, 7) = rsadmin.Fields("admin_password")


rsadmin.MoveNext
Wend
End If
End Sub

Private Sub grddetails_Click()

If rsadmin.State = 1 Then rsadmin.Close
rsadmin.Open "select * from tbl_admin where admin_id=" & grddetails.TextMatrix(grddetails.Row, 0) & "", con, adOpenKeyset, adLockOptimistic
If rsadmin.RecordCount > 0 Then
txtname.Text = rsadmin.Fields("admin_name")
txtname.Text = rsadmin.Fields("admin_uname")
txtname.Text = rsadmin.Fields("admin_contact_no")
txtname.Text = rsadmin.Fields("admin_website")
txtname.Text = rsadmin.Fields("admin_email")
txtname.Text = rsadmin.Fields("admin_address")
txtname.Tag = rsadmin.Fields("admin_id")
txtname.Text = rsadmin.Fields("admin_password")
End If
cmdedit.Enabled = True
cmddelete.Enabled = True
cmdadd.Enabled = False
End Sub

