VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form publisher 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PUBLISHER"
   ClientHeight    =   6570
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   8385
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   8385
   Begin VB.Frame Frame1 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      Begin VB.Frame Frame4 
         Height          =   3135
         Left            =   120
         TabIndex        =   16
         Top             =   3120
         Width           =   7935
         Begin MSFlexGridLib.MSFlexGrid grddetails 
            Height          =   2775
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   7695
            _ExtentX        =   13573
            _ExtentY        =   4895
            _Version        =   393216
         End
      End
      Begin VB.Frame Frame3 
         Height          =   975
         Left            =   120
         TabIndex        =   10
         Top             =   2040
         Width           =   7935
         Begin VB.CommandButton cmdadd 
            Caption         =   "&ADD"
            Height          =   495
            Left            =   840
            TabIndex        =   15
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdupdate 
            Caption         =   "&UPDATE"
            Height          =   495
            Left            =   2040
            TabIndex        =   14
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdedit 
            Caption         =   "&EDIT"
            Height          =   495
            Left            =   3240
            TabIndex        =   13
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmddelete 
            Caption         =   "&DELETE"
            Height          =   495
            Left            =   4440
            TabIndex        =   12
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdcancel 
            Caption         =   "&CANCEL"
            Height          =   495
            Left            =   5640
            TabIndex        =   11
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1695
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   7935
         Begin VB.TextBox txtpubname 
            Appearance      =   0  'Flat
            Height          =   495
            Left            =   720
            TabIndex        =   5
            Top             =   240
            Width           =   2775
         End
         Begin VB.TextBox txtpubemail 
            Appearance      =   0  'Flat
            Height          =   495
            Left            =   720
            TabIndex        =   4
            Top             =   960
            Width           =   2775
         End
         Begin VB.TextBox txtpubweb 
            Appearance      =   0  'Flat
            Height          =   495
            Left            =   4920
            TabIndex        =   3
            Top             =   240
            Width           =   2775
         End
         Begin VB.TextBox txtpubcno 
            Appearance      =   0  'Flat
            Height          =   495
            Left            =   4920
            TabIndex        =   2
            Top             =   960
            Width           =   2775
         End
         Begin VB.Line Line1 
            X1              =   3720
            X2              =   3720
            Y1              =   240
            Y2              =   1440
         End
         Begin VB.Label Label6 
            Caption         =   "Name"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   735
         End
         Begin VB.Label label2 
            Caption         =   "Email"
            Height          =   495
            Index           =   0
            Left            =   120
            TabIndex        =   8
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label label3 
            Caption         =   "Website"
            Height          =   255
            Index           =   1
            Left            =   3960
            TabIndex        =   7
            Top             =   360
            Width           =   855
         End
         Begin VB.Label label4 
            Caption         =   "Contact No"
            Height          =   255
            Left            =   3960
            TabIndex        =   6
            Top             =   1080
            Width           =   855
         End
      End
   End
End
Attribute VB_Name = "publisher"
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
txtpubname.SetFocus
End Sub

Private Sub cmdcancel_Click()
txtpubname.Text = ""
txtpubcno.Text = ""
txtpubemail.Text = ""
txtpubweb.Text = ""
txtpubname.Tag = ""
Frame2.Enabled = False
cmdadd.Enabled = True
cmdupdate.Enabled = False
cmdedit.Enabled = False
cmddelete.Enabled = False
cmdcancel.Enabled = True


End Sub

Private Sub cmddelete_Click()
If rspub.State = 1 Then rspub.Close
rspub.Open "select * from tbl_publisher where pub_id=" & txtpubname.Tag & "", con, adOpenKeyset, adLockOptimistic
rspub.Delete
rspub.Update
rspub.Close

txtpubname.Tag = ""
fillgrid

cmddelete.Enabled = False
cmdadd.Enabled = True
cmdedit.Enabled = False
cmdupdate.Enabled = False
cmdcancel.Enabled = True

txtpubname.Text = ""
txtpubcno.Text = ""
txtpubemail.Text = ""
txtpubweb.Text = ""
End Sub

Private Sub cmdedit_Click()
Frame2.Enabled = True
cmdupdate.Enabled = True
cmddelete.Enabled = False
cmdedit.Enabled = False
txtpubname.SetFocus
End Sub

Private Sub cmdupdate_Click()
If validation Then
If txtpubname.Tag = "" Then
If rspub.State = 1 Then rspub.Close
rspub.Open "select * from tbl_publisher", con, adOpenKeyset, adLockOptimistic
rspub.AddNew
rspub.Fields("pub_name") = txtpubname.Text
rspub.Fields("pub_email") = txtpubemail.Text
rspub.Fields("pub_website") = txtpubweb.Text
rspub.Fields("pub_contactno") = txtpubcno.Text
rspub.Update
rspub.Close

txtpubname.Text = ""
txtpubcno.Text = ""
txtpubemail.Text = ""
txtpubweb.Text = ""
Frame2.Enabled = False
cmdadd.Enabled = True
cmdupdate.Enabled = False
cmdedit.Enabled = False
cmddelete.Enabled = False
cmdcancel.Enabled = True



Else
If rspub.State = 1 Then rspub.Close
rspub.Open "select * from tbl_publisher where pub_id=" & txtpubname.Tag & "", con, adOpenKeyset, adLockOptimistic
rspub.Fields("pub_name") = txtpubname.Text
rspub.Fields("pub_email") = txtpubemail.Text
rspub.Fields("pub_website") = txtpubweb.Text
rspub.Fields("pub_contactno") = txtpubcno.Text
rspub.Update
rspub.Close


txtpubname.Text = ""
txtpubcno.Text = ""
txtpubemail.Text = ""
txtpubweb.Text = ""
txtpubname.Tag = ""
Frame2.Enabled = False
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
If rspub.State = 1 Then rspub.Close
rspub.Open "select * from tbl_publisher where pub_id=" & grddetails.TextMatrix(grddetails.row, 0) & " ", con, adOpenKeyset, adLockOptimistic
If rspub.RecordCount > 0 Then
txtpubname.Text = rspub.Fields("pub_name")
txtpubemail.Text = rspub.Fields("pub_email")
txtpubweb.Text = rspub.Fields("pub_website")
txtpubcno.Text = rspub.Fields("pub_contactno")
txtpubname.Tag = rspub.Fields("pub_id")

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

grddetails.TextMatrix(0, 1) = "publisher Name"
grddetails.TextMatrix(0, 2) = "website"
grddetails.TextMatrix(0, 3) = "contact no"
grddetails.TextMatrix(0, 4) = "email"

If rspub.State = 1 Then rspub.Close
rspub.Open "select * from tbl_publisher", con, adOpenKeyset, adLockOptimistic
If rspub.RecordCount > 0 Then
rspub.MoveFirst
While Not rspub.EOF
grddetails.Rows = grddetails.Rows + 1
grddetails.TextMatrix(grddetails.Rows - 1, 0) = rspub.Fields("pub_id")
grddetails.TextMatrix(grddetails.Rows - 1, 1) = rspub.Fields("pub_name")
grddetails.TextMatrix(grddetails.Rows - 1, 2) = rspub.Fields("pub_website")
grddetails.TextMatrix(grddetails.Rows - 1, 3) = rspub.Fields("pub_contactno")
grddetails.TextMatrix(grddetails.Rows - 1, 4) = rspub.Fields("pub_email")
rspub.MoveNext
Wend
End If
End Function


Public Function validation() As Boolean
If Trim(txtpubname.Text) = "" Then
MsgBox "Enter a name", vbInformation, App.Title
txtpubname.SetFocus
validation = False
Exit Function
End If
If Trim(txtpubemail.Text) = "" Then
MsgBox "Enter the email", vbInformation, App.Title
txtpubemail.SetFocus
validation = False
Exit Function
End If
If Trim(txtpubcno.Text) = "" Then
MsgBox "Enter the contact no", vbInformation, App.Title
txtpubcno.SetFocus
validation = False
Exit Function
End If

If Trim(txtpubweb.Text) = "" Then
MsgBox "Enter the website", vbInformation, App.Title
txtpubweb.SetFocus
validation = False
Exit Function
End If
validation = True

End Function



Private Sub txtpubcno_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = False And Not KeyAscii = 8 Then
MsgBox "Only numbers are allowed"
KeyAscii = 0
End If
End Sub
