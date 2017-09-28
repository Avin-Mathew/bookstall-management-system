VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmloginreg 
   Caption         =   "Login Registration"
   ClientHeight    =   3450
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   ScaleHeight     =   3450
   ScaleWidth      =   8580
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4095
      Begin VB.TextBox txtusername 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1440
         TabIndex        =   10
         Top             =   840
         Width           =   2250
      End
      Begin VB.TextBox txtpassword 
         Appearance      =   0  'Flat
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1440
         PasswordChar    =   "*"
         TabIndex        =   9
         Top             =   1440
         Width           =   2250
      End
      Begin VB.ComboBox cmbtype 
         Height          =   315
         ItemData        =   "frmloginreg.frx":0000
         Left            =   1440
         List            =   "frmloginreg.frx":000A
         TabIndex        =   8
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "UserName"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Password"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Type"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0FF&
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   8415
      Begin VB.CommandButton Command1 
         Caption         =   " "
         Height          =   195
         Left            =   -360
         TabIndex        =   6
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton cmdedit 
         Caption         =   "EDIT"
         Height          =   675
         Left            =   4080
         TabIndex        =   5
         Top             =   240
         Width           =   1035
      End
      Begin VB.CommandButton cmdcancel 
         Caption         =   "CANCEL"
         Height          =   675
         Left            =   6240
         TabIndex        =   4
         Top             =   240
         Width           =   1035
      End
      Begin VB.CommandButton cmddelete 
         Caption         =   "DELETE"
         Height          =   675
         Left            =   5160
         TabIndex        =   3
         Top             =   240
         Width           =   1035
      End
      Begin VB.CommandButton cmdadd 
         Caption         =   "ADD"
         Height          =   675
         Left            =   1920
         TabIndex        =   2
         Top             =   240
         Width           =   1035
      End
      Begin VB.CommandButton cmdupdate 
         Caption         =   "UPDATE"
         Height          =   675
         Left            =   3000
         TabIndex        =   1
         Top             =   240
         Width           =   1035
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexReg 
      Height          =   1935
      Left            =   4320
      TabIndex        =   14
      Top             =   240
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   3413
      _Version        =   393216
      Rows            =   1
      Cols            =   4
      FixedRows       =   0
   End
End
Attribute VB_Name = "frmloginreg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdadd_Click()
Frame1.Enabled = True

cmdedit.Enabled = False
cmddelete.Enabled = False
cmdupdate.Enabled = True
cmdcancel.Enabled = True
cmdadd.Enabled = False
cmbtype.SetFocus
clrtxt
End Sub
Private Sub clrtxt()
cmbtype.Text = "--select--"
txtusername.Text = ""
txtpassword.Text = ""

End Sub

Private Sub cmdcancel_Click()

cmbtype.Text = "--select--"
Frame1.Enabled = False
cmdadd.Enabled = True
cmdupdate.Enabled = False
cmdcancel.Enabled = True
cmdedit.Enabled = False
cmddelete.Enabled = False
txtpassword.Text = ""
txtusername.Text = ""
txtusername.Tag = ""

End Sub

Private Sub cmddelete_Click()
If rsLoginReg.State = 1 Then rsLoginReg.Close
rsLoginReg.Open "select * from tbl_LOGIN where login_id=" & txtusername.Tag & " ", con, adOpenKeyset, adLockOptimistic
rsLoginReg.Delete
rsLoginReg.Update
rsLoginReg.Close
clrtxt
fillgrid
txtusername.Tag = ""
cmdadd.Enabled = True
cmdedit.Enabled = False
cmddelete.Enabled = False
cmdupdate.Enabled = False

End Sub

Private Sub cmdedit_Click()
Frame1.Enabled = True
cmdupdate.Enabled = True
cmddelete.Enabled = False
cmdadd.Enabled = False
cmbtype.SetFocus
cmdedit.Enabled = False
End Sub

Private Sub cmdupdate_Click()
If validation Then
If txtusername.Tag = "" Then
If rsLoginReg.State = 1 Then rsLoginReg.Close
rsLoginReg.Open "select * from tbl_LOGIN ", con, adOpenKeyset, adLockOptimistic
rsLoginReg.AddNew
rsLoginReg.Fields("login_type") = cmbtype.Text
rsLoginReg.Fields("login_username") = txtusername.Text
rsLoginReg.Fields("login_password") = txtpassword.Text
rsLoginReg.Update
rsLoginReg.Close

cmbtype.Text = "--select--"
Frame1.Enabled = False
cmdadd.Enabled = True
cmdupdate.Enabled = False
cmdcancel.Enabled = True
cmdedit.Enabled = False
cmddelete.Enabled = False
txtpassword.Text = ""
txtusername.Text = ""



Else
If rsLoginReg.State = 1 Then rsLoginReg.Close
rsLoginReg.Open "select * from tbl_LOGIN where login_id = " & txtusername.Tag & " ", con, adOpenKeyset, adLockOptimistic
rsLoginReg.Fields("login_type") = cmbtype.Text
rsLoginReg.Fields("login_username") = txtusername.Text
rsLoginReg.Fields("login_password") = txtpassword.Text
rsLoginReg.Update
rsLoginReg.Close

cmbtype.Text = "--select--"
Frame1.Enabled = False
cmdadd.Enabled = True
cmdupdate.Enabled = False
cmdcancel.Enabled = True
cmdedit.Enabled = False
cmddelete.Enabled = False
txtpassword.Text = ""
txtusername.Text = ""
txtusername.Tag = ""

End If
End If

fillgrid


End Sub

Private Sub Form_Load()
cmdadd.Enabled = True
cmdupdate.Enabled = False
cmddelete.Enabled = False
cmdcancel.Enabled = True
cmdedit.Enabled = False
Frame1.Enabled = False
fillgrid
clrtxt
End Sub
Public Function fillgrid()

MSFlexReg.Rows = 1
MSFlexReg.Cols = 4

MSFlexReg.ColWidth(0) = 0
MSFlexReg.ColWidth(1) = 1400
MSFlexReg.ColWidth(2) = 1400
MSFlexReg.ColWidth(3) = 1400

MSFlexReg.TextMatrix(0, 1) = "Type"
MSFlexReg.TextMatrix(0, 2) = "User Name"
MSFlexReg.TextMatrix(0, 3) = "Password"

If rsLoginReg.State = 1 Then rsLoginReg.Close
rsLoginReg.Open " select * from tbl_LOGIN", con, adOpenKeyset, adLockOptimistic
If (rsLoginReg.RecordCount > 0) Then
rsLoginReg.MoveFirst
While Not rsLoginReg.EOF
MSFlexReg.Rows = MSFlexReg.Rows + 1
MSFlexReg.TextMatrix(MSFlexReg.Rows - 1, 0) = rsLoginReg.Fields("login_id")
MSFlexReg.TextMatrix(MSFlexReg.Rows - 1, 1) = rsLoginReg.Fields("login_type")
MSFlexReg.TextMatrix(MSFlexReg.Rows - 1, 2) = rsLoginReg.Fields("login_username")
MSFlexReg.TextMatrix(MSFlexReg.Rows - 1, 3) = rsLoginReg.Fields("login_password")
rsLoginReg.MoveNext
Wend
End If

End Function

Private Sub MSFlexReg_click()
If rsLoginReg.State = 1 Then rsLoginReg.Close
rsLoginReg.Open "select * from tbl_LOGIN where login_id=" & MSFlexReg.TextMatrix(MSFlexReg.row, 0) & " ", con, adOpenKeyset, adLockOptimistic
If (rsLoginReg.RecordCount > 0) Then
cmbtype.Text = rsLoginReg.Fields("login_type")
txtusername.Text = rsLoginReg.Fields("login_username")
txtpassword.Text = rsLoginReg.Fields("login_password")
txtusername.Tag = rsLoginReg.Fields("login_id")
End If
rsLoginReg.Close
cmddelete.Enabled = True
cmdedit.Enabled = True
cmdcancel.Enabled = True
cmdadd.Enabled = False

End Sub




Public Function validation() As Boolean
If Trim(cmbtype.Text) = "" Then
   MsgBox "Enter the type", vbInformation, App.Title
   cmbtype.SetFocus
   validation = False
Exit Function
End If

If Trim(txtusername.Text) = "" Then
  MsgBox "Enter the username", vbInformation, App.Title
  txtusername.SetFocus
  validation = False
  Exit Function
End If

If Trim(txtpassword.Text) = "" Then
   MsgBox "Enter the password", vbInformation, App.Title
   txtpassword.SetFocus
   validation = False
   Exit Function
End If

validation = True
End Function
