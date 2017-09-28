VERSION 5.00
Begin VB.Form Login 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   8850
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":0000
   ScaleHeight     =   8850
   ScaleWidth      =   15060
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Height          =   2175
      Left            =   8280
      TabIndex        =   0
      Top             =   3840
      Width           =   4215
      Begin VB.CommandButton cmdLogin 
         Caption         =   "Login"
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox txtusername 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Text            =   "admin"
         Top             =   360
         Width           =   2250
      End
      Begin VB.TextBox txtpassword 
         Appearance      =   0  'Flat
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   1
         Text            =   "admin"
         Top             =   960
         Width           =   2250
      End
      Begin VB.Label lblmsg 
         Height          =   255
         Left            =   1560
         TabIndex        =   6
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Password"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "UserName"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdLogin_Click()

If rsLogin.State = 1 Then rsLogin.Close
rsLogin.Open "select * from tbl_login where login_username= '" & txtusername.Text & "' and login_password='" & txtpassword.Text & "'", con, adOpenKeyset, adLockOptimistic
If rsLogin.RecordCount > 0 Then
  
  loginid = rsLogin.Fields("login_id")
  logintype = rsLogin.Fields("login_type")
  loginname = rsLogin.Fields("login_username")
  
  If logintype = "Admin" Then
       MDIBookShop.Caption = "welcome " & loginname
       MDIBookShop.Show
       Login.Hide
  Else
       MDIBookShop.Caption = "welcome " & loginname
       MDIBookShop.mnumaster.Enabled = False
       MDIBookShop.Show
       Login.Hide
  End If
 
     
Else
MsgBox "Invalid Login...", vbCritical

End If

End Sub

