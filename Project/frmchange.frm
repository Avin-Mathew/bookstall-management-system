VERSION 5.00
Begin VB.Form frmchange 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   5235
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      Begin VB.TextBox txtOld 
         Appearance      =   0  'Flat
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   360
         Width           =   2535
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Update"
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Top             =   2280
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   2640
         TabIndex        =   3
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox txtnew 
         Appearance      =   0  'Flat
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox txtretype 
         Appearance      =   0  'Flat
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Old Password"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "New Password"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Retype Password"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1560
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmchange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSave_Click()
If validation Then
If rsChange.State = 1 Then rsChange.Close
rsChange.Open "select * from tbl_login where login_username='" & loginname & "' and login_password='" & txtOld.Text & "'", con, adOpenKeyset, adLockOptimistic
If rsChange.RecordCount > 0 Then
 If txtnew.Text = txtretype.Text Then
   rsChange.Fields("login_password") = txtnew.Text
   rsChange.Update
   rsChange.Close
   txtnew.Text = ""
   txtOld.Text = ""
   txtretype.Text = ""
   MsgBox "Password Changed Successfully..."
 Else
   MsgBox "Password mismatch..."
 End If
 End If
Else
  MsgBox "Invalid entry..."
End If
  
End Sub


Public Function validation() As Boolean
If Trim(txtOld.Text) = "" Then
   MsgBox "Enter the old password", vbInformation, App.Title
   txtOld.SetFocus
   validation = False
Exit Function
End If

If Trim(txtnew.Text) = "" Then
  MsgBox "Enter the new password", vbInformation, App.Title
  txtnew.SetFocus
  validation = False
  Exit Function
End If

If Trim(txtretype.Text) = "" Then
   MsgBox "Retype the password", vbInformation, App.Title
   txtretype.SetFocus
   validation = False
   Exit Function
End If

validation = True
End Function
