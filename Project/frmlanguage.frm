VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form language 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LANGUAGE"
   ClientHeight    =   4455
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   7050
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   7050
   Begin VB.Frame Frame1 
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6975
      Begin VB.Frame Frame3 
         Height          =   855
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   6735
         Begin VB.CommandButton cmdcancel 
            Caption         =   "&CANCEL"
            Height          =   495
            Left            =   5400
            TabIndex        =   9
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmddelete 
            Caption         =   "&DELETE"
            Height          =   495
            Left            =   4080
            TabIndex        =   8
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdedit 
            Caption         =   "&EDIT"
            Height          =   495
            Left            =   2760
            TabIndex        =   7
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdupdate 
            Caption         =   "&UPDATE"
            Height          =   495
            Left            =   1440
            TabIndex        =   6
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdadd 
            Caption         =   "&ADD"
            Height          =   495
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1095
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   6735
         Begin VB.TextBox txtlname 
            Height          =   375
            Left            =   1080
            TabIndex        =   2
            Top             =   360
            Width           =   2535
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Language"
            Height          =   195
            Left            =   240
            TabIndex        =   3
            Top             =   360
            Width           =   720
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grddetails 
         Height          =   1815
         Left            =   120
         TabIndex        =   10
         Top             =   2280
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   3201
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "language"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

 

Private Sub cmdadd_Click()
cmdupdate.Enabled = True
cmdcancel.Enabled = True
Frame2.Enabled = True
cmdedit.Enabled = False
cmddelete.Enabled = False
cmdadd.Enabled = False
txtlname.SetFocus

End Sub


Private Sub cmdcancel_Click()
txtlname.Text = ""
txtlname.Tag = ""
Frame2.Enabled = False
cmdadd.Enabled = True
cmdupdate.Enabled = False
cmdedit.Enabled = False
cmddelete.Enabled = False
cmdcancel.Enabled = True
End Sub

Private Sub cmddelete_Click()
If rslang.State = 1 Then rslang.Close
rslang.Open "select * from tbl_language where l_id=" & txtlname.Tag & "", con, adOpenKeyset, adLockOptimistic
rslang.Delete
rslang.Update
txtlname.Tag = ""
fillgrid
cmddelete.Enabled = False
cmdadd.Enabled = True
cmdedit.Enabled = False
txtlname.Text = ""
End Sub

Private Sub cmdedit_Click()
Frame2.Enabled = True
cmdupdate.Enabled = True
cmddelete.Enabled = False
cmdedit.Enabled = False
cmdadd.Enabled = False
txtlname.SetFocus
End Sub

Private Sub cmdupdate_Click()
If validation Then
If (txtlname.Tag = "") Then
If rslang.State = 1 Then rslang.Close
rslang.Open "select * from tbl_language", con, adOpenKeyset, adLockOptimistic
rslang.AddNew
rslang.Fields("l_name") = txtlname.Text
rslang.Update
rslang.Close

txtlname.Text = ""

Frame2.Enabled = False
cmdadd.Enabled = True
cmdupdate.Enabled = False
cmdedit.Enabled = False
cmddelete.Enabled = False
cmdcancel.Enabled = True


Else
If rslang.State = 1 Then rslang.Close
rslang.Open "select * from tbl_language where l_id=" & txtlname.Tag & "", con, adOpenKeyset, adLockOptimistic
rslang.Fields("l_name") = txtlname.Text
rslang.Update
rslang.Close

txtlname.Text = ""
txtlname.Tag = ""
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


cmdupdate.Enabled = False
cmdedit.Enabled = False
cmddelete.Enabled = False
cmdcancel.Enabled = True
Frame2.Enabled = False
fillgrid
End Sub

Private Sub grddetails_Click()
If rslang.State = 1 Then rslang.Close
rslang.Open "select * from tbl_language where l_id=" & grddetails.TextMatrix(grddetails.row, 0) & "", con, adOpenKeyset, adLockOptimistic
If rslang.RecordCount > 0 Then
txtlname.Text = rslang.Fields("l_name")
txtlname.Tag = rslang.Fields("l_id")
End If
cmdedit.Enabled = True
cmddelete.Enabled = True
cmdadd.Enabled = False
End Sub


Public Function fillgrid()
grddetails.Rows = 1
grddetails.Cols = 2

grddetails.ColWidth(0) = 0
grddetails.ColWidth(1) = 1700

grddetails.TextMatrix(0, 1) = "Language"
If rslang.State = 1 Then rslang.Close
rslang.Open "select * from tbl_language", con, adOpenKeyset, adLockOptimistic
If rslang.RecordCount > 0 Then
rslang.MoveFirst
While Not rslang.EOF
grddetails.Rows = grddetails.Rows + 1
grddetails.TextMatrix(grddetails.Rows - 1, 0) = rslang.Fields("l_id")
grddetails.TextMatrix(grddetails.Rows - 1, 1) = rslang.Fields("l_name")
rslang.MoveNext
Wend
End If
End Function

Public Function validation() As Boolean
If Trim(txtlname.Text) = "" Then
   MsgBox "Enter a Language", vbInformation, App.Title
   txtlname.SetFocus
   validation = False
   Exit Function
End If

validation = True

End Function

