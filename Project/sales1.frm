VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form category 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CATEGORY"
   ClientHeight    =   4095
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   7035
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   7035
   Begin VB.Frame Frame1 
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6975
      Begin VB.Frame Frame2 
         Height          =   1095
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   6735
         Begin VB.TextBox txtcname 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1080
            TabIndex        =   9
            Top             =   360
            Width           =   2535
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Catgeory"
            Height          =   195
            Left            =   240
            TabIndex        =   10
            Top             =   360
            Width           =   630
         End
      End
      Begin VB.Frame Frame3 
         Height          =   855
         Left            =   120
         TabIndex        =   1
         Top             =   1320
         Width           =   6735
         Begin VB.CommandButton cmdadd 
            Caption         =   "&ADD"
            Height          =   495
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdupdate 
            Caption         =   "&UPDATE"
            Height          =   495
            Left            =   1440
            TabIndex        =   5
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdedit 
            Caption         =   "&EDIT"
            Height          =   495
            Left            =   2760
            TabIndex        =   4
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmddelete 
            Caption         =   "&DELETE"
            Height          =   495
            Left            =   4080
            TabIndex        =   3
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdcancel 
            Caption         =   "&CANCEL"
            Height          =   495
            Left            =   5400
            TabIndex        =   2
            Top             =   240
            Width           =   1095
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grddetails 
         Height          =   1455
         Left            =   120
         TabIndex        =   7
         Top             =   2280
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   2566
         _Version        =   393216
         Rows            =   3
      End
   End
End
Attribute VB_Name = "category"
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
txtcname.SetFocus
End Sub

Private Sub cmdcancel_Click()
txtcname.Text = " "
txtcname.Tag = ""
Frame2.Enabled = False
cmdadd.Enabled = True
cmdedit.Enabled = False
cmddelete.Enabled = False
cmdupdate.Enabled = False
End Sub

Private Sub cmddelete_Click()
If rsCat.State = 1 Then rsCat.Close
rsCat.Open "select * from tbl_category where c_id=" & txtcname.Tag & "", con, adOpenKeyset, adLockOptimistic
rsCat.Delete
rsCat.Update
rsCat.Close

txtcname.Text = ""
txtcname.Tag = ""

fillgrid

cmddelete.Enabled = False
cmdadd.Enabled = True
cmdedit.Enabled = False

End Sub

Private Sub cmdedit_Click()

Frame2.Enabled = True
cmdupdate.Enabled = True
cmddelete.Enabled = False
cmdedit.Enabled = False
txtcname.SetFocus

End Sub


Private Sub cmdupdate_Click()
If vallidation Then
If (txtcname.Tag = "") Then
   If rsCat.State = 1 Then rsCat.Close
   rsCat.Open "select * from tbl_category", con, adOpenKeyset, adLockOptimistic
   rsCat.AddNew
   rsCat.Fields("c_name") = txtcname.Text
   rsCat.Update
   rsCat.Close
   
   txtcname.Text = " "
   Frame2.Enabled = False
   cmdadd.Enabled = True
   cmdedit.Enabled = False
   cmddelete.Enabled = False
   cmdupdate.Enabled = False

Else
If rsCat.State = 1 Then rsCat.Close
rsCat.Open "select * from tbl_category where c_id=" & txtcname.Tag & "", con, adOpenKeyset, adLockOptimistic
rsCat.Fields("c_name") = txtcname.Text
rsCat.Update
rsCat.Close

txtcname.Text = " "
txtcname.Tag = ""
Frame2.Enabled = False
cmdadd.Enabled = True
cmdedit.Enabled = False
cmddelete.Enabled = False
cmdupdate.Enabled = False

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

Public Function fillgrid()
grddetails.Rows = 1
grddetails.Cols = 2

grddetails.ColWidth(0) = 0
grddetails.ColWidth(1) = 1700

grddetails.TextMatrix(0, 1) = "Category Name"
If rsCat.State = 1 Then rsCat.Close
rsCat.Open "select * from tbl_category", con, adOpenKeyset, adLockOptimistic
If rsCat.RecordCount > 0 Then
rsCat.MoveFirst
While Not rsCat.EOF
grddetails.Rows = grddetails.Rows + 1
grddetails.TextMatrix(grddetails.Rows - 1, 0) = rsCat.Fields("c_id")
grddetails.TextMatrix(grddetails.Rows - 1, 1) = rsCat.Fields("c_name")
rsCat.MoveNext
Wend
End If


End Function


Private Sub grddetails_Click()
If rsCat.State = 1 Then rsCat.Close
rsCat.Open "select * from tbl_category where c_id=" & grddetails.TextMatrix(grddetails.row, 0) & "", con, adOpenKeyset, adLockOptimistic
If rsCat.RecordCount > 0 Then
txtcname.Text = rsCat.Fields("c_name")
txtcname.Tag = rsCat.Fields("c_id")
End If
cmdedit.Enabled = True
cmddelete.Enabled = True
cmdadd.Enabled = False
End Sub


Public Function vallidation() As Boolean
If Trim(txtcname.Text) = "" Then
   MsgBox "Enter the subject", vbInformation, App.Title
   txtcname.SetFocus
   vallidation = False
    Exit Function
End If

vallidation = True

End Function
