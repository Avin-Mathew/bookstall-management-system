VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form subject 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Subject"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   6690
   Begin VB.Frame Frame1 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      Begin VB.Frame Frame4 
         Height          =   855
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   6135
         Begin VB.CommandButton cmdadd 
            Caption         =   "&ADD"
            Height          =   375
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdupdate 
            Caption         =   "&UPDATE"
            Height          =   375
            Left            =   1320
            TabIndex        =   8
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdedit 
            Caption         =   "&EDIT"
            Height          =   375
            Left            =   2520
            TabIndex        =   7
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmddelete 
            Caption         =   "&DELETE"
            Height          =   375
            Left            =   3720
            TabIndex        =   6
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdcancel 
            Caption         =   "&CANCEL"
            Height          =   375
            Left            =   4920
            TabIndex        =   5
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1215
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6135
         Begin VB.TextBox txtname 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1320
            TabIndex        =   3
            Top             =   360
            Width           =   3735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Name"
            Height          =   195
            Left            =   360
            TabIndex        =   2
            Top             =   480
            Width           =   420
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grddetails 
         Height          =   1815
         Left            =   120
         TabIndex        =   10
         Top             =   2520
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   3201
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "subject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsssub As New ADODB.Recordset

Private Sub cmdadd_Click()
cmdupdate.Enabled = True
cmdcancel.Enabled = True
cmdadd.Enabled = False
cmdedit.Enabled = False
cmddelete.Enabled = False
Frame2.Enabled = True
cmdedit.Enabled = False
cmddelete.Enabled = False
txtname.SetFocus
End Sub

Private Sub cmdcancel_Click()
txtname.Text = " "
txtname.Tag = ""
cmdadd.Enabled = True
cmdupdate.Enabled = False
cmdedit.Enabled = False
cmddelete.Enabled = False
Frame2.Enabled = False
End Sub

Private Sub cmddelete_Click()
If rssub.State = 1 Then rssub.Close
  rssub.Open "select * from tbl_subject where sub_id=" & txtname.Tag & "", con, adOpenKeyset, adLockOptimistic
  rssub.Delete
  rssub.Update
  rssub.Close
  
txtname.Tag = ""
txtname.Text = ""

fillgrid

Frame2.Enabled = False
cmddelete.Enabled = False
cmdadd.Enabled = True
cmdedit.Enabled = False

End Sub

Private Sub cmdedit_Click()
Frame2.Enabled = True
cmdupdate.Enabled = True
cmddelete.Enabled = False
cmdedit.Enabled = False
txtname.SetFocus
End Sub

Private Sub cmdupdate_Click()

If vallidation Then

If txtname.Tag = "" Then
   If rssub.State = 1 Then rssub.Close
   rssub.Open "select * from tbl_subject", con, adOpenKeyset, adLockOptimistic
   rssub.AddNew
   rssub.Fields("sub_name") = txtname.Text
   rssub.Update
   rssub.Close
   
   txtname.Text = " "

cmdadd.Enabled = True
cmdupdate.Enabled = False
cmdedit.Enabled = False
cmddelete.Enabled = False
Frame2.Enabled = False


Else
   If rssub.State = 1 Then rssub.Close
   rssub.Open "select * from tbl_subject where sub_id=" & txtname.Tag & "", con, adOpenKeyset, adLockOptimistic
   rssub.Fields("sub_name") = txtname.Text
   rssub.Update
   rssub.Close
   
   txtname.Text = " "
txtname.Tag = ""
cmdadd.Enabled = True
cmdupdate.Enabled = False
cmdedit.Enabled = False
cmddelete.Enabled = False
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
If rssub.State = 1 Then rssub.Close
rssub.Open "select * from tbl_subject where sub_id=" & grddetails.TextMatrix(grddetails.row, 0) & "", con, adOpenKeyset, adLockOptimistic
If rssub.RecordCount > 0 Then
txtname.Text = rssub.Fields("sub_name")
txtname.Tag = rssub.Fields("sub_id")
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

grddetails.TextMatrix(0, 1) = "Subject"
If rssub.State = 1 Then rssub.Close
rssub.Open "select * from tbl_subject", con, adOpenKeyset, adLockOptimistic
If rssub.RecordCount > 0 Then
rssub.MoveFirst
While Not rssub.EOF
grddetails.Rows = grddetails.Rows + 1
grddetails.TextMatrix(grddetails.Rows - 1, 0) = rssub.Fields("sub_id")
grddetails.TextMatrix(grddetails.Rows - 1, 1) = rssub.Fields("sub_name")
rssub.MoveNext
Wend
End If

End Function

Public Function vallidation() As Boolean
If Trim(txtname.Text) = "" Then
   MsgBox "Enter the subject", vbInformation, App.Title
   txtname.SetFocus
   vallidation = False
    Exit Function
End If

vallidation = True

End Function
