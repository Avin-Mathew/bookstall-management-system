VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form author 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Author Details"
   ClientHeight    =   6570
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   8835
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   8835
   Begin VB.Frame Frame2 
      Height          =   3135
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   8295
      Begin VB.TextBox txtname 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   960
         TabIndex        =   10
         Top             =   360
         Width           =   2655
      End
      Begin VB.TextBox txtemail 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   960
         TabIndex        =   9
         Top             =   1080
         Width           =   2655
      End
      Begin VB.TextBox txtweb 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   5400
         TabIndex        =   8
         Top             =   2520
         Width           =   2655
      End
      Begin VB.OptionButton optmale 
         Caption         =   "Male"
         Height          =   255
         Left            =   5520
         TabIndex        =   7
         Top             =   480
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optfemale 
         Caption         =   "Female"
         Height          =   255
         Left            =   6480
         TabIndex        =   6
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtaward 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   960
         TabIndex        =   5
         Top             =   1680
         Width           =   2655
      End
      Begin VB.TextBox txtcno 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   5400
         TabIndex        =   4
         Top             =   1080
         Width           =   2655
      End
      Begin VB.TextBox txtcntry 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   5400
         TabIndex        =   3
         Top             =   1800
         Width           =   2655
      End
      Begin VB.TextBox txtabt 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   960
         TabIndex        =   2
         Top             =   2400
         Width           =   2655
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Email"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Website"
         Height          =   195
         Left            =   4560
         TabIndex        =   16
         Top             =   2520
         Width           =   585
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Contact No"
         Height          =   195
         Left            =   4560
         TabIndex        =   15
         Top             =   1080
         Width           =   810
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Gender"
         Height          =   195
         Left            =   4560
         TabIndex        =   14
         Top             =   480
         Width           =   525
      End
      Begin VB.Label Awards 
         Caption         =   "Awards"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Country"
         Height          =   195
         Left            =   4560
         TabIndex        =   12
         Top             =   1800
         Width           =   540
      End
      Begin VB.Label Label7 
         Caption         =   "About"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2400
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8655
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   120
         TabIndex        =   19
         Top             =   3480
         Width           =   8295
         Begin VB.CommandButton cmdadd 
            Caption         =   "&ADD"
            Height          =   375
            Left            =   600
            TabIndex        =   24
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdupdate 
            Caption         =   "&UPDATE"
            Height          =   375
            Left            =   1920
            TabIndex        =   23
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdedit 
            Caption         =   "&EDIT"
            Height          =   375
            Left            =   3240
            TabIndex        =   22
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmddelete 
            Caption         =   "&DELETE"
            Height          =   375
            Left            =   4560
            TabIndex        =   21
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdcancel 
            Caption         =   "&CANCEL"
            Height          =   375
            Left            =   6000
            TabIndex        =   20
            Top             =   240
            Width           =   975
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grddetails 
         Height          =   1815
         Left            =   120
         TabIndex        =   25
         Top             =   4320
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   3201
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "author"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdadd_Click()

cmdupdate.Enabled = True
cmdcancel.Enabled = True
Frame2.Enabled = True
cmdedit.Enabled = False
cmdadd.Enabled = False
cmddelete.Enabled = False

txtname.SetFocus
End Sub

Private Sub cmdcancel_Click()

txtname.Text = ""
txtname.Tag = ""
txtaward.Text = ""
txtabt.Text = ""
txtcntry.Text = ""
txtweb.Text = ""
txtcno.Text = ""
txtemail.Text = ""

cmdadd.Enabled = True
cmddelete.Enabled = False
cmdedit.Enabled = False
cmdupdate.Enabled = False
cmdcancel.Enabled = True
Frame2.Enabled = False

End Sub

Private Sub cmddelete_Click()
If rsatr.State = 1 Then rsatr.Close
rsatr.Open "select * from tbl_author where atr_id=" & txtname.Tag & "", con, adOpenKeyset, adLockOptimistic
rsatr.Delete
rsatr.Update
rsatr.Close
txtname.Tag = ""
fillgrid
cmddelete.Enabled = False
cmdadd.Enabled = True
cmdedit.Enabled = False
cmdupdate.Enabled = False

txtname.Text = ""
txtcno.Text = ""
txtemail.Text = ""
txtweb.Text = ""
txtaward.Text = ""
txtabt.Text = ""
txtcntry.Text = ""

End Sub

Private Sub cmdedit_Click()
Frame2.Enabled = True
cmdupdate.Enabled = True
cmddelete.Enabled = False
cmdedit.Enabled = False
cmdadd.Enabled = False
txtname.SetFocus
End Sub

Private Sub cmdupdate_Click()
If vallidation Then

If txtname.Tag = "" Then

    If rsatr.State = 1 Then rsatr.Close
      rsatr.Open "select * from tbl_author", con, adOpenKeyset, adLockOptimistic
      rsatr.AddNew
      rsatr.Fields("atr_name") = txtname.Text
      rsatr.Fields("atr_email") = txtemail.Text
      rsatr.Fields("atr_website") = txtweb.Text
      rsatr.Fields("atr_contactno") = txtcno.Text
      rsatr.Fields("atr_awards") = txtaward.Text
      rsatr.Fields("atr_about") = txtabt.Text
      rsatr.Fields("atr_country") = txtcntry.Text
      If optmale.Value = True Then
          rsatr.Fields("atr_gender") = "Male"
      Else
          rsatr.Fields("atr_gender") = "Female"
      End If

      rsatr.Update
      rsatr.Close
      
  txtname.Text = ""

txtaward.Text = ""
txtabt.Text = ""
txtcntry.Text = ""
txtweb.Text = ""
txtcno.Text = ""
txtemail.Text = ""

cmdadd.Enabled = True
cmddelete.Enabled = False
cmdedit.Enabled = False
cmdupdate.Enabled = False
cmdcancel.Enabled = True
Frame2.Enabled = False


Else
If rsatr.State = 1 Then rsatr.Close
rsatr.Open "select * from tbl_author where atr_id=" & txtname.Tag & "", con, adOpenKeyset, adLockOptimistic
rsatr.Fields("atr_name") = txtname.Text
rsatr.Fields("atr_email") = txtemail.Text
rsatr.Fields("atr_website") = txtweb.Text
rsatr.Fields("atr_contactno") = txtcno.Text
rsatr.Fields("atr_awards") = txtaward.Text
rsatr.Fields("atr_about") = txtabt.Text
rsatr.Fields("atr_country") = txtcntry.Text
If optmale.Value = True Then
rsatr.Fields("atr_gender") = "Male"
Else
rsatr.Fields("atr_gender") = "Female"
End If
rsatr.Update
rsatr.Close

txtname.Text = ""
txtname.Tag = ""
txtaward.Text = ""
txtabt.Text = ""
txtcntry.Text = ""
txtweb.Text = ""
txtcno.Text = ""
txtemail.Text = ""

cmdadd.Enabled = True
cmddelete.Enabled = False
cmdedit.Enabled = False
cmdupdate.Enabled = False
cmdcancel.Enabled = True
Frame2.Enabled = False


End If


End If

fillgrid

End Sub

Private Sub Form_Load()
Frame2.Enabled = False
cmdadd.Enabled = True
cmdupdate.Enabled = False
cmdedit.Enabled = False
cmddelete.Enabled = False
cmdcancel.Enabled = True
fillgrid
End Sub


Private Sub grddetails_Click()
If rsatr.State = 1 Then rsatr.Close
rsatr.Open "select * from tbl_author where atr_id=" & grddetails.TextMatrix(grddetails.row, 0) & " ", con, adOpenKeyset, adLockOptimistic
If rsatr.RecordCount > 0 Then
txtname.Text = rsatr.Fields("atr_name")
txtemail.Text = rsatr.Fields("atr_email")
txtweb.Text = rsatr.Fields("atr_website")
txtcno.Text = rsatr.Fields("atr_contactno")
txtname.Tag = rsatr.Fields("atr_id")
txtabt.Text = rsatr.Fields("atr_about")
txtaward.Text = rsatr.Fields("atr_awards")
txtcntry.Text = rsatr.Fields("atr_country")
If rsatr.Fields("atr_gender") = "Male" Then
optmale.Value = True
Else
optfemale.Value = True

End If
txtname.Tag = rsatr.Fields("atr_id")
End If
cmdedit.Enabled = True
cmddelete.Enabled = True
cmdadd.Enabled = False
cmdupdate.Enabled = False
End Sub

Public Function fillgrid()
grddetails.Rows = 1
grddetails.Cols = 10


grddetails.ColWidth(0) = 0
grddetails.ColWidth(1) = 1700
grddetails.ColWidth(2) = 1700
grddetails.ColWidth(3) = 1700
grddetails.ColWidth(5) = 1700
grddetails.ColWidth(6) = 1700
grddetails.ColWidth(7) = 1700
grddetails.ColWidth(8) = 1700


grddetails.TextMatrix(0, 1) = " Name"
grddetails.TextMatrix(0, 2) = "Website"
grddetails.TextMatrix(0, 3) = "Contact no"
grddetails.TextMatrix(0, 4) = "Email"
grddetails.TextMatrix(0, 5) = "Awards"
grddetails.TextMatrix(0, 6) = "About"
grddetails.TextMatrix(0, 7) = "Country"
grddetails.TextMatrix(0, 8) = "Gender"

If rsatr.State = 1 Then rsatr.Close
rsatr.Open "select * from tbl_author", con, adOpenKeyset, adLockOptimistic
If rsatr.RecordCount > 0 Then
rsatr.MoveFirst
While Not rsatr.EOF
grddetails.Rows = grddetails.Rows + 1
grddetails.TextMatrix(grddetails.Rows - 1, 0) = rsatr.Fields("atr_id")
grddetails.TextMatrix(grddetails.Rows - 1, 1) = rsatr.Fields("atr_name")
grddetails.TextMatrix(grddetails.Rows - 1, 2) = rsatr.Fields("atr_website")
grddetails.TextMatrix(grddetails.Rows - 1, 3) = rsatr.Fields("atr_contactno")
grddetails.TextMatrix(grddetails.Rows - 1, 4) = rsatr.Fields("atr_email")
grddetails.TextMatrix(grddetails.Rows - 1, 5) = rsatr.Fields("atr_Awards")
grddetails.TextMatrix(grddetails.Rows - 1, 6) = rsatr.Fields("atr_About")
grddetails.TextMatrix(grddetails.Rows - 1, 7) = rsatr.Fields("atr_Country")
grddetails.TextMatrix(grddetails.Rows - 1, 8) = rsatr.Fields("atr_Gender")


rsatr.MoveNext
Wend
End If
End Function

Public Function vallidation() As Boolean

If Trim(txtname.Text) = "" Then
   MsgBox "Enter a name", vbInformation, App.Title
   txtname.SetFocus
   vallidation = False
   Exit Function
End If

If Trim(txtemail.Text) = "" Then
  MsgBox "Enter the Email id", vbInformation, App.Title
  txtemail.SetFocus
  vallidation = False
  Exit Function
End If

If Trim(txtaward.Text) = "" Then
   MsgBox "Enter the Awards", vbInformation, App.Title
   txtaward.SetFocus
   vallidation = False
   Exit Function
 End If
 
If Trim(txtabt.Text) = "" Then
   MsgBox "Write a few words about the author", vbInformation, App.Title
   txtabt.SetFocus
   vallidation = False
   Exit Function
End If
If Trim(txtcno.Text) = "" Then
   MsgBox "Enter the contact no", vbInformation, App.Title
   txtcno.SetFocus
   vallidation = False
   Exit Function
End If
If Trim(txtcntry.Text) = "" Then
   MsgBox "Enter the country", vbInformation, App.Title
   txtcntry.SetFocus
   vallidation = False
   Exit Function
 End If
 If Trim(txtweb.Text) = "" Then
   MsgBox "Enter the website", vbInformation, App.Title
   txtweb.SetFocus
   vallidation = False
   Exit Function
 End If
   

vallidation = True

End Function
