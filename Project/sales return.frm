VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form salesreturn 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SALES RETURN"
   ClientHeight    =   6825
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   6780
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   6615
      Begin VB.TextBox txtbillno 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1560
         TabIndex        =   17
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "View"
         Height          =   375
         Left            =   4080
         TabIndex        =   16
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Bil lNumber"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1815
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   6615
      Begin VB.Label Label2 
         Caption         =   "Customer"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblname 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   960
         TabIndex        =   13
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Date"
         Height          =   375
         Left            =   3360
         TabIndex        =   12
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lbldate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4320
         TabIndex        =   11
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label lblamount 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   960
         TabIndex        =   10
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label7 
         Caption         =   "Pay Type"
         Height          =   255
         Left            =   3360
         TabIndex        =   9
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label lbltype 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4320
         TabIndex        =   8
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "Amount"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2775
      Left            =   120
      TabIndex        =   4
      Top             =   3120
      Width           =   6615
      Begin MSFlexGridLib.MSFlexGrid grddetails 
         Height          =   2295
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   4048
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   5880
      Width           =   6615
      Begin VB.CommandButton cmdretrun 
         Caption         =   "Return"
         Height          =   375
         Left            =   3600
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   255
         Left            =   1560
         TabIndex        =   1
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         _Version        =   393216
         Format          =   91619329
         CurrentDate     =   41538
      End
      Begin VB.Label Label4 
         Caption         =   "Return Date"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "salesreturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tempphead As New ADODB.Recordset
Dim temppdet As New ADODB.Recordset
Dim rssalretrun As New ADODB.Recordset
Dim rsbookstock As New ADODB.Recordset
Dim row As Integer

Private Sub cmdretrun_Click()

Dim phid As Integer
Dim amount As Double
Dim bid As Integer


If rssalretrun.State = 1 Then rssalretrun.Close
rssalretrun.Open "select * from tbl_salesreturn", con, adOpenKeyset, adLockOptimistic
rssalretrun.AddNew
rssalretrun.Fields("sd_id") = row
rssalretrun.Fields("sret_date") = DTPicker1.Value
rssalretrun.Update
rssalretrun.Close

Frame4.Visible = False
'row = 0

If temppdet.State = 1 Then temppdet.Close
temppdet.Open "select * from tbl_salesdetails  where sd_id=" & row & "", con, adOpenKeyset, adLockOptimistic

  phid = temppdet.Fields("sh_id")
  amount = temppdet.Fields("sd_total")
  bid = temppdet.Fields("bd_id")

  temppdet.Delete
  temppdet.Update
  temppdet.Close




If rsbookstock.State = 1 Then rsbookstock.Close
rsbookstock.Open "select * from tbl_bookdetails  where bd_id=" & bid & "", con, adOpenKeyset, adLockOptimistic
rsbookstock.Fields("bd_status") = 1
rsbookstock.Update
rsbookstock.Close



If tempphead.State = 1 Then tempphead.Close
tempphead.Open "select * from tbl_salesheader where sh_id=" & phid & " ", con, adOpenKeyset, adLockOptimistic
If tempphead.RecordCount > 0 Then
 tempphead.Fields("sh_grand_total") = Val(tempphead.Fields("sh_grand_total") - amount)
 tempphead.Update
 tempphead.Close

End If


If tempphead.State = 1 Then tempphead.Close
tempphead.Open "select * from tbl_salesheader p,tbl_customerdetails s where p.cst_id=s.cst_id and p.sh_billno='" & txtbillno.Text & "' ", con, adOpenKeyset, adLockOptimistic
If tempphead.RecordCount > 0 Then
 lblname.Caption = tempphead.Fields("cst_name")
 lbldate.Caption = tempphead.Fields("sh_date")
 lblamount.Caption = tempphead.Fields("sh_grand_total")
 lbltype.Caption = tempphead.Fields("sh_paymenttype")
  
  fillgrid
  Frame2.Visible = True
  Frame3.Visible = True
 
  
Else
  Frame2.Visible = False
  Frame3.Visible = False
  Frame4.Visible = False
  MsgBox "No Record Found..."
End If
tempphead.Close







End Sub

Private Sub cmdView_Click()
If tempphead.State = 1 Then tempphead.Close
tempphead.Open "select * from tbl_salesheader p,tbl_customerdetails s where p.cst_id=s.cst_id and p.sh_billno='" & txtbillno.Text & "' ", con, adOpenKeyset, adLockOptimistic
If tempphead.RecordCount > 0 Then
 lblname.Caption = tempphead.Fields("cst_name")
 lbldate.Caption = tempphead.Fields("sh_date")
 lblamount.Caption = tempphead.Fields("sh_grand_total")
 lbltype.Caption = tempphead.Fields("sh_paymenttype")
  
  fillgrid
  Frame2.Visible = True
  Frame3.Visible = True
 
  
Else
  Frame2.Visible = False
  Frame3.Visible = False
  Frame4.Visible = False
  MsgBox "No Record Found..."
End If
tempphead.Close
End Sub

Private Sub Form_Load()
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = False

End Sub

Private Sub fillgrid()
grddetails.Rows = 1
grddetails.Cols = 6

grddetails.ColWidth(0) = 0
grddetails.ColWidth(1) = 1400
grddetails.ColWidth(2) = 1400
grddetails.ColWidth(3) = 1400
grddetails.ColWidth(4) = 1400
grddetails.ColWidth(5) = 1400

grddetails.TextMatrix(0, 1) = "Book"
grddetails.TextMatrix(0, 2) = "Quantity"
grddetails.TextMatrix(0, 3) = "Rate"
grddetails.TextMatrix(0, 4) = "Total"
grddetails.TextMatrix(0, 5) = "ISBN No"

If temppdet.State = 1 Then temppdet.Close
temppdet.Open "select * from tbl_bookmaster b,tbl_salesheader sh,tbl_salesdetails sd,tbl_bookdetails bd where b.book_id=bd.book_id and sh.sh_id=sd.sh_id and bd.bd_status='True' and  sh.sh_billno='" & txtbillno.Text & "' ", con, adOpenKeyset, adLockOptimistic
If temppdet.RecordCount > 0 Then
  temppdet.MoveFirst
While Not temppdet.EOF
   grddetails.Rows = grddetails.Rows + 1
   grddetails.TextMatrix(grddetails.Rows - 1, 0) = temppdet.Fields("sd_id")
   grddetails.TextMatrix(grddetails.Rows - 1, 1) = temppdet.Fields("book_name")
   grddetails.TextMatrix(grddetails.Rows - 1, 2) = temppdet.Fields("b_qty")
   grddetails.TextMatrix(grddetails.Rows - 1, 3) = temppdet.Fields("book_price")
   grddetails.TextMatrix(grddetails.Rows - 1, 4) = temppdet.Fields("sd_total")
   grddetails.TextMatrix(grddetails.Rows - 1, 5) = temppdet.Fields("ISBNno")
   temppdet.MoveNext
Wend
End If

temppdet.Close
 
End Sub

Private Sub grddetails_DblClick()
 Frame4.Visible = True
 row = grddetails.TextMatrix(grddetails.row, 0)
 
End Sub
