VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form purchasereturn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PURCHASE RETURN"
   ClientHeight    =   6855
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   6840
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   120
      TabIndex        =   15
      Top             =   5880
      Width           =   6615
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   255
         Left            =   1560
         TabIndex        =   18
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         _Version        =   393216
         Format          =   91619329
         CurrentDate     =   41538
      End
      Begin VB.CommandButton cmdretrun 
         Caption         =   "Return"
         Height          =   375
         Left            =   3600
         TabIndex        =   16
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Return Date"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2775
      Left            =   120
      TabIndex        =   13
      Top             =   3120
      Width           =   6615
      Begin MSFlexGridLib.MSFlexGrid grddetails 
         Height          =   2295
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   4048
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1815
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   6615
      Begin VB.Label Label5 
         Caption         =   "Amount"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label lbltype 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4320
         TabIndex        =   11
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label7 
         Caption         =   "Pay Type"
         Height          =   255
         Left            =   3360
         TabIndex        =   10
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label lblamount 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   960
         TabIndex        =   9
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label lbldate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4320
         TabIndex        =   8
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Date"
         Height          =   375
         Left            =   3360
         TabIndex        =   7
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblname 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   960
         TabIndex        =   6
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Supplier"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      Begin VB.CommandButton cmdView 
         Caption         =   "View"
         Height          =   375
         Left            =   4080
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtinvoice 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Invoice Number"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "purchasereturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tempphead As New ADODB.Recordset
Dim temppdet As New ADODB.Recordset
Dim rspurretrun As New ADODB.Recordset
Dim rsbookstock As New ADODB.Recordset
Dim row As Integer

Private Sub cmdretrun_Click()

Dim phid As Integer
Dim amount As Double
Dim bid As Integer
Dim bisb As String

If rspurretrun.State = 1 Then rspurretrun.Close
rspurretrun.Open "select * from tbl_purchasereturn", con, adOpenKeyset, adLockOptimistic
rspurretrun.AddNew
rspurretrun.Fields("pd_id") = row
rspurretrun.Fields("pret_date") = DTPicker1.Value
rspurretrun.Update
rspurretrun.Close

Frame4.Visible = False
'row = 0

If temppdet.State = 1 Then temppdet.Close
temppdet.Open "select * from tbl_purchasedetails  where pd_id=" & row & "", con, adOpenKeyset, adLockOptimistic
'  temppdet.Fields("pd_status") = 1
  phid = temppdet.Fields("ph_id")
  amount = temppdet.Fields("book_amount")
  bid = temppdet.Fields("book_id")
  bisb = temppdet.Fields("book_isbnno")
  temppdet.Delete
  temppdet.Update
  temppdet.Close




If rsbookstock.State = 1 Then rsbookstock.Close
rsbookstock.Open "select * from tbl_bookdetails  where ISBNno='" & bisb & "' and book_id=" & bid & "", con, adOpenKeyset, adLockOptimistic

  rsbookstock.Delete
  rsbookstock.Update
  rsbookstock.Close



If tempphead.State = 1 Then tempphead.Close
tempphead.Open "select * from tbl_purchasehead where ph_id=" & phid & " ", con, adOpenKeyset, adLockOptimistic
If tempphead.RecordCount > 0 Then
 tempphead.Fields("ph_grand_total") = Val(tempphead.Fields("ph_grand_total") - amount)
 tempphead.Update
 tempphead.Close
 
 

End If







If tempphead.State = 1 Then tempphead.Close
tempphead.Open "select * from tbl_purchasehead p,tbl_supplier s where p.sup_id=s.sup_id and p.invoice_no='" & txtinvoice.Text & "' ", con, adOpenKeyset, adLockOptimistic
If tempphead.RecordCount > 0 Then
 lblname.Caption = tempphead.Fields("sup_name")
 lbldate.Caption = tempphead.Fields("ph_date")
 lblamount.Caption = tempphead.Fields("ph_grand_total")
 lbltype.Caption = tempphead.Fields("ph_paytype")
  
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

Private Sub cmdview_Click()

If tempphead.State = 1 Then tempphead.Close
tempphead.Open "select * from tbl_purchasehead p,tbl_supplier s where p.sup_id=s.sup_id and p.invoice_no='" & txtinvoice.Text & "' ", con, adOpenKeyset, adLockOptimistic
If tempphead.RecordCount > 0 Then
 lblname.Caption = tempphead.Fields("sup_name")
 lbldate.Caption = tempphead.Fields("ph_date")
 lblamount.Caption = tempphead.Fields("ph_grand_total")
 lbltype.Caption = tempphead.Fields("ph_paytype")
  
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
temppdet.Open "select * from tbl_bookmaster b,tbl_purchasehead ph,tbl_purchasedetails pd where b.book_id=pd.book_id and ph.ph_id=pd.ph_id  and  ph.invoice_no='" & txtinvoice.Text & "' ", con, adOpenKeyset, adLockOptimistic
If temppdet.RecordCount > 0 Then
  temppdet.MoveFirst
While Not temppdet.EOF
   grddetails.Rows = grddetails.Rows + 1
   grddetails.TextMatrix(grddetails.Rows - 1, 0) = temppdet.Fields("pd_id")
   grddetails.TextMatrix(grddetails.Rows - 1, 1) = temppdet.Fields("book_name")
   grddetails.TextMatrix(grddetails.Rows - 1, 2) = temppdet.Fields("book_qty")
   grddetails.TextMatrix(grddetails.Rows - 1, 3) = temppdet.Fields("book_unitprice")
   grddetails.TextMatrix(grddetails.Rows - 1, 4) = temppdet.Fields("book_amount")
   grddetails.TextMatrix(grddetails.Rows - 1, 5) = temppdet.Fields("book_isbnno")
   temppdet.MoveNext
Wend
End If

temppdet.Close
 
End Sub

Private Sub Form_Load()

Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = False

End Sub





Private Sub grddetails_DblClick()
 Frame4.Visible = True
 row = grddetails.TextMatrix(grddetails.row, 0)
 
End Sub
