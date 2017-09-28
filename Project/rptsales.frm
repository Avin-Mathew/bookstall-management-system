VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rptsales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales Details"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   6270
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      Begin VB.TextBox txtbill 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   600
         Width           =   2535
      End
      Begin VB.CommandButton cmdview 
         Caption         =   "View"
         Height          =   375
         Left            =   4320
         TabIndex        =   1
         Top             =   600
         Width           =   615
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   1200
         Top             =   960
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.Label Label1 
         Caption         =   "Bill Number"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   1215
      End
   End
End
Attribute VB_Name = "rptsales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsrpt As New ADODB.Recordset

Private Sub cmdview_Click()


If vallidation Then

If rsrpt.State = 1 Then rsrpt.Close
   rsrpt.Open "select * from tbl_salesheader where sh_billno='" & txtbill.Text & "'", con, adOpenKeyset, adLockOptimistic
If rsrpt.RecordCount > 0 Then
CrystalReport1.ReportFileName = App.Path & "/Report/SalesDetails.rpt"
CrystalReport1.SelectionFormula = "{tbl_salesheader.sh_billno}='" & txtbill.Text & "'"
CrystalReport1.RetrieveDataFiles
CrystalReport1.WindowState = crptMaximized
CrystalReport1.Action = 1
rsrpt.Close
Else
txtbill.Text = ""
MsgBox "No Record Found.."

End If

End If

End Sub

Public Function vallidation() As Boolean

If Trim(txtbill.Text) = "" Then
   MsgBox "Enter Bill Number", vbInformation, App.Title
   txtbill.SetFocus
   vallidation = False
   Exit Function
End If

vallidation = True

End Function


