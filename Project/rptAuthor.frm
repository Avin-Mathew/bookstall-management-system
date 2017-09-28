VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rptAuthor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Author-Wise"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   6285
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      Begin VB.CommandButton Command1 
         Caption         =   "View"
         Height          =   375
         Left            =   3840
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox cmbauthor 
         Height          =   315
         Left            =   1200
         TabIndex        =   1
         Top             =   360
         Width           =   2415
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
         Caption         =   "Author"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "rptAuthor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
If vallidation Then
CrystalReport1.ReportFileName = App.Path & "/Report/AuthorWise.rpt"
CrystalReport1.SelectionFormula = "{tbl_bookmaster.atr_id}=" & CboData(cmbauthor) & ""
CrystalReport1.RetrieveDataFiles
CrystalReport1.WindowState = crptMaximized
CrystalReport1.Action = 1
End If

End Sub

Private Sub Form_Load()
Fillcombo "tbl_author", cmbauthor, "atr_name", "atr_id"
End Sub

Public Function vallidation() As Boolean

If Trim(cmbauthor.Text) = "--select--" Then
   MsgBox "Select Author", vbInformation, App.Title
   cmbauthor.SetFocus
   vallidation = False
   Exit Function
End If

vallidation = True

End Function
