VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.MDIForm MDIBookShop 
   BackColor       =   &H8000000C&
   Caption         =   "Bookshop Management System"
   ClientHeight    =   10515
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   14295
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   2760
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnuchange 
         Caption         =   "Change Password"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnumaster 
      Caption         =   "Master"
      Begin VB.Menu mnuCAt 
         Caption         =   "Catagory"
      End
      Begin VB.Menu mnusub 
         Caption         =   "Subject"
      End
      Begin VB.Menu mnuLang 
         Caption         =   "Language"
      End
      Begin VB.Menu nnuPubli 
         Caption         =   "Publisher"
      End
      Begin VB.Menu mnuAuhtor 
         Caption         =   "Author"
      End
      Begin VB.Menu mnubook 
         Caption         =   "Books"
         Begin VB.Menu mnubmaster 
            Caption         =   "Book Master"
         End
         Begin VB.Menu mnubkdetials 
            Caption         =   "Book Details"
         End
      End
      Begin VB.Menu mnuSatff 
         Caption         =   "Staff"
      End
      Begin VB.Menu mnuuser 
         Caption         =   "User Registration"
      End
   End
   Begin VB.Menu mnuTrans 
      Caption         =   "Transcations"
      Begin VB.Menu mnuSupplier 
         Caption         =   "Supplier"
      End
      Begin VB.Menu mnuCust 
         Caption         =   "Customer"
      End
      Begin VB.Menu mnupur 
         Caption         =   "Purchase"
      End
      Begin VB.Menu mnuSales 
         Caption         =   "Sales"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "Report"
      Begin VB.Menu mnurptbooks 
         Caption         =   "Book Details"
         Begin VB.Menu mnurptsub 
            Caption         =   "Subject-Wise"
         End
         Begin VB.Menu mnurptauthor 
            Caption         =   "Author-Wise"
         End
         Begin VB.Menu mnurptpub 
            Caption         =   "Publisher-Wise"
         End
         Begin VB.Menu mnurptcat 
            Caption         =   "Category-Wise"
         End
      End
      Begin VB.Menu mnurptsup 
         Caption         =   "Supplier Details"
      End
      Begin VB.Menu mnurptcus 
         Caption         =   "Customer Details"
      End
      Begin VB.Menu mnurptpur 
         Caption         =   "Purchase Details"
      End
      Begin VB.Menu mnurptsales 
         Caption         =   "Sales Details"
      End
   End
End
Attribute VB_Name = "MDIBookShop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuchange_Click()
frmchange.Show
End Sub

Private Sub mnuexit_Click()
Unload Me
End Sub

Private Sub cmdadmin_Click()
Admin.Show
End Sub

Private Sub mnuAuhtor_Click()
Author.Show
End Sub

Private Sub mnubkdetials_Click()
frmbookdetails.Show
End Sub

Private Sub mnubmaster_Click()
Bkdetails.Show
End Sub



Private Sub mnuCAt_Click()
category.Show
End Sub

Private Sub mnuCust_Click()
Customer.Show
End Sub

Private Sub mnuLang_Click()
language.Show
End Sub

Private Sub mnulogin_Click()
Login.Show
End Sub

Private Sub mnupur_Click()
purchase.Show
End Sub

Private Sub mnupurret_Click()
purchasereturn.Show
End Sub

Private Sub mnurptauthor_Click()
rptAuthor.Show
End Sub

Private Sub mnurptcat_Click()
rptcat.Show
End Sub

Private Sub mnurptcus_Click()
CrystalReport1.ReportFileName = App.Path & "/Report/CustomerDetails.rpt"
CrystalReport1.RetrieveDataFiles
CrystalReport1.WindowState = crptMaximized
CrystalReport1.Action = 1
End Sub

Private Sub mnurptpub_Click()
rptpub.Show
End Sub

Private Sub mnurptpur_Click()
rptPurchase.Show
End Sub

Private Sub mnurptsales_Click()
rptsales.Show
End Sub

Private Sub mnurptsub_Click()
rptsubwise.Show
End Sub

Private Sub mnurptsup_Click()
CrystalReport1.ReportFileName = App.Path & "/Report/SupplierDetails.rpt"
CrystalReport1.RetrieveDataFiles
CrystalReport1.WindowState = crptMaximized
CrystalReport1.Action = 1
End Sub

Private Sub mnurtsales_Click()
salesreturn.Show
End Sub

Private Sub mnuSales_Click()
Sales.Show
End Sub

Private Sub mnuSatff_Click()
Staff.Show
End Sub



Private Sub mnusub_Click()
subject.Show
End Sub

Private Sub mnuSupplier_Click()
Suplier.Show
End Sub

Private Sub mnuuser_Click()
frmloginreg.Show
End Sub

Private Sub nnuPubli_Click()
publisher.Show
End Sub
