VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmcategory 
   Caption         =   "form3"
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5610
   LinkTopic       =   "Form3"
   ScaleHeight     =   6195
   ScaleWidth      =   5610
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   1455
         Left            =   240
         TabIndex        =   11
         Top             =   3960
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   2566
         _Version        =   393216
      End
      Begin VB.Frame Frame3 
         Height          =   975
         Left            =   240
         TabIndex        =   5
         Top             =   2880
         Width           =   5055
         Begin VB.CommandButton Command5 
            Caption         =   "CANCEL"
            Height          =   495
            Left            =   3960
            TabIndex        =   10
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Command4 
            Caption         =   "DELETE"
            Height          =   495
            Left            =   3000
            TabIndex        =   9
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Command3 
            Caption         =   "EDIT"
            Height          =   495
            Left            =   2040
            TabIndex        =   8
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Command2 
            Caption         =   "UPDATE"
            Height          =   495
            Left            =   1080
            TabIndex        =   7
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Command1 
            Caption         =   "ADD"
            Height          =   495
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1695
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   5055
         Begin VB.TextBox Text1 
            Height          =   495
            Left            =   2280
            TabIndex        =   4
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label2 
            Caption         =   "Category"
            Height          =   615
            Left            =   600
            TabIndex        =   3
            Top             =   600
            Width           =   1095
         End
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "CATEGORY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   4095
      End
   End
End
Attribute VB_Name = "frmcategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MSFlexGrid1_Click()

End Sub
