VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmauthor 
   Caption         =   "Form3"
   ClientHeight    =   7680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5985
   LinkTopic       =   "Form3"
   ScaleHeight     =   7680
   ScaleWidth      =   5985
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   7455
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   855
         Left            =   360
         TabIndex        =   17
         Top             =   6120
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   1508
         _Version        =   393216
      End
      Begin VB.Frame Frame3 
         Height          =   1215
         Left            =   360
         TabIndex        =   11
         Top             =   4680
         Width           =   5055
         Begin VB.CommandButton Command5 
            Caption         =   "CANCEL"
            Height          =   435
            Left            =   3960
            TabIndex        =   16
            Top             =   360
            Width           =   795
         End
         Begin VB.CommandButton Command4 
            Caption         =   "DELETE"
            Height          =   435
            Left            =   2940
            TabIndex        =   15
            Top             =   360
            Width           =   795
         End
         Begin VB.CommandButton Command3 
            Caption         =   "EDIT"
            Height          =   435
            Left            =   2040
            TabIndex        =   14
            Top             =   360
            Width           =   675
         End
         Begin VB.CommandButton Command2 
            Caption         =   "UPDATE"
            Height          =   435
            Left            =   1020
            TabIndex        =   13
            Top             =   360
            Width           =   795
         End
         Begin VB.CommandButton Command1 
            Caption         =   "ADD"
            Height          =   435
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Width           =   675
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3615
         Left            =   360
         TabIndex        =   2
         Top             =   960
         Width           =   5055
         Begin VB.TextBox Text4 
            Height          =   495
            Left            =   1440
            TabIndex        =   10
            Top             =   2640
            Width           =   1695
         End
         Begin VB.TextBox Text3 
            Height          =   495
            Left            =   1440
            TabIndex        =   9
            Top             =   1800
            Width           =   1695
         End
         Begin VB.TextBox Text2 
            Height          =   495
            Left            =   1440
            TabIndex        =   8
            Top             =   1080
            Width           =   1695
         End
         Begin VB.TextBox Text1 
            Height          =   495
            Left            =   1440
            TabIndex        =   7
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label4 
            Caption         =   "Contactno"
            Height          =   495
            Left            =   240
            TabIndex        =   6
            Top             =   2760
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Website"
            Height          =   495
            Left            =   360
            TabIndex        =   5
            Top             =   1800
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Email"
            Height          =   495
            Left            =   360
            TabIndex        =   4
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Name"
            Height          =   495
            Left            =   360
            TabIndex        =   3
            Top             =   480
            Width           =   1575
         End
      End
      Begin VB.Label frmauthor 
         Alignment       =   2  'Center
         Caption         =   "AUTHOR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   720
         TabIndex        =   1
         Top             =   240
         Width           =   4455
      End
   End
End
Attribute VB_Name = "frmauthor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Frame2_DragDrop(Source As Control, X As Single, Y As Single)

End Sub
