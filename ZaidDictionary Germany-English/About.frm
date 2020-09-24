VERSION 5.00
Begin VB.Form About 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   Picture         =   "About.frx":0000
   ScaleHeight     =   3450
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "About Zaid Dictionary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   120
      Width           =   3615
   End
   Begin VB.Image TransFormButton1 
      Height          =   465
      Left            =   3480
      Picture         =   "About.frx":35EC2
      Top             =   2640
      Width           =   1140
   End
   Begin VB.Image Button1 
      Height          =   240
      Left            =   4200
      Picture         =   "About.frx":37AA0
      Top             =   120
      Width           =   345
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Include 45200 entries"
      ForeColor       =   &H008080FF&
      Height          =   255
      Left            =   3000
      TabIndex        =   5
      Top             =   1000
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright - 2008"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Em@l : zaidmarkabi@yahoo.com"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   2280
      Width           =   3735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Powered with VB6 by ZaidMarkabi"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   2040
      Width           =   3735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "You can use it any time you want ."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   1680
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "This application is freeware for all ."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   1440
      Width           =   3735
   End
   Begin VB.Image Image2 
      Height          =   765
      Left            =   135
      Picture         =   "About.frx":37F62
      Stretch         =   -1  'True
      Top             =   495
      Width           =   4530
   End
   Begin VB.Image Image1 
      Height          =   435
      Left            =   240
      Picture         =   "About.frx":43554
      Top             =   1320
      Width           =   480
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Button1_Click()
Unload Me
End Sub

Private Sub TransFormButton1_Click()
Unload Me
End Sub
