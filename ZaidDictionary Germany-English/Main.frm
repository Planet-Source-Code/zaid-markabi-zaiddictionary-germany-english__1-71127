VERSION 5.00
Begin VB.Form Main 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1305
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   ScaleHeight     =   1305
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   1320
      Top             =   600
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "45,200 entries"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6840
      TabIndex        =   1
      Top             =   1080
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "zaidmarkabi@yahoo.com"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6050
      TabIndex        =   0
      Top             =   0
      Width           =   1785
   End
   Begin VB.Image Image1 
      Height          =   1320
      Left            =   0
      Picture         =   "Main.frx":0000
      Top             =   0
      Width           =   7920
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
Dictionary.Show
Unload Me
End Sub
