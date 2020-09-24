VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form Dictionary 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Zaid Dictionary - Arabic-English"
   ClientHeight    =   6600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5280
   Icon            =   "DataBaseEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "DataBaseEditor.frx":E008
   ScaleHeight     =   6600
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   435
      Left            =   240
      TabIndex        =   11
      Top             =   5760
      Width           =   4815
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   480
      ScaleHeight     =   2025
      ScaleWidth      =   4305
      TabIndex        =   15
      Top             =   2760
      Visible         =   0   'False
      Width           =   4335
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Deep mode"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   225
         Left            =   855
         TabIndex        =   19
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ÇáÈÍË ÈÏÞÉ ÃßËÑ"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   225
         Left            =   840
         TabIndex        =   18
         Top             =   1320
         Width           =   1020
      End
      Begin VB.Image Image2 
         Height          =   435
         Left            =   240
         Picture         =   "DataBaseEditor.frx":7F74A
         Top             =   1320
         Width           =   480
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "There are no results , Do you like to research with deep mode ?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   435
         Left            =   240
         TabIndex        =   17
         Top             =   600
         Width           =   3975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "áã íÊã ÇáÚËæÑ Úáì Ãí äÊíÌÉ , åá ÊÑíÏ ÇÚÇÏÉ ÇáÈÍË ÈÏÞÉ ÃßËÑ ¿"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   225
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   4080
      End
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   3270
      Left            =   2640
      TabIndex        =   12
      Top             =   2520
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Find"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   6960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   3270
      Left            =   240
      TabIndex        =   10
      Top             =   2520
      Width           =   4815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00A6A6A6&
      Caption         =   "Search Option"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   4815
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00A6A6A6&
         Caption         =   "Identity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "Identity"
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H00A6A6A6&
         Caption         =   "Latest Word"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1800
         TabIndex        =   8
         ToolTipText     =   "Latest Word"
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton Option3 
         Appearance      =   0  'Flat
         BackColor       =   &H00A6A6A6&
         Caption         =   "First Word"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "First  Word"
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton Option4 
         Appearance      =   0  'Flat
         BackColor       =   &H00A6A6A6&
         Caption         =   "Any Word Part"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1800
         TabIndex        =   6
         ToolTipText     =   "Any Word Part"
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "About Íæá"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   3960
         TabIndex        =   20
         Top             =   240
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0 Results."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   3720
         TabIndex        =   14
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   435
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   4215
   End
   Begin MSDataGridLib.DataGrid MSHFlexGrid1 
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   6960
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "English"
      Top             =   6960
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1320
      Top             =   6960
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Zaid Dictionary / Germany-English"
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
      TabIndex        =   21
      Top             =   120
      Width           =   3615
   End
   Begin VB.Image TransFormButton1 
      Height          =   285
      Left            =   4320
      Picture         =   "DataBaseEditor.frx":8026C
      Top             =   90
      Width           =   360
   End
   Begin VB.Image Button1 
      Height          =   240
      Left            =   4680
      Picture         =   "DataBaseEditor.frx":80806
      Top             =   120
      Width           =   345
   End
   Begin VB.Image Image1 
      Height          =   435
      Left            =   4560
      Picture         =   "DataBaseEditor.frx":80CC8
      Top             =   1080
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÇÈÍË Úä ÇáäÕ"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Left            =   3240
      TabIndex        =   13
      Top             =   840
      Width           =   1020
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Text Searching"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   840
      Width           =   1305
   End
End
Attribute VB_Name = "Dictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Button1_Click()
End
End Sub

Private Sub Command3_Click()
On Error Resume Next
Set MSHFlexGrid1.DataSource = Nothing

If Option1.Value = True Then
Adodc1.RecordSource = "Select * From " + "Table1" + " Where " + Text4.Text + " = '" & Text5.Text & "'"
End If
If Option2.Value = True Then
Adodc1.RecordSource = "Select * From " + "Table1" + " Where " + Text4.Text + " LIKE '%" & Text5.Text & "';"
End If
If Option3.Value = True Then
Adodc1.RecordSource = "Select * From " + "Table1" + " Where " + Text4.Text + " LIKE '" & Text5.Text & "%';"
End If
If Option4.Value = True Then
Adodc1.RecordSource = "Select * From " + "Table1" + " Where " + Text4.Text + " LIKE '%" & Text5.Text & "%';"
End If

Adodc1.Refresh
Set MSHFlexGrid1.DataSource = Adodc1

Dim IrecI As Long
For IrecI = 0 To MSHFlexGrid1.ApproxCount - 1
List1.AddItem LCase(MSHFlexGrid1.Columns(0).CellValue(MSHFlexGrid1.GetBookmark(IrecI)))
List2.AddItem MSHFlexGrid1.Columns(1).CellValue(MSHFlexGrid1.GetBookmark(IrecI))
Next
Label2.Caption = Format(List1.ListCount) + " Results."
End Sub

Private Sub Form_Load()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + "Dic.mdb" + ";Persist Security Info=False"
DoEvents
Adodc1.RecordSource = "SELECT * FROM " + "Table1"
DoEvents
Adodc1.Refresh
DoEvents
List2.Height = List1.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Image1_Click()
Text5_KeyPress (13)
End Sub

Private Sub Image2_Click()
Picture1.Visible = False
DoEvents
Option4.Value = True
DoEvents
Picture1.Visible = True
Option1.Value = True
Picture1.Visible = False
If List1.ListCount = 0 Then
Picture1.Visible = True
End If
End Sub

Private Sub Label7_Click()
About.Show 1
End Sub

Private Sub List1_Click()
Text1.Text = List2.List(List1.ListIndex)
End Sub

Private Sub List2_Click()
Text1.Text = List1.List(List2.ListIndex)
End Sub

Private Sub Option1_Click()
If Not Picture1.Visible = True Then
Text5.SetFocus
Text5_KeyPress (13)
End If
End Sub

Private Sub Option2_Click()
Text5.SetFocus
Text5_KeyPress (13)
End Sub

Private Sub Option3_Click()
Text5.SetFocus
Text5_KeyPress (13)
End Sub

Private Sub Option4_Click()
Text5.SetFocus
Text5_KeyPress (13)
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Not Text5.Text = "" Then
Picture1.Visible = False
DoEvents
List1.Clear
List2.Clear
Text4.Text = "English"
Command3_Click
Text5.Text = Text5.Text + "."
Command3_Click
Text5.Text = Left(Text5.Text, Len(Text5.Text) - 1)
Text4.Text = "Germany"
Command3_Click
If List1.ListCount > 0 Then
List1.ListIndex = 0
List1_Click
Picture1.Visible = False
Else
Picture1.Visible = True
End If
Text5.SelStart = 0
Text5.SelLength = Len(Text5.Text)
End If
End Sub

Private Sub TransFormButton1_Click()
Me.WindowState = 1
End Sub

