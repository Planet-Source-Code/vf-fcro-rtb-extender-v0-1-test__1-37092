VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Rich Text Box Extender V0.01 by Vanja Fuckar"
   ClientHeight    =   7080
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   7080
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text6 
      Height          =   285
      Index           =   1
      Left            =   6240
      TabIndex        =   28
      Top             =   4800
      Width           =   855
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Index           =   0
      Left            =   4920
      TabIndex        =   27
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00FF8080&
      Caption         =   "New Document"
      Height          =   375
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Undo Last"
      Height          =   375
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   0
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   600
      Top             =   6360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Index           =   1
      Left            =   3480
      TabIndex        =   21
      Top             =   6120
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Index           =   0
      Left            =   2160
      TabIndex        =   19
      Top             =   6120
      Width           =   855
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Render To HDC"
      Height          =   375
      Left            =   2040
      TabIndex        =   18
      Top             =   6480
      Width           =   2415
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Whole Word"
      Height          =   255
      Left            =   4800
      TabIndex        =   17
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Case Sensitive"
      Height          =   255
      Left            =   4800
      TabIndex        =   16
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Find Text"
      Height          =   375
      Left            =   4800
      TabIndex        =   15
      Top             =   5520
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   4800
      TabIndex        =   14
      Top             =   5160
      Width           =   2415
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Get Zoom"
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Zoom"
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   5280
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   480
      TabIndex        =   11
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Open Document"
      Height          =   375
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "UnFreeze"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Freeze"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H008080FF&
      Caption         =   "Save Document"
      Height          =   375
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Insert From Image/Picture"
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   4800
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3360
      TabIndex        =   3
      Top             =   3960
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2160
      TabIndex        =   2
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Insert From File"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   4320
      Width           =   2415
   End
   Begin RichTextLib.RichTextBox RTB1 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   5530
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":0000
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "End Pos"
      Height          =   255
      Left            =   6240
      TabIndex        =   30
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Start Pos"
      Height          =   255
      Left            =   4920
      TabIndex        =   29
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Type: -1,-1 To Insert Picture At Current Position!"
      Height          =   375
      Left            =   2040
      TabIndex        =   24
      Top             =   5280
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   5415
   End
   Begin VB.Shape Shape4 
      Height          =   1095
      Left            =   1920
      Top             =   5880
      Width           =   2655
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "LineTo"
      Height          =   255
      Index           =   1
      Left            =   3480
      TabIndex        =   22
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "LineFrom"
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   20
      Top             =   5880
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   1455
      Left            =   240
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Shape Shape2 
      Height          =   2415
      Left            =   4680
      Top             =   3600
      Width           =   2655
   End
   Begin VB.Shape Shape1 
      Height          =   2175
      Left            =   1920
      Top             =   3600
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   1515
      Left            =   8160
      Picture         =   "Form1.frx":0082
      Top             =   4080
      Visible         =   0   'False
      Width           =   6225
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Char Pos:"
      Height          =   255
      Left            =   3360
      TabIndex        =   5
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Line Num:"
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   3720
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RICH1 As New RichExtender

Private Sub Command1_Click()
cd1.DialogTitle = "Load Picture Format"
cd1.ShowOpen

If cd1.FileName = "" Then Exit Sub
If Text1 = "" Or Text2 = "" Or Not IsNumeric(Text1) Or Not IsNumeric(Text2) Then MsgBox "Error", , "Info": Exit Sub
RICH1.InsertPictureFromFile(Text1, Text2) = cd1.FileName
End Sub

Private Sub Command10_Click()
If Text6(0) = "" Or Not IsNumeric(Text6(0)) Or Text6(1) = "" Or Not IsNumeric(Text6(1)) Then MsgBox "Error", , "Info": Exit Sub
Dim RT As CharRange
RT = RICH1.FindText(Text4, Text6(0), Text6(1), CBool(Check1.Value), CBool(Check2.Value))
If RT.Min < 0 Then
MsgBox "Couldn't Find...", , "Info"
Else
MsgBox "Find At:" & vbCrLf & "Start Position:" & RT.Min & vbCrLf & "End Position:" & RT.Max - 1, , "Info"
End If
End Sub



Private Sub Command11_Click()
Form2.Show

Dim FR As RECT
FR.Bottom = Form2.ScaleHeight
FR.Right = Form2.ScaleWidth

RICH1.Render Form2.hDC, FR, Text5(0), Text5(1)
End Sub

Private Sub Command12_Click()
RICH1.Undo
End Sub



Private Sub Command13_Click()
RICH1.NewDocument
End Sub

Private Sub Command2_Click()
If Text1 = "" Or Text2 = "" Or Not IsNumeric(Text1) Or Not IsNumeric(Text2) Then MsgBox "Error", , "Info": Exit Sub

RICH1.InsertPicture(Text1, Text2) = Image1
End Sub

Private Sub Command3_Click()
cd1.ShowSave
If cd1.FileName = "" Then Exit Sub
RICH1.SaveRTF cd1.FileName, OpenSaveRTF Or OpenSaveOpenExisting, 0
End Sub

Private Sub Command4_Click()
RICH1.Freeze
End Sub

Private Sub Command5_Click()
RICH1.UnFreeze
End Sub

Private Sub Command6_Click()
cd1.DialogTitle = "Open RTF File"
cd1.ShowOpen

If cd1.FileName = "" Then Exit Sub
RICH1.OpenRTF cd1.FileName, OpenSaveRTF, 0

End Sub





Private Sub Command8_Click()
If Text3 = "" Or Not IsNumeric(Text3) Or CLng(Text3) > 500 Then MsgBox "Error", , "Info": Exit Sub

RICH1.Zoom = Text3
End Sub

Private Sub Command9_Click()
MsgBox "Zoom:" & RICH1.Zoom & " %", , "Info"
End Sub

Private Sub Form_Load()
Top = (Screen.Height - Height) / 2
Left = (Screen.Width - Me.Width) / 2
RICH1.WorkingControl = RTB1.hWnd
RICH1.OpenRTF App.Path & "\test.rtf", OpenSaveOpenExisting Or OpenSaveRTF, 0
End Sub

Private Sub RTB1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim INFOX As INFOFROMPOS
INFOX = RICH1.MousePosition(x / 15, y / 15)
Label4 = "Line Number:" & INFOX.LineNumber & ",Character Position:" & INFOX.CharNumber
End Sub


