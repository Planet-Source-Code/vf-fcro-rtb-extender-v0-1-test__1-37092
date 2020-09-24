VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RichExtender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Enum WindowStyles
 WS_MAXIMIZEBOX = &H10000
 WS_MINIMIZEBOX = &H20000
 WS_THICKFRAME = &H40000
 WS_SYSMENU = &H80000
 WS_HSCROLL = &H100000
 WS_VSCROLL = &H200000
 WS_DLGFRAME = &H400000
 WS_BORDER = &H800000
 WS_MAXIMIZE = &H1000000
 WS_CLIPCHILDREN = &H2000000
 WS_CLIPSIBLINGS = &H4000000
 WS_DISABLED = &H8000000
 WS_VISIBLE = &H10000000
 WS_MINIMIZE = &H20000000
 WS_CHILD = &H40000000
 WS_POPUP = &H80000000
End Enum

Public Enum EditStyles
ES_AUTOHSCROLL = &H80&
ES_AUTOVSCROLL = &H40&
ES_CENTER = &H1&
ES_CONTINUOUS = (&H80000000)
ES_DISABLENOSCROLL = &H2000
ES_DISPLAY_REQUIRED = (&H2)
ES_EX_NOCALLOLEINIT = &H1000000
ES_LEFT = &H0&
ES_LOWERCASE = &H10&
ES_MULTILINE = &H4&
ES_NOHIDESEL = &H100&
ES_NOIME = &H80000
ES_NOOLEDRAGDROP = &H8
ES_NUMBER = &H2000&
ES_OEMCONVERT = &H400&
ES_PASSWORD = &H20&
ES_READONLY = &H800&
ES_RIGHT = &H2&
ES_SAVESEL = &H8000
ES_SELECTIONBAR = &H1000000
ES_SELFIME = &H40000
ES_SUNKEN = &H4000
ES_UPPERCASE = &H8&
ES_VERTICAL = &H400000
ES_WANTRETURN = &H1000&
End Enum



' Flags used to open/save files
Public Enum OpenSaveFlags

   OpenSaveRTF = tomRTF                       ' Open/Save RTF
   OpenSaveText = tomText                     ' Open/Save plain text
   OpenSaveHTML = tomHTML                     ' Open/Save HTML. Not supported by
                                                ' RichEdit
   OpenSaveWord = tomWordDocument             ' Open/Save Word document. Not
                                                ' supported by RichEdit

   OpenSaveReplaceSelection = tomPasteFile    ' Replace selection

   OpenSaveReadOnly                           ' Read only
   OpenSaveShareDenyRead = tomShareDenyRead   ' Other programs can't read
   OpenSaveShareDenyWrite = tomShareDenyWrite ' Other programs can't write

   OpenSaveCreateNew = tomCreateNew           ' Creates new file. Fails if
                                                ' the file already exists
   OpenSaveCreateAlways = tomCreateAlways     ' Creates new file. Destroys
                                                ' any previous file
   OpenSaveOpenExisting = tomOpenExisting     ' Opens existing file, fails
                                                ' if the file doesn't exists
   OpenSaveOpenAlways = tomOpenAlways         ' Opens existing file, otherwise
                                                ' creates a new file
OpenSaveTruncateExisting = tomTruncateExisting   ' Open existing file but
                                                      ' truncates it to zero length
End Enum

Public Type CharRange
Min As Long
Max As Long
End Type

Public Type FindTextStuct
ChrgIn As CharRange
Text As String
ChrgOut As CharRange
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


Public Type FORMATRANGE
DestHDC As Long
SrcHDC As Long
DestRC As RECT 'Twips
SrcRC As RECT 'Twips
ChrgIn As CharRange
End Type


Public Type INFOFROMPOS
LineNumber As Long
CharNumber As Long
End Type


Private Handle As Long


Private IRange As ITextRange
Private IPara As ITextPara
Private IFont As ITextFont
Private IStory As ITextStoryRanges
Private ISelection As ITextSelection
Private IDocument As ITextDocument

Private IOBJ As IUnknown

Public Property Let WorkingControl(ByVal hwnd As Long)
Handle = hwnd
SendMessage Handle, EM_GETOLEINTERFACE, 0&, IOBJ
Set IDocument = IOBJ
Set IOBJ = Nothing
End Property
Public Property Get WorkingControl() As Long
WorkingControl = Handle
End Property

Public Sub SaveRTF(ByVal Filename As String, OpenSave As OpenSaveFlags, ByVal CodePage As Long)
IDocument.Save CStr(Filename), OpenSave, CodePage
End Sub
Public Sub OpenRTF(ByVal Filename As String, OpenSave As OpenSaveFlags, ByVal CodePage As Long)
IDocument.Open CStr(Filename), OpenSave, CodePage
End Sub

Public Property Get Mask() As Long
Mask = SendMessage(Handle, EM_GETEVENTMASK, ByVal 0&, ByVal 0&)
End Property
Public Property Let Mask(ByVal Value As Long)
Call SendMessage(Handle, EM_SETEVENTMASK, ByVal 0&, ByVal Value)
End Property


Public Property Let Zoom(ByVal Value As Long)
 If Value < 2 Or Value > 500 Then Exit Property
Dim Numerator As Long
Dim Denominator As Long
If Value = 100 Then
Numerator = 0
Denominator = 0
Else
If Value < 100 Then
Numerator = Value
Denominator = 100
Else
Numerator = Value / 1.960784
Denominator = 51
End If
End If
SendMessage Handle, EM_SETZOOM, ByVal Numerator, ByVal Denominator
End Property
Public Property Get Zoom() As Long
Dim Numerator As Long
Dim Denominator As Long
SendMessage Handle, EM_GETZOOM, ByVal VarPtr(Numerator), Denominator
If Numerator = 0 Or Denominator = 0 Then
Zoom = 100
Else
Zoom = MulDiv(100, Numerator, Denominator)
End If
End Property

Public Sub Freeze()
IDocument.Freeze
End Sub
Public Sub UnFreeze()
IDocument.UnFreeze
End Sub

Private Property Get LineIndex(ByVal LineNumber As Long) As Long
LineIndex = SendMessage(Handle, EM_LINEINDEX, ByVal LineNumber, ByVal &O0)
End Property


Public Property Let InsertPictureFromFile(ByVal LineNumber As Long, ByVal CharPosition As Long, ByVal Filename As String)
Dim FF As Long
FF = FreeFile
If Dir(Filename) = "" Or Filename = "" Then Exit Property
Dim PicData() As Byte
Open Filename For Binary Access Read As #FF
ReDim PicData(LOF(FF) - 1)
Get #FF, , PicData
Close #FF
Dim Pic As StdPicture
Set Pic = GetPicture(PicData)
If Pic = 0 Then Exit Property
If LineNumber = -1 Or CharPosition = -1 Then GoTo dalje
Dim LI As Long
LI = LineIndex(LineNumber) + CharPosition
Call SendMessage(Handle, EM_SETSEL, ByVal LI, ByVal LI)
dalje:
Clipboard.Clear
Clipboard.SetData Pic
Call SendMessage(Handle, WM_PASTE, ByVal 0&, ByVal 0&)
Set Pic = Nothing
End Property
Public Property Let InsertPicture(ByVal LineNumber As Long, ByVal CharPosition As Long, ByVal NewPicture As StdPicture)
If NewPicture = 0 Then Exit Property
If LineNumber = -1 Or CharPosition = -1 Then GoTo dalje2
Dim LI As Long
LI = LineIndex(LineNumber) + CharPosition
Call SendMessage(Handle, EM_SETSEL, ByVal LI, ByVal LI)
dalje2:
Clipboard.Clear
Clipboard.SetData NewPicture
Call SendMessage(Handle, WM_PASTE, ByVal 0&, ByVal 0&)
End Property

Public Function FindText(ByVal Text As String, ByVal StartPos As Long, ByVal EndPos As Long, Optional CaseSensitive As Boolean, Optional WholeWord As Boolean, Optional DownUnderRange As Boolean) As CharRange
Dim ret(1) As Long
Dim FT As FindTextStuct
FT.ChrgIn.Min = StartPos
FT.ChrgIn.Max = EndPos
FT.Text = Text
Dim FindFlag As Long
If CaseSensitive Then FindFlag = FindFlag Or FR_MATCHCASE
If WholeWord Then FindFlag = FindFlag Or FR_WHOLEWORD
If DownUnderRange Then FindFlag = FindFlag Or FR_DOWN
Call SendMessage(Handle, EM_FINDTEXTEX, ByVal FindFlag, FT)
CopyMemory FindText, FT.ChrgOut, 8
End Function

Public Property Get MousePosition(ByVal X As Long, ByVal Y As Long) As INFOFROMPOS
Dim Coord(1) As Long
Coord(0) = X
Coord(1) = Y
MousePosition.CharNumber = SendMessage(Handle, EM_CHARFROMPOS, ByVal 0&, Coord(0))
MousePosition.LineNumber = SendMessage(Handle, EM_LINEFROMCHAR, ByVal MousePosition.CharNumber, ByVal 0&)
End Property
Public Sub Render(ByVal DestHDC As Long, DestRECT As RECT, ByVal LineFrom As Long, ByVal LineTo As Long)
Dim FRG As FORMATRANGE
FRG.ChrgIn.Min = LineIndex(LineFrom)
FRG.ChrgIn.Max = LineIndex(LineTo)
Dim PDC As Long
PDC = GetDC(Handle)
FRG.DestHDC = DestHDC
FRG.SrcHDC = PDC
CopyMemory FRG.DestRC, DestRECT, Len(DestRECT)
Call SendMessage(Handle, EM_FORMATRANGE, ByVal 1&, FRG)
Call SendMessage(Handle, EM_FORMATRANGE, ByVal 0&, ByVal 0&)
ReleaseDC Handle, PDC
End Sub
Public Sub NewDocument()
IDocument.New
End Sub
Public Sub Undo()
Call SendMessage(Handle, EM_UNDO, ByVal 0&, ByVal 0&)
End Sub
