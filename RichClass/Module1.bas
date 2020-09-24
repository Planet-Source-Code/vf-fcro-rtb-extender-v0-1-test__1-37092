Attribute VB_Name = "Module1"
Public Const WM_USER = &H400
Public Const EM_GETAUTOURLDETECT = (WM_USER + 92)
Public Const EM_GETBIDIOPTIONS = (WM_USER + 201)
Public Const EM_GETCHARFORMAT = (WM_USER + 58)
Public Const EM_GETPARAFORMAT = (WM_USER + 61)
Public Const EM_GETRECT = &HB2
Public Const EM_GETTYPOGRAPHYOPTIONS = (WM_USER + 203)
Public Const EM_SETBIDIOPTIONS = (WM_USER + 200)
Public Const EM_SETBKGNDCOLOR = (WM_USER + 67)
Public Const EM_SETCHARFORMAT = (WM_USER + 68)
Public Const EM_SETEVENTMASK = (WM_USER + 69)
Public Const EM_SETFONTSIZE = (WM_USER + 223)
Public Const EM_SETLANGOPTIONS = (WM_USER + 120)
Public Const EM_SETPALETTE = (WM_USER + 93)
Public Const EM_SETPARAFORMAT = (WM_USER + 71)
Public Const EM_SETRECT = &HB3
Public Const EM_SETTYPOGRAPHYOPTIONS = (WM_USER + 202)

Public Const EM_CHARFROMPOS = &HD7
Public Const EM_EXGETSEL = (WM_USER + 52)
Public Const EM_EXSETSEL = (WM_USER + 55)
Public Const EM_GETFIRSTVISIBLELINE = &HCE
Public Const EM_GETSEL = &HB0
Public Const EM_HIDESELECTION = (WM_USER + 63)
Public Const EM_POSFROMCHAR = (WM_USER + 38)
Public Const EM_SELECTIONTYPE = (WM_USER + 66)
Public Const EM_SETSEL = &HB1
Public Const EM_EXLIMITTEXT = (WM_USER + 53)
Public Const EM_FINDTEXT = (WM_USER + 56)
Public Const EM_FINDTEXTEX = (WM_USER + 79)
Public Const EM_FINDTEXTEXW = (WM_USER + 124)
Public Const EM_GETLIMITTEXT = (WM_USER + 37)
Public Const EM_GETSELTEXT = (WM_USER + 62)
Public Const EM_GETTEXTEX = (WM_USER + 94)
Public Const EM_GETTEXTLENGTHEX = (WM_USER + 95)
Public Const EM_GETTEXTMODE = (WM_USER + 90)
Public Const EM_GETTEXTRANGE = (WM_USER + 75)
Public Const EM_REPLACESEL = &HC2
Public Const EM_LIMITTEXT = &HC5
Public Const EM_SETLIMITTEXT = EM_LIMITTEXT
Public Const EM_SETTEXTEX = (WM_USER + 97)
Public Const EM_SETTEXTMODE = (WM_USER + 89)


Public Const EM_EXLINEFROMCHAR = (WM_USER + 54)
Public Const EM_FINDWORDBREAK = (WM_USER + 76)
Public Const EM_GETWORDBREAKPROC = &HD1
Public Const EM_GETWORDBREAKPROCEX = (WM_USER + 80)
Public Const EM_GETWORDWRAPMODE = (WM_USER + 103)
Public Const EM_SETWORDBREAKPROC = &HD0
Public Const EM_SETWORDBREAKPROCEX = (WM_USER + 81)
Public Const EM_SETWORDWRAPMODE = (WM_USER + 102)

Public Const EM_GETLINE = &HC4
Public Const EM_GETLINECOUNT = &HBA
Public Const EM_GETSCROLLPOS = (WM_USER + 221)
Public Const EM_GETTHUMB = &HBE
Public Const EM_LINEFROMCHAR = &HC9
Public Const EM_LINEINDEX = &HBB
Public Const EM_LINELENGTH = &HC1
Public Const EM_LINESCROLL = &HB6
Public Const EM_SCROLL = &HB5
Public Const EM_SCROLLCARET = &HB7
Public Const EM_SETSCROLLPOS = (WM_USER + 222)
Public Const EM_SHOWSCROLLBAR = (WM_USER + 96)

Public Const EM_CANPASTE = (WM_USER + 50)
Public Const EM_CANREDO = (WM_USER + 85)
Public Const EM_CANUNDO = &HC6
Public Const EM_EMPTYUNDOBUFFER = &HCD
Public Const EM_GETEDITSTYLE = (WM_USER + 205)
Public Const EM_GETREDONAME = (WM_USER + 87)
Public Const EM_GETUNDONAME = (WM_USER + 86)
Public Const EM_PASTESPECIAL = (WM_USER + 64)
Public Const EM_RECONVERSION = (WM_USER + 125)
Public Const EM_REDO = (WM_USER + 84)
Public Const EM_SETUNDOLIMIT = (WM_USER + 82)
Public Const EM_STOPGROUPTYPING = (WM_USER + 88)
Public Const EM_UNDO = &HC7

Public Const EM_STREAMIN = (WM_USER + 73)
Public Const EM_STREAMOUT = (WM_USER + 74)
Public Const EM_DISPLAYBAND = (WM_USER + 51)
Public Const EM_FORMATRANGE = (WM_USER + 57)
Public Const EM_SETTARGETDEVICE = (WM_USER + 72)
Public Const EM_REQUESTRESIZE = (WM_USER + 65)

Public Const EM_GETOLEINTERFACE = (WM_USER + 60)
Public Const EM_SETOLECALLBACK = (WM_USER + 70)
Public Const EM_GETEVENTMASK = (WM_USER + 59)
Public Const EM_GETIMECOLOR = (WM_USER + 105)
Public Const EM_GETIMECOMPMODE = (WM_USER + 122)
Public Const EM_GETIMEMODEBIAS = (WM_USER + 127)
Public Const EM_GETIMEOPTIONS = (WM_USER + 107)
Public Const EM_GETLANGOPTIONS = (WM_USER + 121)
Public Const EM_GETIMESTATUS = &HD9
Public Const EM_GETMODIFY = &HB8
Public Const EM_GETOPTIONS = (WM_USER + 78)
Public Const EM_GETPUNCTUATION = (WM_USER + 101)
Public Const EM_GETZOOM = (WM_USER + 224)
Public Const EM_SETIMECOLOR = (WM_USER + 104)
Public Const EM_SETIMEOPTIONS = (WM_USER + 106)
Public Const EM_SETIMEMODEBIAS = (WM_USER + 126)
Public Const EM_SETIMESTATUS = &HD8
Public Const EM_SETOPTIONS = (WM_USER + 77)
Public Const EM_SETPUNCTUATION = (WM_USER + 100)
Public Const EM_SETREADONLY = &HCF
Public Const EM_SETZOOM = (WM_USER + 225)
Public Const EM_SETEDITSTYLE = (WM_USER + 204)


' Notification masks
 Public Const ENM_NONE = &H0
Public Const ENM_CHANGE = &H1
 Public Const ENM_UPDATE = &H2
Public Const ENM_SCROLL = &H4
Public Const ENM_SCROLLEVENTS = &H8
 Public Const ENM_DRAGDROPDONE = &H10
 Public Const ENM_PARAGRAPHEXPANDED = &H20
Public Const ENM_KEYEVENTS = &H10000
Public Const ENM_MOUSEEVENTS = &H20000
 Public Const ENM_REQUESTRESIZE = &H40000
Public Const ENM_SELCHANGE = &H80000
Public Const ENM_DROPFILES = &H100000
Public Const ENM_PROTECTED = &H200000
 Public Const ENM_CORRECTTEXT = &H400000
 Public Const ENM_LANGCHANGE = &H1000000
 Public Const ENM_OBJECTPOSITIONS = &H2000000
Public Const ENM_LINK = &H4000000



Public Const WM_PASTE = &H302



Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, wParam As Any, lParam As Any) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)


Public Type GUID
  dwData1 As Long
  wData2 As Integer
  wData3 As Integer
  abData4(7) As Byte
End Type

Declare Function CLSIDFromString Lib "ole32" (ByVal lpsz As Any, pclsid As GUID) As Long
Const sIID_IPicture = "{7BF80980-BF32-101A-8BBB-00AA00300CAB}"

Const GMEM_MOVEABLE = &H2
Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByValdwBytes As Long) As Long
Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function CreateStreamOnHGlobal Lib "ole32" _
                              (ByVal hGlobal As Long, _
                              ByVal fDeleteOnRelease As Long, _
                              ppstm As Any) As Long

Declare Function OleLoadPicture Lib "olepro32" _
                              (pStream As Any, _
                              ByVal lSize As Long, _
                              ByVal fRunmode As Long, _
                              riid As GUID, _
                              ppvObj As Any) As Long

Public Const CP_UTF8 = 65001
Declare Function MulDiv Lib "kernel32" (ByVal Mul As Long, ByVal Nom As Long, ByVal Den As Long) As Long


Public Const FR_DOWN = &H1
Public Const FR_MATCHCASE = &H4
Public Const FR_WHOLEWORD = &H2


Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long


Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long




Public Function GetPicture(dataXY() As Byte) As IPicture
Dim hMem  As Long
Dim lpMem  As Long
Dim IID_IPicture As GUID
Dim istm As stdole.IUnknown
Dim ipic As IPicture
hMem = GlobalAlloc(GMEM_MOVEABLE, UBound(dataXY) + 1)
lpMem = GlobalLock(hMem)
CopyMemory ByVal lpMem, dataXY(0), UBound(dataXY) + 1
Call GlobalUnlock(hMem)
Call CreateStreamOnHGlobal(hMem, 1, istm)
Call CLSIDFromString(StrPtr(sIID_IPicture), IID_IPicture)
Call OleLoadPicture(ByVal ObjPtr(istm), UBound(dataXY) + 1, 0, IID_IPicture, GetPicture)
Call GlobalFree(hMem)
End Function


