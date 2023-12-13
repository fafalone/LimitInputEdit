Attribute VB_Name = "modLimitInput"
Option Explicit

'********************************************************************
'modLimitInput                                                      '
'                                                                   '
'Declarations for SHLimitInputEditWithFlags, an API entirely        '
'undocumented until now. It's much more useful than the documented  '
'function that partially wraps it, SHLimitInputEdit (which, by the  '
'way, does *not* require your class to implement IShellFolder, only '
'IItemNameLimits). AFAIK, it's never been exported by name, so keep '
'the ordinal number alias in the API declare.                       '
'                                                                   '
'Function documented by Jon Johnson (fafalone)                      '
'Last update: 13 Dec 2023; x64 port for twinBASIC.                  '
'********************************************************************

#If VBA7 Then
Private Declare PtrSafe Function DefSubclassProc Lib "comctl32" Alias "#413" (ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Function SetWindowSubclass Lib "comctl32" Alias "#410" (ByVal hWnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr, Optional ByVal dwRefData As LongPtr) As Long
Private Declare PtrSafe Function GetWindowSubclass Lib "comctl32" Alias "#411" (ByVal hWnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr, dwRefData As LongPtr) As Long
Private Declare PtrSafe Function RemoveWindowSubclass Lib "comctl32" Alias "#412" (ByVal hWnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Public Declare PtrSafe Function SHLimitInputEditWithFlags Lib "shell32" Alias "#754" (ByVal hwndEdit As LongPtr, pil As LIMITINPUT) As Long
Public Type LIMITINPUT
    cbSize As Long          'Size of structure. Must set.
    dwMask As LI_Mask       'LIM_* values.
    dwFlags As LI_Flags     'LIF_* values.
    hInst As LongPtr        'App.hInstance or loaded module hInstance.
    pszFilter As LongPtr    'String via StrPtr, LICF_* category, LPSTR_TEXTCALLBACK to set via LIN_GETDISPINFO, or resource id in .hInst.
    pszTitle As LongPtr     'Optional. String via StrPtr, LPSTR_TEXTCALLBACK to set via LIN_GETDISPINFO, or resource id in .hInst.
    pszMessage As LongPtr   'Ignore if tooltip disabled. String via StrPtr, LPSTR_TEXTCALLBACK to set via LIN_GETDISPINFO, or resource id in .hInst.
    hIcon As LongPtr        'See TTM_SETTITLE. Can be TTI_* default icon, hIcon, or I_ICONCALLBACK to set via LIN_GETDISPINFO.
    hwndNotify As LongPtr   'Window to send notifications to. Must specify if any callbacks used or bad character notifications enabled.
    iTimeout As Long        'Timeout in milliseconds. Defaults to 10000 if not set.
    cxTipWidth As Long      'Tooltip width. Default 500px.
End Type

Private Type NMHDR
  hWndFrom As LongPtr   'Window handle of control sending message
  IDFrom As LongPtr     'Identifier of control sending message
  Code  As Long         'Specifies the notification code
End Type
       
Private Type NMLIBADCHAR
    hdr As NMHDR
    wParam As LongPtr 'WM_CHAR wParam (Char code)
    lParam As LongPtr 'WM_CHAR lParam (see MSDN for details)
End Type
#Else
Private Declare Function SetWindowSubclass Lib "comctl32.dll" Alias "#410" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, Optional ByVal dwRefData As Long) As Long
Private Declare Function RemoveWindowSubclass Lib "comctl32.dll" Alias "#412" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long) As Long
Private Declare Function DefSubclassProc Lib "comctl32.dll" Alias "#413" (ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function SHLimitInputEditWithFlags Lib "shell32" Alias "#754" (ByVal hwndEdit As Long, pil As LIMITINPUT) As Long
Public Type LIMITINPUT
    cbSize As Long       'Size of structure. Must set.
    dwMask As LI_Mask    'LIM_* values.
    dwFlags As LI_Flags  'LIF_* values.
    hInst As Long        'App.hInstance or loaded module hInstance.
    pszFilter As Long    'String via StrPtr, LICF_* category, LPSTR_TEXTCALLBACK to set via LIN_GETDISPINFO, or resource id in .hInst.
    pszTitle As Long     'Optional. String via StrPtr, LPSTR_TEXTCALLBACK to set via LIN_GETDISPINFO, or resource id in .hInst.
    pszMessage As Long   'Ignore if tooltip disabled. String via StrPtr, LPSTR_TEXTCALLBACK to set via LIN_GETDISPINFO, or resource id in .hInst.
    hIcon As Long        'See TTM_SETTITLE. Can be TTI_* default icon, hIcon, or I_ICONCALLBACK to set via LIN_GETDISPINFO.
    hwndNotify As Long   'Window to send notifications to. Must specify if any callbacks used or bad character notifications enabled.
    iTimeout As Long     'Timeout in milliseconds. Defaults to 10000 if not set.
    cxTipWidth As Long   'Tooltip width. Default 500px.
End Type

Private Type NMHDR
  hWndFrom As Long   'Window handle of control sending message
  IDFrom As Long     'Identifier of control sending message
  Code  As Long      'Specifies the notification code
End Type
       
Private Type NMLIBADCHAR
    hdr As NMHDR
    wParam As Long 'WM_CHAR wParam (Char code)
    lParam As Long 'WM_CHAR lParam (see MSDN for details)
End Type
#End If


'Values for LIMITINPUT.dwMask
Public Enum LI_Mask
    LIM_FLAGS = &H1      'dwFlags used
    LIM_FILTER = &H2     'pszFilter used
    LIM_HINST = &H8      'hinst contains valid data. Generally must be set.
    LIM_TITLE = &H10     'pszTitle used. Tooltip title.
    LIM_MESSAGE = &H20   'pszMessage used. Tooltip main message.
    LIM_ICON = &H40      'hicon used. Can use default icons e.g. IDI_HAND. Loaded from .hInst.
    LIM_NOTIFY = &H80    'hwndNotify used. NOTE: Must be set to receive notifications. Automatic finding of parent broken.
    LIM_TIMEOUT = &H100  'iTimeout used. Default timeout=10000.
    LIM_TIPWIDTH = &H200 'cxTipWidth used. Default 500px.
End Enum

'Values for LIMITINPUT.dwFlags
Public Enum LI_Flags
    LIF_INCLUDEFILTER = &H0     'Default: pszFilter specifies what to include.
    LIF_EXCLUDEFILTER = &H1     'pszFilter specifies what to exclude.
    LIF_CATEGORYFILTER = &H2    'pszFilter uses LICF_* categories, not a string of chars.

    LIF_WARNINGBELOW = &H0      'Default: Tooltip below.
    LIF_WARNINGABOVE = &H4      'Tooltip above.
    LIF_WARNINGCENTERED = &H8   'Tooltip centered.
    LIF_WARNINGOFF = &H10       'Disable tooltip.

    LIF_FORCEUPPERCASE = &H20   'Makes chars uppercase.
    LIF_FORCELOWERCASE = &H40   'Makes chars lowercase. (This and forceupper mutually exclusive)

    LIF_MESSAGEBEEP = &H0       'Default: System default beep played.
    LIF_SILENT = &H80           'No beep.

    LIF_NOTIFYONBADCHAR = &H100 'Send WM_NOTIFY LIN_NOTIFYBADCHAR. NOTE: Must set LIM_NOTIFY flag and .hwndNotify member.
    LIF_HIDETIPONVALID = &H200  'Timeout tooltip early if valid char entered.

    LIF_PASTESKIP = &H0         'Default: Paste any allowed characters, skip disallowed.
    LIF_PASTESTOP = &H400       'Paste until first disallowed character encountered.
    LIF_PASTECANCEL = &H800     'Cancel paste entirely if any disallowed character.

    LIF_KEEPCLIPBOARD = &H1000  'If not set, modifies clipboard to what was pasted after paste flags executed.
End Enum

'Filters support CT_TYPE1 categories:
Public Const LICF_UPPER = &H1
Public Const LICF_LOWER = &H2
Public Const LICF_DIGIT = &H4
Public Const LICF_SPACE = &H8
Public Const LICF_PUNCT = &H10  'Punctuation
Public Const LICF_CNTRL = &H20  'Control characters
Public Const LICF_BLANK = &H40
Public Const LICF_XDIGIT = &H80  'Hexadecimal values, 0-9 and A-F.
Public Const LICF_ALPHA = &H100  'Any CT_TYPE1 linguistic character. Includes non-Latin alphabets.
'Custom categories
Public Const LICF_BINARYDIGIT = &H10000
Public Const LICF_OCTALDIGIT = &H20000 'Base 8; 0-7.
Public Const LICF_ATOZUPPER = &H100000 'ASCII A to Z
Public Const LICF_ATOZLOWER = &H200000 'ASCII a to z
Public Const LICF_ATOZ = (LICF_ATOZUPPER Or LICF_ATOZLOWER)

'Notification codes
Public Const LIN_GETDISPINFO = &H1   'Need tooltip display info (pszTitle and pszMessage).
Public Const LIN_GETFILTERINFO = &H2 'Need pszFilter and dwMask if modifying it.
Public Const LIN_BADCHAR = &H3       'Bad character notification from LIF_NOTIFYONBADCHAR



Public Type NMLIDISPINFO
    hdr As NMHDR
    li As LIMITINPUT 'Set all values requested in dwMask.
End Type

'Misc support
Public Const I_ICONCALLBACK = (-1)
Public Const LPSTR_TEXTCALLBACK = (-1)

Public Const TTI_NONE = 0
Public Const TTI_INFO = 1
Public Const TTI_WARNING = 2
Public Const TTI_ERROR = 3
Public Const TTI_INFO_LARGE = 4
Public Const TTI_WARNING_LARGE = 5
Public Const TTI_ERROR_LARGE = 6

Private Const WM_NOTIFY = &H4E
Private Const WM_DESTROY = &H2
Private Const WM_USER = &H400

#If VBA7 Then
Public Function Subclass2(hWnd As LongPtr, lpFN As LongPtr, Optional uId As LongPtr = 0&, Optional dwRefData As LongPtr = 0&) As Boolean
If uId = 0 Then uId = hWnd
    Subclass2 = SetWindowSubclass(hWnd, lpFN, uId, dwRefData):      Debug.Assert Subclass2
End Function

Public Function UnSubclass2(hWnd As LongPtr, ByVal lpFN As LongPtr, pid As LongPtr) As Boolean
    UnSubclass2 = RemoveWindowSubclass(hWnd, lpFN, pid)
End Function

Private Function PtrFormWndProc() As LongPtr
PtrFormWndProc = FARPROC(AddressOf FormWndProc)
End Function

Private Function FARPROC(pfn As LongPtr) As LongPtr
  FARPROC = pfn
End Function
#Else
Public Function Subclass2(hWnd As Long, lpFN As Long, Optional uId As Long = 0&, Optional dwRefData As Long = 0&) As Boolean
If uId = 0 Then uId = hWnd
    Subclass2 = SetWindowSubclass(hWnd, lpFN, uId, dwRefData):      Debug.Assert Subclass2
End Function

Public Function UnSubclass2(hWnd As Long, ByVal lpFN As Long, pid As Long) As Boolean
    UnSubclass2 = RemoveWindowSubclass(hWnd, lpFN, pid)
End Function

Private Function PtrFormWndProc() As Long
PtrFormWndProc = FARPROC(AddressOf FormWndProc)
End Function

Private Function FARPROC(pfn As Long) As Long
  FARPROC = pfn
End Function
#End If

#If VBA7 Then
Public Function FormWndProc(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal uIdSubclass As LongPtr, ByVal dwRefData As LongPtr) As LongPtr
#Else
Public Function FormWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
#End If


Select Case uMsg
        Case WM_NOTIFY
            Dim hdr As NMHDR
            CopyMemory hdr, ByVal lParam, LenB(hdr)
            If hdr.hWndFrom = Form1.Text1.hWnd Then
                If hdr.Code = LIN_BADCHAR Then
                    Dim nmlibc As NMLIBADCHAR
                    CopyMemory nmlibc, ByVal lParam, Len(nmlibc)
                    Debug.Print "NotifyBadChar " & Chr$(CLng(nmlibc.wParam)) & " (0x" & Hex$(wParam) & ")"
                    Form1.BadInput Chr$(CLng(nmlibc.wParam))
                End If
            End If
        Case WM_DESTROY
            Call UnSubclass2(hWnd, PtrFormWndProc(), uIdSubclass)
End Select
FormWndProc = DefSubclassProc(hWnd, uMsg, wParam, lParam)
End Function
