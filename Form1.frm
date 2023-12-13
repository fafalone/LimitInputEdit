VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Undocumented Input Limiter"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6240
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   6240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Limit length"
      Height          =   360
      Left            =   5115
      TabIndex        =   33
      Top             =   5430
      Width           =   990
   End
   Begin VB.TextBox Text5 
      Height          =   315
      Left            =   4695
      TabIndex        =   31
      Text            =   "8"
      Top             =   5460
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Simple: Numbers only"
      Height          =   390
      Left            =   2550
      TabIndex        =   30
      Top             =   4380
      Width           =   1740
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Custom set:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   90
      TabIndex        =   27
      Top             =   1620
      Width           =   1455
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Custom set:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   90
      TabIndex        =   26
      Top             =   1190
      Width           =   1455
   End
   Begin VB.OptionButton Option2 
      Caption         =   "No space/num"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   90
      TabIndex        =   25
      Top             =   760
      Width           =   1605
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Hexadecimal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   90
      TabIndex        =   24
      Top             =   330
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.CheckBox Check6 
      Caption         =   "Enable notifications"
      Height          =   255
      Left            =   4335
      TabIndex        =   23
      Top             =   0
      Value           =   1  'Checked
      Width           =   1785
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2505
      TabIndex        =   21
      Top             =   1635
      Width           =   1725
   End
   Begin VB.Frame Frame2 
      Caption         =   "Paste behavior"
      Height          =   2175
      Left            =   3600
      TabIndex        =   15
      Top             =   2100
      Width           =   2595
      Begin VB.OptionButton Option1 
         Caption         =   "Skip bad characters on paste"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Value           =   -1  'True
         Width           =   2655
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Paste until first bad char"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   510
         Width           =   2445
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Block paste if any bad chars"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   2985
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Don't modify clipboard"
         Height          =   195
         Left            =   135
         TabIndex        =   16
         Top             =   1275
         Value           =   1  'Checked
         Width           =   2115
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Flags"
      Height          =   2205
      Left            =   135
      TabIndex        =   8
      Top             =   2085
      Width           =   3405
      Begin VB.CheckBox Check7 
         Caption         =   "Include icon"
         Height          =   255
         Left            =   105
         TabIndex        =   28
         Top             =   1830
         Value           =   1  'Checked
         Width           =   2475
      End
      Begin VB.CheckBox Check1 
         Caption         =   "No beep"
         Height          =   225
         Left            =   90
         TabIndex        =   13
         Top             =   240
         Width           =   1695
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Show tooltip above instead of below"
         Height          =   255
         Left            =   90
         TabIndex        =   12
         Top             =   720
         Width           =   2985
      End
      Begin VB.CheckBox Check3 
         Caption         =   "No bad character tooltip"
         Height          =   225
         Left            =   90
         TabIndex        =   11
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   60
         TabIndex        =   10
         Text            =   "0"
         Top             =   990
         Width           =   375
      End
      Begin VB.CheckBox Check4 
         Caption         =   "If valid input received, hide tip immediately (otherwise waits for timeout)"
         Height          =   525
         Left            =   105
         TabIndex        =   9
         Top             =   1290
         Value           =   1  'Checked
         Width           =   3255
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Custom tip timeout (ms) - 0 to not set"
         Height          =   195
         Left            =   480
         TabIndex        =   14
         Top             =   1050
         Width           =   2700
      End
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2505
      TabIndex        =   7
      Top             =   1215
      Width           =   1725
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   4365
      TabIndex        =   4
      Top             =   270
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Apply Full Settings"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   135
      TabIndex        =   2
      Top             =   4350
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   2115
      TabIndex        =   0
      Top             =   4860
      Width           =   3555
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Setting a length limit is done through normal edit APIs, but included for reference."
      Height          =   450
      Left            =   75
      TabIndex        =   32
      Top             =   5370
      Width           =   4290
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line1 
      X1              =   525
      X2              =   5685
      Y1              =   5340
      Y2              =   5340
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "or"
      Height          =   195
      Left            =   2400
      TabIndex        =   29
      Top             =   4470
      Width           =   150
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Limited Input Box:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   22
      Top             =   4920
      Width           =   1725
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exclude:"
      Height          =   195
      Left            =   1755
      TabIndex        =   20
      Top             =   1695
      Width           =   735
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Allow only: "
      Height          =   195
      Left            =   1605
      TabIndex        =   6
      Top             =   1245
      Width           =   915
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exclude spaces and numbers, allow everything else."
      Height          =   420
      Left            =   1740
      TabIndex        =   5
      Top             =   735
      Width           =   2640
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Limit to valid hex characters and automatically convert to uppercase."
      Height          =   390
      Left            =   1695
      TabIndex        =   3
      Top             =   300
      Width           =   2655
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select demo to apply:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   75
      TabIndex        =   1
      Top             =   30
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'********************************************************************
'LimitInput                                                         '
'                                                                   '
'Implementation of SHLimitInputEditWithFlags, an API entirely       '
'undocumented until now. It's much more useful than the documented  '
'function that partially wraps it, SHLimitInputEdit (which, by the  '
'way, does *not* require your class to implement IShellFolder, only '
'IItemNameLimits). AFAIK, it's never been exported by name, so keep '
'the ordinal number alias in the API declare.                       '
'                                                                   '
'Function documented by Jon Johnson (fafalone)                      '
'Last update: 13 Dec 2023; x64 port for twinBASIC.                  '
'********************************************************************


'Other options not implemented:
'-You can set properties in a callback instead of up front... e.g. you'd set
' the tip title and message to LPSTR_TEXTCALLBACK, then would receive a LIN_GETDISPINFO
' message with a NMLIDISPINFO structure to set those properties then.
'
'-I didn't implement using LIM_TIPWIDTH and .cxTipWidth to set a custom width.
'
'-Most of these are optional; e.g. you don't need to specify an icon, or title/message
' if you're not using the tooltip.
'
'-All declarations that existed as of XP SP1 are included.
'
'-SHLimitInputEndSubclass is exported at ordinal #888 in Vista+, but I don't know how to
' restructure things to reapply a new filter. If you want to experiment, by all means.
'
'
Private sTitle As String
Private sMsg As String
Private sFilter As String
Private Const EM_LIMITTEXT = &HC5
#If VBA7 Then
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
#Else
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
#End If


Public Sub BadInput(szChar As String)
List1.AddItem "Bad char: " & szChar
End Sub

Private Sub Command1_Click()
Dim tli As LIMITINPUT
tli.cbSize = LenB(tli)
tli.dwMask = LIM_FILTER Or LIM_FLAGS Or LIM_TITLE Or LIM_MESSAGE Or LIM_HINST
tli.hInst = App.hInstance
sTitle = "Bad character"
If Check6.Value = vbChecked Then
    tli.dwMask = tli.dwMask Or LIM_NOTIFY
    tli.dwFlags = tli.dwFlags Or LIF_NOTIFYONBADCHAR
    tli.hwndNotify = Form1.hWnd 'It *should* have sent it to the parent. But it wasn't working without manually specifying.
End If
If Option2(0).Value = True Then
    tli.dwFlags = tli.dwFlags Or LIF_CATEGORYFILTER Or LIF_FORCEUPPERCASE
    tli.pszFilter = LICF_XDIGIT Or LICF_CNTRL 'NOTE: When LIF_CATEGORYFILTER is present, this represents a LICF_ category.
                                '      Without that flag, you can set a custom String of allowed (or excluded) chars.
    sMsg = "Only hexadecimal (0-9, A-F) allowed."
ElseIf Option2(1).Value = True Then
    tli.dwFlags = tli.dwFlags Or LIF_EXCLUDEFILTER Or LIF_CATEGORYFILTER
    tli.pszFilter = LICF_SPACE Or LICF_DIGIT Or LICF_CNTRL
    sMsg = "No spaces or numbers allowed."
ElseIf Option2(2).Value = True Then
    sFilter = StrConv(Text3.Text, vbFromUnicode) '& Chr(8)
    tli.pszFilter = StrPtr(sFilter)
    sMsg = "Only " & Text3.Text & " are allowed."
ElseIf Option2(3).Value = True Then
    tli.dwFlags = tli.dwFlags Or LIF_EXCLUDEFILTER
    sFilter = StrConv(Text4.Text, vbFromUnicode)
    tli.pszFilter = StrPtr(sFilter)
    sMsg = Text4.Text & " are not allowed."
End If
tli.pszTitle = StrPtr(sTitle)
tli.pszMessage = StrPtr(sMsg)
If Check7.Value = vbChecked Then
    tli.dwMask = tli.dwMask Or LIM_ICON
    tli.hIcon = TTI_WARNING_LARGE
End If
If Check1.Value = vbChecked Then
    tli.dwFlags = tli.dwFlags Or LIF_SILENT
End If
If Check2.Value = vbChecked Then
    tli.dwFlags = tli.dwFlags Or LIF_WARNINGABOVE
Else
    If Check3.Value = vbChecked Then
        tli.dwFlags = tli.dwFlags Or LIF_WARNINGOFF
    End If
End If
If Check4.Value = vbChecked Then
    tli.dwFlags = tli.dwFlags Or LIF_HIDETIPONVALID
End If
If Check5.Value = vbChecked Then
    tli.dwFlags = tli.dwFlags Or LIF_KEEPCLIPBOARD
End If
If Option1(1).Value = True Then
    tli.dwFlags = tli.dwFlags Or LIF_PASTESTOP
ElseIf Option1(2).Value = True Then
    tli.dwFlags = tli.dwFlags Or LIF_PASTECANCEL
End If
If Text2.Text <> "0" Then
    tli.dwMask = tli.dwMask Or LIM_TIMEOUT
    tli.iTimeout = CLng(Text2.Text)
End If
Dim hr As Long
Dim lerr As Long
hr = SHLimitInputEditWithFlags(Text1.hWnd, tli)
lerr = Err.LastDllError
Debug.Print "Result (0/0 indicates success): hr=0x" & Hex$(hr) & ", err=0x" & Hex$(lerr)
If Check6.Value = vbChecked Then
    Subclass2 Form1.hWnd, AddressOf FormWndProc
End If
Text1.SetFocus
End Sub

Private Sub Command2_Click()
Dim tli As LIMITINPUT
tli.cbSize = LenB(tli)
tli.dwMask = LIM_FILTER Or LIM_FLAGS
tli.dwFlags = LIF_CATEGORYFILTER Or LIF_WARNINGOFF
tli.pszFilter = LICF_DIGIT
Dim lerr As Long
Dim hr As Long
hr = SHLimitInputEditWithFlags(Text1.hWnd, tli)
lerr = Err.LastDllError
Debug.Print "Result (0,0 indicates success): hr=0x" & Hex$(hr) & ", err=0x" & Hex$(lerr)

End Sub

Private Sub Command3_Click()
SendMessage Text1.hWnd, EM_LIMITTEXT, CLng(Text5.Text), ByVal 0&
End Sub

