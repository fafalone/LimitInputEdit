# LimitInputEdit
## SHLimitInputWithFlags Demo
### Easily apply category filters, paste handling, and automated tooltips to an edit control.

![image](https://github.com/fafalone/LimitInputEdit/assets/7834493/33f2b7ba-dd27-460e-b8e1-a94ba649750b)

---

This is an updated, x64 compatible version of my VB6 demo of the shell32.dll API `SHLimitInputEditWithFlags`, an API entirely undocumented either by Microsoft or 3rd parties until my project. There are three versions of the project in this repository, LimitInputEditWithFlags.twinproj, a full twinBASIC version that uses `Return` and `Handles` syntax etc, and a universal compatibility version (VB6, VBA6, VBA7 32bit, VBA7 64bit, twinBASIC 32bit, twinBASIC 64bit) in both VB6 and twinBASIC form with identical code.

Original description:


Microsoft being Microsoft, they only begrudgingly documented a function called `SHLimitInputEdit` for the DOJ settlement, and did so poorly. This is a weird function; it takes an edit hwnd, and an object that implements `IShellFolder` and `IItemNameLimits`. The former doesn't even matter (unless it's been implemented in newer versions of Windows; I haven't checked). When you implement `IItemNameLimits`, you get a single call to `GetValidCharacters`, where you can supply a string of either included or excluded characters (only 1 can be used, so if you specify any excluded characters, included becomes null). It's an odd way of doing things.

But it turns out, that's a front end for an actually much more interesting and useful but completely undocumented, `SHLimitInputEditWithFlags`, an API Geoff Chappell found as exported at ordinal #754 in shell32.dll (it's still ordinal only in Windows 10, even though it's been kicking around since Windows XP).

This function allows a wide variety of options for limits and a tooltip that pops up upon bad input. Instead of just being able to specify an exact string, you can use `CT_TYPE1` categories, which in addition to the standard upper, lower, digits... has some handy options like categories for hexadecimal, punctuation, or control characters. It also implements custom categories; binary, octal, and ASCII a-z/A-Z. It also provides control over a tooltip that pops up when you attempt to enter a bad character-- you can have no tooltip, or specify the title, message, and icon (a `TTI_*` default icon or custom hIcon), and set alignment, width, and timeout (including timing out immediately if a valid input is received). It also handles pasting in several different ways; filtering in the valid chars, pasting until the 1st invalid char, or canceling the paste. If the paste is modified, it puts what was pasted on the clipboard (optionally). The pasting options and automatic control over the tooltip is what really makes this worthwhile over just manually checking KeyPress events or `WM_CHAR` messages.

**Requirements**\
-No dependencies.\
-Function present on Windows XP through at least Windows 10 (I haven't checked 11).

```vb6
#If VBA7 Then
Public Declare PtrSafe Function SHLimitInputEditWithFlags Lib "shell32" Alias "#754" (ByVal hwndEdit As LongPtr, pil As LIMITINPUTSTRUCT) As Long
'...
#Else
Public Declare Function SHLimitInputEditWithFlags Lib "shell32" Alias "#754" (ByVal hwndEdit As Long, pil As LIMITINPUTSTRUCT) As Long
'...
```

On all Windows versions, this function is exported by ordinal only, so you can't remove the Alias. `SHLimitInputEditWithFlags` takes two arguments, an hWnd for an edit control, and an (until this post) undocumented structure. Here's the members and a description:

```vb6
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
```

`dwMask` is just a list of which of the remaining members should be used:

```vb6
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
```

Now we'll get into the core of it with the flags for `dwFlags`:

```vb6
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
```
If you do not use the `LIF_CATEGORYFILTER` flag, the `.pszFilter` member must be set to `StrPtr(value)` where value is a non-delimited string of which characters to allow (by default) or disallow (if `LIF_EXCLUDEFILTER` flag is included). If you do use the flag, the following categories are valid:

```vb6
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
```

From there, you're all set to apply basic input limits to an edit control. Remember, if you don't want a tooltip you don't need to set the title, message, and icon, but in that case you must include the LIF_WARNINGOFF flag, or the function will fail. If you are going to have a tooltip, you must as a minimum specify the message.

### Advanced

There's a couple flags for advanced options. `LIF_NOTIFYONBADCHAR` will send hWnd specified by the .hwndNotify member a `LIN_BADCHAR` notification code in a `WM_NOTIFY` message. You must subclass the specified hWnd to receive the message (on Windows 10, it will not automatically send them to the parent, but directly to the provided hWnd. That automatic behavior may work on earlier versions, but manually specifying it works on all). From there it has it's own NM structure to copy:

```vb6
Private Type NMLIBADCHAR
    hdr As NMHDR
    wParam As LongPtr 'WM_CHAR wParam (Char code)
    lParam As LongPtr 'WM_CHAR lParam (see MSDN for details)
End Type
```

That gives you the WM_CHAR message.

There's also special handling for WM_PASTE operations built in. The default behavior is to paste whatever characters from the clipboard are allowed, then set the contents of the clipboard to the filtered result. You can change that behavior to only pasting up until the first disallowed character with the `LIF_PASTESTOP` flag, or to cancel the paste entirely with `LIF_PASTECANCEL`.

### Callbacks

I didn't implement this option in the demo because I don't see a lot of utility for it, but you can specify `LPSTR_TEXTCALLBACK` for the text fields, and `I_ICONCALLBACK` for the icon field, and the control will send a `LIN_GETDISPINFO` message for the tooltip text and `LIN_GETFILTERINFO` for the filter. I'm not going to detail it, but it works exactly like `LVN_GETDISPINFO` callbacks for the ListView control, and there's plenty of documentation for that. The constants and structure are included in the Demo if you did want to explore this.

### Sample Project

The demo pictured at the top of this post implements a wide array of features, including subclassing for the bad character notifications, but also includes a simple 'Set to numbers only' to show how simple calls to this function can be:

```vb6
Dim tli As LIMITINPUTSTRUCT
tli.cbSize = Len(tli)
tli.dwMask = LIM_FILTER Or LIM_FLAGS
tli.dwFlags = LIF_CATEGORYFILTER Or LIF_WARNINGOFF
tli.pszFilter = LICF_DIGIT

SHLimitInputEditWithFlags Text1.hWnd, tli
```

That's all you need to do to have a textbox take only numbers, with no tooltip.

And that's it! Enjoy this undocumented treasure from the Windows API.

---

**Thanks:** Thanks ToddB, it is super cool and should have been a documented tool the world knew

>[!IMPORTANT]
>This is an undocumented, internal API, with all the issues that involves. There may be small variations in functionality between Windows versions, stability is not guaranteed, and it may be removed at any time from future versions, or have it's ordinal changed.

**UPDATE**\
02 Jul 2023: UndocEditLimit-R2.zip (Revision 2) corrects an odd bug where the custom filters weren't working because despite being declared as a Unicode string type (LPWSTR), the API handled it as an ANSI string, and thus is was necessary to convert first.\
13 Dec 2023: Project updated for x64 compatibility and re-released as a universal compatibility version.
