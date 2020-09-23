Attribute VB_Name = "Module1"
Option Explicit
'gets position and size of window relative to the parent
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
'releases dc and frees system resources
Declare Function ReleaseDC Lib "user32" (ByVal hwnd%, ByVal hdc%) As Integer
'this api call returns the parent of a given hwnd
Declare Function GetParent& Lib "user32" (ByVal hwnd As Long)
'get class name of a window using its hwnd
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
'this function is used to draw the rectangle around a window
Declare Function DrawFocusRect& Lib "user32" (ByVal hdc As Long, lpRect As RECT)
'used to load the cursor position into pointapi structure
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'gets handle of a window from its x and y coordinates
Declare Function WindowFromPoint& Lib "user32" (ByVal X As Long, ByVal Y As Long)
'obtains device context for a window given its handle
Declare Function GetDC Lib "user32" (ByVal hwnd%) As Integer
'this function loads the pointapi structure with upper left
'coordinates of a window give the window device context
Declare Function GetWindowOrgEx& Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI)
'used to copy one rect structure into another
Declare Function CopyRect& Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT)
'get the length of the text to allocate the buffer
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
'get the actual text of a window given the hwnd
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
'used to find topmost windows
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'used to find child windows
Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

'api rect structure is used for getting position
'and size of windows
Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

'the pointapi structure contains the x and y coordinates
'of the mouse
Public Type POINTAPI
    X As Long
    Y As Long
End Type

'do not touch these, they are used to draw focus rectangle
Global oldhwnd&, oldbox As RECT, olddc&
Function GetTopParent&(hwnd&)
'this function gets the handle of the top window
'in the hiarchy

Dim temp&
temp& = hwnd

Do: DoEvents
    'get the parent of a window
    temp& = GetParent(temp&)
    If temp& = 0 Then
        'if it doesnt have a parent then stop searching
        Exit Do
    Else
        'if it does have a parent then update handle
        GetTopParent = temp&
    End If
Loop
End Function
Sub DrawFocus()
'this function draws a rectangle around a window

Dim where As POINTAPI, box As RECT, hdc&
Dim difx&, dify&, retval&, hwnd&

'get new structure

'load cursor position into where pointapi structure
Call GetCursorPos(where)
'get the handle of the window using where coordinates
hwnd& = WindowFromPoint(where.X, where.Y)
'get device context of the window using the handle
hdc& = GetDC(hwnd&)
'get dimensions of the window using handle and load them
'into box which is RECT structure
retval& = GetWindowRect(hwnd, box)
'get upper left corner coordinates of a window
Call GetWindowOrgEx(hdc, where)

'check if its same as old

'check to see if the mouse moved to new window
'if yes then erase the old focus rectangle
If hwnd& = oldhwnd& Then
    'mouse hasnt moved
    Exit Sub
Else
    'mouse moved
    'erase rectangle
    Call DrawFocusRect(olddc, oldbox)
    'free resources
    retval = ReleaseDC(oldhwnd&, olddc&)
End If

'check if the window has a DC
If hdc& = 0 Then Exit Sub

'get dimensions

'get the difference between the actualy coordinates
'and coordinates relative to the parent
difx = box.Left - where.X
dify = box.Top - where.Y

'set the top left corner to actual coordinates
box.Left = where.X
box.Top = where.Y

'resize the rectangle
box.Right = box.Right - difx
box.Bottom = box.Bottom - dify

'draw the rectangle which indicates focus
Call DrawFocusRect(hdc, box)

'copy old structure
oldhwnd& = hwnd&
olddc = hdc&
Call CopyRect(oldbox, box)

End Sub
Function AlphaNumeric(ByRef what$) As Boolean
'check if a string consists of only numbers and letters
Dim i%, temp%
'cut the spaces
what$ = TrimSpaces(LCase(what$))

'this for loop goes through all the letters in the string
For i = 1 To Len(what$)
    'get the ascii value of a letter
    temp% = Asc(Mid(what$, i, 1))
    'check if its a letter or a number
    If (temp% >= 48 And temp% <= 57) Or (temp% >= 97 And temp% <= 122) Then
        AlphaNumeric = True
    Else
        'if its not then we stop checking and return false
        AlphaNumeric = False
        Exit Function
    End If
Next i
End Function
Function ReplaceAlpha(text As String)
'replace everything except for the aplha numeric characters
Dim i, temp%, tt$
tt$ = text$
tt$ = LCase(TrimSpaces(tt))
For i = 1 To Len(text)
    If AlphaNumeric(Mid(text, i, 1)) = True Then
        ReplaceAlpha = ReplaceAlpha & Mid(text, i, 1)
    End If
Next i
End Function
Function ReplaceString(ByVal text As String, ByVal what As String, ByVal WithWhat As String, Optional first As Boolean) As String
'set first to false to make it replace all otherswise it will
'replace only the first match
Dim pos As Integer
Dim ttemp As String, ttemp2 As String, ttemp3 As String
'check if the string you want to replace exists
pos = InStr(text, what)
'if no then exit function
If pos = 0 Then ReplaceString = text: Exit Function
ttemp = text
'if replacing all the matches
If first = False Then
    'loop until everything has been replaced
    Do: DoEvents
        pos = InStr(ttemp, what)
        If pos = 0 Then Exit Do
        ttemp2 = Left(ttemp, pos - 1)
        ttemp3 = Mid(ttemp, pos + Len(what))
        ttemp$ = ttemp2 & WithWhat & ttemp3
    Loop
'if replacing only the first match
Else
    pos = InStr(ttemp, what)
    ttemp2 = Left(ttemp, pos - 1)
    ttemp3 = Mid(ttemp, pos + Len(what))
    ttemp$ = ttemp2 & WithWhat & ttemp3
End If
ReplaceString$ = ttemp$
End Function
Function TrimSpaces(text$)
'remove all the spaces in a string
TrimSpaces = ReplaceString(text$, " ", "")
End Function
Public Function GetCaption(WindowHandle As Long) As String
'get caption of the windowhandle
Dim buffer As String, TextLength As Long
'get the length of the text
TextLength& = GetWindowTextLength(WindowHandle&)
'allocate the buffer
buffer$ = String(TextLength&, 0&)
'get the actual text
Call GetWindowText(WindowHandle&, buffer$, TextLength& + 1)
GetCaption$ = buffer$
End Function
Function GetWindowOrder%(hwnd&)
'this function checks if classname of a handle
'is unique in its hiarchy
Dim parent&, temp&, cname$, win&, name$
'get parent of the given window
parent& = GetParent(hwnd&)
'set total found to 0
GetWindowOrder% = 0

'get classname of the window
name$ = String(255, " ")
Call GetClassName(hwnd&, name$, 255)
name$ = Left(name$, InStr(name$, "    ") - 2)

Do: DoEvents
    'find all the windows with that class name
    temp& = FindWindowEx(parent&, win&, name$, vbNullString)
    If temp& <> 0 Then
        'if window found then update total found and
        'hwnd2 for findwindowex function
        win& = temp&
        GetWindowOrder = GetWindowOrder + 1
        If win& = hwnd& Then Exit Do
    Else
        'if no more windows then exit
        Exit Do
    End If
Loop
End Function
