VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "API Spy By Phrost"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7155
   Icon            =   "apispy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   7155
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3000
      Top             =   2280
   End
   Begin VB.TextBox Text1 
      Height          =   1575
      Left            =   1560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   3360
      Width           =   5535
   End
   Begin VB.Frame Frame3 
      Caption         =   "Drag and Drop"
      Height          =   1575
      Left            =   0
      TabIndex        =   1
      Top             =   3360
      Width           =   1455
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         Height          =   1185
         Left            =   120
         Picture         =   "apispy.frx":0442
         ScaleHeight     =   1125
         ScaleWidth      =   1170
         TabIndex        =   2
         Top             =   240
         Width           =   1230
      End
   End
   Begin MSComctlLib.TreeView tree 
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   5741
      _Version        =   393217
      Indentation     =   4
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   4800
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "apispy.frx":49A8
            Key             =   "CONSULTANTS"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "apispy.frx":5284
            Key             =   "CLIENTS"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "apispy.frx":5B60
            Key             =   "JOBORDERS"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "apispy.frx":643C
            Key             =   "ENGAGEMENTS"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "apispy.frx":6D18
            Key             =   "LOGIN"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "apispy.frx":75FC
            Key             =   "ACCOUNTS"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "apispy.frx":7EE0
            Key             =   "COLORS"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "apispy.frx":81FC
            Key             =   "EXIT"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5040
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "apispy.frx":8AF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "apispy.frx":93D0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu code 
      Caption         =   "Code"
      Begin VB.Menu clipcopy 
         Caption         =   "Copy to Clipboard"
         Shortcut        =   ^C
      End
      Begin VB.Menu generate 
         Caption         =   "Generate Code"
         Shortcut        =   {F2}
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'force variable declaration
Option Explicit
'true/false used to tell program if it should
'draw focus rectangle and get the handle
Dim draggin As Boolean
Private Sub clipcopy_Click()
If Text1.text <> "" Then
    'clear clipboard
    Clipboard.Clear
    'copy generated code to cpipboard
    Call Clipboard.SetText(Text1.text)
End If
End Sub
Private Sub generate_Click()
Dim hwnd&, temp$, cname$, thandle&, tvar$(), i%
Dim name$, capt$, clname$, temp2$, old&, msg$, added As Boolean
'this boolean indicates if it should add i% to the DIM string
added = False
Text1.text = ""
On Error GoTo 1
'handle for which to generate code
hwnd = Mid(tree.Nodes.Item(tree.Nodes.Count).text, InStr(tree.Nodes.Item(tree.Nodes.Count).text, ":") + 1)
'if that window is the topmost window then we do the following
If GetParent(hwnd&) = 0 Then
    'get class name
    temp$ = String(255, " ")
    Call GetClassName(hwnd, temp$, 255)
    temp$ = Left(temp$, InStr(temp$, "           ") - 2)
    'get variable name
    cname$ = ReplaceAlpha(temp$)
    'generate code to find window using findwindow api call
    Text1.text = "DIM " & cname$ & "&" & vbCrLf & vbCrLf & cname$ & "& = FindWindow(""" & temp$ & """, vbNullString)"
    Exit Sub
Else
'if its not parent window then we have to go backwards
'through the windows hiarchy
    thandle& = hwnd&
    'size of dynamic array with variable declares
    i = 0
    Do: DoEvents
        'if the topmost parent window has been found
        If GetParent(thandle&) = 0 Then
            
            'load class name into temp2
            ReDim Preserve tvar(i)
            temp2$ = String(255, " ")
            Call GetClassName(thandle, temp2$, 255)
            temp2$ = Left(temp2$, InStr(temp2$, "           ") - 2)
            
            'get variable declaration name
            cname$ = ReplaceAlpha(temp2$)
            
            'update the array
            tvar(i) = cname$
            temp$ = "DIM "
            
            'generate declaration statement
            'preventing dupes
            For i = 0 To UBound(tvar)
                If InStr(temp$, tvar(UBound(tvar) - i)) = 0 Then temp$ = temp$ & tvar(UBound(tvar) - i) & "&, "
            Next i
            'check if loop was used to add the i% declaration
            If added = True Then temp$ = temp$ & "i%, "
            temp$ = Left(temp$, Len(temp$) - 2)
            
            'update the generated code
            Text1.text = temp$ & vbCrLf & vbCrLf & cname$ & "& = FindWindow(""" & temp2$ & """, vbNullString)" & vbCrLf & vbCrLf & Text1.text
            Text1.text = Left(Text1.text, Len(Text1.text) - 2)
            Exit Sub
        Else
            'get class name and variable declaration name
            temp$ = String(255, " ")
            Call GetClassName(thandle, temp$, 255)
            temp$ = Left(temp$, InStr(temp$, "           ") - 2)
            clname$ = temp$
            name$ = ReplaceAlpha(temp$)
            'this checks if window has a caption
            'if no it sets caption variable to vbnullstring
            capt$ = """" & GetCaption(thandle) & """"
            If capt = """""" Then capt = "vbnullstring"
            
            'update handle to search
            old& = thandle
            thandle = GetParent(thandle)
            
            'get new class name
            temp$ = String(255, " ")
            Call GetClassName(thandle, temp$, 255)
            temp$ = Left(temp$, InStr(temp$, "           ") - 2)
            cname$ = ReplaceAlpha(temp$)
            
            'resize array and assign variable names
            ReDim Preserve tvar(i)
            tvar(i) = name$
            i = i + 1
            
            'clear temp variable
            msg$ = ""
            
            'check if the window is the only window
            'with that classname in the hiarchy
            If GetWindowOrder(old&) = 1 Then
                'if it is then generate code using findwindowex
                msg = name$ & "& = FindWindowEx(" & cname$ & "&, 0, """ & clname$ & """, " & capt$ & ")" & vbCrLf & vbCrLf & msg
            Else
                'if its not then
                'find the first match for class name
                msg$ = name$ & "& = FindWindowEx(" & cname$ & "&, 0, """ & clname$ & """, " & capt$ & ")" & vbCrLf & vbCrLf & msg
                'check if there is more than 2 windows
                'with that class name, if yes then
                If GetWindowOrder(old) - 1 > 1 Then
                    'add the i% declaration
                    added = True
                    'generate a loop code and findwindowex code
                    msg$ = msg$ & "for i = 1 to " & GetWindowOrder(old) - 1 & vbCrLf & vbCrLf
                    msg$ = msg$ & name$ & "& = FindWindowEx(" & name$ & "&, 0, """ & clname$ & """, " & capt$ & ")" & vbCrLf & vbCrLf
                    msg$ = msg$ & "next i" & vbCrLf & vbCrLf
                Else
                    'if there is only 2 windows with that classname
                    'then second is found using the handle of the first
                    msg$ = msg$ & name$ & "& = FindWindowEx(" & name$ & "&, 0, """ & clname$ & """, " & capt$ & ")" & vbCrLf & vbCrLf
                End If
            End If
            'update the generated code
            Text1.text = msg$ & Text1.text
        End If
    Loop
End If
1 End Sub
Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'change mouse pointer
Form1.MousePointer = 10
'tell the program to start drawing focus rectangle
draggin = True
End Sub
Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim where As POINTAPI, hwnd&, cname$, tname$
Dim thwnd&, pname$, phwnd&, temp$, ptext$
Dim kname$, ktext$, nodx As Node

'reset pointer and treeview
Form1.MousePointer = 0
tree.Nodes.Clear

'erase the focus rectangle and reset global variables
Call DrawFocusRect(olddc, oldbox)
Call ReleaseDC(oldhwnd&, olddc&)
oldhwnd = 0
olddc = 0
draggin = False

'get cursor positon and window handle
Call GetCursorPos(where)
hwnd& = WindowFromPoint(where.X, where.Y)

'main hwnd
'get class name and add the node
cname$ = String(255, " ")
Call GetClassName(hwnd, cname$, 255)
temp$ = Left(cname$, InStr(cname$, "          ") - 2)
Set nodx = tree.Nodes.add(, , temp$, temp$, 1)

'top most
'get class name and add three nodes
'caption, class name, and handle
'add topmost node
Set nodx = tree.Nodes.add(temp$, tvwChild, "t", "Top Most Window", 2)
'get the topmost parent in the hiarchy
thwnd& = GetTopParent(hwnd)
'get class name
cname$ = String(255, " ")
Call GetClassName(thwnd, cname$, 255)
tname$ = Left(cname$, InStr(cname$, "           ") - 2)
'add 3 child nodes to main topmost node
Set nodx = tree.Nodes.add("t", tvwChild, "tcaption", "Caption: " & GetCaption(thwnd), 2)
Set nodx = tree.Nodes.add("t", tvwChild, "tname", "Class Name: " & tname$, 2)
Set nodx = tree.Nodes.add("t", tvwChild, "thwnd", "HWND: " & thwnd&, 2)
'make sure the branch is expanded
nodx.EnsureVisible

'parent
'get class name and add three nodes
'caption, class name, and handle
'add parent node
Set nodx = tree.Nodes.add(temp$, tvwChild, "p", "Parent Window", 2)
'get parent
phwnd& = GetParent(hwnd&)
'get class name
cname$ = String(255, " ")
Call GetClassName(phwnd, cname$, 255)
pname$ = Left(cname$, InStr(cname$, "            ") - 2)
'get caption
ptext$ = GetCaption(phwnd)
'add 3 child nodes to parent node
Set nodx = tree.Nodes.add("p", tvwChild, "pcaption", "Caption: " & ptext$, 2)
Set nodx = tree.Nodes.add("p", tvwChild, "pname", "Class Name: " & pname$, 2)
Set nodx = tree.Nodes.add("p", tvwChild, "phwnd", "HWND: " & phwnd&, 2)
'make sure the branch is expanded
nodx.EnsureVisible

'current
'get class name and add three nodes
'caption, class name, and handle
'add current node
Set nodx = tree.Nodes.add(temp$, tvwChild, "c", "Current Window", 2)
'get the class name
cname$ = String(255, " ")
Call GetClassName(hwnd, cname$, 255)
kname$ = Left(cname$, InStr(cname$, "         ") - 2)
'get caption
ktext$ = GetCaption(hwnd)
'add 3 child nodes to parent node
Set nodx = tree.Nodes.add("c", tvwChild, "ccaption", "Caption: " & ktext$, 2)
Set nodx = tree.Nodes.add("c", tvwChild, "cname", "Class Name: " & kname$, 2)
Set nodx = tree.Nodes.add("c", tvwChild, "chwnd", "HWND: " & hwnd&, 2)
'make sure the branch is expanded
nodx.EnsureVisible
End Sub
Private Sub Timer1_Timer()
'check if it focus rectangle should be draw
If draggin = False Then Exit Sub
'draw the focus rectangle
Call DrawFocus
End Sub
