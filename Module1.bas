Attribute VB_Name = "Module1"
Option Explicit
'=============================================================================
' ByPassMSOutlook
' Written by Mohamad Rawey Chek Ani, Perlis, Malaysia
' mrawey@yahoo.com.my
' If you use this application then please mention me in your credits
' Thanks and enjoy
' Please also vote for me :-)

'This program will clicks the Yes button on behalf of you, when
'Outlook's Security Guard opens prompt dialog
'saying that a program is trying to send an email with Outlook or
'access its address book.

'This had been tested for
'   Outlook 2000 SP1+SR1
'   Outlook 2000 SP2
'   Outlook 2002

Declare Function FindWindow Lib "user32" _
        Alias "FindWindowA" (ByVal lpClassName As Any, _
        ByVal lpWindowName As Any) As Long

Declare Function SendMessage Lib "user32" _
        Alias "SendMessageA" (ByVal hwnd As Long, _
        ByVal wMsg As Long, ByVal wParam As Long, _
        lParam As Any) As Long
Declare Function FindWindowEx Lib "user32" Alias _
        "FindWindowExA" (ByVal hWnd1 As Long, _
        ByVal hWnd2 As Long, ByVal lpsz1 As String, _
        ByVal lpsz2 As String) As Long
Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Declare Function IsWindowEnabled Lib "user32" _
        (ByVal hwnd As Long) As Long
Declare Function EnableWindow Lib "user32" _
        (ByVal hwnd As Long, ByVal fEnable As Long) As Long
        
Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, _
        ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Declare Function EnumChildWindows Lib "user32" _
        (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, _
        ByVal lParam As Long) As Long
Declare Function GetWindow Lib "user32" _
        (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function SefocusAPI Lib "user32" Alias "SetFocus" _
        (ByVal hwnd As Long) As Long

        
Public Const KEYEVENTF_EXTENDEDKEY = &H1
Public Const KEYEVENTF_KEYUP = &H2
Public Const VK_ENTER = &HD
Public Const VK_TAB = &H9

Public Const GW_HWNDNEXT = 2
Public Const GW_CHILD = 5
Public Const WM_GETTEXT = &HD
