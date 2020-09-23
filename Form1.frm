VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   1725
   LinkTopic       =   "Form1"
   ScaleHeight     =   690
   ScaleWidth      =   1725
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub SomeProc()
    Dim wnd&, Aimwindow&, wnd_chld As Long
    Dim uClickYes As Long
    Dim Res As Long
    Dim wasenabled As Long
    Dim a As String, j As Long
    
    DoEvents
    ' is there dialog box appear yet !
    Let wnd = FindWindow(vbNullString, "Microsoft Outlook")
    If (wnd = 0) Then Exit Sub
        Let wnd_chld = GetWindow(wnd, GW_CHILD)
            If (wnd_chld = 0) Then Exit Sub
    
        ' is dialog box enable yet
                wasenabled = IsWindowEnabled(wnd_chld)
                    If wasenabled = 0 Then Exit Sub
                            Do While wnd_chld
                                DoEvents
                                a = Space$(128)
                                ' get child box
                                j = SendMessage(wnd_chld, WM_GETTEXT, 100, ByVal a)
                                If (Len(Trim(a)) <> 0) Then
                                    a = Left(a, InStr(a, Chr$(0)) - 1)
                                    If (UCase(a) = "YES") Then
                                        ' is button YES enable yet
                                        wasenabled = IsWindowEnabled(wnd_chld)
                                        If wasenabled <> 0 Then
                                               Beep
                                               SefocusAPI wnd_chld
                                               keybd_event VK_TAB, 0, 0, 0
                                               keybd_event VK_TAB, 0, KEYEVENTF_KEYUP, 0
                        
                                               keybd_event VK_TAB, 0, 0, 0
                                               keybd_event VK_TAB, 0, KEYEVENTF_KEYUP, 0
                        
                                               keybd_event VK_ENTER, 0, 0, 0
                                               keybd_event VK_ENTER, 0, KEYEVENTF_KEYUP, 0
                                        End If
                                    End If
                                End If
                                DoEvents
                                wnd_chld = GetWindow(wnd_chld, GW_HWNDNEXT)
                            Loop
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
DoEvents
SomeProc
Timer1.Enabled = True
End Sub
