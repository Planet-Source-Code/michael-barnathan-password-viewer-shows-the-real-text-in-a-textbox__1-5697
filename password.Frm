VERSION 4.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Password finder"
   ClientHeight    =   810
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   3000
   Height          =   1215
   Icon            =   "password.frx":0000
   Left            =   1080
   LinkTopic       =   "Form1"
   ScaleHeight     =   810
   ScaleWidth      =   3000
   Top             =   1170
   Width           =   3120
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2760
      Top             =   2760
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Last password decoded: None"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   2235
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Child being scanned for password: None"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2925
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function FindWindow& Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String)
Private Declare Function GetWindow& Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long)
Private Declare Function Sendmessagebynum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Private Declare Function SendMessageByString& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String)
Private Declare Function GetCursorPos& Lib "user32" (lpPoint As POINTAPI)
Private Declare Function WindowFromPoint& Lib "user32" (ByVal x As Long, ByVal y As Long)
Private Declare Function ChildWindowFromPoint& Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long)
Private Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function UpdateWindow& Lib "user32" (ByVal hwnd As Long)
Const SWP_NOACTIVATE = &H10
Const SWP_NOREDRAW = &H8
Const SWP_NOSIZE = &H1
Const SWP_NOZORDER = &H4
Const SWP_NOMOVE = &H2
Const HWND_TOPMOST = -1
Const HWND_BOTTOM = 1
Const SWP_HIDEWINDOW = &H80
Const WM_SETTEXT = &HC
Const WM_GETTEXT = &HD
Const WM_CHAR = &H102
Const WM_CLEAR = &H303
Const GW_CHILD = 5
Const GW_HWNDNEXT = 2
Const EM_SETPASSWORDCHAR = &HCC
Const EM_GETPASSWORDCHAR = &HD2
Const EN_CHANGE = &H300
Dim Abort, LastWindow&, LastCaption$
Private Type POINTAPI
    x As Long
    y As Long
End Type
Function GetCaption(hwnd) As String
'hwndlength% = GetWindowTextLength(hwnd)
'hwndTitle$ = String$(hwndlength%, 0)
'A% = GetWindowText(hwnd, hwndTitle$, (hwndlength% + 1))
'GetCaption = hwndTitle$
Capt$ = Space$(255)
TChars$ = GetWindowText(hwnd, Capt$, 255)
GetCaption = Left$(Capt$, TChars$)
End Function

Function GetText(hwnd) As String
GetTrim = Sendmessagebynum(hwnd, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
GetString = SendMessageByString(hwnd, 13, GetTrim + 1, TrimSpace$)

GetText = TrimSpace$
End Function

Private Sub Form_Load()
Call SetWindowPos(Form1.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_NOMOVE Or SWP_NOSIZE)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Abort = 1
End Sub

Private Sub Timer1_Timer()
Dim mypoint As POINTAPI
Call GetCursorPos(mypoint)
A& = WindowFromPoint(mypoint.x, mypoint.y)
If LastWindow& <> A& And LastWindow& <> 0 Then Call Sendmessagebynum(LastWindow&, EM_SETPASSWORDCHAR, Asc("*"), 0&): DoEvents: Call SendMessageByString(LastWindow&, WM_SETTEXT, 0&, LastCaption$): LastWindow& = 0 ': LastWindow& = A&: LastCaption$ = GetText(A&)
B& = ChildWindowFromPoint(A&, mypoint.x, mypoint.y)
If A& = Form1.hwnd Then Exit Sub
Label2.Caption = GetCaption(A&)
Label3.Caption = GetText(A&)
If Sendmessagebynum(A&, EM_GETPASSWORDCHAR, 0&, 0&) <> 0 Then Call Sendmessagebynum(A&, EM_SETPASSWORDCHAR, 0&, 0&): LastWindow& = A&: LastCaption$ = GetText(A&): DoEvents: LastWindow& = A&: LastCaption$ = GetText(A&): Call SendMessageByString(A&, WM_SETTEXT, 0&, Label3.Caption)
End Sub


