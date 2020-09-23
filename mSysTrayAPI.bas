Attribute VB_Name = "mSysTrayAPI"
Option Explicit

''-------------------------------------------------------------------------------------
''Application defined variables
''-------------------------------------------------------------------------------------
Public frm As Form
Public IconObject As Object
Public lngPrevWndProc As Long 'Original WNDPROC address.
Public lngWndID As Long 'Our unique icon identifier.
Public lngHwnd As Long 'The hwnd of frmTray.
Public Notify As NOTIFYICONDATA
Public BarData As APPBARDATA
''-------------------------------------------------------------------------------------
''Application defined enumerations
''-------------------------------------------------------------------------------------
Public Enum ZoomTypes
    ZOOM_FROM_TRAY
    ZOOM_TO_TRAY
End Enum

''-------------------------------------------------------------------------------------
''WIN32 structures
''-------------------------------------------------------------------------------------
Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type

Type APPBARDATA
        cbSize As Long
        hwnd As Long
        uCallbackMessage As Long
        uEdge As Long
        rc As RECT
        lParam As Long
End Type
''-------------------------------------------------------------------------------------
''WIN32 API Constants
''-------------------------------------------------------------------------------------
Public Const GW_CHILD = 5
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GWL_WNDPROC = (-4)
Public Const IDANI_OPEN = &H1
Public Const IDANI_CLOSE = &H2
Public Const IDANI_CAPTION = &H3
Public Const NIF_TIP = &H4
Public Const NIM_ADD = 0&
Public Const NIM_DELETE = 2&
Public Const NIM_MODIFY = 1&
Public Const NIF_ICON = 2&
Public Const NIF_MESSAGE = 1&
Public Const ABM_GETTASKBARPOS = &H5&
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_MOUSEMOVE = &H200
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_USER = &H400

''-------------------------------------------------------------------------------------
''WIN32 DLL Declares
''-------------------------------------------------------------------------------------
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long

Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, _
    ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

Declare Function DrawAnimatedRects Lib "user32" (ByVal hwnd As Long, ByVal idAni As Long, lprcFrom As RECT, _
    lprcTo As RECT) As Long

Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
    ByVal hwnd As Long, ByVal Msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long

Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, _
    ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Declare Function SHAppBarMessage Lib "shell32.dll" (ByVal dwMessage As Long, pData As APPBARDATA) As Long

''-------------------------------------------------------------------------------------
''Application functions
''-------------------------------------------------------------------------------------
Public Function ZoomForm(zoomToWhere As ZoomTypes, hwnd As Long) As Boolean
'This function 'zooms' a window.
    Dim rctFrom As RECT
    Dim rctTo As RECT
    Dim lngTrayHand As Long
    Dim lngStartMenuHand As Long
    Dim lngChildHand As Long
    Dim strClass As String * 255
    Dim lngClassNameLen As Long
    Dim lngRetVal As Long

    'Select the type of zoom to do.
    Select Case zoomToWhere
        'Zoom the window into the tray.
        Case ZOOM_FROM_TRAY
            'Get the handle to the start menu.
            lngStartMenuHand = FindWindow("Shell_TrayWnd", vbNullString)

            'Get the handle to the first child window of the start menu.
            lngChildHand = GetWindow(lngStartMenuHand, GW_CHILD)

            'Loop through all siblings until we find the 'System Tray' (A.K.A. --> TrayNotifyWnd)
            Do
                lngClassNameLen = GetClassName(lngChildHand, strClass, Len(strClass))

                'If it is the tray then store the handle.
                If InStr(1, strClass, "TrayNotifyWnd") Then
                    lngTrayHand = lngChildHand
                    Exit Do
                End If
                'If we didn't find it, go to the next sibling.
                lngChildHand = GetWindow(lngChildHand, GW_HWNDNEXT)
            Loop

            'Get the RECT of  our form.
            lngRetVal = GetWindowRect(hwnd, rctFrom)

            'Get the RECT of the Tray.
            lngRetVal = GetWindowRect(lngTrayHand, rctTo)

            'Zoom from the tray to where our form is.
            lngRetVal = DrawAnimatedRects(frm.hwnd, IDANI_CLOSE Or IDANI_CAPTION, rctTo, rctFrom)

        Case ZOOM_TO_TRAY

            'Get the handle to the start menu.
            lngStartMenuHand = FindWindow("Shell_TrayWnd", vbNullString)

            'Get the handle to the first child window of the start menu.
            lngChildHand = GetWindow(lngStartMenuHand, GW_CHILD)

            'Loop through all siblings until we find the 'System Tray' (A.K.A. --> TrayNotifyWnd)
            Do
                lngClassNameLen = GetClassName(lngChildHand, strClass, Len(strClass))
                'If it is the tray then store the handle.
                If InStr(1, strClass, "TrayNotifyWnd") Then
                    lngTrayHand = lngChildHand
                    Exit Do
                End If
                'If we didn't find it, go to the next sibling.
                lngChildHand = GetWindow(lngChildHand, GW_HWNDNEXT)
            Loop
            'Get the RECT of  our form.
            lngRetVal = GetWindowRect(hwnd, rctFrom)

            'Get the RECT of the Tray.
            lngRetVal = GetWindowRect(lngTrayHand, rctTo)

            'Zoom from where our form is to the tray .
            lngRetVal = DrawAnimatedRects(frm.hwnd, IDANI_OPEN Or IDANI_CAPTION, rctFrom, rctTo)
    End Select
End Function

Public Sub modIcon(Form1 As Form, IconID As Long, Icon As Object, ToolTip As String)
'Modify an existing icon on the system tray
    Dim Result As Long
    Notify.cbSize = 88&
    Notify.hwnd = Form1.hwnd
    Notify.uID = IconID
    Notify.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
    Notify.uCallbackMessage = WM_MOUSEMOVE
    Notify.hIcon = Icon
    Notify.szTip = ToolTip & Chr$(0)
    Result = Shell_NotifyIcon(NIM_MODIFY, Notify)
End Sub

Public Sub AddIcon(Form1 As Form, IconID As Long, Icon As Object, ToolTip As String)
'Create an icon on the system tray
    Dim Result As Long
    BarData.cbSize = 36&
    Result = SHAppBarMessage(ABM_GETTASKBARPOS, BarData)
    Notify.cbSize = 88&
    Notify.hwnd = Form1.hwnd
    Notify.uID = IconID
    Notify.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
    Notify.uCallbackMessage = WM_MOUSEMOVE
    Notify.hIcon = Icon
    Notify.szTip = ToolTip & Chr$(0)
    Result = Shell_NotifyIcon(NIM_ADD, Notify)
End Sub

Public Sub delIcon(IconID As Long)
'Remove an icon from the system tray
    Dim Result As Long
    Notify.uID = IconID
    Result = Shell_NotifyIcon(NIM_DELETE, Notify)
End Sub







