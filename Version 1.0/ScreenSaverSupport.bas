Attribute VB_Name = "modScreenSaverSupport"

'--------------------------------------'
'            Ariad Development Library '
'                          Version 3.0 '
'--------------------------------------'
'          Screen Saver Support Module '
'                          Version 1.0 '
'--------------------------------------'
'Copyright © 2000 by Ariad Software. All Rights Reserved

'Created        : 06/07/2000
'Completed      : 06/07/2000
'Last Updated   :


Option Explicit
DefInt A-Z

Private Type RECT
 Left    As Long
 Top     As Long
 Right   As Long
 Bottom  As Long
End Type

Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
Private Declare Function ShowCursor& Lib "user32" (ByVal bShow&)
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long

Private Const SPI_GETSCREENSAVEACTIVE = 16
Private Const SPI_SETSCREENSAVEACTIVE = 17

Private Const GWL_HWNDPARENT = (-8)
Private Const GWL_STYLE = (-16)

Private Const WS_CHILD = &H40000000

Private Const HWND_TOP = 0&

Private Const SWP_NOZORDER = &H4
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40

Public PreviewMode As Boolean
Property Let CursorVisible(ByVal CursorVisible As Boolean)
Attribute CursorVisible.VB_Description = "Returns or sets if the mouse cursor is visible"
 ShowCursor CLng(Abs(CursorVisible))
End Property ' Property Let CursorVisible

'----------------------------------------------------------------------
'Name        : PreviewSaver
'Created     : 06/07/2000 21:15
'----------------------------------------------------------------------
'Author      : Richard James Moss
'Organisation: Ariad Software
'----------------------------------------------------------------------
'Description : Positions the screensaver form within the specified
'              preview window handle
'----------------------------------------------------------------------
'Updates     :
'
'----------------------------------------------------------------------
'                              Ariad Procedure Builder Add-In 1.00.0036
Public Sub PreviewSaver(SSForm As Form, ByVal hWndPreview As Long)
Attribute PreviewSaver.VB_Description = "Positions the screensaver form within the specified preview window handle"
 '##BD Positions the screensaver form within the specified preview window handle
 Dim DispRect As RECT
 Dim hWnd As Long
 Dim Style As Long
 PreviewMode = -1
 'Get display rectangle dimensions
 GetClientRect hWndPreview, DispRect
 'Load form for preview
 Load SSForm
 'Get HWND for display form
 hWnd = SSForm.hWnd
 'Get current window style
 Style = GetWindowLong(hWnd, GWL_STYLE)
 'Append "WS_CHILD" style to the current window style
 Style = Style Or WS_CHILD
 'Add new style to display window
 SetWindowLong hWnd, GWL_STYLE, Style
 'Set display window as parent window
 SetParent hWnd, hWndPreview
 'Save the parent hWnd in the display form's window structure.
 SetWindowLong hWnd, GWL_HWNDPARENT, hWndPreview
 'Preview screensaver in the window...
 'With DispRect
 'SSForm.Move 0, 0, .Right * Screen.TwipsPerPixelX, .Bottom * Screen.TwipsPerPixelY
 'End With
 SetWindowPos hWnd, HWND_TOP, 0&, 0&, DispRect.Right, DispRect.Bottom, SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub '(Public) Sub PreviewSaver ()



'-------------------------------------------------------------------
'Name        : ScreenSaverActive
'Created     : 06/07/2000 20:35
'-------------------------------------------------------------------
'Author      : Richard James Moss
'Organisation: Ariad Software
'-------------------------------------------------------------------
'Description : Returns or sets if a screen saver is active
'-------------------------------------------------------------------
'Returns     : Returns True if the property is set, otherwise False
'-------------------------------------------------------------------
'Updates     :
'
'-------------------------------------------------------------------
'                           Ariad Procedure Builder Add-In 1.00.0036
Public Property Get ScreenSaverActive() As Boolean
Attribute ScreenSaverActive.VB_Description = "Returns or sets if a screen saver is active"
 '##BD Returns or sets if a screen saver is active
 Dim IsActive As Long
 SystemParametersInfo SPI_GETSCREENSAVEACTIVE, 0, IsActive, 0
 ScreenSaverActive = IsActive
End Property '(Public) Property Get ScreenSaverActive () As Boolean

Property Let ScreenSaverActive(ByVal ScreenSaverActive As Boolean)
 SystemParametersInfo SPI_SETSCREENSAVEACTIVE, ByVal CLng(Abs(ScreenSaverActive)), ByVal 0&, 0
End Property ' Property Let ScreenSaverActive

