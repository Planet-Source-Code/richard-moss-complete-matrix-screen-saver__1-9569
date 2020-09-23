VERSION 5.00
Begin VB.UserControl asxColourSelect 
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1230
   ScaleHeight     =   25
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   82
   ToolboxBitmap   =   "ColourSelect.ctx":0000
End
Attribute VB_Name = "asxColourSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'-----------------------------------------'
'            Ariad Development Components '
'-----------------------------------------'
'                ColourSelect UserControl '
'                             Version 1.0 '
'-----------------------------------------'
'Copyright © 1999 by Ariad Software. All Rights Reserved.

'Created        : 06/10/1999
'Completed      : 06/10/1999
'Last Updated   :


Option Explicit
DefInt A-Z

Private Type RECT
 Left       As Long
 Top        As Long
 Right      As Long
 Bottom     As Long
End Type

Private Type TCHOOSECOLOR
 lStructSize        As Long
 hWndOwner          As Long
 hInstance          As Long
 rgbResult          As Long
 lpCustColors       As Long
 Flags              As Long
 lCustData          As Long
 lpfnHook           As Long
 lpTemplateName     As Long
End Type

Private Declare Function ChooseColor Lib "COMDLG32.DLL" Alias "ChooseColorA" (Color As TCHOOSECOLOR) As Long
Private Declare Function DrawFocusRect& Lib "user32" (ByVal hDC As Long, lpRect As RECT)
Private Declare Function DrawFrameControl Lib "user32" (ByVal hDC&, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Boolean
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long

Private InFocus As Boolean
Private IsPushed As Boolean

Private pColour As OLE_COLOR

Public Event Change(Color As OLE_COLOR)
Private Function SelectColor(DefColor As Long, Optional ShowExpDlg As Boolean = 0, Optional InitCustomColours As Boolean = -1) As Long
 Dim I
 Dim C As Long
 Dim CC As TCHOOSECOLOR
 Dim CT$
 Dim CustomColors(16) As Long
 'Initialise Custom Colours
 If InitCustomColours Then
  For I = 0 To 15
   CT$ = GetSetting$("Ariad Non-ADL User Settings", "CustomColours", CStr(I))
   CustomColors(I) = IIf(Len(CT$), Val(CT$), QBColor(15))
  Next
 End If
 'Show Dialog
 With CC
  .rgbResult = DefColor
  .hWndOwner = hWnd
  .lpCustColors = VarPtr(CustomColors(0))
  .Flags = &H101
  If ShowExpDlg Then .Flags = .Flags Or &H2
  .lStructSize = Len(CC)
  C = ChooseColor(CC)
  If C Then
   SelectColor = .rgbResult
  Else
   SelectColor = -1
  End If
 End With
End Function

'------------------------------------------------------
'Name        : Colour
'Created     : 06/10/1999 19:41
'------------------------------------------------------
'Author      : Richard James Moss
'Organisation: Ariad Software
'------------------------------------------------------
'Description : Returns/sets the colour of the control.
'------------------------------------------------------
'Returns     : Returns an OLE_COLOR Variable
'------------------------------------------------------
'Updates     :
'
'------------------------------------------------------
'              Ariad Procedure Builder Add-In 1.00.0027
Public Property Get Colour() As OLE_COLOR
 Colour = pColour
End Property '(Public) Property Get Colour () As OLE_COLOR

Property Let Colour(ByVal Colour As OLE_COLOR)
 pColour = Colour
 PropertyChanged "Colour"
 Refresh
End Property ' Property Let Colour

'--------------------------------------------------------
'Name        : Refresh
'Created     : 06/10/1999 19:38
'--------------------------------------------------------
'Author      : Richard James Moss
'Organisation: Ariad Software
'--------------------------------------------------------
'Description : Forces a complete repaint of the control.
'--------------------------------------------------------
'Updates     :
'
'--------------------------------------------------------
'                Ariad Procedure Builder Add-In 1.00.0027
Public Sub Refresh()
 Dim Flags As Long
 Dim R As RECT
 Dim Z
 Const FR = 3
 Const CB = 5
 Z = Abs(IsPushed)
 Flags = 16
 If IsPushed Then Flags = Flags Or 512
 Line (-1, -1)-(ScaleWidth + 1, ScaleHeight + 1), vbButtonFace, BF
 'border
 R.Right = ScaleWidth
 R.Bottom = ScaleHeight
 DrawFrameControl hDC, R, 4, Flags
 'colour
 Line (CB + Z, CB + Z)-(ScaleWidth + Z - (CB + 1), ScaleHeight + Z - (CB + 1)), pColour, BF
 Line (CB + Z, CB + Z)-(ScaleWidth + Z - (CB + 1), ScaleHeight + Z - (CB + 1)), vbWindowText, B
 'focus
 If InFocus Then
  With R
   .Left = FR + Z
   .Top = FR + Z
   .Bottom = ScaleHeight - (FR - Z)
   .Right = ScaleWidth - (FR - Z)
  End With
  DrawFocusRect hDC, R
 End If
End Sub '(Public) Sub Refresh ()


Private Sub UserControl_Click()
 Dim C As Long, D As Long
 OleTranslateColor pColour, 0, D
 C = SelectColor(D)
 If C <> -1 Then
  Colour = C
  RaiseEvent Change(C)
 End If
End Sub

Private Sub UserControl_GotFocus()
 InFocus = -1
 Refresh
End Sub

Private Sub UserControl_Initialize()
 AutoRedraw = -1
End Sub


Private Sub UserControl_InitProperties()
 pColour = vbWhite
 Refresh
End Sub


Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 32 Then
  IsPushed = -1
  Refresh
 End If
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = 32 Then
  IsPushed = 0
  Refresh
  UserControl_Click
 End If
End Sub


Private Sub UserControl_LostFocus()
 InFocus = 0
 Refresh
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 InFocus = -1
 IsPushed = -1
 Refresh
End Sub


Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 IsPushed = 0
 Refresh
End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
 With PropBag
  pColour = .ReadProperty("Colour", vbWhite)
 End With
 Refresh
End Sub

Private Sub UserControl_Resize()
 Refresh
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
 With PropBag
  .WriteProperty "Colour", pColour, vbWhite
 End With
 Refresh
End Sub


