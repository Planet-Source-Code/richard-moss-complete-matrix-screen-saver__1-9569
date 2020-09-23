Attribute VB_Name = "modMatrixScr"

'--------------------------------'
'            Matrix Screen Saver '
'                    Version 1.0 '
'--------------------------------'
'Copyright © 2000 by Ariad Software. All Rights Reserved

'Created        : 06/07/2000
'Completed      :
'Last Updated   :


Option Explicit
DefInt A-Z

Public BackgroundClr          As OLE_COLOR
Public HighlightTextClr       As OLE_COLOR
Public DimmedTextClr          As OLE_COLOR

Public Speed                  As Long

Public CharacterSet

Public CharacterSetChar$
Public FontData$
'---------------------------------------------------
'Name        : FontToString
'Created     : 26/05/2000 09:07
'---------------------------------------------------
'Author      : Richard Moss
'Organisation: Ariad Software
'---------------------------------------------------
'Description : Converts a font object into a string
'---------------------------------------------------
'Returns     : Returns a String Variable
'---------------------------------------------------
'Updates     :
'
'---------------------------------------------------
'           Ariad Procedure Builder Add-In 1.00.0036
Public Function FontToString(ByVal FontData As StdFont) As String
 '##BD Converts a font object into a string
 With FontData
  FontToString = .Name & "," & .Size & "," & Abs(.Bold) & "," & Abs(.Italic) & "," & Abs(.Underline) & "," & Abs(.Strikethrough)
 End With
End Function '(Public) Function FontToString () As String
'---------------------------------------------------
'Name        : StringToFont
'Created     : 26/05/2000 08:51
'---------------------------------------------------
'Author      : Richard Moss
'Organisation: Ariad Software
'---------------------------------------------------
'Description : Converts a string into a font object
'---------------------------------------------------
'Returns     : Returns a StdFont Object
'---------------------------------------------------
'Updates     :
'
'---------------------------------------------------
'           Ariad Procedure Builder Add-In 1.00.0036
Public Function StringToFont(ByVal FontData$) As StdFont
 '##BD Converts a string into a font object
 Dim Dat$()
 'Parse FontData$, Dat$(), ",", ptsmASSplit, True, 10
 Dat$ = Split(FontData$, ",")
 Set StringToFont = New StdFont
 With StringToFont
  .Name = Dat$(0)
  .Size = Val(Dat$(1))
  .Bold = Val(Dat$(2))
  .Italic = Val(Dat$(3))
  .Underline = Val(Dat$(4))
  .Strikethrough = Val(Dat$(5))
 End With
End Function '(Public) Function StringToFont () As StdFont

'-------------------------------------------------------------------
'Name        : About
'Created     : 06/07/2000 21:06
'-------------------------------------------------------------------
'Author      : Richard James Moss
'Organisation: Ariad Software
'-------------------------------------------------------------------
'Description : Displays version, copyright and contact information.
'-------------------------------------------------------------------
'Updates     :
'
'-------------------------------------------------------------------
'                           Ariad Procedure Builder Add-In 1.00.0036
Public Sub About()
Attribute About.VB_Description = "Displays version, copyright and contact information."
Attribute About.VB_UserMemId = -552
 '##BD Displays version, copyright and contact information.
 On Error Resume Next
  Dim Frm As frmAbout
  Set Frm = New frmAbout
  Frm.Show 1
 On Error GoTo 0
End Sub '(Public) Sub About ()

'---------------------------------
'Name        : Main
'Created     : 06/07/2000 20:27
'---------------------------------
'Author      : Richard James Moss
'Organisation: Ariad Software
'---------------------------------
'Description : Startup
'---------------------------------
'Updates     :
'
'---------------------------------
'           AS-PROCBUILD 1.00.0036
Public Sub Main()
 '##BD Startup
 If ScreenSaverActive() = 0 Then 'And App.PrevInstance = 0 Then
  PreviewMode = -1
  'load default settings
  BackgroundClr = GetSetting("Matrix ScreenSaver", "Version 1.0", "BackgroundColour", QBColor(0))
  HighlightTextClr = GetSetting("Matrix ScreenSaver", "Version 1.0", "HighlightTextColour", QBColor(10))
  DimmedTextClr = GetSetting("Matrix ScreenSaver", "Version 1.0", "DimmedTextColour", QBColor(2))
  Speed = GetSetting("Matrix ScreenSaver", "Version 1.0", "Speed", 75)
  CharacterSet = GetSetting("Matrix ScreenSaver", "Version 1.0", "CharacterSet", 1)
  CharacterSetChar = GetSetting("Matrix ScreenSaver", "Version 1.0", "CharacterSetChar")
  FontData$ = GetSetting("Matrix ScreenSaver", "Version 1.0", "Font", "Courier New,15,0,0,0,0")
  'ensure that neither the ss is running or a second copy of this app
  Select Case Left$(UCase$(Command$), 2)
   Case "/A"         'change password
    'no password support (yet)
    'so why not show about dialog?
    About
   Case "/C"         'config
    frmConfig.Show 1
   Case "/P"         'preview
    PreviewSaver frmScreenSaver, Val(Trim$(Mid$(Command$, 3)))
   Case "/S"         'display
    PreviewMode = 0
    Load frmScreenSaver
  End Select
 End If
End Sub '(Public) Sub Main ()

