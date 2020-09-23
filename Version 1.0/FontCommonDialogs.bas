Attribute VB_Name = "modFontCommonDialogs"
Attribute VB_HelpID = 3216

'--------------------------------------'
'            Ariad Development Library '
'                          Version 3.0 '
'--------------------------------------'
'              API Font Common Dialogs '
'                          Version 1.0 '
'--------------------------------------'
'Copyright © 1999 by Ariad Software. All Rights Reserved.

'Based on original code by Steve McMahon
'of vbAccelerator (http://www.vbaccelerator.com)

'Created        : 24/09/1999
'Completed      : 24/09/1999
'Last Updated   :


Option Explicit
DefInt A-Z

Private Declare Sub CopyMemoryStr Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, ByVal lpvSource As String, ByVal cbCopy As Long)

Private Declare Function ChooseFont Lib "COMDLG32" Alias "ChooseFontA" (chfont As TCHOOSEFONT) As Long
Private Declare Function CommDlgExtendedError Lib "COMDLG32.DLL" () As Long

Private Const LF_FACESIZE = 32

Private Type TCHOOSEFONT
 lStructSize        As Long     ' Filled with UDT size
 hWndOwner          As Long     ' Caller's window handle
 hDC                As Long     ' Printer DC/IC or NULL
 lpLogFont          As Long     ' Pointer to LOGFONT
 iPointSize         As Long     ' 10 * size in points of font
 Flags              As Long     ' Type flags
 rgbColors          As Long     ' Returned text color
 lCustData          As Long     ' Data passed to hook function
 lpfnHook           As Long     ' Pointer to hook function
 lpTemplateName     As Long     ' Custom template name
 hInstance          As Long     ' Instance handle for template
 lpszStyle          As String   ' Return style field
 nFontType          As Integer  ' Font type bits
 iAlign             As Integer  ' Filler
 nSizeMin           As Long     ' Minimum point size allowed
 nSizeMax           As Long     ' Maximum point size allowed
End Type

Private Type LOGFONT
 lfHeight                   As Long
 lfWidth                    As Long
 lfEscapement               As Long
 lfOrientation              As Long
 lfWeight                   As Long
 lfItalic                   As Byte
 lfUnderline                As Byte
 lfStrikeOut                As Byte
 lfCharSet                  As Byte
 lfOutPrecision             As Byte
 lfClipPrecision            As Byte
 lfQuality                  As Byte
 lfPitchAndFamily           As Byte
 lfFaceName(LF_FACESIZE)    As Byte
End Type

Enum FDFontFlags
 CF_SCREENFONTS = &H1
 CF_PRINTERFONTS = &H2
 CF_BOTH = &H3
 CF_FONTSHOWHELP = &H4
 CF_USESTYLE = &H80
 CF_EFFECTS = &H100
 CF_ANISONLY = &H400
 CF_NOVECTORFONTS = &H800
 CF_NOOEMFONTS = CF_NOVECTORFONTS
 CF_NOSIMULATIONS = &H1000
 CF_LIMITSIZE = &H2000
 CF_FIXEDPITCHONLY = &H4000
 CF_WYSIWYG = &H8000  ' Must also have ScreenFonts And PrinterFonts
 CF_FORCEFONTEXIST = &H10000
 CF_SCALABLEONLY = &H20000
 CF_TTONLY = &H40000
 CF_NOFACESEL = &H80000
 CF_NOSTYLESEL = &H100000
 CF_NOSIZESEL = &H200000
End Enum

Public Const CF_INITTOLOGFONTSTRUCT = &H40
Public Const CF_FONTNOTSUPPORTED = &H238

Public ApiReturn As Long, ExtendedError As Long
Attribute ApiReturn.VB_VarHelpID = 3219
Attribute ExtendedError.VB_VarHelpID = 3220

'----------------------------------------------------------------------
'Name        : SelectFont
'Created     : 24/09/1999 21:11
'----------------------------------------------------------------------
'Author      : Richard James Moss
'Organisation: Ariad Software
'----------------------------------------------------------------------
'Description : Allows the user to select a font from an API create
'              common dialog.
'----------------------------------------------------------------------
'Returns     : Returns True on success
'----------------------------------------------------------------------
'Updates     :
'
'----------------------------------------------------------------------
'                              Ariad Procedure Builder Add-In 1.00.0027
Public Function SelectFont(ByVal hWndOwner As Long, CurFont As StdFont, Optional Colour As OLE_COLOR = -1, Optional MinSize As Long = 0, Optional MaxSize As Long = 0, Optional Flags As FDFontFlags = CF_FORCEFONTEXIST Or CF_SCREENFONTS) As Boolean
Attribute SelectFont.VB_HelpID = 3221
 Dim CF As TCHOOSEFONT
 Dim Fnt As LOGFONT
 ApiReturn = 0
 ExtendedError = 0
 If Colour <> -1 Then Flags = Flags Or CF_EFFECTS
 If MinSize Then Flags = Flags Or CF_LIMITSIZE
 If MaxSize Then Flags = Flags Or CF_LIMITSIZE
 Flags = (Flags Or CF_INITTOLOGFONTSTRUCT) And Not CF_FONTNOTSUPPORTED
 ' Initialize LOGFONT variable
 Fnt.lfHeight = -(CurFont.Size * ((1440 / 72) / Screen.TwipsPerPixelY))
 Fnt.lfWeight = CurFont.Weight
 Fnt.lfItalic = CurFont.Italic
 Fnt.lfUnderline = CurFont.Underline
 Fnt.lfStrikeOut = CurFont.Strikethrough
 ' Other fields zero
 StrToBytes Fnt.lfFaceName, CurFont.Name
 ' Initialize TCHOOSEFONT variable
 CF.lStructSize = Len(CF)
 CF.hWndOwner = hWndOwner
 CF.lpLogFont = VarPtr(Fnt)
 CF.iPointSize = CurFont.Size * 10
 CF.Flags = Flags
 CF.rgbColors = Colour
 CF.nSizeMin = MinSize
 CF.nSizeMax = MaxSize
 ' All other fields zero
 ApiReturn = ChooseFont(CF)
 Select Case ApiReturn
  Case 1
   ' Success
   SelectFont = -1
   Flags = CF.Flags
   Colour = CF.rgbColors
   CurFont.Bold = CF.nFontType And &H100
   'CurFont.Italic = cf.nFontType And Italic_FontType
   CurFont.Italic = Fnt.lfItalic
   CurFont.Strikethrough = Fnt.lfStrikeOut
   CurFont.Underline = Fnt.lfUnderline
   CurFont.Weight = Fnt.lfWeight
   CurFont.Size = CF.iPointSize / 10
   CurFont.Name = BytesToStr(Fnt.lfFaceName)
  Case 0
   ' Cancelled
   SelectFont = 0
  Case Else
   ' Extended error
   ExtendedError = CommDlgExtendedError()
   SelectFont = 0
 End Select
End Function '(Public) Function SelectFont () As StdFont

Private Function BytesToStr(ab() As Byte) As String
Attribute BytesToStr.VB_HelpID = 3222
 BytesToStr = StrConv(ab, vbUnicode)
End Function

Private Sub StrToBytes(ab() As Byte, s As String)
Attribute StrToBytes.VB_HelpID = 3223
 Dim Cab As Long
 If IsArrayEmpty(ab) Then
  ' Assign to empty array
  ab = StrConv(s, vbFromUnicode)
 Else
  ' Copy to existing array, padding or truncating if necessary
  Cab = UBound(ab) - LBound(ab) + 1
  If Len(s) < Cab Then s = s & String$(Cab - Len(s), 0)
  CopyMemoryStr ab(LBound(ab)), s, Cab
 End If
End Sub

Private Function IsArrayEmpty(Arr As Variant) As Boolean
Attribute IsArrayEmpty.VB_HelpID = 3224
 Dim V As Variant
 On Error Resume Next
  V = Arr(LBound(Arr))
  IsArrayEmpty = (Err <> 0)
 On Error GoTo 0
End Function

