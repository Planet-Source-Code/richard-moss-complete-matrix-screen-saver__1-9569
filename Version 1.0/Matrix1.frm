VERSION 5.00
Begin VB.Form frmScreenSaver 
   BorderStyle     =   0  'None
   ClientHeight    =   5670
   ClientLeft      =   2370
   ClientTop       =   1575
   ClientWidth     =   6585
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Matrix1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   378
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   439
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrUpdate 
      Interval        =   75
      Left            =   2925
      Top             =   2070
   End
End
Attribute VB_Name = "frmScreenSaver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefInt A-Z

Private LastX        As Single
Private LastY        As Single

Private ScrW%, ScrH%
Private TxtHght%, TxtWdth%
Private hMemDc&, hBmp&, hBmpOld&
Private hFont&, hFontOld&
Private MaxHeight
Private MinHeight

Private Type RECT
     rLeft As Long
     rTop As Long
     rRight As Long
     rBottom As Long
End Type

Private Rct As RECT

Private Type StringData
     CurX As Integer
     CurY As Integer
     Dy As Integer
     NumChars As Integer
End Type

Private Mtrx(1 To 100) As StringData   ' One Hundred Output Strings.

Private Declare Function BitBlt& Lib "gdi32" (ByVal hDestDC&, ByVal x1&, ByVal y1&, ByVal nWidth&, ByVal nHeight&, ByVal hSrcDC&, ByVal xSrc&, ByVal ySrc&, ByVal dwRop&)
Private Declare Function CreateCompatibleBitmap& Lib "gdi32" (ByVal hDC&, ByVal nWidth&, ByVal nHeight&)
Private Declare Function CreateCompatibleDC& Lib "gdi32" (ByVal hDC&)
Private Declare Function CreateSolidBrush& Lib "gdi32" (ByVal crColor As Long)
Private Declare Function DeleteDC& Lib "gdi32" (ByVal hDC&)
Private Declare Function DeleteObject& Lib "gdi32" (ByVal hObject&)
Private Declare Function FillRect& Lib "user32" (ByVal hDC&, lpRect As RECT, ByVal hBrush&)
Private Declare Function GetSystemMetrics& Lib "user32" (ByVal nIndex&)
Private Declare Function SelectObject& Lib "gdi32" (ByVal hDC&, ByVal hObject&)
Private Declare Function SendMessage& Lib "user32" Alias "SendMessageA" (ByVal hWnd&, ByVal wMsg&, ByVal wParam&, lParam As Any)
Private Declare Function SetBkMode& Lib "gdi32" (ByVal hDC&, ByVal nBkMode&)
Private Declare Function SetRect& Lib "user32" (lpRect As RECT, ByVal x1&, ByVal y1&, ByVal x2&, ByVal y2&)
Private Declare Function SetTextColor& Lib "gdi32" (ByVal hDC&, ByVal crColor&)
Private Declare Function TextOut& Lib "gdi32" Alias "TextOutA" (ByVal hDC&, ByVal x1&, ByVal y1&, ByVal lpString$, ByVal nCount&)

Private Const TRANSPARENT = 1

Private Const WM_GETFONT = &H31
'--------------------------------------------------
'Name        : UpdateFont
'Created     : 07/07/2000 08:07
'--------------------------------------------------
'Author      : Richard James Moss
'Organisation: Ariad Software
'--------------------------------------------------
'Description : Updates the font of the back buffer
'--------------------------------------------------
'Updates     :
'
'--------------------------------------------------
'          Ariad Procedure Builder Add-In 1.00.0036
Public Sub UpdateFont()
Attribute UpdateFont.VB_Description = "Updates the font of the back buffer"
 '##BD Updates the font of the back buffer
 If hFontOld Then
  DeleteObject SelectObject(hMemDc, hFontOld)
 End If
 ' Get The Form's Font (Courier, Regular, 15)... (Just Call Me Spock!).
 hFont = SendMessage(hWnd, WM_GETFONT, 0, 0&)
 ' Select It Into Our Back Buffer So We Can Output Text.
 hFontOld = SelectObject(hMemDc, hFont)
End Sub '(Public) Sub UpdateFont ()


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 Form_KeyPress KeyCode
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 If PreviewMode = 0 Then
  Unload Me
 End If
End Sub


Private Sub Form_Load()
 Dim Cols
 Dim K
 'setup values
 BackColor = BackgroundClr
 tmrUpdate.Interval = Speed
 Set Font = StringToFont(FontData$)
 'now screensaver
    ' Aquire The Screen Width And Height In Pixels.
    ScrW = GetSystemMetrics(0)
    ScrH = GetSystemMetrics(1)

    ' Setup A RECT Structure The Size Of The Screen.
    ' This Will Be Used Later With The API Function "FillRect"
    ' To Clear The Back Buffer.
    SetRect Rct, 0, 0, ScrW, ScrH
    
    ' Create An Off Screen Drawing Area In Memory (Back Buffer)... (Backbuffer,.. That Picture NoOne Can See).
    hMemDc = CreateCompatibleDC(0)
    hBmp = CreateCompatibleBitmap(hDC, ScrW, ScrH)
    hBmpOld = SelectObject(hMemDc, hBmp)
    SetBkMode hMemDc, TRANSPARENT

    UpdateFont
    
    TxtWdth = TextWidth("A")
    TxtHght = TextHeight("A")
    MaxHeight = ScrH - TxtHght

    ' Seed Random Number Generator.
    Randomize

    For K = 1 To 100
     Cols = Int(ScrW / TxtWdth)
        Mtrx(K).CurX = Int(Rnd * Cols) * TxtWdth 'Rnd * (ScrW - TxtWdth)
        Mtrx(K).NumChars = Int((20 - 5 + 1) * Rnd + 5)
        Mtrx(K).Dy = TxtHght + Rnd * TxtHght
        MinHeight = -2 * Mtrx(K).Dy * Mtrx(K).NumChars
        Mtrx(K).CurY = Int((MaxHeight - MinHeight + 1) * Rnd + MinHeight)
    Next 'showtime...
 If PreviewMode = 0 Then
  ScreenSaverActive = -1
  WindowState = 2
  CursorVisible = 0
  Show
 End If
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If PreviewMode = 0 Then
'  Unload Me
 End If
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If PreviewMode = 0 Then
  If (LastX = 0 And LastY = 0) Or (Abs(LastX - X) < 2 And Abs(LastY - Y) < 2) Then
   ' Small Mouse Movement...
   LastX = X
   LastY = Y
  Else
   ' Massive Mouse-Movement (Rat'ssssssssss)... End.
   Unload Me
  End If
 End If
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 ' Delete The Font We Created.
 DeleteObject SelectObject(hMemDc, hFontOld)
 ' Delete The Back Buffer.
 DeleteObject SelectObject(hMemDc, hBmpOld)
 DeleteDC hMemDc
 CursorVisible = -1
 ScreenSaverActive = 0
 End
End Sub


Private Sub tmrUpdate_Timer()
 Dim hBrush As Long
 Dim Char$
 Dim Cols
 Dim K, N
 Dim CY
 Dim MX
 ' Clear The BackBuffer.
 hBrush = CreateSolidBrush(BackgroundClr)
 FillRect hMemDc, Rct, hBrush
 DeleteObject hBrush
 ' Output Our Strings.
 For K = 1 To 100
  CY = Mtrx(K).CurY
  MX = Mtrx(K).NumChars
  For N = 1 To MX
   If N = MX Then ' Last Char In String.
    SetTextColor hMemDc, HighlightTextClr  ' The Brightest Letter.
   Else
    SetTextColor hMemDc, DimmedTextClr   ' The Darker Letters.
   End If
   ' OutPut The Character On The Back Buffer.
   Select Case CharacterSet
    Case 0           'complete
     Char$ = Chr$(Int((255 - 33 + 1) * Rnd + 33))
    Case 1           'binary
     Char$ = Chr$((Rnd * 1) + 48)
    Case Else        'custom
     If Len(CharacterSetChar) Then
      Char$ = Mid$(CharacterSetChar, Int(Rnd * (Len(CharacterSetChar & " ") - 1) + 1), 1)
     Else
      Char$ = Chr$((Rnd * 1) + 48)
     End If
   End Select
   TextOut hMemDc, Mtrx(K).CurX, CY, Char$, 1
   'End If
   CY = CY + Mtrx(K).Dy
  Next
  Mtrx(K).CurY = Mtrx(K).CurY + Mtrx(K).Dy
  If Mtrx(K).CurY > ScrH Then
   ' A String Has Now Left The Screen So
   ' Need To Initialize Another One.
   Cols = Int(ScrW / TxtWdth)
   Mtrx(K).CurX = Int(Rnd * Cols) * TxtWdth 'Rnd * (ScrW - TxtWdth)
'    Mtrx(K).CurX = Rnd * (ScrW - TxtWdth)
   Mtrx(K).NumChars = Int((20 - 5 + 1) * Rnd + 5)
   Mtrx(K).Dy = TxtHght + Rnd * (TxtHght \ 2)
   Mtrx(K).CurY = -2 * Mtrx(K).Dy * Mtrx(K).NumChars
  End If
 Next
 ' Now That The Off Screen Drawing Is Complete,
 ' Blit The Finished Frame Onto The Screen.
 BitBlt hDC, 0, 0, ScrW, ScrH, hMemDc, 0, 0, vbSrcCopy
End Sub


