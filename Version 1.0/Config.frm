VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConfig 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Matrix Configuration"
   ClientHeight    =   4740
   ClientLeft      =   1305
   ClientTop       =   2370
   ClientWidth     =   7800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Config.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   316
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAction 
      Caption         =   "Cha&nge..."
      Height          =   330
      Index           =   3
      Left            =   4995
      TabIndex        =   16
      Top             =   1530
      Width           =   1140
   End
   Begin VB.TextBox txtFont 
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   1215
      Width           =   2895
   End
   Begin VB.PictureBox picPreview 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1680
      Left            =   3570
      ScaleHeight     =   112
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   152
      TabIndex        =   20
      Top             =   2235
      Width           =   2280
   End
   Begin MSComctlLib.Slider sldSpeed 
      Height          =   420
      Left            =   3240
      TabIndex        =   13
      Top             =   360
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   741
      _Version        =   393216
      Min             =   1
      Max             =   150
      SelStart        =   75
      TickFrequency   =   10
      Value           =   75
   End
   Begin VB.Frame Frame2 
      Caption         =   "Character Set:"
      Height          =   1905
      Left            =   180
      TabIndex        =   7
      Top             =   1980
      Width           =   2895
      Begin VB.TextBox txtChar 
         Height          =   285
         Left            =   540
         TabIndex        =   11
         Top             =   1350
         Width           =   2040
      End
      Begin VB.OptionButton optChar 
         Caption         =   "C&ustom"
         Height          =   285
         Index           =   2
         Left            =   270
         TabIndex        =   10
         Top             =   1080
         Width           =   1230
      End
      Begin VB.OptionButton optChar 
         Caption         =   "&Binary"
         Height          =   285
         Index           =   1
         Left            =   270
         TabIndex        =   9
         Top             =   720
         Value           =   -1  'True
         Width           =   1230
      End
      Begin VB.OptionButton optChar 
         Caption         =   "&Complete"
         Height          =   285
         Index           =   0
         Left            =   270
         TabIndex        =   8
         Top             =   360
         Width           =   1230
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Colours:"
      Height          =   1770
      Left            =   180
      TabIndex        =   0
      Top             =   135
      Width           =   2895
      Begin MatrixScr.asxColourSelect colSelect 
         Height          =   330
         Index           =   0
         Left            =   1440
         TabIndex        =   2
         Top             =   315
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   582
         Colour          =   0
      End
      Begin MatrixScr.asxColourSelect colSelect 
         Height          =   330
         Index           =   1
         Left            =   1440
         TabIndex        =   4
         Top             =   720
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   582
         Colour          =   65280
      End
      Begin MatrixScr.asxColourSelect colSelect 
         Height          =   330
         Index           =   2
         Left            =   1440
         TabIndex        =   6
         Top             =   1125
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   582
         Colour          =   32768
      End
      Begin VB.Label lblHdr 
         AutoSize        =   -1  'True
         Caption         =   "&Dimmed Text:"
         Height          =   195
         Index           =   2
         Left            =   270
         TabIndex        =   5
         Top             =   1170
         Width           =   990
      End
      Begin VB.Label lblHdr 
         AutoSize        =   -1  'True
         Caption         =   "&Highlight Text:"
         Height          =   195
         Index           =   1
         Left            =   270
         TabIndex        =   3
         Top             =   765
         Width           =   1050
      End
      Begin VB.Label lblHdr 
         AutoSize        =   -1  'True
         Caption         =   "Back&ground:"
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   1
         Top             =   360
         Width           =   900
      End
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "&About..."
      Height          =   330
      Index           =   2
      Left            =   6480
      TabIndex        =   19
      Top             =   990
      Width           =   1140
   End
   Begin VB.CommandButton cmdAction 
      Cancel          =   -1  'True
      Caption         =   "Cance&l"
      Height          =   330
      Index           =   1
      Left            =   6480
      TabIndex        =   18
      Top             =   495
      Width           =   1140
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   330
      Index           =   0
      Left            =   6480
      TabIndex        =   17
      Top             =   135
      Width           =   1140
   End
   Begin VB.Label lblHdr 
      AutoSize        =   -1  'True
      Caption         =   "&Font:"
      Height          =   195
      Index           =   4
      Left            =   3240
      TabIndex        =   14
      Top             =   990
      Width           =   390
   End
   Begin VB.Label lblHdr 
      AutoSize        =   -1  'True
      Caption         =   "&Speed:"
      Height          =   195
      Index           =   3
      Left            =   3240
      TabIndex        =   12
      Top             =   135
      Width           =   510
   End
   Begin VB.Image Image1 
      Height          =   2535
      Left            =   3330
      Picture         =   "Config.frx":0442
      Stretch         =   -1  'True
      Top             =   1980
      Width           =   2775
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefInt A-Z

Private SaverForm As frmScreenSaver

Private DispFont As StdFont
Private Sub cmdAction_Click(Index As Integer)
 Dim I
 Dim TempFont As StdFont
 Select Case Index
  Case 0
   'save settings and close
   BackgroundClr = colSelect(0).Colour
   HighlightTextClr = colSelect(1).Colour
   DimmedTextClr = colSelect(2).Colour
   Speed = sldSpeed.Value
   For I = 0 To optChar.Count - 1
    If optChar(I) Then
     CharacterSet = I
    End If
   Next
   SaveSetting "Matrix ScreenSaver", "Version 1.0", "BackgroundColour", BackgroundClr
   SaveSetting "Matrix ScreenSaver", "Version 1.0", "HighlightTextColour", HighlightTextClr
   SaveSetting "Matrix ScreenSaver", "Version 1.0", "DimmedTextColour", DimmedTextClr
   SaveSetting "Matrix ScreenSaver", "Version 1.0", "Speed", Speed
   SaveSetting "Matrix ScreenSaver", "Version 1.0", "CharacterSet", CharacterSet
   SaveSetting "Matrix ScreenSaver", "Version 1.0", "CharacterSetChar", CharacterSetChar
   SaveSetting "Matrix ScreenSaver", "Version 1.0", "Font", FontData$
   Unload Me
  Case 1: Unload Me
  Case 2: About
  Case 3
   'change font
   Set TempFont = StringToFont(FontData$)
   If SelectFont(hWnd, TempFont) Then
    FontData$ = FontToString(TempFont)
    Set SaverForm.Font = StringToFont(FontData$)
    SaverForm.UpdateFont
    Set DispFont = StringToFont(FontData$)
    txtFont = DispFont.Size & "pt " & DispFont.Name
   End If
 End Select
End Sub


Private Sub colSelect_Change(Index As Integer, Color As stdole.OLE_COLOR)
 Select Case Index
  Case 0
   BackgroundClr = Color
   SaverForm.BackColor = Color
  Case 1: HighlightTextClr = Color
  Case 2: DimmedTextClr = Color
 End Select
End Sub

Private Sub Form_Load()
 'create preview saver form
 Set SaverForm = New frmScreenSaver
 'set values
 colSelect(0).Colour = BackgroundClr
 colSelect(1).Colour = HighlightTextClr
 colSelect(2).Colour = DimmedTextClr
 sldSpeed.Value = Speed
 optChar(CharacterSet) = -1
 txtChar = CharacterSetChar$
 Set DispFont = StringToFont(FontData$)
 txtFont = DispFont.Size & "pt " & DispFont.Name
 'show preview (cool!)
 PreviewSaver SaverForm, picPreview.hWnd
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 Unload SaverForm
 Set SaverForm = Nothing
End Sub


Private Sub optChar_Click(Index As Integer)
 CharacterSet = Index
End Sub

Private Sub sldSpeed_Change()
 Speed = sldSpeed
 SaverForm.tmrUpdate.Interval = sldSpeed
End Sub


Private Sub sldSpeed_Scroll()
 sldSpeed_Change
End Sub


Private Sub txtChar_Change()
 CharacterSetChar = txtChar
End Sub


