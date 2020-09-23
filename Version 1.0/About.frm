VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   3750
   ClientLeft      =   2550
   ClientTop       =   2985
   ClientWidth     =   5250
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   250
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   350
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   330
      Left            =   4005
      TabIndex        =   0
      Top             =   3330
      Width           =   1140
   End
   Begin VB.Label lblWeb 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "support@ariad-software.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   1
      Left            =   945
      MouseIcon       =   "About.frx":1042
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Tag             =   "mailto:support@ariad-software.com?subject=%APP%"
      ToolTipText     =   "Click to send email to Ariad Technical Support"
      Top             =   2835
      Width           =   2475
   End
   Begin VB.Label lblWeb 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "http://www.ariad-software.com/"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   945
      MouseIcon       =   "About.frx":134C
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Tag             =   "http://www.ariad-software.com/"
      ToolTipText     =   "Click to visit Ariad Software Online on the web"
      Top             =   2610
      Width           =   2790
   End
   Begin VB.Image imgAriad 
      Height          =   480
      Left            =   150
      Picture         =   "About.frx":1656
      Top             =   135
      Width           =   480
   End
   Begin VB.Label lblApp 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Application Name and Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   945
      TabIndex        =   2
      Top             =   135
      Width           =   2475
   End
   Begin VB.Label lblCopy 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Application Comments, Copyright and Trademarks"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   945
      TabIndex        =   1
      Top             =   450
      Width           =   4155
      WordWrap        =   -1  'True
   End
   Begin VB.Line lneShad 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   0
      X2              =   370
      Y1              =   213
      Y2              =   213
   End
   Begin VB.Line lneHigh 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   0
      X2              =   370
      Y1              =   214
      Y2              =   214
   End
   Begin VB.Shape shpAbout 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   645
      Index           =   0
      Left            =   0
      Top             =   3195
      Width           =   5550
   End
   Begin VB.Line lneShad 
      BorderColor     =   &H80000010&
      Index           =   1
      X1              =   54
      X2              =   54
      Y1              =   216
      Y2              =   -9
   End
   Begin VB.Shape shpAbout 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   3660
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   825
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_HelpID = 2338

'-----------------------------------------'
'            Ariad Development Components '
'-----------------------------------------'
'                     Simple About Dialog '
'                             Version 1.0 '
'-----------------------------------------'
'Copyright © 2000 by Ariad Software. All Rights Reserved.

'Created        : 23/02/2000
'Completed      : 23/02/2000
'Last Updated   :

'Nice slimline dialog, instead of the new Ariad logo
'dialog with a massive 365KB bitmap!!! (Thankfully now reduced to 29KB rle!)

Option Explicit
DefInt A-Z

Private Declare Function ShellExecute& Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long)

Private Sub cmdOK_Click()
Attribute cmdOK_Click.VB_HelpID = 2341
 Unload Me
End Sub


Private Sub Form_Activate()
Attribute Form_Activate.VB_HelpID = 2444
 Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
 Dim Comment$
 On Error Resume Next
  Set Icon = Nothing
  lblApp = App.Title & " " & App.Major & "." & App.Minor
  If App.Revision Then
   lblApp = lblApp & " (Build " & Format$(App.Revision, "000)")
  End If
  Comment$ = App.Comments
  If Len(Comment$) Then Comment$ = Comment$ & vbCr & vbCr
  lblCopy = Comment$ & App.LegalCopyright & vbCr & vbCr & App.LegalTrademarks
 On Error GoTo 0
End Sub


Private Sub lblWeb_Click(Index As Integer)
Attribute lblWeb_Click.VB_HelpID = 2343
 Dim Ret As Long
 Dim Cmnd$
 On Error Resume Next
  Cmnd$ = lblWeb(Index).Tag
  Cmnd$ = Replace$(Replace$(Cmnd$, "%APP%", lblApp.Caption), " ", "%20")
  Ret = ShellExecute(hWnd, "Open", Cmnd$, "", "", 5)
  If Err Then
   MsgBox Error$ & " (" & Err & ")", vbCritical, "Web Access"
  ElseIf Ret <= 32 Then
   MsgBox "Unable to run web document (" & Ret & ")", vbCritical, "Web Access"
  End If
 On Error GoTo 0
End Sub


