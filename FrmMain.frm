VERSION 5.00
Begin VB.Form FrmMain 
   BackColor       =   &H000000FF&
   Caption         =   "Window Transparency Settings"
   ClientHeight    =   3345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3345
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll3 
      Height          =   255
      LargeChange     =   10
      Left            =   120
      Max             =   255
      TabIndex        =   8
      Top             =   2880
      Width           =   2175
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      LargeChange     =   10
      Left            =   120
      Max             =   255
      TabIndex        =   6
      Top             =   2280
      Width           =   2175
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   10
      Left            =   120
      Max             =   255
      TabIndex        =   4
      Top             =   1680
      Value           =   255
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Text            =   "40"
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blue: 0"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Green: 0"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "% transparent"
      Height          =   195
      Left            =   1920
      TabIndex        =   3
      Top             =   240
      Width           =   990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Red: 255"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   660
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Transparency:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1020
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const WS_EX_TRANSPARENT = &H20&
Private Const LWA_ALPHA = &H2&

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long




Private Sub Form_Load()
Me.Show
DoEvents
Dim NormalWindowStyle As Long
NormalWindowStyle = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
SetWindowLong Me.hwnd, GWL_EXSTYLE, NormalWindowStyle Or WS_EX_LAYERED

SetLayeredWindowAttributes Me.hwnd, 0, 155, LWA_ALPHA

'0 to 255. 255 is 100% visible. 0 is 0% visible
End Sub

Private Sub HScroll1_Change()
Label2.Caption = "Red: " & HScroll1.Value
Label4.Caption = "Green: " & HScroll2.Value
Label5.Caption = "Blue: " & HScroll3.Value
Me.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
End Sub

Private Sub HScroll2_Change()
Label2.Caption = "Red: " & HScroll1.Value
Label4.Caption = "Green: " & HScroll2.Value
Label5.Caption = "Blue: " & HScroll3.Value
Me.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
End Sub

Private Sub HScroll3_Change()
Label2.Caption = "Red: " & HScroll1.Value
Label4.Caption = "Green: " & HScroll2.Value
Label5.Caption = "Blue: " & HScroll3.Value
Me.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
End Sub

Private Sub Text1_Change()
If Val(Text1.Text) <= 100 Then SetLayeredWindowAttributes Me.hwnd, 0, 255 * (1 - (Val(Text1.Text) / 100)), LWA_ALPHA
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If Val(Chr$(KeyAscii)) = 0 And Chr$(KeyAscii) <> "0" And KeyAscii <> 8 And KeyAscii <> 9 Then KeyAscii = 0
End Sub
