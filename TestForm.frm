VERSION 5.00
Begin VB.Form TestForm 
   AutoRedraw      =   -1  'True
   Caption         =   "Reflection Demo"
   ClientHeight    =   2550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7305
   LinkTopic       =   "Form1"
   ScaleHeight     =   170
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   487
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
   Begin VB.Timer tmrTextAnimate 
      Interval        =   500
      Left            =   600
      Top             =   600
   End
   Begin VB.Timer tmrAnimate 
      Interval        =   25
      Left            =   120
      Top             =   600
   End
   Begin VB.PictureBox picCurvy 
      BorderStyle     =   0  'None
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   1455
      Left            =   5760
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   97
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
   End
   Begin VB.PictureBox picSpiky 
      BorderStyle     =   0  'None
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   1455
      Left            =   120
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   97
      TabIndex        =   2
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox txtDemo 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1470
      HideSelection   =   0   'False
      Left            =   1680
      TabIndex        =   1
      Text            =   "Reflect"
      Top             =   1080
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "Just some demonstration controls here to show that the reflection works whatever is on the form, and  is updated continuously."
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   7215
   End
End
Attribute VB_Name = "TestForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' The content of this file is not essential to the functioning of the Reflection, it's just an example.
' The only important parts are the Form_Load() and the Form_Unload() events

Private Declare Function Polygon Lib "gdi32.dll" (ByVal hdc As Long, ByRef lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function PolyBezier Lib "gdi32.dll" (ByVal hdc As Long, ByRef lppt As POINTAPI, ByVal cPoints As Long) As Long
Private Declare Function BeginPath Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function EndPath Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function FillPath Lib "gdi32.dll" (ByVal hdc As Long) As Long

Private Type POINTAPI
    x As Long
    y As Long
End Type

Dim pts(0 To 9) As POINTAPI
Dim polypts1(0 To 15) As POINTAPI
Dim polypts2(0 To 15) As POINTAPI

' =========== Begin essential code ==============
Private Sub Form_Load()
    On Error GoTo NO_REFLECTION
    Reflection.Attach Me.hwnd
    Exit Sub
    
NO_REFLECTION:
    MsgBox "Could not initialize the reflection." & vbCrLf & Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Reflection.Detach
End Sub
' =========== End essential code ==============

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub Form_Initialize()
    ' Sets up the polygon points
    tmrAnimate_Timer
End Sub

Private Sub picSpiky_Paint()
    picSpiky.Cls
    
    BeginPath picSpiky.hdc
    PolyBezier picSpiky.hdc, polypts1(0), 16
    EndPath picSpiky.hdc
    
    FillPath picSpiky.hdc
    PolyBezier picSpiky.hdc, polypts1(0), 16
End Sub

Private Sub picCurvy_Paint()
    picCurvy.Cls
    
    BeginPath picCurvy.hdc
    PolyBezier picCurvy.hdc, polypts2(0), 16
    EndPath picCurvy.hdc
    
    FillPath picCurvy.hdc
    PolyBezier picCurvy.hdc, polypts2(0), 16
End Sub

Private Sub tmrAnimate_Timer()
    Static phase As Double
    Dim i As Long
    Dim a As Double, r As Double, da As Double
    a = phase
    da = (2 * Atn(1) * 4) / 10#
    For i = 0 To 9
        If i Mod 2 = 0 Then
            r = 40# + Sin(phase) * 10
        Else
            r = 20# + Sin(phase * 2) * 10
        End If
        ' Geometric points
        pts(i).x = 47 + Int(Cos(a) * r + 0.5)
        pts(i).y = 47 - Int(Sin(a) * r + 0.5)
        '
        a = a + da
    Next
    ' Points for curved star 1 "spiky"
    polypts1(0) = pts(0)
    polypts1(1) = pts(1)
    polypts1(2) = pts(1)
    polypts1(3) = pts(2)
    polypts1(4) = pts(3)
    polypts1(5) = pts(3)
    polypts1(6) = pts(4)
    polypts1(7) = pts(5)
    polypts1(8) = pts(5)
    polypts1(9) = pts(6)
    polypts1(10) = pts(7)
    polypts1(11) = pts(7)
    polypts1(12) = pts(8)
    polypts1(13) = pts(9)
    polypts1(14) = pts(9)
    polypts1(15) = pts(0)
    ' Points for curved star 2 "curvy"
    polypts2(0) = pts(1)
    polypts2(1) = pts(2)
    polypts2(2) = pts(2)
    polypts2(3) = pts(3)
    polypts2(4) = pts(4)
    polypts2(5) = pts(4)
    polypts2(6) = pts(5)
    polypts2(7) = pts(6)
    polypts2(8) = pts(6)
    polypts2(9) = pts(7)
    polypts2(10) = pts(8)
    polypts2(11) = pts(8)
    polypts2(12) = pts(9)
    polypts2(13) = pts(0)
    polypts2(14) = pts(0)
    polypts2(15) = pts(1)
    ' Refresh both
    picSpiky_Paint
    picCurvy_Paint
    
    
    phase = phase + 0.075
End Sub

Private Sub tmrTextAnimate_Timer()
    Static Count As Long
    
    Debug.Print "Tick"
    
    txtDemo.SelStart = Count Mod (Len(txtDemo.Text) - 1)
    txtDemo.SelLength = 1
    
    Count = Count + 1
End Sub

Private Sub txtDemo_GotFocus()
    tmrTextAnimate.Enabled = False
End Sub

Private Sub txtDemo_LostFocus()
    tmrTextAnimate.Enabled = True
End Sub
