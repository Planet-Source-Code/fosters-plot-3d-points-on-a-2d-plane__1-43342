VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10125
   LinkTopic       =   "Form1"
   ScaleHeight     =   8985
   ScaleWidth      =   10125
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Auto"
      Height          =   195
      Index           =   1
      Left            =   2400
      TabIndex        =   13
      Top             =   8760
      Value           =   1  'Checked
      Width           =   675
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Auto"
      Height          =   195
      Index           =   0
      Left            =   2400
      TabIndex        =   12
      Top             =   8400
      Value           =   1  'Checked
      Width           =   675
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Text"
      Height          =   195
      Index           =   2
      Left            =   8520
      TabIndex        =   11
      Top             =   8580
      Width           =   675
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Galaxy"
      Height          =   195
      Index           =   1
      Left            =   7560
      TabIndex        =   10
      Top             =   8580
      Width           =   795
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Cube"
      Height          =   195
      Index           =   0
      Left            =   6720
      TabIndex        =   9
      Top             =   8580
      Value           =   -1  'True
      Width           =   675
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   3420
      Top             =   8460
   End
   Begin VB.TextBox Theta_Ang 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   1500
      TabIndex        =   4
      Text            =   "30"
      Top             =   8280
      Width           =   775
   End
   Begin VB.TextBox Alt_Ang 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   1500
      TabIndex        =   3
      Text            =   "340"
      Top             =   8640
      Width           =   775
   End
   Begin VB.TextBox Size_Factor 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   4200
      TabIndex        =   2
      Text            =   "10000"
      Top             =   8565
      Width           =   885
   End
   Begin VB.TextBox Perspective_Factor 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   5160
      TabIndex        =   1
      Text            =   "10000"
      Top             =   8565
      Width           =   1320
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   8115
      Left            =   120
      ScaleHeight     =   537
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   653
      TabIndex        =   0
      Top             =   120
      Width           =   9855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Dir of Eye (theta)"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   120
      TabIndex        =   8
      Top             =   8340
      Width           =   1365
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Altitude of Eye"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   180
      TabIndex        =   7
      Top             =   8700
      Width           =   1275
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Size Factor"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   3900
      TabIndex        =   6
      Top             =   8340
      Width           =   1185
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Perspective"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   5205
      TabIndex        =   5
      Top             =   8340
      Width           =   1275
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

'my definitions start
Const Pi As Single = 3.14159265358978
Const iMaxStars As Integer = 300
Dim lX(iMaxStars) As Long
Dim lY(iMaxStars) As Long
Dim lZ(iMaxStars) As Long
Dim lMaxDistance As Integer
Dim bCube As Boolean
Dim sngPower As Single
'my definitions end

Private Sub PlotPoint(PicCtl As Control, x As Long, y As Long, z As Long, Theta As Long, Alt As Long, Size As Long, Perspective As Long, lColor As Long)
Dim cX As Single, cY As Single
Dim vX As Single, vY As Single, vZ As Single
Dim pX As Single, pY As Single
Dim Phi As Single
Dim Sin_Theta As Single, Cos_Theta As Single, Sin_Phi   As Single, Cos_Phi   As Single
    
    Phi = 90 - Alt
        
    cX = (PicCtl.Width / Screen.TwipsPerPixelX) / 2
    cY = (PicCtl.Height / Screen.TwipsPerPixelY) / 2
    
    Sin_Theta = Sine(Theta)
    Cos_Theta = Cosine(Theta)
      Sin_Phi = Sine(Phi)
      Cos_Phi = Cosine(Phi)
    
    vX = -x * Sin_Theta _
        + y * Cos_Theta
    
    vY = -x * Cos_Theta * Cos_Phi _
        - y * Sin_Theta * Cos_Phi _
        + z * Sin_Phi
    
    vZ = -x * Cos_Theta * Sin_Phi _
        - y * Sin_Theta * Sin_Phi _
        - z * Cos_Phi + Perspective
    
    pX = cX + Size * vX / vZ
    pY = cY - Size * vY / vZ
    
    SetPixel PicCtl.hdc, pX, pY, lColor

End Sub
Function Sine(Degrees_Arg)
Sine = sIn(Degrees_Arg * Atn(1) / 45)
End Function

Function Cosine(Degrees_Arg)
Cosine = Cos(Degrees_Arg * Atn(1) / 45)
End Function

' 3d stuff above. my stuff below ===========================================================================




Function GimmeX(ByVal aIn As Single, lIn As Integer) As Integer
    GimmeX = sIn(aIn * (Pi / 180)) * lIn

End Function
Function GimmeY(ByVal aIn As Single, lIn As Integer) As Integer
    GimmeY = Cos(aIn * (Pi / 180)) * lIn
End Function



Sub DrawCube(bColor As Boolean)
Dim x As Long

    For x = 0 To 360 Step 12
        PlotPoint Picture1, GimmeX(x, 80), GimmeY(x, 80), 0, CLng(Theta_Ang), CLng(Alt_Ang), CLng(Size_Factor), CLng(Perspective_Factor), IIf(bColor, vbWhite, vbBlack)
        PlotPoint Picture1, GimmeX(x, 100), GimmeY(x, 100), 0, CLng(Theta_Ang), CLng(Alt_Ang), CLng(Size_Factor), CLng(Perspective_Factor), IIf(bColor, vbWhite, vbBlack)

        PlotPoint Picture1, GimmeX(x, 80), 0, GimmeY(x, 80), CLng(Theta_Ang), CLng(Alt_Ang), CLng(Size_Factor), CLng(Perspective_Factor), IIf(bColor, vbWhite, vbBlack)
        PlotPoint Picture1, GimmeX(x, 100), 0, GimmeY(x, 100), CLng(Theta_Ang), CLng(Alt_Ang), CLng(Size_Factor), CLng(Perspective_Factor), IIf(bColor, vbWhite, vbBlack)

        PlotPoint Picture1, 0, GimmeX(x, 80), GimmeY(x, 80), CLng(Theta_Ang), CLng(Alt_Ang), CLng(Size_Factor), CLng(Perspective_Factor), IIf(bColor, vbWhite, vbBlack)
        PlotPoint Picture1, 0, GimmeX(x, 100), GimmeY(x, 100), CLng(Theta_Ang), CLng(Alt_Ang), CLng(Size_Factor), CLng(Perspective_Factor), IIf(bColor, vbWhite, vbBlack)
    
    Next x
    
    PlotPoint Picture1, -100, -100, 100, CLng(Theta_Ang), CLng(Alt_Ang), CLng(Size_Factor), CLng(Perspective_Factor), IIf(bColor, vbWhite, vbBlack)
    PlotPoint Picture1, 100, -100, 100, CLng(Theta_Ang), CLng(Alt_Ang), CLng(Size_Factor), CLng(Perspective_Factor), IIf(bColor, vbWhite, vbBlack)
    PlotPoint Picture1, 100, 100, 100, CLng(Theta_Ang), CLng(Alt_Ang), CLng(Size_Factor), CLng(Perspective_Factor), IIf(bColor, vbWhite, vbBlack)
    PlotPoint Picture1, -100, 100, 100, CLng(Theta_Ang), CLng(Alt_Ang), CLng(Size_Factor), CLng(Perspective_Factor), IIf(bColor, vbWhite, vbBlack)

    PlotPoint Picture1, -100, -100, -100, CLng(Theta_Ang), CLng(Alt_Ang), CLng(Size_Factor), CLng(Perspective_Factor), IIf(bColor, vbWhite, vbBlack)
    PlotPoint Picture1, 100, -100, -100, CLng(Theta_Ang), CLng(Alt_Ang), CLng(Size_Factor), CLng(Perspective_Factor), IIf(bColor, vbWhite, vbBlack)
    PlotPoint Picture1, 100, 100, -100, CLng(Theta_Ang), CLng(Alt_Ang), CLng(Size_Factor), CLng(Perspective_Factor), IIf(bColor, vbWhite, vbBlack)
    PlotPoint Picture1, -100, 100, -100, CLng(Theta_Ang), CLng(Alt_Ang), CLng(Size_Factor), CLng(Perspective_Factor), IIf(bColor, vbWhite, vbBlack)
End Sub




Private Sub Form_Load()
    sngPower = 1.1
    PrepareGalaxy
End Sub

Private Sub Timer1_Timer()
    Picture1.Cls
    'DrawCube False 'if you've got a quick machine
    If Check1(0).Value = vbChecked Then
        Theta_Ang = CInt(Theta_Ang + 3)
        If CInt(Theta_Ang) = 360 Then Theta_Ang = 0
    End If
    If Check1(1).Value = vbChecked Then
        Alt_Ang = CInt(Alt_Ang + 1)
        If CInt(Alt_Ang) = 360 Then Alt_Ang = 0
    End If
    Select Case True
    Case Option1(0)
        DrawCube True
    Case Option1(1)
        Drawgalaxy
    Case Option1(2)
        PlotText
    End Select
End Sub
Sub Drawgalaxy()
Dim x As Integer
    For x = 0 To iMaxStars - 1
        PlotPoint Picture1, lX(x), lY(x), lZ(x), CLng(Theta_Ang), CLng(Alt_Ang), CLng(Size_Factor), CLng(Perspective_Factor), vbWhite
    Next x
End Sub
Sub PrepareGalaxy()
Dim x As Integer
Dim lRND As Long
    lMaxDistance = 10
    For x = 0 To iMaxStars - 1
        lRND = Int(Rnd * lMaxDistance) + 1
        lX(x) = lRND - (lMaxDistance \ 2)
        lRND = Int(Rnd * lMaxDistance) + 1
        lY(x) = lRND - (lMaxDistance \ 2)
        lRND = Int(Rnd * lMaxDistance) + 1
        lZ(x) = lRND - (lMaxDistance \ 2)
        lMaxDistance = lMaxDistance + Int(sngPower ^ 2)
        sngPower = sngPower + 0.001
    Next x
End Sub
Sub PlotText()
Dim ix As Long
Dim lSpace As Long
    lSpace = 5
    ix = 0
    'For ix = -1 To 1 'loop for 3d text
        Plot2D "..    ..  ..  ..  ..  .....", -3, ix, lSpace
        Plot2D "...  ...  ..  .. ..   ..   ", -2, ix, lSpace
        Plot2D "........  ..  ....    ..   ", -1, ix, lSpace
        Plot2D ".. .. ..  ..  ...     .....", 0, ix, lSpace
        Plot2D "..    ..  ..  ....    ..   ", 1, ix, lSpace
        Plot2D "..    ..  ..  .. ..   ..   ", 2, ix, lSpace
        Plot2D "..    ..  ..  ..  ..  .....", 3, ix, lSpace
    'Next ix
End Sub
Sub Plot2D(sIn As String, y As Long, z As Long, lSpaceFactor As Long)
Dim x As Integer
    For x = 1 To Len(sIn)
        If Mid(sIn, x, 1) <> " " Then PlotPoint Picture1, (0 - ((Len(sIn) * lSpaceFactor) / 2)) + (x * lSpaceFactor), y * lSpaceFactor, z * lSpaceFactor, CLng(Theta_Ang), CLng(Alt_Ang), CLng(Size_Factor), CLng(Perspective_Factor), vbWhite
    Next x
End Sub
