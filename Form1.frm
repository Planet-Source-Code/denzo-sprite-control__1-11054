VERSION 5.00
Object = "*\ASpriteCtl.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6900
   LinkTopic       =   "Form1"
   ScaleHeight     =   4035
   ScaleWidth      =   6900
   StartUpPosition =   3  'Windows Default
   Begin SpriteCtl.Sprite Sprite1 
      Height          =   2055
      Left            =   4320
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
      _extentx        =   1931
      _extenty        =   3625
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   3615
      Left            =   120
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   237
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   261
      TabIndex        =   0
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Frame As Integer          'Tells if the frame is 1 or 2 in the current direction
Dim XPos As Integer, YPos As Integer    'Position holders for the position of the Wolf Man on the Form
Dim LeftX As Integer, TopY As Integer   'Tells where the frame of the animation is
Const Speed = 5  'Change this number to how fast you want him to move





Private Sub Form_Load()
    'Sprite Control Initialization!
    Set Sprite1.Picture = Picture1                  'Specify sprite back ground
    Sprite1.LoadSprite (App.Path + "\WolfMan.bmp")  'Load Sprite
    Sprite1.DimX = 32                               'Define image dimension
    Sprite1.DimY = 32
    
    
    'Game initialization!
    XPos = Picture1.ScaleWidth / 2                  'Startup position on the form, is Center X
    YPos = Picture1.ScaleHeight / 2                 'Startup position on the form, is Center Y
    LeftX = 0                                       'TopLeft X Coord of the Wolf Man Frame 1
    TopY = 0                                        'TopLeft Y Coord of the Wolf Man Frame 1
    moveit
    
End Sub


Sub moveit()
    Picture1.Cls
    Sprite1.PasteSprite XPos, YPos, LeftX, TopY     'Specify where to paste and the index of image to paste

End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyLeft Then    'The user pressed the Left arrow key
    If Frame = 0 Then
        Frame = 1
        LeftX = 0
        TopY = 3
        XPos = XPos - Speed
        moveit
    ElseIf Frame = 1 Then
        Frame = 0
        LeftX = 1
        TopY = 3
        XPos = XPos - Speed
        moveit
    End If
ElseIf KeyCode = vbKeyRight Then    'The user pressed the Right arrow key
    If Frame = 1 Then
        Frame = 0
        LeftX = 0
        TopY = 1
        XPos = XPos + Speed
        moveit
    ElseIf Frame = 0 Then
        Frame = 1
        LeftX = 1
        TopY = 1
        XPos = XPos + Speed
        moveit
    End If
ElseIf KeyCode = vbKeyUp Then    'The user pressed the Up arrow key
        If Frame = 1 Then
        Frame = 0
        LeftX = 0
        TopY = 0
        YPos = YPos - Speed
        moveit
    ElseIf Frame = 0 Then
        Frame = 1
        LeftX = 1
        TopY = 0
        YPos = YPos - Speed
        moveit
    End If
ElseIf KeyCode = vbKeyDown Then    'The user pressed the Down arrow key
    If Frame = 0 Then
        Frame = 1
        LeftX = 0
        TopY = 2
        YPos = YPos + Speed
        moveit
    ElseIf Frame = 1 Then
        Frame = 0
        LeftX = 1
        TopY = 2
        YPos = YPos + Speed
        moveit
    End If
End If
End Sub
