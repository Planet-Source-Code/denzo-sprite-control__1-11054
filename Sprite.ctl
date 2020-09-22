VERSION 5.00
Begin VB.UserControl Sprite 
   ClientHeight    =   3300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5415
   ScaleHeight     =   3300
   ScaleWidth      =   5415
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   2055
      Index           =   2
      Left            =   3000
      ScaleHeight     =   133
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   85
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   2055
      Index           =   1
      Left            =   1560
      ScaleHeight     =   133
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   85
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Height          =   1980
      Index           =   0
      Left            =   0
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   0
      Top             =   0
      Width           =   1020
   End
End
Attribute VB_Name = "Sprite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Picture As PictureBox
Public DimX As Long
Public DimY As Long
Private LocalMaskColor As Long

Public Sub PasteSprite(ByVal X As Long, ByVal Y As Long, Optional ByVal SpriteNumX As Integer = 0, Optional ByVal SpriteNumY As Integer = 0)
    Call BitBlt(Picture.hdc, X, Y, DimX, DimY, Picture1(1).hdc, DimX * SpriteNumX, DimY * SpriteNumY, SRCAND)
    Call BitBlt(Picture.hdc, X, Y, DimX, DimY, Picture1(2).hdc, DimX * SpriteNumX, DimY * SpriteNumY, SRCINVERT)
    'Refresh must be always be written to tell the program you want to see all the picture information
    Picture.Refresh
End Sub


Public Sub LoadSprite(ByVal PathFile As String)
    'Use this method to load runtime your sprites
    Set Picture1(0) = LoadPicture(PathFile, , vbLPColor)
    CreateSprite
End Sub

Private Sub CreateSprite()
    Picture1(1).Width = Picture1(0).Width           'Set size for Mask and Sprites
    Picture1(2).Width = Picture1(0).Width
    Picture1(1).Height = Picture1(0).Height
    Picture1(2).Height = Picture1(0).Height
    
    LocalMaskColor = Picture1(0).Point(0, 0)         'Calculate Mask Color From the upper left corner (but you set it by yourself)
    a = Mask(Picture1(0), Picture1(1), LocalMaskColor)
    a = Sprite(Picture1(0), Picture1(2), LocalMaskColor)
    
End Sub

Public Property Let MaskColor(ByVal Color As Long)
    'Use this property to change the default mask color!
    LocalMaskColor = MaskColor
    a = Mask(Picture1(0), Picture1(1), LocalMaskColor)
    a = Sprite(Picture1(0), Picture1(2), LocalMaskColor)

End Property

