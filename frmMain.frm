VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "API Gaming Techniques Vol. 1 (Fast Sprite Manipulation)"
   ClientHeight    =   9270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   618
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   838
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Select BG Color"
      Height          =   735
      Left            =   4560
      TabIndex        =   6
      Top             =   1080
      Width           =   735
   End
   Begin VB.PictureBox picInfoA 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5400
      ScaleHeight     =   1335
      ScaleWidth      =   4335
      TabIndex        =   4
      Top             =   120
      Width           =   4335
   End
   Begin VB.PictureBox picInfoC 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      ScaleHeight     =   1335
      ScaleWidth      =   4335
      TabIndex        =   3
      Top             =   120
      Width           =   4335
   End
   Begin VB.HScrollBar sScaleArbiter 
      Height          =   255
      LargeChange     =   10
      Left            =   5400
      Max             =   200
      TabIndex        =   2
      Top             =   1560
      Value           =   100
      Width           =   4335
   End
   Begin VB.HScrollBar sScaleCarrier 
      Height          =   255
      LargeChange     =   10
      Left            =   120
      Max             =   200
      TabIndex        =   1
      Top             =   1560
      Value           =   100
      Width           =   4335
   End
   Begin VB.PictureBox picField 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   7260
      Left            =   120
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   0
      Top             =   1920
      Width           =   9660
   End
   Begin VB.Label lbDesc 
      Caption         =   $"frmMain.frx":0000
      Height          =   5775
      Left            =   9840
      TabIndex        =   5
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Force variables to be defined

'Declarations
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function PlgBlt Lib "gdi32" (ByVal hdcDest As Long, lpPoint As POINTAPI, ByVal hdcSrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hbmMask As Long, ByVal xMask As Long, ByVal yMask As Long) As Long
Private Declare Function CHOOSECOLOR Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long

'Types
Private Type POINTAPI
   X As Long
   Y As Long
End Type

Private Type CHOOSECOLOR 'Used for selecting the backcolor in this example
  lStructSize As Long
  hwndOwner As Long
  hInstance As Long
  rgbResult As Long
  lpCustColors As String
  flags As Long
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type

Dim Carrier As Long                               'DC for Carrier sprite strip
Dim Arbiter As Long                               'DC for Arbiter sprite strip
Dim CarrierPos As POINTAPI                        'Center point for Carrier
Dim ArbiterPos As POINTAPI                        'Center point for Arbiter
Dim cHeight As Single, cWidth As Single           'Dimensions of Carrier after scaling
Dim aHeight As Single, aWidth As Single           'Dimensions of Arbiter after scaling
Dim oldX As Single, oldY As Single                'Previous mouse positions


Private Const Carrier_Height = 97                 'Default Carrier height
Private Const Carrier_Width = 124                 'Default Carrier width
Private Const Arbiter_Height = 63                 'Default Arbiter height
Private Const Arbiter_Width = 74                  'Default Arbiter width
Private Const PI = 3.14159265238                  'PI

'Loads an image on your hard drive to DC memory
Public Function LoadPic(DestID As Long, Path As String) As Long
  DestID = CreateCompatibleDC(GetDC(0))
  LoadPic = SelectObject(DestID, LoadPicture(Path))
End Function

'Unloads an image stored in DC memory to prevent memory leaks
Public Function UnloadPic(SrcID As Long) As Long
  UnloadPic = DeleteDC(SrcID)
End Function

'Distance Formula - Gets the distance between two points...(x1,y1) & (x2,y2)
Public Function GetDistance(ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single) As Single
  GetDistance = Sqr(((X2 - X1) ^ 2) + ((Y2 - Y1) ^ 2))
End Function

'Returns the angle between 3 points
Public Function GetAngle(ByVal Ax As Single, ByVal Ay As Single, ByVal Bx As Single, ByVal By As Single, ByVal Cx As Single, ByVal Cy As Single, Optional ByVal FullRotate As Boolean = True) As Single
  Dim PointGrid(1) As Single
  Dim AngleOut As Single
  
  'First we have to get the dot product of our 3 points
  PointGrid(0) = (Ax - Bx) * (Cx - Bx) + (Ay - By) * (Cy - By) 'Dot Product
  'Now we get the cross product of our 3 points
  PointGrid(1) = (Ax - Bx) * (Cy - By) - (Ay - By) * (Cx - Bx) 'Cross Product
  
  'Next we need to get the ArcTangent of our points
  If Abs(PointGrid(0)) < 0.0001 Then
    AngleOut = PI / 2
  Else
    AngleOut = Abs(Atn(PointGrid(1) / PointGrid(0)))
  End If
  
  'If our angle is between 0 and 180
  If PointGrid(0) < 0 Then
    AngleOut = PI - AngleOut
  End If
  
  'If our angle if between 180 and 360
  'Returns a negative angle if above 180....e.g. 300 degrees = -60 degrees
  If PointGrid(1) < 0 Then
    AngleOut = -AngleOut
  End If
  
  'Get the whole numbers of the angle
  AngleOut = Format$(AngleOut / PI * 180, "0")
  
  'If you want the full 360 degree without using the negative angles
  'we have to subtract the negative angle from 360
  If FullRotate And AngleOut < 0 Then
    AngleOut = 360 + AngleOut
  End If
  
  GetAngle = AngleOut
End Function

'For setting the background color of our play field
Private Sub Command1_Click()
  Dim NewColor As Long
  
  NewColor = ShowColor 'calls the ShowColor() function
  
  If NewColor <> -1 Then 'if a color was chosen
    picField.BackColor = NewColor 'set backcolor
    Call picField_MouseMove(0, 0, oldX, oldY) 'Refresh our sprites
  End If
End Sub

'Opens a Choose Color dialog box
Private Function ShowColor() As Long
    Dim cc As CHOOSECOLOR
    Dim Custcolor(16) As Long
    Dim lReturn As Long

    'set the structure size
    cc.lStructSize = Len(cc)
    'Set the owner
    cc.hwndOwner = Me.hwnd
    'set the application's instance
    cc.hInstance = App.hInstance
    'set the custom colors (converted to Unicode)
    cc.lpCustColors = StrConv(CustomColors, vbUnicode)
    'no extra flags
    cc.flags = 0

    'Show the 'Select Color'-dialog
    If CHOOSECOLOR(cc) <> 0 Then
        ShowColor = cc.rgbResult
        CustomColors = StrConv(cc.lpCustColors, vbFromUnicode)
    Else
        ShowColor = -1
    End If
End Function

Private Sub Form_Load()
  lbDesc.Caption = Replace(lbDesc.Caption, "#", vbCrLf)
  LoadPic Carrier, App.Path & "\Protoss - Carrier.bmp" 'Load the Carrier sprite strip into memory
  LoadPic Arbiter, App.Path & "\Protoss - Arbiter.bmp" 'Load the Arbiter sprite strip into memory
  CarrierPos.X = (picField.ScaleWidth / 3) 'Sets the Carrier x position to 1/3rd into the workspace
  CarrierPos.Y = (picField.ScaleHeight / 2) - (Carrier_Height / 2) 'Sets the Carrier y position to the center
  ArbiterPos.X = ((picField.ScaleWidth / 3) * 2) 'Set the Arbiter x position to 2/3rds into the workspace
  ArbiterPos.Y = (picField.ScaleHeight / 2) - (Arbiter_Height / 2) 'Sets the Arbiter y position to the center
  cWidth = Carrier_Width: cHeight = Carrier_Height 'Sets Carrier to default width/height
  aWidth = Arbiter_Width: aHeight = Arbiter_Height 'Sets Arbiter to default width/height
  Call picField_MouseMove(0, 0, 0, 0) 'Draw our sprites
End Sub

'unloads our sprite strips from memory
Private Sub Form_Terminate()
  UnloadPic Carrier
  UnloadPic Arbiter
End Sub

'same as above
Private Sub Form_Unload(Cancel As Integer)
  UnloadPic Carrier
  UnloadPic Arbiter
End Sub

'This is where our main coding is done...You can do this in a separate sub if needed...This is just an example
Private Sub picField_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim cAngle As Single, cOffset As Single
  Dim aAngle As Single, aOffset As Single
  Dim sTime As Long, eTime As Long
  
  oldX = X: oldY = Y 'set old x,y positions
  
  'determine the angle of the carrier ship from it's center to the mouse cursor
  cAngle = GetAngle(CarrierPos.X + 1, CarrierPos.Y, CarrierPos.X, CarrierPos.Y, X, Y) 'gets the angle
  cOffset = Round(cAngle / 11.25) '360 degrees DIVIDED BY 32 Frames in our sprite strip = 11.25 degrees in between frames
  If cOffset = 32 Then cOffset = 0 'Error correction
  
  'same as above except for the arbiter ship
  aAngle = GetAngle(ArbiterPos.X + 1, ArbiterPos.Y, ArbiterPos.X, ArbiterPos.Y, X, Y)
  aOffset = Round(aAngle / 11.25)
  If aOffset = 32 Then aOffset = 0
  
  picField.Cls 'clear our workspace
  
  'Transparent Blit both the ships at their positions
  'We take from the sprite strip loaded into memory starting at the position determined above.
  'e.g. offset * default_width
  TransparentBlt picField.hdc, CarrierPos.X - (cWidth / 2), CarrierPos.Y - (cHeight / 2), cWidth, cHeight, Carrier, (cOffset * Carrier_Width), 0, Carrier_Width, Carrier_Height, vbCyan
  TransparentBlt picField.hdc, ArbiterPos.X - (aWidth / 2), ArbiterPos.Y - (aHeight / 2), aWidth, aHeight, Arbiter, (aOffset * Arbiter_Width), 0, Arbiter_Width, Arbiter_Height, vbCyan
  
  picField.Refresh 'refresh our workspace
  
  'This is just debugging information for the purpose of this tutorial
  picInfoC.Cls
  picInfoA.Cls
  
  picInfoC.FontBold = True
  picInfoC.Print "Carrier Information"
  picInfoC.FontBold = False
  picInfoC.Print "Mouse Angle......" & cAngle
  picInfoC.Print "Mouse Distance..." & Round(GetDistance(CarrierPos.X, CarrierPos.Y, X, Y))
  picInfoC.Print "Sprite Frame....." & cOffset + 1
  picInfoC.Print "Width............" & cWidth
  picInfoC.Print "Height..........." & cHeight
  picInfoC.Print "Position (x,y)..." & CarrierPos.X & "," & CarrierPos.Y
  picInfoC.Print "Scale............" & sScaleCarrier.Value & "%"
  
  picInfoA.FontBold = True
  picInfoA.Print "Arbiter Information"
  picInfoA.FontBold = False
  picInfoA.Print "Mouse Angle......" & aAngle
  picInfoA.Print "Mouse Distance..." & Round(GetDistance(ArbiterPos.X, ArbiterPos.Y, X, Y))
  picInfoA.Print "Sprite Frame....." & aOffset + 1
  picInfoA.Print "Width............" & aWidth
  picInfoA.Print "Height..........." & aHeight
  picInfoA.Print "Position (x,y)..." & ArbiterPos.X & "," & ArbiterPos.Y
  picInfoA.Print "Scale............" & sScaleArbiter.Value & "%"
  
  picInfoC.Refresh
  picInfoA.Refresh
End Sub

'Sets the (x,y) coordinates of our 2 ships
Private Sub picField_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbLeftButton Then
    CarrierPos.X = X: CarrierPos.Y = Y 'if the left button is clicked move the Carrier to that position
  ElseIf Button = vbRightButton Then
    ArbiterPos.X = X: ArbiterPos.Y = Y 'if the right button is clicked move the Arbiter to that position
  End If
End Sub

'This is to set the scale our ships are painted at
'(Ship's Width) * (ScalePercent / 100)
Private Sub sScaleCarrier_Change()
  cWidth = Round(Carrier_Width * (sScaleCarrier.Value / 100))
  cHeight = Round(Carrier_Height * (sScaleCarrier.Value / 100))
  Call picField_MouseMove(0, 0, oldX, oldY)
End Sub

Private Sub sScaleCarrier_Scroll()
  cWidth = Round(Carrier_Width * (sScaleCarrier.Value / 100))
  cHeight = Round(Carrier_Height * (sScaleCarrier.Value / 100))
  Call picField_MouseMove(0, 0, oldX, oldY)
End Sub

Private Sub sScaleArbiter_Change()
  aWidth = Round(Arbiter_Width * (sScaleArbiter.Value / 100))
  aHeight = Round(Arbiter_Height * (sScaleArbiter.Value / 100))
  Call picField_MouseMove(0, 0, oldX, oldY)
End Sub

Private Sub sScaleArbiter_Scroll()
  aWidth = Round(Arbiter_Width * (sScaleArbiter.Value / 100))
  aHeight = Round(Arbiter_Height * (sScaleArbiter.Value / 100))
  Call picField_MouseMove(0, 0, oldX, oldY)
End Sub
