VERSION 5.00
Begin VB.UserControl HeaderPicture 
   ClientHeight    =   255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2595
   PropertyPages   =   "HeaderPicture.ctx":0000
   ScaleHeight     =   255
   ScaleWidth      =   2595
   Begin VB.PictureBox picHeader 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   250
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   2595
      TabIndex        =   0
      Top             =   0
      Width           =   2595
   End
End
Attribute VB_Name = "HeaderPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

'By Jim K in April, 2004.
'UserControl Portions of this code Peter Hart, April 2004.
'Use PictureBox as guiding info headers.

'Properties
'----------
' Caption
' Font
' FontSize
' FontColor
' Gradient            (Vertical/Horizontal)
' GradientStart       (Color)
' GradientFinish      (Color)
' GradientFinishStyle (Transparent/Opaque)
' Shape               (Rectangle/Rounded/RoundedTop)

Private sCaption As String
Private SGC As Long
Private EGC As Long

Public Enum Style                 ' Border Shape
    Rectangle = 0
    Rounded = 1
    RoundedTop = 2
End Enum

Public Enum Direction             ' Gradient
    Horizontal = 0
    Vertical = 1
End Enum

Public Enum BackStyle             ' Gradient Finish Style
    Opaque = 0
    Transparent = 1
End Enum

Private Type UserControlProps
    GradientDirection       As Direction
    Shape                   As Style
    GradientBackStyle       As BackStyle
End Type

Private myProps             As UserControlProps  ' cached ctrl properties

Private Declare Function GetSysColor Lib "user32" ( _
                   ByVal nIndex As Long) As Long
                                
Public Property Let Caption(str As String)
sCaption = str
DrawInfoHeader
PropertyChanged "Caption"
End Property
Public Property Get Caption() As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = "Caption"
Caption = sCaption
End Property
 
Public Property Get Font() As Font
Set Font = picHeader.Font
End Property
Public Property Set Font(ByVal NewFont As Font)
Set picHeader.Font = NewFont
DrawInfoHeader
PropertyChanged "Font"
End Property

Public Property Get FontSize() As Integer
FontSize = picHeader.FontSize
End Property
Public Property Let FontSize(i As Integer)
picHeader.FontSize = i
DrawInfoHeader
PropertyChanged "FontSize"
End Property

Public Property Let FontColor(nColor As OLE_COLOR)
picHeader.ForeColor = nColor
DrawInfoHeader
PropertyChanged "FontColor"
End Property
Public Property Get FontColor() As OLE_COLOR
FontColor = picHeader.ForeColor
End Property

Public Property Let Gradient(Styles As Direction)
'Gradient Direction
myProps.GradientDirection = Styles
DrawInfoHeader
PropertyChanged "Gradient"
End Property
Public Property Get Gradient() As Direction
Gradient = myProps.GradientDirection
End Property

Public Property Let GradientStart(nColor As OLE_COLOR)
'Starting Gradient Color
SGC = nColor
DrawInfoHeader
PropertyChanged "GradientStart"
End Property
Public Property Get GradientStart() As OLE_COLOR
GradientStart = SGC
End Property

Public Property Let GradientFinish(nColor As OLE_COLOR)
'Finishing Gradient Color
EGC = nColor
DrawInfoHeader
PropertyChanged "GradientFinish"
End Property
Public Property Get GradientFinish() As OLE_COLOR
GradientFinish = EGC
End Property

Public Property Let GradientFinishStyle(Styles As BackStyle)
'Sets whether or not the finish color is opaque
myProps.GradientBackStyle = Styles
DrawInfoHeader
PropertyChanged "GradientFinishStyle"
End Property
Public Property Get GradientFinishStyle() As BackStyle
GradientFinishStyle = myProps.GradientBackStyle
End Property

Public Property Let Shape(Styles As Style)
myProps.Shape = Styles
DrawInfoHeader
PropertyChanged "Shape"
End Property
Public Property Get Shape() As Style
Shape = myProps.Shape
End Property

Private Sub UserControl_InitProperties()
Caption = "Title Here"
Font = UserControl.Parent.Font
FontSize = 8
FontColor = vbWhite
Gradient = Horizontal
GradientStart = vbBlue
GradientFinish = vbWhite
GradientFinishStyle = Opaque
Shape = Rounded
UserControl.BackColor = Parent.BackColor
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
     Caption = .ReadProperty("Caption", "Title Here")
     Font = .ReadProperty("Font", UserControl.Parent.Font)
     FontSize = .ReadProperty("FontSize", 8)
     FontColor = .ReadProperty("FontColor", vbWhite)
     Gradient = .ReadProperty("Gradient", 0)
     GradientStart = .ReadProperty("GradientStart", vbBlue)
     GradientFinish = .ReadProperty("GradientFinish", vbWhite)
     GradientFinishStyle = .ReadProperty("GradientFinishStyle", 0)
     Shape = .ReadProperty("Shape", 1)
End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
     .WriteProperty "Caption", Caption, "Title Here"
     .WriteProperty "Font", Font
     .WriteProperty "FontSize", FontSize
     .WriteProperty "FontColor", FontColor, vbWhite
     .WriteProperty "Gradient", Gradient, 0
     .WriteProperty "GradientStart", SGC, vbBlue
     .WriteProperty "GradientFinish", GradientFinish, vbWhite
     .WriteProperty "GradientFinishStyle", GradientFinishStyle, 0
     .WriteProperty "Shape", Shape, 1
End With
End Sub

Private Sub UserControl_Resize()
picHeader.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
DrawInfoHeader
End Sub

Private Function ConvertColor(tColor As Long) As Long
'Converts VB color constants to real color values
If tColor < 0 Then
   ConvertColor = GetSysColor(tColor And &HFF&)
 Else
   ConvertColor = tColor
End If
End Function

Public Sub DrawInfoHeader()
Dim x1 As Integer
Dim R, G, b, dr, dg, db As Single
Dim y1 As Integer
Dim r1 As Integer
Dim g1 As Integer
Dim b1 As Integer
Dim r2 As Integer
Dim g2 As Integer
Dim b2 As Integer
Dim r3 As Integer
Dim g3 As Integer
Dim b3 As Integer
Dim BC As Long
Dim i As Integer
Dim cnrInt As Integer
Dim EGCCopy As Long

BC = ConvertColor(Parent.BackColor)

If myProps.GradientBackStyle = Transparent Then
   EGCCopy = BC
 Else
   EGCCopy = EGC
End If

DrawGrad SGC, r1, g1, b1
DrawGrad EGCCopy, r2, g2, b2
DrawGrad BC, r3, g3, b3

With picHeader
   
   Select Case myProps.GradientDirection
   
     Case Horizontal
        'Gradient Horizontal
        dr = (r2 - r1) / (.ScaleWidth / 15)
        dg = (g2 - g1) / (.ScaleWidth / 15)
        db = (b2 - b1) / (.ScaleWidth / 15)
        R = r1
        G = g1
        b = b1
        For x1 = 0 To .ScaleWidth Step 15
            picHeader.Line (x1, 0)-(x1, .ScaleHeight), RGB(R, G, b) + &H2000000
            R = R + dr
            G = G + dg
            b = b + db
        Next x1
     
     Case Vertical
        'Gradient Vertical
        dr = (r2 - r1) / (.ScaleHeight / 15)
        dg = (g2 - g1) / (.ScaleHeight / 15)
        db = (b2 - b1) / (.ScaleHeight / 15)
        R = r1
        G = g1
        b = b1
        For y1 = 0 To .ScaleHeight Step 15
            picHeader.Line (0, y1)-(.ScaleWidth, y1), RGB(R, G, b) + &H2000000
            R = R + dr
            G = G + dg
            b = b + db
        Next y1
   End Select
   
    'Top Corners
    If myProps.Shape = Rounded Or myProps.Shape = RoundedTop Then
       'Left
       cnrInt = 20
       i = 15
       For x1 = 0 To cnrInt
           picHeader.Line (x1, 0)-(x1, i), RGB(r3, g3, b3) + &H2000000
           i = i - 1
       Next x1
       'Right
       i = 0
       For x1 = .ScaleWidth - cnrInt To .ScaleWidth
           picHeader.Line (x1, 0)-(x1, i), RGB(r3, g3, b3) + &H2000000
           i = i + 1
       Next x1
    End If
    
    'Bottom Corners
    If myProps.Shape = Rounded Then
       'Right
       i = 0
       For x1 = .ScaleWidth - cnrInt To .ScaleWidth
           picHeader.Line (x1, .ScaleHeight - i)-(x1, .ScaleHeight), _
                                               RGB(r3, g3, b3) + &H2000000
           i = i + 1
       Next x1
       'Left
       i = 15
       For x1 = 0 To cnrInt
           picHeader.Line (x1, .ScaleHeight - i)-(x1, .ScaleHeight), _
                                               RGB(r3, g3, b3) + &H2000000
           i = i - 1
       Next x1
    End If
    
   'Caption
   .FontBold = True
   .CurrentX = 60
   .CurrentY = 35
    picHeader.Print sCaption
   .Refresh
End With
End Sub
    
Function DrawGrad(Color As Long, R As Integer, G As Integer, b As Integer)
R = Color Mod 256 'red
G = (Color \ 256) Mod 256 'green
b = Color \ 65536 'blue
End Function
