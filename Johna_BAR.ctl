VERSION 5.00
Begin VB.UserControl Johna_BAR 
   ClientHeight    =   1455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   PropertyPages   =   "Johna_BAR.ctx":0000
   ScaleHeight     =   1455
   ScaleWidth      =   4800
   ToolboxBitmap   =   "Johna_BAR.ctx":0014
   Begin VB.PictureBox PicBAR 
      Height          =   1215
      Left            =   120
      ScaleHeight     =   1155
      ScaleWidth      =   4380
      TabIndex        =   0
      Top             =   120
      Width           =   4440
   End
End
Attribute VB_Name = "Johna_BAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================
' this class was created 10/10/2001 version 1.00
'  by johna
'  it handles a pictureBox as a progressBar
'  it can be used freely and can be
'  redistibuted only as this
'  Any modifications or Extra GFX routines added
'  write me on Johna.pop@caramail.com
'
' If someone plan to compile it into an ActiveX Dll
' send me the new version Created at Johna.pop@caramail.com
' anyone can aDD special FX boutton
' by using The 3d flags
'
'    11/10/2001  version 1.01
'      Add a buffer for back drawing before refresh
'        the pictureBox
'      -Too fluide drawing
'      -New Effects
'
'
'   15/10/2001  version 1.03
'       add Min Max,Value property
'       Add text percentage
'
'=================================================



'Useful APIz
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long




'Pour les frame et bouttons
Private Declare Function Drawjohna_EDGE Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal johna_EDGE As Long, ByVal grfFlags As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long



Public Enum johna_EDGE_3D
    SUNKEN_INNER = &H8
    SUNKEN_OUTER = &H2
    RAISED_INNER = &H4
    RAISED_OUTER = &H1
    EDGE_BUMP = (RAISED_OUTER Or SUNKEN_INNER)
    EDGE_ETCHED = (SUNKEN_OUTER Or RAISED_INNER)
    EDGE_RAISED = (RAISED_OUTER Or RAISED_INNER)
    EDGE_SUNKEN = (SUNKEN_INNER Or SUNKEN_OUTER)
End Enum

Public Enum Flags_3D

    CB_ADJUST = &H2000
    CB_BOTTOM = &H8
    CB_DIAGONAL = &H10
    CB_FLAT = &H4000
    CB_LEFT = &H1
    CB_MIDDLE = &H800
    CB_MONO = &H8000
    CB_RIGHT = &H4
    CB_SOFT = &H1000
    CB_TOP = &H2
    CB_BOTTOMLEFT = (CB_BOTTOM Or CB_LEFT)
    CB_BOTTOMRIGHT = (CB_BOTTOM Or CB_RIGHT)
    CB_DIAGONAL_ENDBOTTOMLEFT = (CB_DIAGONAL Or CB_BOTTOM Or CB_LEFT)
    CB_DIAGONAL_ENDBOTTOMRIGHT = (CB_DIAGONAL Or CB_BOTTOM Or CB_RIGHT)
    CB_DIAGONAL_ENDTOPLEFT = (CB_DIAGONAL Or CB_TOP Or CB_LEFT)
    CB_DIAGONAL_ENDTOPRIGHT = (CB_DIAGONAL Or CB_TOP Or CB_RIGHT)
    CB_RECT = (CB_LEFT Or CB_TOP Or CB_RIGHT Or CB_BOTTOM)
    CB_TOPLEFT = (CB_TOP Or CB_LEFT)
    CB_TOPRIGHT = (CB_TOP Or CB_RIGHT)

End Enum


Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type


Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, lpRect As RECT) As Long



'for text caption marking the percentage
Private Declare Function SetSysColors Lib "user32" (ByVal nChanges As Long, lpSysColor As Long, lpColorValues As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long


Enum Esys_COLOR_const
 COLOR_SCROLLBAR = 0 'The Scrollbar colour
 COLOR_BACKGROUND = 1 'Colour of the background with no wallpaper
 COLOR_ACTIVECAPTION = 2 'Caption of Active Window
 COLOR_INACTIVECAPTION = 3 'Caption of Inactive window
 COLOR_MENU = 4 'Menu
 COLOR_WINDOW = 5 'Windows background
 COLOR_WINDOWFRAME = 6 'Window frame
 COLOR_MENUTEXT = 7 'Window Text
 COLOR_WINDOWTEXT = 8 '3D dark shadow (Win95)
 COLOR_CAPTIONTEXT = 9 'Text in window caption
 COLOR_ACTIVEBORDER = 10 'Border of active window
 COLOR_INACTIVEBORDER = 11 'Border of inactive window
 COLOR_APPWORKSPACE = 12 'Background of MDI desktop
 COLOR_HIGHLIGHT = 13 'Selected item background
 COLOR_HIGHLIGHTTEXT = 14 'Selected menu item
 COLOR_BTNFACE = 15 'Button
 COLOR_BTNSHADOW = 16 '3D shading of button
 COLOR_GRAYTEXT = 17 'Grey text, of zero if dithering is used.
 COLOR_BTNTEXT = 18 'Button text
 COLOR_INACTIVECAPTIONTEXT = 19 'Text of inactive window
 COLOR_BTNHIGHLIGHT = 20 '3D highlight of button
 COLOR_2NDACTIVECAPTION = 27 'Win98 only: 2nd active window color
 COLOR_2NDINACTIVECAPTION = 28 'Win98 only: 2nd inactive window color

End Enum














Private Type Tfx_BAR
  BOX As RECT
  Width As Integer
  Height As Integer
  COLOR As Long
  percent As Single
  lp_PICTURE As PictureBox 'une sorte de pointeur vers Un controle picture box
  DRAW_TYPE As johna_EDGE_3D
  MIN As Integer
  MAX As Integer
  Step As Integer
  Value As Integer
  
  
End Type


Public Enum Johna_Bar_Orientation
    Johna_Horizontal
    Johna_Vertical
End Enum

Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long




'For HDC for avoiding flashing screen
Private Const BI_RGB = 0&
Private Const DIB_RGB_COLORS = 0 '  color table in RGBs
Private Type BITMAPINFOHEADER '40 bytes
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type
Private Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type
Private Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors As RGBQUAD
End Type
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal HBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Dim BarDC As Long

Dim BITmap As BITMAPINFO 'For holding the BITMAP
Dim HBitmap As Long




Private TBAR As Tfx_BAR

Private ORIENTATIO As Johna_Bar_Orientation



Private Sub INIT_BAR()
Set TBAR.lp_PICTURE = PicBAR
TBAR.COLOR = RGB(110, 80, 185)
GetWindowRect TBAR.lp_PICTURE.hwnd, TBAR.BOX

With TBAR.BOX
 TBAR.Width = .Right - .Left
 TBAR.Height = .Bottom - .Top
 
End With


TBAR.DRAW_TYPE = EDGE_RAISED
Me.Orientation = Johna_Horizontal
TBAR.MIN = 0
TBAR.MAX = 100
CreateDC

End Sub


Public Sub About()
   MsgBox "JohnaProgress bar ActiveX" + Chr(10) + "John Company .Inc (c) 2001-2002", vbInformation, "About JohnaProgressBar Controle"
   
   
End Sub



Sub Refresh_BAR()

GetWindowRect TBAR.lp_PICTURE.hwnd, TBAR.BOX

With TBAR.BOX
 TBAR.Width = (.Right - .Left)
 TBAR.Height = .Bottom - .Top
 
End With
CreateDC
 


UPdate_BAR TBAR.percent
End Sub
Private Sub CreateDC()

With BITmap.bmiHeader
        .biBitCount = 24
        .biCompression = BI_RGB
        .biPlanes = 1
        .biSize = Len(BITmap.bmiHeader)
        .biWidth = TBAR.Width
        .biHeight = TBAR.Height
    End With
    
    'For primary Dc
    
    BarDC = CreateCompatibleDC(0)
    
    HBitmap = CreateDIBSection(BarDC, BITmap, DIB_RGB_COLORS, ByVal 0&, ByVal 0&, ByVal 0&)
    
    SelectObject BarDC, HBitmap
    
    
   
End Sub




Private Sub UPdate_BAR(percent As Single)

Dim L As Integer, T As Integer, b As Integer, R As Integer

Dim REST
REST = 1
If TBAR.lp_PICTURE.Appearance > 0 Then REST = 4
If Me.Orientation = Johna_Horizontal Then
    L = 0
    R = Int((TBAR.Width - REST) * percent / 100)
    b = TBAR.Height - REST
    T = 0
ElseIf Me.Orientation = Johna_Vertical Then
    L = 0
    R = TBAR.Width - REST
    b = TBAR.Height - REST
    T = ((1 - percent / 100) * TBAR.Height)
End If


Dim Rec As RECT
Dim BRUSH As Long
'Dim LB As LOGBRUSH
'crée un pinceau
BRUSH = CreateSolidBrush(TBAR.COLOR)


If percent < TBAR.percent Or TBAR.percent <= 1 Then _
 TBAR.lp_PICTURE.Cls   'BitBlt GetDC(TBAR.lp_PICTURE.hwnd), 0, 0, TBAR.Width, TBAR.Height, ClearDC, 0, 0, vbSrcCopy

'TBAR.lp_PICTURE.AutoRedraw = True


  
  TBAR.percent = percent
  
    BitBlt BarDC, 0, 0, TBAR.Width, TBAR.Height, GetDC(TBAR.lp_PICTURE.hwnd), 0, 0, vbSrcCopy

If Me.Orientation = Johna_Horizontal Then
  Call SetRect(Rec, L, T, R, b)

  'ajoute la couleur
  FillRect BarDC, Rec, BRUSH
  'éfface le pinceau
   DeleteObject BRUSH
  'déssine le bord 3d
  Draw_3DEX BarDC, L, T, R, b, TBAR.DRAW_TYPE, CB_RECT
  Draw_3DEX BarDC, L + 1, T + 1, R - 1, b - 1, TBAR.DRAW_TYPE, CB_BOTTOMRIGHT
  Draw_3DEX BarDC, L + 2, T + 2, R - 2, b - 2, EDGE_ETCHED, CB_RECT
  
  BitBlt GetDC(TBAR.lp_PICTURE.hwnd), 0, 0, R - L, b - T, BarDC, 0, 0, vbSrcCopy

ElseIf Me.Orientation = Johna_Vertical Then
 
 'ajoute la couleur
 Call SetRect(Rec, L, T, R, b)
 
 
  FillRect BarDC, Rec, BRUSH
  'éfface le pinceau
  DeleteObject BRUSH
  
  'déssine le bord 3d
  Draw_3DEX BarDC, L, T, R, b, TBAR.DRAW_TYPE, CB_RECT
  Draw_3DEX BarDC, L + 1, T + 1, R - 1, b - 1, TBAR.DRAW_TYPE, CB_BOTTOMRIGHT
  Draw_3DEX BarDC, L + 2, T + 2, R - 2, b - 2, EDGE_ETCHED, CB_RECT
  
  BitBlt GetDC(TBAR.lp_PICTURE.hwnd), 0, 0, R - L, b, BarDC, 0, 0, vbSrcCopy

End If
  
  
  
 



'TBAR.lp_PICTURE.Refresh
'TBAR.lp_PICTURE.AutoRedraw = False



End Sub


Private Function KOLOR(COLa As Esys_COLOR_const) As Long
   KOLOR = GetSysColor(COLa)
   
End Function






Private Function Draw_3D(hwnd As Long, Left As Integer, Top As Integer, Right As Integer, Bottom As Integer, type_of_johna_EDGE As johna_EDGE_3D, type_of_border As Flags_3D)
Dim rect_ As RECT
Dim retval As Integer
Dim hdc As Long

rect_.Left = Left
rect_.Top = Top
rect_.Right = Right
rect_.Bottom = Bottom

retval = DrawEdge(GetDC(hwnd), rect_, type_of_johna_EDGE, type_of_border)

End Function


Private Function Draw_3DEX(DC As Long, Left As Integer, Top As Integer, Right As Integer, Bottom As Integer, type_of_johna_EDGE As johna_EDGE_3D, type_of_border As Flags_3D)
Dim rect_ As RECT
Dim retval As Integer
Dim hdc As Long

rect_.Left = Left
rect_.Top = Top
rect_.Right = Right
rect_.Bottom = Bottom

retval = DrawEdge(DC, rect_, type_of_johna_EDGE, type_of_border)

End Function


Sub DrawTXT(St As String, DC As Long, X, Y, lColor As Long)
  Dim hBrush As Long, oldBrush As Long
  hBrush = CreateSolidBrush(lColor)
  oldBrush = SelectObject(DC, hBrush)
  
  TextOut DC, X, Y, St, Len(St)
  
    SelectObject DC, oldBrush
    'delete our  brush
    DeleteObject hBrush

End Sub



Private Sub Class_Initialize()
 ORIENTATIO = Johna_Horizontal
End Sub

Private Sub Class_Terminate()
 If TBAR.Width > 0 Then Set TBAR.lp_PICTURE = Nothing
 If BarDC > 0 Then DeleteDC BarDC
 If HBitmap > 0 Then DeleteObject HBitmap
 


End Sub


Public Property Get Bar_COLOR() As Long
       Bar_COLOR = TBAR.COLOR
End Property

Public Property Let Bar_COLOR(ByVal vNewColor As Long)
   TBAR.COLOR = vNewColor
   PropertyChanged "Bar_COLOR"
   TBAR.lp_PICTURE.Cls
   UPdate_BAR TBAR.percent
   
End Property

Public Property Get Orientation() As Johna_Bar_Orientation
   Orientation = ORIENTATIO
End Property

Public Property Let Orientation(ByVal vNewOrientation As Johna_Bar_Orientation)
   ORIENTATIO = vNewOrientation
   PropertyChanged "Orientation"
   TBAR.lp_PICTURE.Cls
   UPdate_BAR TBAR.percent
End Property


Private Sub PicBAR_Paint()
  Me.Refresh_BAR
End Sub

Private Sub UserControl_Initialize()
  INIT_BAR
End Sub

Private Sub UserControl_Paint()
  Me.Refresh_BAR
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
 MIN = PropBag.ReadProperty("MIN", 0)
 MAX = PropBag.ReadProperty("MAX", 100)
 Bar_COLOR = PropBag.ReadProperty("Bar_COLOR", RGB(110, 80, 185))
 Value = PropBag.ReadProperty("Value", 0)
 Border_Type = PropBag.ReadProperty("Border_Type", Border_Type)
 Orientation = PropBag.ReadProperty("Orientation", Johna_Horizontal)



End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

   PropBag.WriteProperty "MIN", MIN, 0
   PropBag.WriteProperty "MAX", MAX, 100
   PropBag.WriteProperty "Value", Value, MIN
   PropBag.WriteProperty "Border_Type", Border_Type, EDGE_RAISED
   PropBag.WriteProperty "Bar_COLOR", Bar_COLOR, RGB(110, 80, 185)
   PropBag.WriteProperty "Orientation", Orientation, Johna_Horizontal
   
   
                         

End Sub



Private Sub UserControl_Resize()
 PicBAR.Width = UserControl.Width
 PicBAR.Height = UserControl.Height
 PicBAR.Left = 0
 PicBAR.Top = 0
 Call Refresh_BAR
End Sub

Public Property Get Value() As Integer
   

    Value = TBAR.Value
End Property

Public Property Let Value(ByVal VnewVAL As Integer)
      
   Dim LastVAL
   Dim Ecart
   Dim Lev
   LastVAL = TBAR.Value
   TBAR.Value = VnewVAL
   
   'if min=-50;max=-2
If (SIGNE(TBAR.MIN) = -1 And SIGNE(TBAR.MAX) = -1) Or (SIGNE(TBAR.MIN) = 1 And SIGNE(TBAR.MAX) = 1) Then
   Ecart = Abs(TBAR.MAX) - Abs(TBAR.MIN)
   Lev = Abs(TBAR.Value) - Abs(TBAR.MIN)
   TBAR.percent = (Lev / Ecart) * 100
   
ElseIf (SIGNE(TBAR.MIN) = -1 And SIGNE(TBAR.MAX) = 1) Then
   Dim POS
   Ecart = Abs(TBAR.MIN) + Abs(TBAR.MAX)
   
   If TBAR.Value < 0 Then
     POS = Abs(TBAR.MIN) - Abs(TBAR.Value)
     Lev = POS
   ElseIf TBAR.Value >= 0 Then
     POS = Abs(TBAR.MIN) + TBAR.Value
     Lev = POS
   End If
   
   TBAR.percent = (Lev / Ecart) * 100
 

End If

If TBAR.Value < LastVAL Then TBAR.lp_PICTURE.Cls

   UPdate_BAR TBAR.percent
   PropertyChanged "Value"
   
   
End Property

Private Function SIGNE(Val) As Integer
  If Val >= 0 Then SIGNE = 1
  If Val < 0 Then SIGNE = -1
 
  
End Function



Public Property Get Border_Type() As johna_EDGE_3D
    Border_Type = TBAR.DRAW_TYPE
End Property

Public Property Let Border_Type(ByVal vNewBorder3d As johna_EDGE_3D)
   TBAR.DRAW_TYPE = vNewBorder3d
   PropertyChanged "Border_Type"
   UPdate_BAR TBAR.percent
End Property

Public Property Get MIN() As Integer
   MIN = TBAR.MIN
End Property

Public Property Let MIN(ByVal vNewMin As Integer)
  TBAR.MIN = vNewMin
  PropertyChanged "MIN"
  UPdate_BAR TBAR.percent
End Property


Public Property Get MAX() As Integer
   MAX = TBAR.MAX
   
End Property

Public Property Let MAX(ByVal vNewmax As Integer)
  TBAR.MAX = vNewmax
  PropertyChanged "MAX"
  UPdate_BAR TBAR.percent
End Property



