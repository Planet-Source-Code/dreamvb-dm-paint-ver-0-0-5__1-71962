Attribute VB_Name = "ImageTools"
Option Explicit

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type ImageProp
    iWidth As Long
    iHeight As Long
    iBackColor As OLE_COLOR
End Type

Public Type TextProp
    ffName As String
    fSize As Integer
    fText As String
    fStyle As Integer
End Type

Public Type cRgb
    Red As Integer
    Green As Integer
    Blue As Integer
End Type

Private Type GRADIENT_RECT
    UpperLeft As Long
    LowerRight As Long
End Type

Private Type TRIVERTEX
    X As Long
    Y As Long
    Red As Integer
    Green As Integer
    Blue As Integer
    Alpha As Integer
End Type

Private Type BLENDFUNCTION
  BlendOp As Byte
  BlendFlags As Byte
  SourceConstantAlpha As Byte
  AlphaFormat As Byte
End Type

Enum GRADIENT_DIR
    Horizontal = &H0
    Vertical = &H1
End Enum

Public mImage As ImageProp
Public mTextProp As TextProp
Public StrokeW As Long
Public StrokeC As OLE_COLOR
Public mBrightness As Integer
Public mSharpen As Single
Public mMaskFile As String
Public mRGBSwapIndex As Integer
'Custom image fliter variables
Public CustomFilter(5, 5) As Integer
Public FilterNorm As Integer
Public FilterBias As Integer
Public FliterSize As Integer
'Gradient Alpha variables
Public mAlphaLevel As Integer
Public AlphaDir As Integer

Private Declare Function BeginPath Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function StrokeAndFillPath Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function EndPath Lib "gdi32.dll" (ByVal hdc As Long) As Long
Public Declare Function SetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function SetStretchBltMode Lib "gdi32.dll" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Public Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByRef lColorRef As Long) As Long
Private Declare Function GradientFill Lib "msimg32" (ByVal hdc As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Public Declare Function Ellipse Lib "gdi32.dll" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function RoundRect Lib "gdi32.dll" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function ExtFloodFill Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Public Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetPixelV Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function CreatePatternBrush Lib "gdi32.dll" (ByVal hBitmap As Long) As Long
Public Declare Function CreateHatchBrush Lib "gdi32.dll" (ByVal nIndex As Long, ByVal crColor As Long) As Long
Public Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Public Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Long
Public Declare Function TextOut Lib "gdi32.dll" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function OpenClipboard Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function EmptyClipboard Lib "user32.dll" () As Long
Private Declare Function SetClipboardData Lib "user32.dll" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function CloseClipboard Lib "user32.dll" () As Long
Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal BLENDFUNCT As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (Destination As Any, Source As Any, ByVal Length As Long)

Private Const CF_BITMAP As Long = 2

Public Sub OutLineText(pBox As PictureBox, ByVal X As Integer, ByVal Y As Integer, ByVal sText As String)
Dim iRet As Long
    'Begin the path
    iRet = BeginPath(pBox.hdc)
    'Set the text
    iRet = TextOut(pBox.hdc, X, Y, sText, Len(sText))
    'End the path
    iRet = EndPath(pBox.hdc)
    'Outline the path
    iRet = StrokeAndFillPath(pBox.hdc)
    pBox.Refresh
End Sub

Public Sub PictureBoxToClipboard(pBox As PictureBox)
Dim hBmpC As Long
Dim hBmp As Long
Dim hOldBmp As Long
Dim pDc As Long
Dim iRet As Long

    'Create DC Device
    pDc = CreateCompatibleDC(pBox.hdc)
    If (pDc = 0) Then Exit Sub
    'Create Bitmap object
    hBmp = CreateCompatibleBitmap(pBox.hdc, pBox.ScaleWidth, pBox.ScaleHeight)
    If (hBmp = 0) Then Exit Sub
    
    'Place picturebox data onto hbmp
    hOldBmp = SelectObject(pDc, hBmp)
    iRet = BitBlt(pDc, 0, 0, pBox.Width, pBox.ScaleHeight, pBox.hdc, 0, 0, vbSrcCopy)
    
    Call SelectObject(pDc, hOldBmp)
    'Open Clipabord
    iRet = OpenClipboard(0)
    'Empty Clipboard data
    iRet = EmptyClipboard
    'Set Clipboard data as bitmap
    iRet = SetClipboardData(CF_BITMAP, hBmp)
    'Close the clipboard
    iRet = CloseClipboard
    'Delete DC Object
    Call DeleteObject(pDc)
End Sub

Public Sub StrokeImage(pBox As PictureBox, sWidth As Long, sColor As OLE_COLOR)
Dim Count As Long
    'Draws a outline around the image with a selected color
    With pBox
        For Count = 0 To sWidth
            pBox.Line (Count - 1, Count - 1)-(.Width - Count, .Height - Count), sColor, B
        Next Count
        .Refresh
    End With
End Sub

Public Sub ScaleImage(pBox As PictureBox, sBox As PictureBox, ByVal iWidth As Long, ByVal iHeight As Long)
    'Scale the image
    With pBox
        'Resize src picturebox
        sBox.Cls
        sBox.Width = iWidth
        sBox.Height = iHeight
        'Set Palette mode
        SetStretchBltMode sBox.hdc, vbPaletteModeNone
        'Resize pBox onto temp sBox
        StretchBlt sBox.hdc, 0, 0, iWidth, iHeight, .hdc, 0, 0, .Width, .Height, vbSrcCopy
        Set .Picture = Nothing
        'Resize src picturebox
        .Width = iWidth
        .Height = iHeight
        'Copy the new image to the src picturebox
        BitBlt .hdc, 0, 0, iWidth, iHeight, sBox.hdc, 0, 0, vbSrcCopy
        sBox.Cls
        .Refresh
    End With
End Sub

Public Sub FlipImage(pBox As PictureBox, ByVal FlipOp As Integer)
    'Flip an image Vertical or horizontal
    With pBox
        If (FlipOp = 0) Then
            'Flip Vertical
            StretchBlt .hdc, (.Width - 1), 0, -.Width, .Height, _
            .hdc, 0, 0, .Width, .Height, vbSrcCopy
        ElseIf (FlipOp = 1) Then
            'Flip horizontal
            StretchBlt .hdc, 0, (.Height - 1), .Width, -.Height, _
            .hdc, 0, 0, .Width, .Height, vbSrcCopy
        Else
            'Flip Both
            StretchBlt .hdc, 0, 0, .Width, .Height, _
            .hdc, .Width, .Height, -.Width, -.Height, vbSrcCopy
        End If
        .Refresh
    End With
    
End Sub

Private Function TranslateColor(OleClr As OLE_COLOR, Optional hPal As Integer = 0) As Long
    ' used to return the correct color value of OleClr as a long
    If OleTranslateColor(OleClr, hPal, TranslateColor) Then
        TranslateColor = &HFFFF&
    End If
End Function

Private Sub setTriVertexColor(tTV As TRIVERTEX, oColor As Long)
    Dim lRed As Long
    Dim lGreen As Long
    Dim lBlue As Long
    
    lRed = (oColor And &HFF&) * &H100&
    lGreen = (oColor And &HFF00&)
    lBlue = (oColor And &HFF0000) \ &H100&
    
    setTriVertexColorComponent tTV.Red, lRed
    setTriVertexColorComponent tTV.Green, lGreen
    setTriVertexColorComponent tTV.Blue, lBlue
End Sub

Private Sub setTriVertexColorComponent(ByRef oColor As Integer, ByVal lComponent As Long)
    If (lComponent And &H8000&) = &H8000& Then
        oColor = (lComponent And &H7F00&)
        oColor = oColor Or &H8000
    Else
        oColor = lComponent
    End If
End Sub

Public Sub GDI_GradientFill(hdc As Long, mRect As RECT, mStartColor As OLE_COLOR, mEndColor As OLE_COLOR, gDir As GRADIENT_DIR)
Dim gRect As GRADIENT_RECT
Dim tTV(0 To 1) As TRIVERTEX

    setTriVertexColor tTV(1), TranslateColor(mEndColor)
    tTV(0).X = mRect.Left
    tTV(0).Y = mRect.Top
    
    setTriVertexColor tTV(0), TranslateColor(mStartColor)
    tTV(1).X = mRect.Right
    tTV(1).Y = mRect.Bottom

    gRect.UpperLeft = 0
    gRect.LowerRight = 1
    
    GradientFill hdc, tTV(0), 2, gRect, 1, gDir
    
End Sub

Public Sub GradientAlpha(pBox1 As PictureBox, pBox2 As PictureBox, AlphaLevel As Integer, vColor1 As OLE_COLOR, _
vColor2 As OLE_COLOR, Optional ByVal vDir As GRADIENT_DIR = Horizontal)
Dim iRet As Long
Dim rc As RECT
Dim BF As BLENDFUNCTION, lBF As Long

    With pBox2
        .Width = pBox1.Width
        .Height = pBox1.Height
        iRet = SetRect(rc, 0, 0, .Width, .Height)
        Call GDI_GradientFill(.hdc, rc, vColor1, vColor2, vDir)
        .Refresh
        'Fill alpha type
        With BF
            .BlendOp = &H0
            .BlendFlags = 0
            .SourceConstantAlpha = AlphaLevel
            .AlphaFormat = 0
        End With

        Call RtlMoveMemory(lBF, BF, 4)
        
        AlphaBlend pBox1.hdc, 0, 0, .Width, .Height, .hdc, 0, 0, pBox1.ScaleWidth, pBox1.ScaleHeight, lBF
        pBox1.Refresh
        Set .Picture = Nothing
    End With

End Sub

