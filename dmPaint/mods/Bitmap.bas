Attribute VB_Name = "Bitmap"
Option Explicit

Private Type TRGB
    R As Integer
    G As Integer
    B As Integer
End Type

Private BmpBits() As Long
Public TmpRgb As TRGB

Private Sub ResetPixels()
    Erase BmpBits
End Sub

Private Sub GetBitmapBits(pBox As PictureBox)
Dim iRet As Long
Dim X As Long
Dim Y As Long
Dim lClr As Long

    'Resize BmpBits to hold pixels
    ReDim BmpBits(0 To 2, 0 To pBox.ScaleWidth, 0 To pBox.ScaleHeight) As Long
    
    For X = 0 To pBox.ScaleWidth
        For Y = 0 To pBox.ScaleHeight
            'Get color
            lClr = GetPixel(pBox.hDC, X, Y)
            'Store Pixels and Red
            BmpBits(0, X, Y) = (lClr Mod 256)
            'Store Pixels and Green
            BmpBits(1, X, Y) = ((lClr And &HFF00) / 256) Mod 256
            'Store Pixels and Blue
            BmpBits(2, X, Y) = (lClr And &HFF0000) / 65536
        Next Y
    Next X
End Sub

Public Sub InvertImage(pBox As PictureBox, Index As Integer)
Dim X As Long
Dim Y As Long
Dim mRgb As TRGB
    
    'Get Pixels
    Call GetBitmapBits(pBox)
    
    For X = 0 To pBox.ScaleWidth
        For Y = 0 To pBox.ScaleHeight
            mRgb.R = BmpBits(0, X, Y)
            mRgb.G = BmpBits(1, X, Y)
            mRgb.B = BmpBits(2, X, Y)
            'Invert colors
            
            Select Case Index
                Case 0
                    mRgb.R = (255 - mRgb.R)
                    mRgb.G = (255 - mRgb.G)
                    mRgb.B = (255 - mRgb.B)
                Case 1
                    mRgb.R = (255 - mRgb.R)
                Case 2
                    mRgb.G = (255 - mRgb.G)
                Case 3
                    mRgb.B = (255 - mRgb.B)
            End Select

            If (mRgb.R < 0) Then mRgb.R = 0
            If (mRgb.G < 0) Then mRgb.G = 0
            If (mRgb.B < 0) Then mRgb.B = 0
            
            'Set pixels
            SetPixelV pBox.hDC, X, Y, RGB(mRgb.R, mRgb.G, mRgb.B)
        Next Y
    Next X
    Call ResetPixels
    pBox.Refresh
End Sub

Public Sub RemoveRGB(pBox As PictureBox, Index As Integer)
Dim X As Long
Dim Y As Long
Dim mRgb As TRGB
    
    'Get Pixels
    Call GetBitmapBits(pBox)
    
    For X = 0 To pBox.ScaleWidth - 1
        For Y = 0 To pBox.ScaleHeight - 1
            mRgb.R = BmpBits(0, X, Y)
            mRgb.G = BmpBits(1, X, Y)
            mRgb.B = BmpBits(2, X, Y)
            'Remove Colors
            Select Case Index
                Case 0: mRgb.R = 0
                Case 1: mRgb.G = 0
                Case 2: mRgb.B = 0
            End Select
            'Set pixels
            SetPixelV pBox.hDC, X, Y, RGB(mRgb.R, mRgb.G, mRgb.B)
        Next Y
    Next X
    Call ResetPixels
    pBox.Refresh
End Sub

Public Sub GrayImage(pBox As PictureBox)
Dim X As Long
Dim Y As Long
Dim Gray As Long
Dim mRgb As TRGB
    
    'Get Pixels
    Call GetBitmapBits(pBox)
    
    For X = 0 To pBox.ScaleWidth - 1
        For Y = 0 To pBox.ScaleHeight - 1
            mRgb.R = BmpBits(0, X, Y)
            mRgb.G = BmpBits(1, X, Y)
            mRgb.B = BmpBits(2, X, Y)
            'Gray Color
            ''Gray = (mRgb.R + mRgb.G + mRgb.B) \ 3
            Gray = Round(mRgb.R * 0.56 + mRgb.G * 0.33 + mRgb.B * 0.11)
            'Set pixels
            SetPixelV pBox.hDC, X, Y, RGB(Gray, Gray, Gray)
        Next Y
    Next X
    Call ResetPixels
    pBox.Refresh
End Sub

Public Sub Noise(pBox As PictureBox, NoiseVal As Integer)
Dim X As Long
Dim Y As Long
Dim mRgb As TRGB

    'Get Pixels
    Call GetBitmapBits(pBox)
    
    For X = 0 To pBox.ScaleWidth - 1
        For Y = 0 To pBox.ScaleHeight - 1
            'Edit RGB Colors
            mRgb.R = Abs(BmpBits(0, X, Y)) + ((NoiseVal * 2 + 1) * Rnd - NoiseVal)
            mRgb.G = Abs(BmpBits(1, X, Y)) + ((NoiseVal * 2 + 1) * Rnd - NoiseVal)
            mRgb.B = Abs(BmpBits(2, X, Y)) + ((NoiseVal * 2 + 1) * Rnd - NoiseVal)
            
            If (mRgb.R < 0) Then mRgb.R = 0
            If (mRgb.G < 0) Then mRgb.G = 0
            If (mRgb.B < 0) Then mRgb.B = 0
            'Set the pixels
            SetPixelV pBox.hDC, X, Y, RGB(mRgb.R, mRgb.G, mRgb.B)
        Next
    Next
    
    Call ResetPixels
    pBox.Refresh
End Sub

Public Sub LightenDarken(pBox As PictureBox, lValue As Integer)
Dim X As Long
Dim Y As Long
Dim mRgb As TRGB
    
    'Get Pixels
    Call GetBitmapBits(pBox)
    
    For X = 0 To pBox.ScaleWidth - 1
        For Y = 0 To pBox.ScaleHeight - 1
            mRgb.R = BmpBits(0, X, Y)
            mRgb.G = BmpBits(1, X, Y)
            mRgb.B = BmpBits(2, X, Y)
            
            mRgb.R = (mRgb.R + lValue)
            mRgb.G = (mRgb.G + lValue)
            mRgb.B = (mRgb.B + lValue)

            If (mRgb.R > 255) Then mRgb.R = 255
            If (mRgb.G > 255) Then mRgb.G = 255
            If (mRgb.B > 255) Then mRgb.B = 255
            
            If (mRgb.R < 0) Then mRgb.R = 0
            If (mRgb.G < 0) Then mRgb.G = 0
            If (mRgb.B < 0) Then mRgb.B = 0
            
            'Set pixels
            SetPixelV pBox.hDC, X, Y, RGB(mRgb.R, mRgb.G, mRgb.B)
        Next Y
    Next X
    Call ResetPixels
    pBox.Refresh
End Sub

Public Sub BlurImage(pBox As PictureBox)
Dim X As Long
Dim Y As Long
Dim mRgb As TRGB
    
    'Get Pixels
    Call GetBitmapBits(pBox)
    
    For X = 1 To pBox.ScaleWidth - 2
        For Y = 1 To pBox.ScaleHeight - 2
            
            mRgb.R = BmpBits(0, X - 1, Y - 1) + BmpBits(0, X - 1, Y) + BmpBits(0, X - 1, Y + 1) + _
            BmpBits(0, X, Y - 1) + BmpBits(0, X, Y) + BmpBits(0, X, Y + 1) + _
            BmpBits(0, X + 1, Y - 1) + BmpBits(0, X + 1, Y) + BmpBits(0, X + 1, Y + 1)
            
            mRgb.G = BmpBits(1, X - 1, Y - 1) + BmpBits(1, X - 1, Y) + BmpBits(1, X - 1, Y + 1) + _
            BmpBits(1, X, Y - 1) + BmpBits(1, X, Y) + BmpBits(1, X, Y + 1) + _
            BmpBits(1, X + 1, Y - 1) + BmpBits(1, X + 1, Y) + BmpBits(1, X + 1, Y + 1)
            
            mRgb.B = BmpBits(2, X - 1, Y - 1) + BmpBits(2, X - 1, Y) + BmpBits(2, X - 1, Y + 1) + _
            BmpBits(2, X, Y - 1) + BmpBits(2, X, Y) + BmpBits(2, X, Y + 1) + _
            BmpBits(2, X + 1, Y - 1) + BmpBits(2, X + 1, Y) + BmpBits(2, X + 1, Y + 1)
            
            'Set pixels
            SetPixelV pBox.hDC, X, Y, RGB(mRgb.R / 9, mRgb.G / 9, mRgb.B / 9)
        Next Y
    Next X
    Call ResetPixels
    pBox.Refresh
End Sub

Public Sub Sharpen(pBox As PictureBox, lValue As Single)
Dim X As Long
Dim Y As Long
Dim mRgb As TRGB
    
    'Get Pixels
    Call GetBitmapBits(pBox)
    
    For X = 1 To pBox.ScaleWidth - 2
        For Y = 1 To pBox.ScaleHeight - 2
            mRgb.R = BmpBits(0, X, Y)
            mRgb.G = BmpBits(1, X, Y)
            mRgb.B = BmpBits(2, X, Y)
            'Sharpen colors
            mRgb.R = BmpBits(0, X, Y) + lValue * (BmpBits(0, X, Y) - BmpBits(0, X - 1, Y - 1))
            mRgb.G = BmpBits(1, X, Y) + lValue * (BmpBits(1, X, Y) - BmpBits(1, X - 1, Y - 1))
            mRgb.B = BmpBits(2, X, Y) + lValue * (BmpBits(2, X, Y) - BmpBits(2, X - 1, Y - 1))
            
            If (mRgb.R < 0) Then mRgb.R = 0
            If (mRgb.G < 0) Then mRgb.G = 0
            If (mRgb.B < 0) Then mRgb.B = 0
            'Set pixels
            SetPixelV pBox.hDC, X, Y, RGB(mRgb.R, mRgb.G, mRgb.B)
        Next Y
    Next X
    Call ResetPixels
    pBox.Refresh
End Sub

Public Sub Diffuse(pBox As PictureBox)
Dim X As Long
Dim Y As Long
Dim mRgb As TRGB
Dim Rnd1 As Double

    'Get Pixels
    Call GetBitmapBits(pBox)
    
    For X = 2 To pBox.ScaleWidth - 3
        For Y = 2 To pBox.ScaleHeight - 3
            'Diffuse value
            Rnd1 = (Rnd * 2) - 2
            
            mRgb.R = Abs(BmpBits(0, X + Rnd1, Y + Rnd1))
            mRgb.G = Abs(BmpBits(1, X + Rnd1, Y + Rnd1))
            mRgb.B = Abs(BmpBits(2, X + Rnd1, Y + Rnd1))
        
            'Set pixels
            SetPixelV pBox.hDC, X, Y, RGB(mRgb.R, mRgb.G, mRgb.B)
        Next Y
    Next X
    Call ResetPixels
    pBox.Refresh
End Sub

Public Sub AdjustRGB(pBox As PictureBox, R As Integer, G As Integer, B As Integer)
Dim X As Long
Dim Y As Long
Dim mRgb As TRGB
    
    'Get Pixels
    Call GetBitmapBits(pBox)
    
    For X = 0 To pBox.ScaleWidth - 1
        For Y = 0 To pBox.ScaleHeight - 1
            mRgb.R = BmpBits(0, X, Y)
            mRgb.G = BmpBits(1, X, Y)
            mRgb.B = BmpBits(2, X, Y)
            
            mRgb.R = mRgb.R + R
            mRgb.G = mRgb.G + G
            mRgb.B = mRgb.B + B
            
            If (mRgb.R < 0) Then mRgb.R = 0
            If (mRgb.G < 0) Then mRgb.G = 0
            If (mRgb.B < 0) Then mRgb.B = 0
            
            'Set pixels
            SetPixelV pBox.hDC, X, Y, RGB(mRgb.R, mRgb.G, mRgb.B)
        Next Y
    Next X
    Call ResetPixels
    pBox.Refresh
End Sub

Public Sub PixelsEffect(pBox As PictureBox, PixelVal As Integer)
Dim X As Long
Dim Y As Long
Dim mRgb As TRGB
    
    'Get Pixels
    Call GetBitmapBits(pBox)
    For X = 0 To pBox.ScaleWidth - 1 Step PixelVal
        For Y = 0 To pBox.ScaleHeight - 1 Step PixelVal
            mRgb.R = BmpBits(0, X, Y)
            mRgb.G = BmpBits(1, X, Y)
            mRgb.B = BmpBits(2, X, Y)
            'Set pixels
            pBox.Line (X, Y)-(X + PixelVal, Y + PixelVal), RGB(mRgb.R, mRgb.G, mRgb.B), BF
            'SetPixelV pBox.hDC, X + Cos(Y), Y + Sin(X), RGB(mRgb.R, mRgb.G, mRgb.B)
        Next Y
    Next X
    Call ResetPixels
    pBox.Refresh
End Sub

Public Sub Mirror(pBox As PictureBox, Index As Integer)

    If (Index = 0) Then
        'Mirror left
        StretchBlt pBox.hDC, -pBox.Width, 0, -pBox.Width, pBox.Height, pBox.hDC, 0, 0, pBox.Width, pBox.Height, vbSrcCopy
        pBox.Picture = pBox.image
        pBox.PaintPicture pBox, pBox.Width, 0, -pBox.Width / 2, pBox.Height, 0, 0, pBox.Width / 2
    Else
        'Mirror right
        StretchBlt pBox.hDC, pBox.Width, 0, -pBox.Width, pBox.Height, pBox.hDC, 0, 0, pBox.Width, pBox.Height, vbSrcCopy
        pBox.Picture = pBox.image
        pBox.PaintPicture pBox, pBox.Width, 0, -pBox.Width / 2, pBox.Height, 0, 0, pBox.Width / 2
    End If
    
End Sub

Public Sub SwapRGB(pBox As PictureBox, Index As Integer)
Dim X As Long
Dim Y As Long
Dim mRgb As TRGB
Dim Tmp1 As Integer
Dim Tmp2 As Integer
Dim Tmp3 As Integer

    'Get Pixels
    Call GetBitmapBits(pBox)
    
    For X = 0 To pBox.ScaleWidth - 1
        For Y = 0 To pBox.ScaleHeight - 1
            mRgb.R = BmpBits(0, X, Y)
            mRgb.G = BmpBits(1, X, Y)
            mRgb.B = BmpBits(2, X, Y)
            
            Select Case Index
                Case 0
                    'RGB->RBG
                    Tmp1 = mRgb.B
                    mRgb.B = mRgb.G
                    mRgb.G = Tmp1
                Case 1
                    'RGB->GRB
                    Tmp1 = mRgb.R
                    mRgb.R = mRgb.G
                    mRgb.G = Tmp1
                Case 2
                    'RGB->GBR
                    Tmp1 = mRgb.B
                    Tmp2 = mRgb.G
                    mRgb.R = mRgb.G
                    mRgb.G = Tmp1
                    mRgb.B = Tmp2
                Case 3
                    'RGB->BRG
                    Tmp1 = mRgb.B
                    Tmp2 = mRgb.R
                    Tmp3 = mRgb.G
                    mRgb.R = Tmp1
                    mRgb.G = Tmp2
                    mRgb.B = Tmp3
                Case 4
                    'RGB->BGR
                    Tmp1 = mRgb.B
                    Tmp2 = mRgb.R
                    mRgb.R = Tmp1
                    mRgb.B = Tmp2
            End Select
            
            'Set pixels
            SetPixelV pBox.hDC, X, Y, RGB(mRgb.R, mRgb.G, mRgb.B)
        Next Y
    Next X
    Call ResetPixels
    pBox.Refresh
End Sub

Public Sub BlackAndWhite(pBox As PictureBox)
Dim X As Long
Dim Y As Long
Dim mRgb As TRGB
Dim lCol As Long

    'Get Pixels
    Call GetBitmapBits(pBox)
    
    For X = 0 To pBox.ScaleWidth - 1
        For Y = 0 To pBox.ScaleHeight - 1
            mRgb.R = BmpBits(0, X, Y)
            mRgb.G = BmpBits(1, X, Y)
            mRgb.B = BmpBits(2, X, Y)
            
            If (mRgb.R > 194) Then
                lCol = vbWhite
            ElseIf (mRgb.G > 194) Then
                lCol = vbWhite
            ElseIf (mRgb.B > 194) Then
                lCol = vbWhite
            Else
                lCol = vbBlack
            End If
            
            'Set pixels
            SetPixelV pBox.hDC, X, Y, lCol
        Next Y
    Next X
    Call ResetPixels
    pBox.Refresh
End Sub

Public Sub Flatten(pBox As PictureBox, Value As Integer)
Dim X As Long
Dim Y As Long
Dim mRgb As TRGB
Dim lCol As Long
On Error Resume Next

    'Get Pixels
    Call GetBitmapBits(pBox)
    
    For X = 0 To pBox.ScaleWidth - 1
        For Y = 0 To pBox.ScaleHeight - 1
            'Imboss
            mRgb.R = Abs(BmpBits(0, X, Y) - BmpBits(0, X + 1, Y + 1) + Value)
            mRgb.G = Abs(BmpBits(1, X, Y) - BmpBits(1, X + 1, Y + 1) + Value)
            mRgb.B = Abs(BmpBits(2, X, Y) - BmpBits(2, X + 1, Y + 1) + Value)
            
            'Set pixels
            SetPixelV pBox.hDC, X, Y, RGB(mRgb.R, mRgb.G, mRgb.B)
        Next Y
    Next X
    Call ResetPixels
    pBox.Refresh
End Sub

Public Sub ColorBurn(pBox As PictureBox)
Dim X As Long
Dim Y As Long
Dim mRgb As TRGB
On Error Resume Next

    'Get Pixels
    Call GetBitmapBits(pBox)
    
    For X = 0 To pBox.ScaleWidth - 1
        For Y = 0 To pBox.ScaleHeight - 1
            'Imboss
            mRgb.R = BmpBits(0, X, Y) Xor 32
            mRgb.G = BmpBits(1, X, Y) Xor 32
            mRgb.B = BmpBits(2, X, Y) Xor 32
            'Set pixels
            
            SetPixelV pBox.hDC, X, Y, RGB(mRgb.R, mRgb.G, mRgb.B)
        Next Y
    Next X
    Call ResetPixels
    pBox.Refresh
End Sub

Public Sub Highlight(pBox As PictureBox, R As Integer, G As Integer, B As Integer)
Dim X As Long
Dim Y As Long
Dim mRgb As TRGB
Dim tmp As TRGB
    
    'Fill RGB Type
    tmp.R = R
    tmp.G = G
    tmp.B = B
    
    'Get Pixels
    Call GetBitmapBits(pBox)
    
    For X = 0 To pBox.ScaleWidth - 1
        For Y = 0 To pBox.ScaleHeight - 1
            'Highlight color
            mRgb.R = (tmp.R * BmpBits(0, X, Y)) / 255
            mRgb.G = (tmp.G * BmpBits(1, X, Y)) / 255
            mRgb.B = (tmp.B * BmpBits(2, X, Y)) / 255
            'Set pixels
            SetPixelV pBox.hDC, X, Y, RGB(mRgb.R, mRgb.G, mRgb.B)
        Next Y
    Next X
    Call ResetPixels
    pBox.Refresh
End Sub

Public Sub MeltEffect(pBox As PictureBox)
Dim X As Long
Dim Y As Long
Dim mRgb As TRGB
On Error Resume Next

    'Get Pixels
    Call GetBitmapBits(pBox)
    
    For X = 0 To pBox.ScaleWidth - 1
        For Y = 0 To pBox.ScaleHeight - 1
            'Imboss
            mRgb.R = BmpBits(0, X, Y)
            mRgb.G = BmpBits(1, X, Y)
            mRgb.B = BmpBits(2, X, Y)
            'Set pixels
            SetPixelV pBox.hDC, X - Cos(Y / 8) + Sin(Y / 8), Y - Cos(X / 8) + Sin(X / 8), RGB(mRgb.R, mRgb.G, mRgb.B)
        Next Y
    Next X
    Call ResetPixels
    pBox.Refresh
End Sub

Public Sub XorColor(pBox As PictureBox, ByVal iValue As Integer)
Dim X As Long
Dim Y As Long
Dim mRgb As TRGB
Dim lCol As Long

    'Get Pixels
    Call GetBitmapBits(pBox)
    
    For X = 0 To pBox.ScaleWidth - 1
        For Y = 0 To pBox.ScaleHeight - 1
            mRgb.R = BmpBits(0, X, Y) Xor iValue
            mRgb.G = BmpBits(1, X, Y) Xor iValue
            mRgb.B = BmpBits(2, X, Y) Xor iValue
            
            If (mRgb.R > 255) Then mRgb.R = 255
            If (mRgb.G > 255) Then mRgb.G = 255
            If (mRgb.B > 255) Then mRgb.B = 255
            
            If (mRgb.R < 0) Then mRgb.R = 0
            If (mRgb.G < 0) Then mRgb.G = 0
            If (mRgb.B < 0) Then mRgb.B = 0
            'Set pixels
            SetPixelV pBox.hDC, X, Y, RGB(mRgb.R, mRgb.G, mRgb.B)
        Next Y
    Next X
    Call ResetPixels
    pBox.Refresh
End Sub

Public Sub CustomFlilter(pBox As PictureBox)
On Error Resume Next
Dim Rgb1 As TRGB
Dim Rgb2 As TRGB
Dim Offset As Integer
Dim X As Long
Dim Y As Long
Dim X1 As Long
Dim Y1 As Long

    'Get Pixels
    Call GetBitmapBits(pBox)
    
    If FilterNorm = 0 Then FilterNorm = 1
    Offset = IIf(FliterSize = 1, 1, 2)

    For X = Offset To pBox.ScaleWidth - Offset - 1
        For Y = Offset To pBox.ScaleHeight - Offset - 1
            'Reset rgb values
            Rgb1.R = 0
            Rgb1.G = 0
            Rgb1.B = 0

            For X1 = -Offset To Offset
                For Y1 = -Offset To Offset
                    'Preform custom fillter
                    Rgb1.R = Rgb1.R + BmpBits(0, X + X1, Y + Y1) * CustomFilter(X1 + 2, Y1 + 2)
                    Rgb1.G = Rgb1.G + BmpBits(1, X + X1, Y + Y1) * CustomFilter(X1 + 2, Y1 + 2)
                    Rgb1.B = Rgb1.B + BmpBits(2, X + X1, Y + Y1) * CustomFilter(X1 + 2, Y1 + 2)
                Next Y1
            Next X1
            
            Rgb1.R = Abs(Rgb1.R / FilterNorm + FilterBias)
            Rgb1.G = Abs(Rgb1.G / FilterNorm + FilterBias)
            Rgb1.B = Abs(Rgb1.B / FilterNorm + FilterBias)
            'Set pixels
            SetPixelV pBox.hDC, X, Y, RGB(Rgb1.R, Rgb1.G, Rgb1.B)
        Next
        DoEvents
    Next
    
    Call ResetPixels
    pBox.Refresh
    
End Sub

Public Sub Gamma(pBox As PictureBox, ByVal iValue As Integer)
Dim X As Long
Dim Y As Long
Dim mRgb As TRGB
    
    'Get Pixels
    Call GetBitmapBits(pBox)
    'Check for divBy0
    If (iValue <= 0) Then iValue = 1
    
    For X = 0 To pBox.ScaleWidth - 1
        For Y = 0 To pBox.ScaleHeight - 1
            mRgb.R = BmpBits(0, X, Y)
            mRgb.G = BmpBits(1, X, Y)
            mRgb.B = BmpBits(2, X, Y)
            'Gamma
            mRgb.R = ((mRgb.R / 255) ^ (1 / iValue)) * 255
            mRgb.G = ((mRgb.G / 255) ^ (1 / iValue)) * 255
            mRgb.B = ((mRgb.B / 255) ^ (1 / iValue)) * 255
            'Set pixels
            SetPixelV pBox.hDC, X, Y, RGB(mRgb.R, mRgb.G, mRgb.B)
        Next Y
    Next X
    Call ResetPixels
    pBox.Refresh
End Sub

Public Sub Contrast(pBox As PictureBox, ByVal iValue As Integer)
Dim X As Long
Dim Y As Long
Dim mRgb As TRGB
    
    'Get Pixels
    Call GetBitmapBits(pBox)
    
    For X = 0 To pBox.ScaleWidth - 1
        For Y = 0 To pBox.ScaleHeight - 1
            'Contrast
            mRgb.R = ((BmpBits(0, X, Y) - 128) * iValue) + 128
            mRgb.G = ((BmpBits(1, X, Y) - 128) * iValue) + 128
            mRgb.B = ((BmpBits(2, X, Y) - 128) * iValue) + 128
            
            'Check for overflow and underflow
            If (mRgb.R > 255) Then mRgb.R = 255
            If (mRgb.G > 255) Then mRgb.G = 255
            If (mRgb.B > 255) Then mRgb.B = 255
            If (mRgb.R < 0) Then mRgb.R = 0
            If (mRgb.G < 0) Then mRgb.G = 0
            If (mRgb.B < 0) Then mRgb.B = 0

            'Set pixels
            SetPixelV pBox.hDC, X, Y, RGB(mRgb.R, mRgb.G, mRgb.B)
        Next Y
    Next X
    Call ResetPixels
    pBox.Refresh
End Sub

Public Sub ReplaceColor(pBox As PictureBox, FindColor As OLE_COLOR, ReplaceColor As OLE_COLOR)
Dim X As Long
Dim Y As Long
Dim mRgb As TRGB
Dim lCol As OLE_COLOR
    'Get Pixels
    Call GetBitmapBits(pBox)
    
    For X = 0 To pBox.ScaleWidth - 1
        For Y = 0 To pBox.ScaleHeight - 1
            'Contrast
            mRgb.R = BmpBits(0, X, Y)
            mRgb.G = BmpBits(1, X, Y)
            mRgb.B = BmpBits(2, X, Y)
            lCol = RGB(mRgb.R, mRgb.G, mRgb.B)
            
            'Check if FindColor is found
            If lCol = FindColor Then
                lCol = ReplaceColor
            End If
            
            'Set pixels
            SetPixelV pBox.hDC, X, Y, lCol
        Next Y
    Next X
    Call ResetPixels
    pBox.Refresh
End Sub

Public Sub Snow(pBox As PictureBox, iVal As Integer)
Dim X As Long
Dim Y As Long
Dim mRgb As TRGB
Dim R As Integer

    'Get Pixels
    Call GetBitmapBits(pBox)
    
    For X = 0 To pBox.ScaleWidth - 1
        For Y = 0 To pBox.ScaleHeight - 1
            
            R = Int(Rnd * iVal)
            
            mRgb.R = BmpBits(0, X, Y) + R
            mRgb.G = BmpBits(1, X, Y) + R
            mRgb.B = BmpBits(2, X, Y) + R
            
            If (mRgb.R < 0) Then mRgb.R = 0
            If (mRgb.G < 0) Then mRgb.G = 0
            If (mRgb.B < 0) Then mRgb.B = 0
            
            If (mRgb.R > 255) Then mRgb.R = 255
            If (mRgb.G > 255) Then mRgb.G = 255
            If (mRgb.B > 255) Then mRgb.B = 255

            'Set pixels
            SetPixelV pBox.hDC, X, Y, RGB(mRgb.R, mRgb.G, mRgb.B)
        Next Y
    Next X
    Call ResetPixels
    pBox.Refresh
End Sub

