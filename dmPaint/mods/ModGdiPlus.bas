Attribute VB_Name = "ModGdiPlus"
'I never wrote this code just edited parts to work with DMPaint but like to say thanks to who ever did write it.
Option Explicit

Private Const GdiPlusVersion As Long = 1

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type IID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7)  As Byte
End Type

Private Type PICTDESC
    cbSizeOfStruct As Long
    picType As Long
    hgdiObj As Long
    hPalOrXYExt As Long
End Type

Private Type ImageCodecInfo
    Clsid As GUID
    FormatID As GUID
    CodecNamePtr As Long
    DllNamePtr As Long
    FormatDescriptionPtr As Long
    FilenameExtensionPtr As Long
    MimeTypePtr As Long
    flags As Long
    Version As Long
    SigCount As Long
    SigSize As Long
    SigPatternPtr As Long
    SigMaskPtr As Long
End Type

Private Type GdiplusStartupOutput
    NotificationHook As Long
    NotificationUnhook As Long
End Type

Private Type GDIPlusStartupInput
    GdiPlusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

' GDI+ Status
Private Enum Status
    OK = 0
    GenericError = 1
    InvalidParameter = 2
    OutOfMemory = 3
    ObjectBusy = 4
    InsufficientBuffer = 5
    NotImplemented = 6
    Win32Error = 7
    WrongState = 8
    Aborted = 9
    FileNotFound = 10
    ValueOverflow = 11
    AccessDenied = 12
    UnknownImageFormat = 13
    FontFamilyNotFound = 14
    FontStyleNotFound = 15
    NotTrueTypeFont = 16
    UnsupportedGdiplusVersion = 17
    GdiplusNotInitialized = 18
    PropertyNotFound = 19
    PropertyNotSupported = 20
    ProfileNotFound = 21
End Enum

Private Declare Function GdipCreateBitmapFromFile Lib "gdiplus" _
    (ByVal Filename As Long, ByRef Bitmap As Long) As Status

Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" _
    (ByVal hbm As Long, ByVal hPal As Long, _
    ByRef Bitmap As Long) As Status

Private Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" _
    (ByVal Bitmap As Long, ByRef hbmReturn As Long, _
    ByVal background As Long) As Status

Private Declare Function GdipDisposeImage Lib "gdiplus" _
    (ByVal image As Long) As Status

Private Declare Function GdipGetImageEncoders Lib "gdiplus" _
    (ByVal numEncoders As Long, ByVal Size As Long, _
    ByRef Encoders As Any) As Status

Private Declare Function GdipGetImageEncodersSize Lib "gdiplus" _
    (ByRef numEncoders As Long, ByRef Size As Long) As Status

Private Declare Function GdiplusShutdown Lib "gdiplus" _
    (ByVal token As Long) As Status

Private Declare Function GdiplusStartup Lib "gdiplus" _
    (ByRef token As Long, ByRef lpInput As GDIPlusStartupInput, _
    Optional ByRef lpOutput As Any) As Status

Private Declare Function GdipSaveImageToFile Lib "gdiplus" _
    (ByVal image As Long, ByVal Filename As Long, _
    ByRef clsidEncoder As GUID, _
    ByRef encoderParams As Any) As Status

Private Declare Sub OleCreatePictureIndirect Lib "oleaut32.dll" _
    (lpPictDesc As PICTDESC, riid As IID, _
    ByVal fOwn As Boolean, lplpvObj As Object)

Private Declare Function lstrcpyW Lib "kernel32" _
    (lpString1 As Any, lpString2 As Any) As Long

Private Declare Function lstrlenW Lib "kernel32" _
    (lpString As Any) As Long

Dim GdipToken As Long
Dim GdipInitialized As Boolean

Private Function Execute(ByVal lReturn As Status) As Status
    Dim lCurErr As Status
    If lReturn = Status.OK Then
        lCurErr = Status.OK
    Else
        lCurErr = lReturn
        MsgBox GdiErrorString(lReturn) & " GDI+ Error:" _
        & lReturn, vbOKOnly, "GDI Error"
    End If
    Execute = lCurErr
End Function

Private Function GdiErrorString(ByVal lError As Status) As String
    Dim s As String
    
    Select Case lError
    Case GenericError:              s = "Generic Error."
    Case InvalidParameter:          s = "Invalid Parameter."
    Case OutOfMemory:               s = "Out Of Memory."
    Case ObjectBusy:                s = "Object Busy."
    Case InsufficientBuffer:        s = "Insufficient Buffer."
    Case NotImplemented:            s = "Not Implemented."
    Case Win32Error:                s = "Win32 Error."
    Case WrongState:                s = "Wrong State."
    Case Aborted:                   s = "Aborted."
    Case FileNotFound:              s = "File Not Found."
    Case ValueOverflow:             s = "Value Overflow."
    Case AccessDenied:              s = "Access Denied."
    Case UnknownImageFormat:        s = "Unknown Image Format."
    Case FontFamilyNotFound:        s = "FontFamily Not Found."
    Case FontStyleNotFound:         s = "FontStyle Not Found."
    Case NotTrueTypeFont:           s = "Not TrueType Font."
    Case UnsupportedGdiplusVersion: s = "Unsupported Gdiplus Version."
    Case GdiplusNotInitialized:     s = "Gdiplus Not Initialized."
    Case PropertyNotFound:          s = "Property Not Found."
    Case PropertyNotSupported:      s = "Property Not Supported."
    Case Else:                      s = "Unknown GDI+ Error."
    End Select
    
    GdiErrorString = s
End Function

Private Function GetEncoderClsid(MimeType As String, _
    pClsid As GUID) As Boolean
    
    Dim num As Long
    Dim Size As Long
    Dim pImageCodecInfo() As ImageCodecInfo
    Dim j As Long
    Dim buffer As String
    
    Call GdipGetImageEncodersSize(num, Size)
    
    If (Size = 0) Then
        GetEncoderClsid = False
        Exit Function
    End If
    
    ReDim pImageCodecInfo(0 To Size \ Len(pImageCodecInfo(0)) - 1)
    
    Call GdipGetImageEncoders(num, Size, pImageCodecInfo(0))
    
    For j = 0 To num - 1
        
        buffer = _
        Space$(lstrlenW(ByVal pImageCodecInfo(j).MimeTypePtr))
        
        Call lstrcpyW(ByVal StrPtr(buffer), _
        ByVal pImageCodecInfo(j).MimeTypePtr)
        
        If (StrComp(buffer, MimeType, vbTextCompare) = 0) Then
            pClsid = pImageCodecInfo(j).Clsid
            Erase pImageCodecInfo
            GetEncoderClsid = True
            Exit Function
        End If
    Next j
    
    Erase pImageCodecInfo
    GetEncoderClsid = False
End Function

Private Function HandleToPicture(ByVal hGDIHandle As Long, _
    ByVal ObjectType As PictureTypeConstants, _
    Optional ByVal hPal As Long = 0) As StdPicture
    
    Dim tPictDesc As PICTDESC
    Dim IID_IPicture As IID
    Dim oPicture As IPicture
    
    With tPictDesc
        .cbSizeOfStruct = Len(tPictDesc)
        .picType = ObjectType
        .hgdiObj = hGDIHandle
        .hPalOrXYExt = hPal
    End With

    With IID_IPicture
        .Data1 = &H7BF80981
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(3) = &HAA
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With
    
    OleCreatePictureIndirect tPictDesc, IID_IPicture, _
    True, oPicture

    Set HandleToPicture = oPicture
    
End Function

Public Function LoadPicturePlus( _
    ByVal sFileName As String) As StdPicture
    
    Dim lBitmap As Long
    Dim hBitmap As Long
    

    If Execute(GdipCreateBitmapFromFile(StrPtr(sFileName), _
    lBitmap)) = OK Then
        
        If Execute(GdipCreateHBITMAPFromBitmap(lBitmap, _
        hBitmap, 0)) = OK Then
        
            Set LoadPicturePlus = HandleToPicture(hBitmap, _
            vbPicTypeBitmap)
        End If
    
        Call Execute(GdipDisposeImage(lBitmap))
    End If
End Function

Public Function GDISavePicture(pBox As StdPicture, ByVal sFileName As String, ByVal sMineType As String) As Boolean
Dim lBitmap As Long
Dim tPicEncoder As GUID
    'Check if GDI is ready
    If (Not GdipInitialized) Then Exit Function
    
    If Execute(GdipCreateBitmapFromHBITMAP(pBox.Handle, 0, lBitmap)) = OK Then
        If GetEncoderClsid(sMineType, tPicEncoder) = True Then
            If Execute(GdipSaveImageToFile(lBitmap, _
            StrPtr(sFileName), tPicEncoder, ByVal 0)) = OK Then
                GDISavePicture = True
            Else
                GDISavePicture = False
            End If
        Else
            GDISavePicture = False
        End If
        
        Call Execute(GdipDisposeImage(lBitmap))
        
    End If
End Function

Private Function StartUpGDIPlus(ByVal GdipVersion As Long) As Status
'GDI Start Up
Dim tGdipStartupInput As GDIPlusStartupInput
Dim tGdipStartupOutput As GdiplusStartupOutput
    
    tGdipStartupInput.GdiPlusVersion = GdipVersion
    StartUpGDIPlus = GdiplusStartup(GdipToken, _
    tGdipStartupInput, tGdipStartupOutput)
End Function

Private Function ShutDownGDIPlus() As Status
    'Close down GDI
    ShutDownGDIPlus = GdiplusShutdown(GdipToken)
End Function

Public Sub GDILoadPicture(ByVal Filename As String, pBox As PictureBox)
On Error GoTo GdiErr:
    
    'Check if GDI is ready
    If (GdipInitialized) Then
        'Load Picture
        pBox.Picture = LoadPicturePlus(Filename)
    End If
    
    Exit Sub
    'Error Flag
GdiErr:
    If Err Then Err.Clear
End Sub

Public Sub InitGDIPlus()
    GdipInitialized = False
    
    'Check if GDI can be loaded
    If Execute(StartUpGDIPlus(GdiPlusVersion)) = OK Then
        GdipInitialized = True
    Else
        MsgBox "GDI+ not inizialized.", vbOKOnly, "GDI Error"
    End If
    
End Sub

Public Sub GDIShutdown()
    If GdipInitialized = True Then
        Call Execute(ShutDownGDIPlus)
    End If
End Sub

