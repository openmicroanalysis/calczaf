Attribute VB_Name = "GDIPlus_Interface"
'***************************************************************************
' GDI+ Interface
' Added in summer 2018 by Tanner Helland (tannerhelland.com)
'
' Redistribution and use in source and binary forms, with or without modification, are permitted
' provided that the following conditions are met:
'
' - Redistributions of source code must retain the above copyright notice, this list of conditions
'    and the following disclaimer.
' - Redistributions in binary form are not required to reproduce the above copyright notice, but all
'    terms of the following disclaimer remain in effect.
'
' THIS SOFTWARE IS PROVIDED "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO,
' THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO
' EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL,
' EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
' SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF
' LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN
' ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
'
' These functions were developed with help from the following third-party code samples:
' Avery P's initial GDI+ deconstruction: http://planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=37541&lngWId=1
' Carles P.V.'s iBMP implementation: http://planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=42376&lngWId=1
' Many thanks to these individuals for their work on VB-compatible GDI+ interfaces.
'***************************************************************************

Option Explicit

Public Enum GP_Result
    GP_OK = 0
    GP_GenericError = 1
    GP_InvalidParameter = 2
    GP_OutOfMemory = 3
    GP_ObjectBusy = 4
    GP_InsufficientBuffer = 5
    GP_NotImplemented = 6
    GP_Win32Error = 7
    GP_WrongState = 8
    GP_Aborted = 9
    GP_FileNotFound = 10
    GP_ValueOverflow = 11
    GP_AccessDenied = 12
    GP_UnknownImageFormat = 13
    GP_FontFamilyNotFound = 14
    GP_FontStyleNotFound = 15
    GP_NotTrueTypeFont = 16
    GP_UnsupportedGDIPlusVersion = 17
    GP_GDIPlusNotInitialized = 18
    GP_PropertyNotFound = 19
    GP_PropertyNotSupported = 20
End Enum

#If False Then
    Private Const GP_OK = 0, GP_GenericError = 1, GP_InvalidParameter = 2, GP_OutOfMemory = 3, GP_ObjectBusy = 4, GP_InsufficientBuffer = 5, GP_NotImplemented = 6, GP_Win32Error = 7, GP_WrongState = 8, GP_Aborted = 9, GP_FileNotFound = 10, GP_ValueOverflow = 11, GP_AccessDenied = 12, GP_UnknownImageFormat = 13
    Private Const GP_FontFamilyNotFound = 14, GP_FontStyleNotFound = 15, GP_NotTrueTypeFont = 16, GP_UnsupportedGDIPlusVersion = 17, GP_GDIPlusNotInitialized = 18, GP_PropertyNotFound = 19, GP_PropertyNotSupported = 20
#End If

Private Type GDIPlusStartupInput
    GDIPlusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

Private Enum GP_DebugEventLevel
    GP_DebugEventLevelFatal = 0
    GP_DebugEventLevelWarning = 1
End Enum

#If False Then
    Private Const GP_DebugEventLevelFatal = 0, GP_DebugEventLevelWarning = 1
#End If

'Drawing-related enums

Public Enum GP_QualityMode      'Note that many other settings just wrap these default Quality Mode values
    GP_QM_Invalid = -1
    GP_QM_Default = 0
    GP_QM_Low = 1
    GP_QM_High = 2
End Enum

#If False Then
    Private Const GP_QM_Invalid = -1, GP_QM_Default = 0, GP_QM_Low = 1, GP_QM_High = 2
#End If

Public Enum GP_BitmapLockMode
    GP_BLM_Read = &H1
    GP_BLM_Write = &H2
    GP_BLM_UserInputBuf = &H4
End Enum

#If False Then
    Private Const GP_BLM_Read = &H1, GP_BLM_Write = &H2, GP_BLM_UserInputBuf = &H4
#End If

Public Enum GP_BrushType        'IMPORTANT NOTE!  This enum is *not* the same as PD's internal 2D brush modes!
    GP_BT_SolidColor = 0
    GP_BT_HatchFill = 1
    GP_BT_TextureFill = 2
    GP_BT_PathGradient = 3
    GP_BT_LinearGradient = 4
End Enum

#If False Then
    Private Const GP_BT_SolidColor = 0, GP_BT_HatchFill = 1, GP_BT_TextureFill = 2, GP_BT_PathGradient = 3, GP_BT_LinearGradient = 4
#End If

'Color adjustments are handled internally, at present, so we don't need to expose them to other objects
Private Enum GP_ColorAdjustType
    GP_CAT_Default = 0
    GP_CAT_Bitmap = 1
    GP_CAT_Brush = 2
    GP_CAT_Pen = 3
    GP_CAT_Text = 4
    GP_CAT_Count = 5
    GP_CAT_Any = 6
End Enum

#If False Then
    Private Const GP_CAT_Default = 0, GP_CAT_Bitmap = 1, GP_CAT_Brush = 2, GP_CAT_Pen = 3, GP_CAT_Text = 4, GP_CAT_Count = 5, GP_CAT_Any = 6
#End If

Private Enum GP_ColorMatrixFlags
    GP_CMF_Default = 0
    GP_CMF_SkipGrays = 1
    GP_CMF_AltGray = 2
End Enum

#If False Then
    Private Const GP_CMF_Default = 0, GP_CMF_SkipGrays = 1, GP_CMF_AltGray = 2
#End If

Public Enum GP_CombineMode
    GP_CM_Replace = 0
    GP_CM_Intersect = 1
    GP_CM_Union = 2
    GP_CM_Xor = 3
    GP_CM_Exclude = 4
    GP_CM_Complement = 5
End Enum

#If False Then
    Private Const GP_CM_Replace = 0, GP_CM_Intersect = 1, GP_CM_Union = 2, GP_CM_Xor = 3, GP_CM_Exclude = 4, GP_CM_Complement = 5
#End If

'Compositing mode is the closest GDI+ comes to offering "blend modes".  The default mode alpha-blends the source
' with the destination; "copy" mode overwrites the destination completely.
Public Enum GP_CompositingMode
    GP_CM_SourceOver = 0
    GP_CM_SourceCopy = 1
End Enum

#If False Then
    Private Const GP_CM_SourceOver = 0, GP_CM_SourceCopy = 1
#End If

'Alpha compositing qualities, which affects how GDI+ blends pixels.  Use with caution, as gamma-corrected blending
' yields non-inutitive results!
Public Enum GP_CompositingQuality
    GP_CQ_Invalid = GP_QM_Invalid
    GP_CQ_Default = GP_QM_Default
    GP_CQ_HighSpeed = GP_QM_Low
    GP_CQ_HighQuality = GP_QM_High
    GP_CQ_GammaCorrected = 3&
    GP_CQ_AssumeLinear = 4&
End Enum

#If False Then
    Private Const GP_CQ_Invalid = GP_QM_Invalid, GP_CQ_Default = GP_QM_Default, GP_CQ_HighSpeed = GP_QM_Low, GP_CQ_HighQuality = GP_QM_High, GP_CQ_GammaCorrected = 3&, GP_CQ_AssumeLinear = 4&
#End If

Public Enum GP_DashCap
    GP_DC_Flat = 0
    GP_DC_Square = 0     'This is not a typo; it's supplied as a convenience enum to match supported GP_LineCap values (which differentiate between flat and square, as they should)
    GP_DC_Round = 2
    GP_DC_Triangle = 3
End Enum

#If False Then
    Private Const GP_DC_Flat = 0, GP_DC_Square = 0, GP_DC_Round = 2, GP_DC_Triangle = 3
#End If

Public Enum GP_DashStyle
    GP_DS_Solid = 0&
    GP_DS_Dash = 1&
    GP_DS_Dot = 2&
    GP_DS_DashDot = 3&
    GP_DS_DashDotDot = 4&
    GP_DS_Custom = 5&
End Enum

#If False Then
    Private Const GP_DS_Solid = 0&, GP_DS_Dash = 1&, GP_DS_Dot = 2&, GP_DS_DashDot = 3&, GP_DS_DashDotDot = 4&, GP_DS_Custom = 5&
#End If

Public Enum GP_EncoderValueType
    GP_EVT_Byte = 1
    GP_EVT_ASCII = 2
    GP_EVT_Short = 3
    GP_EVT_Long = 4
    GP_EVT_Rational = 5
    GP_EVT_LongRange = 6
    GP_EVT_Undefined = 7
    GP_EVT_RationalRange = 8
    GP_EVT_Pointer = 9
End Enum

#If False Then
    Private Const GP_EVT_Byte = 1, GP_EVT_ASCII = 2, GP_EVT_Short = 3, GP_EVT_Long = 4, GP_EVT_Rational = 5, GP_EVT_LongRange = 6, GP_EVT_Undefined = 7, GP_EVT_RationalRange = 8, GP_EVT_Pointer = 9
#End If

Public Enum GP_EncoderValue
    GP_EV_ColorTypeCMYK = 0
    GP_EV_ColorTypeYCCK = 1
    GP_EV_CompressionLZW = 2
    GP_EV_CompressionCCITT3 = 3
    GP_EV_CompressionCCITT4 = 4
    GP_EV_CompressionRle = 5
    GP_EV_CompressionNone = 6
    GP_EV_ScanMethodInterlaced = 7
    GP_EV_ScanMethodNonInterlaced = 8
    GP_EV_VersionGif87 = 9
    GP_EV_VersionGif89 = 10
    GP_EV_RenderProgressive = 11
    GP_EV_RenderNonProgressive = 12
    GP_EV_TransformRotate90 = 13
    GP_EV_TransformRotate180 = 14
    GP_EV_TransformRotate270 = 15
    GP_EV_TransformFlipHorizontal = 16
    GP_EV_TransformFlipVertical = 17
    GP_EV_MultiFrame = 18
    GP_EV_LastFrame = 19
    GP_EV_Flush = 20
    GP_EV_FrameDimensionTime = 21
    GP_EV_FrameDimensionResolution = 22
    GP_EV_FrameDimensionPage = 23
    GP_EV_ColorTypeGray = 24
    GP_EV_ColorTypeRGB = 25
End Enum

#If False Then
    Private Const GP_EV_ColorTypeCMYK = 0, GP_EV_ColorTypeYCCK = 1, GP_EV_CompressionLZW = 2, GP_EV_CompressionCCITT3 = 3, GP_EV_CompressionCCITT4 = 4, GP_EV_CompressionRle = 5, GP_EV_CompressionNone = 6, GP_EV_ScanMethodInterlaced = 7, GP_EV_ScanMethodNonInterlaced = 8, GP_EV_VersionGif87 = 9, GP_EV_VersionGif89 = 10
    Private Const GP_EV_RenderProgressive = 11, GP_EV_RenderNonProgressive = 12, GP_EV_TransformRotate90 = 13, GP_EV_TransformRotate180 = 14, GP_EV_TransformRotate270 = 15, GP_EV_TransformFlipHorizontal = 16, GP_EV_TransformFlipVertical = 17, GP_EV_MultiFrame = 18, GP_EV_LastFrame = 19, GP_EV_Flush = 20
    Private Const GP_EV_FrameDimensionTime = 21, GP_EV_FrameDimensionResolution = 22, GP_EV_FrameDimensionPage = 23, GP_EV_ColorTypeGray = 24, GP_EV_ColorTypeRGB = 25
#End If

Public Enum GP_FillMode
    GP_FM_Alternate = 0&
    GP_FM_Winding = 1&
End Enum

#If False Then
    Private Const GP_FM_Alternate = 0&, GP_FM_Winding = 1&
#End If

Private Enum GP_FlushIntention
    GP_FI_Flush = 0&
    GP_FI_Sync = 1&
End Enum

#If False Then
    Private Const GP_FI_Flush = 0&, GP_FI_Sync = 1&
#End If

Public Enum GP_ImageFlags
    GP_IF_None = 0
    GP_IF_Scalable = &H1&
    GP_IF_HasAlpha = &H2&
    GP_IF_HasTranslucent = &H4&
    GP_IF_PartiallyScalable = &H8&
    GP_IF_ColorSpaceRGB = &H10&
    GP_IF_ColorSpaceCMYK = &H20&
    GP_IF_ColorSpaceGRAY = &H40&
    GP_IF_ColorSpaceYCBCR = &H80&
    GP_IF_ColorSpaceYCCK = &H100&
    GP_IF_HasRealDPI = &H1000&
    GP_IF_HasRealPixelSize = &H2000&
    GP_IF_ReadOnly = &H10000
    GP_IF_Caching = &H20000
End Enum

#If False Then
    Private Const GP_IF_None = 0, GP_IF_Scalable = &H1, GP_IF_HasAlpha = &H2, GP_IF_HasTranslucent = &H4, GP_IF_PartiallyScalable = &H8, GP_IF_ColorSpaceRGB = &H10, GP_IF_ColorSpaceCMYK = &H20, GP_IF_ColorSpaceGRAY = &H40, GP_IF_ColorSpaceYCBCR = &H80, GP_IF_ColorSpaceYCCK = &H100, GP_IF_HasRealDPI = &H1000, GP_IF_HasRealPixelSize = &H2000, GP_IF_ReadOnly = &H10000, GP_IF_Caching = &H20000
#End If

Public Enum GP_ImageFormat
    GP_IF_BMP = 0
    GP_IF_GIF = 1
    GP_IF_JPEG = 2
    GP_IF_PNG = 3
    GP_IF_TIFF = 4
End Enum

#If False Then
    Private Const GP_IF_BMP = 0, GP_IF_GIF = 1, GP_IF_JPEG = 2, GP_IF_PNG = 3, GP_IF_TIFF = 4
#End If

Public Enum GP_ImageType
    GP_IT_Unknown = 0
    GP_IT_Bitmap = 1
    GP_IT_Metafile = 2
End Enum

#If False Then
    Private Const GP_IT_Unknown = 0, GP_IT_Bitmap = 1, GP_IT_Metafile = 2
#End If

Public Enum GP_InterpolationMode
    GP_IM_Invalid = GP_QM_Invalid
    GP_IM_Default = GP_QM_Default
    GP_IM_LowQuality = GP_QM_Low
    GP_IM_HighQuality = GP_QM_High
    GP_IM_Bilinear = 3
    GP_IM_Bicubic = 4
    GP_IM_NearestNeighbor = 5
    GP_IM_HighQualityBilinear = 6
    GP_IM_HighQualityBicubic = 7
End Enum

#If False Then
    Private Const GP_IM_Invalid = GP_QM_Invalid, GP_IM_Default = GP_QM_Default, GP_IM_LowQuality = GP_QM_Low, GP_IM_HighQuality = GP_QM_High, GP_IM_Bilinear = 3, GP_IM_Bicubic = 4, GP_IM_NearestNeighbor = 5, GP_IM_HighQualityBilinear = 6, GP_IM_HighQualityBicubic = 7
#End If

Public Enum GP_LineCap
    GP_LC_Flat = 0&
    GP_LC_Square = 1&
    GP_LC_Round = 2&
    GP_LC_Triangle = 3&
    GP_LC_NoAnchor = &H10
    GP_LC_SquareAnchor = &H11
    GP_LC_RoundAnchor = &H12
    GP_LC_DiamondAnchor = &H13
    GP_LC_ArrowAnchor = &H14
    GP_LC_Custom = &HFF
End Enum

#If False Then
    Private Const GP_LC_Flat = 0, GP_LC_Square = 1, GP_LC_Round = 2, GP_LC_Triangle = 3, GP_LC_NoAnchor = &H10, GP_LC_SquareAnchor = &H11, GP_LC_RoundAnchor = &H12, GP_LC_DiamondAnchor = &H13, GP_LC_ArrowAnchor = &H14, GP_LC_Custom = &HFF
#End If

Public Enum GP_LineJoin
    GP_LJ_Miter = 0&
    GP_LJ_Bevel = 1&
    GP_LJ_Round = 2&
End Enum

#If False Then
    Private Const GP_LJ_Miter = 0&, GP_LJ_Bevel = 1&, GP_LJ_Round = 2&
#End If

Public Enum GP_MatrixOrder
    GP_MO_Prepend = 0&
    GP_MO_Append = 1&
End Enum

#If False Then
    Private Const GP_MO_Prepend = 0&, GP_MO_Append = 1&
#End If

'EMFs can be converted between various formats.  GDI+ prefers "EMF+", which supports GDI+ primitives as well
Public Enum GP_MetafileType
    GP_MT_Invalid = 0
    GP_MT_Wmf = 1
    GP_MT_WmfPlaceable = 2
    GP_MT_Emf = 3              'Old-style EMF consisting only of GDI commands
    GP_MT_EmfPlus = 4          'New-style EMF+ consisting only of GDI+ commands
    GP_MT_EmfDual = 5          'New-style EMF+ with GDI fallbacks for legacy rendering
End Enum

#If False Then
    Private Const GP_MT_Invalid = 0, GP_MT_Wmf = 1, GP_MT_WmfPlaceable = 2, GP_MT_Emf = 3, GP_MT_EmfPlus = 4, GP_MT_EmfDual = 5
#End If

Public Enum GP_PatternStyle
    GP_PS_Horizontal = 0
    GP_PS_Vertical = 1
    GP_PS_ForwardDiagonal = 2
    GP_PS_BackwardDiagonal = 3
    GP_PS_Cross = 4
    GP_PS_DiagonalCross = 5
    GP_PS_05Percent = 6
    GP_PS_10Percent = 7
    GP_PS_20Percent = 8
    GP_PS_25Percent = 9
    GP_PS_30Percent = 10
    GP_PS_40Percent = 11
    GP_PS_50Percent = 12
    GP_PS_60Percent = 13
    GP_PS_70Percent = 14
    GP_PS_75Percent = 15
    GP_PS_80Percent = 16
    GP_PS_90Percent = 17
    GP_PS_LightDownwardDiagonal = 18
    GP_PS_LightUpwardDiagonal = 19
    GP_PS_DarkDownwardDiagonal = 20
    GP_PS_DarkUpwardDiagonal = 21
    GP_PS_WideDownwardDiagonal = 22
    GP_PS_WideUpwardDiagonal = 23
    GP_PS_LightVertical = 24
    GP_PS_LightHorizontal = 25
    GP_PS_NarrowVertical = 26
    GP_PS_NarrowHorizontal = 27
    GP_PS_DarkVertical = 28
    GP_PS_DarkHorizontal = 29
    GP_PS_DashedDownwardDiagonal = 30
    GP_PS_DashedUpwardDiagonal = 31
    GP_PS_DashedHorizontal = 32
    GP_PS_DashedVertical = 33
    GP_PS_SmallConfetti = 34
    GP_PS_LargeConfetti = 35
    GP_PS_ZigZag = 36
    GP_PS_Wave = 37
    GP_PS_DiagonalBrick = 38
    GP_PS_HorizontalBrick = 39
    GP_PS_Weave = 40
    GP_PS_Plaid = 41
    GP_PS_Divot = 42
    GP_PS_DottedGrid = 43
    GP_PS_DottedDiamond = 44
    GP_PS_Shingle = 45
    GP_PS_Trellis = 46
    GP_PS_Sphere = 47
    GP_PS_SmallGrid = 48
    GP_PS_SmallCheckerBoard = 49
    GP_PS_LargeCheckerBoard = 50
    GP_PS_OutlinedDiamond = 51
    GP_PS_SolidDiamond = 52
End Enum

#If False Then
    Private Const GP_PS_Horizontal = 0, GP_PS_Vertical = 1, GP_PS_ForwardDiagonal = 2, GP_PS_BackwardDiagonal = 3, GP_PS_Cross = 4, GP_PS_DiagonalCross = 5, GP_PS_05Percent = 6, GP_PS_10Percent = 7, GP_PS_20Percent = 8, GP_PS_25Percent = 9, GP_PS_30Percent = 10, GP_PS_40Percent = 11, GP_PS_50Percent = 12, GP_PS_60Percent = 13, GP_PS_70Percent = 14, GP_PS_75Percent = 15, GP_PS_80Percent = 16, GP_PS_90Percent = 17, GP_PS_LightDownwardDiagonal = 18, GP_PS_LightUpwardDiagonal = 19, GP_PS_DarkDownwardDiagonal = 20, GP_PS_DarkUpwardDiagonal = 21, GP_PS_WideDownwardDiagonal = 22, GP_PS_WideUpwardDiagonal = 23, GP_PS_LightVertical = 24, GP_PS_LightHorizontal = 25
    Private Const GP_PS_NarrowVertical = 26, GP_PS_NarrowHorizontal = 27, GP_PS_DarkVertical = 28, GP_PS_DarkHorizontal = 29, GP_PS_DashedDownwardDiagonal = 30, GP_PS_DashedUpwardDiagonal = 31, GP_PS_DashedHorizontal = 32, GP_PS_DashedVertical = 33, GP_PS_SmallConfetti = 34, GP_PS_LargeConfetti = 35, GP_PS_ZigZag = 36, GP_PS_Wave = 37, GP_PS_DiagonalBrick = 38, GP_PS_HorizontalBrick = 39, GP_PS_Weave = 40, GP_PS_Plaid = 41, GP_PS_Divot = 42, GP_PS_DottedGrid = 43, GP_PS_DottedDiamond = 44, GP_PS_Shingle = 45, GP_PS_Trellis = 46, GP_PS_Sphere = 47, GP_PS_SmallGrid = 48, GP_PS_SmallCheckerBoard = 49, GP_PS_LargeCheckerBoard = 50
    Private Const GP_PS_OutlinedDiamond = 51, GP_PS_SolidDiamond = 52
#End If

Public Enum GP_PenAlignment
    GP_PA_Center = 0&
    GP_PA_Inset = 1&
End Enum

#If False Then
    Private Const GP_PA_Center = 0&, GP_PA_Inset = 1&
#End If

'GDI+ pixel format IDs use a bitfield system:
' [0, 7] = format index
' [8, 15] = pixel size (in bits)
' [16, 23] = flags
' [24, 31] = reserved (currently unused)

'Note also that pixel format is *not* 100% reliable.  Behavior differs between OSes, even for the "same"
' major GDI+ version.  (See http://stackoverflow.com/questions/5065371/how-to-identify-cmyk-images-in-asp-net-using-c-sharp)
Public Enum GP_PixelFormat
    GP_PF_Indexed = &H10000         'Image uses a palette to define colors
    GP_PF_GDI = &H20000             'Is a format supported by GDI
    GP_PF_Alpha = &H40000           'Alpha channel present
    GP_PF_PreMultAlpha = &H80000    'Alpha is premultiplied (not always correct; manual verification should be used)
    GP_PF_HDR = &H100000            'High bit-depth colors are in use (e.g. 48-bpp or 64-bpp; behavior is unpredictable on XP)
    GP_PF_Canonical = &H200000      'Canonical formats: 32bppARGB, 32bppPARGB, 64bppARGB, 64bppPARGB
    
    GP_PF_32bppCMYK = &H200F        'CMYK is never returned on XP or Vista; ImageFlags can be checked as a failsafe
                                    ' (Conversely, ImageFlags are unreliable on Win 7 - this is the shit we deal with
                                    '  as Windows developers!)
    
    GP_PF_1bppIndexed = &H30101
    GP_PF_4bppIndexed = &H30402
    GP_PF_8bppIndexed = &H30803
    GP_PF_16bppGreyscale = &H101004
    GP_PF_16bppRGB555 = &H21005
    GP_PF_16bppRGB565 = &H21006
    GP_PF_16bppARGB1555 = &H61007
    GP_PF_24bppRGB = &H21808
    GP_PF_32bppRGB = &H22009
    GP_PF_32bppARGB = &H26200A
    GP_PF_32bppPARGB = &HE200B
    GP_PF_48bppRGB = &H10300C
    GP_PF_64bppARGB = &H34400D
    GP_PF_64bppPARGB = &H1C400E
End Enum

#If False Then
    Private Const GP_PF_Indexed = &H10000, GP_PF_GDI = &H20000, GP_PF_Alpha = &H40000, GP_PF_PreMultAlpha = &H80000, GP_PF_HDR = &H100000, GP_PF_Canonical = &H200000, GP_PF_32bppCMYK = &H200F
    Private Const GP_PF_1bppIndexed = &H30101, GP_PF_4bppIndexed = &H30402, GP_PF_8bppIndexed = &H30803, GP_PF_16bppGreyscale = &H101004, GP_PF_16bppRGB555 = &H21005, GP_PF_16bppRGB565 = &H21006
    Private Const GP_PF_16bppARGB1555 = &H61007, GP_PF_24bppRGB = &H21808, GP_PF_32bppRGB = &H22009, GP_PF_32bppARGB = &H26200A, GP_PF_32bppPARGB = &HE200B, GP_PF_48bppRGB = &H10300C, GP_PF_64bppARGB = &H34400D, GP_PF_64bppPARGB = &H1C400E
#End If

'PixelOffsetMode controls how GDI+ calculates positioning.  Normally, each a pixel is treated as a unit square that covers
' the area between [0, 0] and [1, 1].  However, for point-based objects like paths, GDI+ can treat coordinates as if they
' are centered over [0.5, 0.5] offsets within each pixel.  This typically yields prettier path renders, at some consequence
' to rendering performance.  (See http://drilian.com/2008/11/25/understanding-half-pixel-and-half-texel-offsets/)
Public Enum GP_PixelOffsetMode
    GP_POM_Invalid = GP_QM_Invalid
    GP_POM_Default = GP_QM_Default
    GP_POM_HighSpeed = GP_QM_Low
    GP_POM_HighQuality = GP_QM_High
    GP_POM_None = 3&
    GP_POM_Half = 4&
End Enum

#If False Then
    Private Const GP_POM_Invalid = QualityModeInvalid, GP_POM_Default = QualityModeDefault, GP_POM_HighSpeed = QualityModeLow, GP_POM_HighQuality = QualityModeHigh, GP_POM_None = 3, GP_POM_Half = 4
#End If

'Property tags describe image metadata.  Metadata is very complicated to read and/or write, because tags are encoded
' in a variety of ways.  Refer to https://msdn.microsoft.com/en-us/library/ms534416(v=vs.85).aspx for details.
' pd2D uses these sparingly; do not expect it to perform full metadata preservation.
Public Enum GP_PropertyTag
'    GP_PT_Artist = &H13B&
'    GP_PT_BitsPerSample = &H102&
'    GP_PT_CellHeight = &H109&
'    GP_PT_CellWidth = &H108&
'    GP_PT_ChrominanceTable = &H5091&
'    GP_PT_ColorMap = &H140&
'    GP_PT_ColorTransferFunction = &H501A&
'    GP_PT_Compression = &H103&
'    GP_PT_Copyright = &H8298&
'    GP_PT_DateTime = &H132&
'    GP_PT_DocumentName = &H10D&
'    GP_PT_DotRange = &H150&
'    GP_PT_EquipMake = &H10F&
'    GP_PT_EquipModel = &H110&
'    GP_PT_ExifAperture = &H9202&
'    GP_PT_ExifBrightness = &H9203&
'    GP_PT_ExifCfaPattern = &HA302&
'    GP_PT_ExifColorSpace = &HA001&
'    GP_PT_ExifCompBPP = &H9102&
'    GP_PT_ExifCompConfig = &H9101&
'    GP_PT_ExifDTDigitized = &H9004&
'    GP_PT_ExifDTDigSS = &H9292&
'    GP_PT_ExifDTOrig = &H9003&
'    GP_PT_ExifDTOrigSS = &H9291&
'    GP_PT_ExifDTSubsec = &H9290&
'    GP_PT_ExifExposureBias = &H9204&
'    GP_PT_ExifExposureIndex = &HA215&
'    GP_PT_ExifExposureProg = &H8822&
'    GP_PT_ExifExposureTime = &H829A&
'    GP_PT_ExifFileSource = &HA300&
'    GP_PT_ExifFlash = &H9209&
'    GP_PT_ExifFlashEnergy = &HA20B&
'    GP_PT_ExifFNumber = &H829D&
'    GP_PT_ExifFocalLength = &H920A&
'    GP_PT_ExifFocalResUnit = &HA210&
'    GP_PT_ExifFocalXRes = &HA20E&
'    GP_PT_ExifFocalYRes = &HA20F&
'    GP_PT_ExifFPXVer = &HA000&
'    GP_PT_ExifIFD = &H8769&
'    GP_PT_ExifInterop = &HA005&
'    GP_PT_ExifISOSpeed = &H8827&
'    GP_PT_ExifLightSource = &H9208&
'    GP_PT_ExifMakerNote = &H927C&
'    GP_PT_ExifMaxAperture = &H9205&
'    GP_PT_ExifMeteringMode = &H9207&
'    GP_PT_ExifOECF = &H8828&
'    GP_PT_ExifPixXDim = &HA002&
'    GP_PT_ExifPixYDim = &HA003&
'    GP_PT_ExifRelatedWav = &HA004&
'    GP_PT_ExifSceneType = &HA301&
'    GP_PT_ExifSensingMethod = &HA217&
'    GP_PT_ExifShutterSpeed = &H9201&
'    GP_PT_ExifSpatialFR = &HA20C&
'    GP_PT_ExifSpectralSense = &H8824&
'    GP_PT_ExifSubjectDist = &H9206&
'    GP_PT_ExifSubjectLoc = &HA214&
'    GP_PT_ExifUserComment = &H9286&
'    GP_PT_ExifVer = &H9000&
'    GP_PT_ExtraSamples = &H152&
'    GP_PT_FillOrder = &H10A&
'    GP_PT_FrameDelay = &H5100&
'    GP_PT_FreeByteCounts = &H121&
'    GP_PT_FreeOffset = &H120&
'    GP_PT_Gamma = &H301&
'    GP_PT_GlobalPalette = &H5102&
'    GP_PT_GpsAltitude = &H6&
'    GP_PT_GpsAltitudeRef = &H5&
'    GP_PT_GpsDestBear = &H18&
'    GP_PT_GpsDestBearRef = &H17&
'    GP_PT_GpsDestDist = &H1A&
'    GP_PT_GpsDestDistRef = &H19&
'    GP_PT_GpsDestLat = &H14&
'    GP_PT_GpsDestLatRef = &H13&
'    GP_PT_GpsDestLong = &H16&
'    GP_PT_GpsDestLongRef = &H15&
'    GP_PT_GpsGpsDop = &HB&
'    GP_PT_GpsGpsMeasureMode = &HA&
'    GP_PT_GpsGpsSatellites = &H8&
'    GP_PT_GpsGpsStatus = &H9&
'    GP_PT_GpsGpsTime = &H7&
'    GP_PT_GpsIFD = &H8825&
'    GP_PT_GpsImgDir = &H11&
'    GP_PT_GpsImgDirRef = &H10&
'    GP_PT_GpsLatitude = &H2&
'    GP_PT_GpsLatitudeRef = &H1&
'    GP_PT_GpsLongitude = &H4&
'    GP_PT_GpsLongitudeRef = &H3&
'    GP_PT_GpsMapDatum = &H12&
'    GP_PT_GpsSpeed = &HD&
'    GP_PT_GpsSpeedRef = &HC&
'    GP_PT_GpsTrack = &HF&
'    GP_PT_GpsTrackRef = &HE&
'    GP_PT_GpsVer = &H0&
'    GP_PT_GrayResponseCurve = &H123&
'    GP_PT_GrayResponseUnit = &H122&
'    GP_PT_GridSize = &H5011&
'    GP_PT_HalftoneDegree = &H500C&
'    GP_PT_HalftoneHints = &H141&
'    GP_PT_HalftoneLPI = &H500A&
'    GP_PT_HalftoneLPIUnit = &H500B&
'    GP_PT_HalftoneMisc = &H500E&
'    GP_PT_HalftoneScreen = &H500F&
'    GP_PT_HalftoneShape = &H500D&
'    GP_PT_HostComputer = &H13C&
    GP_PT_ICCProfile = &H8773&
    GP_PT_ICCProfileDescriptor = &H302&
'    GP_PT_ImageDescription = &H10E&
'    GP_PT_ImageHeight = &H101&
'    GP_PT_ImageTitle = &H320&
'    GP_PT_ImageWidth = &H100&
'    GP_PT_IndexBackground = &H5103&
'    GP_PT_IndexTransparent = &H5104&
'    GP_PT_InkNames = &H14D&
'    GP_PT_InkSet = &H14C&
'    GP_PT_JPEGACTables = &H209&
'    GP_PT_JPEGDCTables = &H208&
'    GP_PT_JPEGInterFormat = &H201&
'    GP_PT_JPEGInterLength = &H202&
'    GP_PT_JPEGLosslessPredictors = &H205&
'    GP_PT_JPEGPointTransforms = &H206&
'    GP_PT_JPEGProc = &H200&
'    GP_PT_JPEGQTables = &H207&
'    GP_PT_JPEGQuality = &H5010&
'    GP_PT_JPEGRestartInterval = &H203&
'    GP_PT_LoopCount = &H5101&
'    GP_PT_LuminanceTable = &H5090&
'    GP_PT_MaxSampleValue = &H119&
'    GP_PT_MinSampleValue = &H118&
'    GP_PT_NewSubfileType = &HFE&
'    GP_PT_NumberOfInks = &H14E&
    GP_PT_Orientation = &H112&
    GP_PT_PageName = &H11D&
    GP_PT_PageNumber = &H129&
'    GP_PT_PaletteHistogram = &H5113&
'    GP_PT_PhotometricInterp = &H106&
'    GP_PT_PixelPerUnitX = &H5111&
'    GP_PT_PixelPerUnitY = &H5112&
'    GP_PT_PixelUnit = &H5110&
'    GP_PT_PlanarConfig = &H11C&
'    GP_PT_Predictor = &H13D&
'    GP_PT_PrimaryChromaticities = &H13F&
'    GP_PT_PrintFlags = &H5005&
'    GP_PT_PrintFlagsBleedWidth = &H5008&
'    GP_PT_PrintFlagsBleedWidthScale = &H5009&
'    GP_PT_PrintFlagsCrop = &H5007&
'    GP_PT_PrintFlagsVersion = &H5006&
'    GP_PT_REFBlackWhite = &H214&
'    GP_PT_ResolutionUnit = &H128&
'    GP_PT_ResolutionXLengthUnit = &H5003&
'    GP_PT_ResolutionXUnit = &H5001&
'    GP_PT_ResolutionYLengthUnit = &H5004&
'    GP_PT_ResolutionYUnit = &H5002&
'    GP_PT_RowsPerStrip = &H116&
'    GP_PT_SampleFormat = &H153&
'    GP_PT_SamplesPerPixel = &H115&
'    GP_PT_SMaxSampleValue = &H155&
'    GP_PT_SMinSampleValue = &H154&
'    GP_PT_SoftwareUsed = &H131&
'    GP_PT_SRGBRenderingIntent = &H303&
'    GP_PT_StripBytesCount = &H117&
'    GP_PT_StripOffsets = &H111&
'    GP_PT_SubfileType = &HFF&
'    GP_PT_T4Option = &H124&
'    GP_PT_T6Option = &H125&
'    GP_PT_TargetPrinter = &H151&
'    GP_PT_ThreshHolding = &H107&
'    GP_PT_ThumbnailArtist = &H5034&
'    GP_PT_ThumbnailBitsPerSample = &H5022&
'    GP_PT_ThumbnailColorDepth = &H5015&
'    GP_PT_ThumbnailCompressedSize = &H5019&
'    GP_PT_ThumbnailCompression = &H5023&
'    GP_PT_ThumbnailCopyRight = &H503B&
'    GP_PT_ThumbnailData = &H501B&
'    GP_PT_ThumbnailDateTime = &H5033&
'    GP_PT_ThumbnailEquipMake = &H5026&
'    GP_PT_ThumbnailEquipModel = &H5027&
'    GP_PT_ThumbnailFormat = &H5012&
'    GP_PT_ThumbnailHeight = &H5014&
'    GP_PT_ThumbnailImageDescription = &H5025&
'    GP_PT_ThumbnailImageHeight = &H5021&
'    GP_PT_ThumbnailImageWidth = &H5020&
'    GP_PT_ThumbnailOrientation = &H5029&
'    GP_PT_ThumbnailPhotometricInterp = &H5024&
'    GP_PT_ThumbnailPlanarConfig = &H502F&
'    GP_PT_ThumbnailPlanes = &H5016&
'    GP_PT_ThumbnailPrimaryChromaticities = &H5036&
'    GP_PT_ThumbnailRawBytes = &H5017&
'    GP_PT_ThumbnailRefBlackWhite = &H503A&
'    GP_PT_ThumbnailResolutionUnit = &H5030&
'    GP_PT_ThumbnailResolutionX = &H502D&
'    GP_PT_ThumbnailResolutionY = &H502E&
'    GP_PT_ThumbnailRowsPerStrip = &H502B&
'    GP_PT_ThumbnailSamplesPerPixel = &H502A&
'    GP_PT_ThumbnailSize = &H5018&
'    GP_PT_ThumbnailSoftwareUsed = &H5032&
'    GP_PT_ThumbnailStripBytesCount = &H502C&
'    GP_PT_ThumbnailStripOffsets = &H5028&
'    GP_PT_ThumbnailTransferFunction = &H5031&
'    GP_PT_ThumbnailWhitePoint = &H5035&
'    GP_PT_ThumbnailWidth = &H5013&
'    GP_PT_ThumbnailYCbCrCoefficients = &H5037&
'    GP_PT_ThumbnailYCbCrPositioning = &H5039&
'    GP_PT_ThumbnailYCbCrSubsampling = &H5038&
'    GP_PT_TileByteCounts = &H145&
'    GP_PT_TileLength = &H143&
'    GP_PT_TileOffset = &H144&
'    GP_PT_TileWidth = &H142&
'    GP_PT_TransferFunction = &H12D&
'    GP_PT_TransferRange = &H156&
'    GP_PT_WhitePoint = &H13E&
'    GP_PT_XPosition = &H11E&
    GP_PT_XResolution = &H11A&
'    GP_PT_YCbCrCoefficients = &H211&
'    GP_PT_YCbCrPositioning = &H213&
'    GP_PT_YCbCrSubsampling = &H212&
'    GP_PT_YPosition = &H11F&
    GP_PT_YResolution = &H11B&
End Enum

#If False Then
    Private Const GP_PT_Artist = &H13B, GP_PT_BitsPerSample = &H102, GP_PT_CellHeight = &H109, GP_PT_CellWidth = &H108, GP_PT_ChrominanceTable = &H5091, GP_PT_ColorMap = &H140, GP_PT_ColorTransferFunction = &H501A, GP_PT_Compression = &H103, GP_PT_Copyright = &H8298, GP_PT_DateTime = &H132, GP_PT_DocumentName = &H10D, GP_PT_DotRange = &H150, GP_PT_EquipMake = &H10F, GP_PT_EquipModel = &H110, GP_PT_ExifAperture = &H9202, GP_PT_ExifBrightness = &H9203, GP_PT_ExifCfaPattern = &HA302, GP_PT_ExifColorSpace = &HA001
    Private Const GP_PT_ExifCompBPP = &H9102, GP_PT_ExifCompConfig = &H9101, GP_PT_ExifDTDigitized = &H9004, GP_PT_ExifDTDigSS = &H9292, GP_PT_ExifDTOrig = &H9003, GP_PT_ExifDTOrigSS = &H9291, GP_PT_ExifDTSubsec = &H9290, GP_PT_ExifExposureBias = &H9204, GP_PT_ExifExposureIndex = &HA215, GP_PT_ExifExposureProg = &H8822, GP_PT_ExifExposureTime = &H829A, GP_PT_ExifFileSource = &HA300, GP_PT_ExifFlash = &H9209, GP_PT_ExifFlashEnergy = &HA20B, GP_PT_ExifFNumber = &H829D, GP_PT_ExifFocalLength = &H920A
    Private Const GP_PT_ExifFocalResUnit = &HA210, GP_PT_ExifFocalXRes = &HA20E, GP_PT_ExifFocalYRes = &HA20F, GP_PT_ExifFPXVer = &HA000, GP_PT_ExifIFD = &H8769, GP_PT_ExifInterop = &HA005, GP_PT_ExifISOSpeed = &H8827, GP_PT_ExifLightSource = &H9208, GP_PT_ExifMakerNote = &H927C, GP_PT_ExifMaxAperture = &H9205, GP_PT_ExifMeteringMode = &H9207, GP_PT_ExifOECF = &H8828, GP_PT_ExifPixXDim = &HA002, GP_PT_ExifPixYDim = &HA003, GP_PT_ExifRelatedWav = &HA004, GP_PT_ExifSceneType = &HA301
    Private Const GP_PT_ExifSensingMethod = &HA217, GP_PT_ExifShutterSpeed = &H9201, GP_PT_ExifSpatialFR = &HA20C, GP_PT_ExifSpectralSense = &H8824, GP_PT_ExifSubjectDist = &H9206, GP_PT_ExifSubjectLoc = &HA214, GP_PT_ExifUserComment = &H9286, GP_PT_ExifVer = &H9000, GP_PT_ExtraSamples = &H152, GP_PT_FillOrder = &H10A, GP_PT_FrameDelay = &H5100, GP_PT_FreeByteCounts = &H121, GP_PT_FreeOffset = &H120, GP_PT_Gamma = &H301, GP_PT_GlobalPalette = &H5102, GP_PT_GpsAltitude = &H6
    Private Const GP_PT_GpsAltitudeRef = &H5, GP_PT_GpsDestBear = &H18, GP_PT_GpsDestBearRef = &H17, GP_PT_GpsDestDist = &H1A, GP_PT_GpsDestDistRef = &H19, GP_PT_GpsDestLat = &H14, GP_PT_GpsDestLatRef = &H13, GP_PT_GpsDestLong = &H16, GP_PT_GpsDestLongRef = &H15, GP_PT_GpsGpsDop = &HB, GP_PT_GpsGpsMeasureMode = &HA, GP_PT_GpsGpsSatellites = &H8, GP_PT_GpsGpsStatus = &H9, GP_PT_GpsGpsTime = &H7, GP_PT_GpsIFD = &H8825, GP_PT_GpsImgDir = &H11, GP_PT_GpsImgDirRef = &H10, GP_PT_GpsLatitude = &H2
    Private Const GP_PT_GpsLatitudeRef = &H1, GP_PT_GpsLongitude = &H4, GP_PT_GpsLongitudeRef = &H3, GP_PT_GpsMapDatum = &H12, GP_PT_GpsSpeed = &HD, GP_PT_GpsSpeedRef = &HC, GP_PT_GpsTrack = &HF, GP_PT_GpsTrackRef = &HE, GP_PT_GpsVer = &H0, GP_PT_GrayResponseCurve = &H123, GP_PT_GrayResponseUnit = &H122, GP_PT_GridSize = &H5011, GP_PT_HalftoneDegree = &H500C, GP_PT_HalftoneHints = &H141, GP_PT_HalftoneLPI = &H500A, GP_PT_HalftoneLPIUnit = &H500B, GP_PT_HalftoneMisc = &H500E, GP_PT_HalftoneScreen = &H500F
    Private Const GP_PT_HalftoneShape = &H500D, GP_PT_HostComputer = &H13C, GP_PT_ICCProfile = &H8773, GP_PT_ICCProfileDescriptor = &H302, GP_PT_ImageDescription = &H10E, GP_PT_ImageHeight = &H101, GP_PT_ImageTitle = &H320, GP_PT_ImageWidth = &H100, GP_PT_IndexBackground = &H5103, GP_PT_IndexTransparent = &H5104, GP_PT_InkNames = &H14D, GP_PT_InkSet = &H14C, GP_PT_JPEGACTables = &H209, GP_PT_JPEGDCTables = &H208, GP_PT_JPEGInterFormat = &H201, GP_PT_JPEGInterLength = &H202, GP_PT_JPEGLosslessPredictors = &H205
    Private Const GP_PT_JPEGPointTransforms = &H206, GP_PT_JPEGProc = &H200, GP_PT_JPEGQTables = &H207, GP_PT_JPEGQuality = &H5010, GP_PT_JPEGRestartInterval = &H203, GP_PT_LoopCount = &H5101, GP_PT_LuminanceTable = &H5090, GP_PT_MaxSampleValue = &H119, GP_PT_MinSampleValue = &H118, GP_PT_NewSubfileType = &HFE, GP_PT_NumberOfInks = &H14E, GP_PT_Orientation = &H112, GP_PT_PageName = &H11D, GP_PT_PageNumber = &H129, GP_PT_PaletteHistogram = &H5113, GP_PT_PhotometricInterp = &H106, GP_PT_PixelPerUnitX = &H5111
    Private Const GP_PT_PixelPerUnitY = &H5112, GP_PT_PixelUnit = &H5110, GP_PT_PlanarConfig = &H11C, GP_PT_Predictor = &H13D, GP_PT_PrimaryChromaticities = &H13F, GP_PT_PrintFlags = &H5005, GP_PT_PrintFlagsBleedWidth = &H5008, GP_PT_PrintFlagsBleedWidthScale = &H5009, GP_PT_PrintFlagsCrop = &H5007, GP_PT_PrintFlagsVersion = &H5006, GP_PT_REFBlackWhite = &H214, GP_PT_ResolutionUnit = &H128, GP_PT_ResolutionXLengthUnit = &H5003, GP_PT_ResolutionXUnit = &H5001, GP_PT_ResolutionYLengthUnit = &H5004
    Private Const GP_PT_ResolutionYUnit = &H5002, GP_PT_RowsPerStrip = &H116, GP_PT_SampleFormat = &H153, GP_PT_SamplesPerPixel = &H115, GP_PT_SMaxSampleValue = &H155, GP_PT_SMinSampleValue = &H154, GP_PT_SoftwareUsed = &H131, GP_PT_SRGBRenderingIntent = &H303, GP_PT_StripBytesCount = &H117, GP_PT_StripOffsets = &H111, GP_PT_SubfileType = &HFF, GP_PT_T4Option = &H124, GP_PT_T6Option = &H125, GP_PT_TargetPrinter = &H151, GP_PT_ThreshHolding = &H107, GP_PT_ThumbnailArtist = &H5034, GP_PT_ThumbnailBitsPerSample = &H5022
    Private Const GP_PT_ThumbnailColorDepth = &H5015, GP_PT_ThumbnailCompressedSize = &H5019, GP_PT_ThumbnailCompression = &H5023, GP_PT_ThumbnailCopyRight = &H503B, GP_PT_ThumbnailData = &H501B, GP_PT_ThumbnailDateTime = &H5033, GP_PT_ThumbnailEquipMake = &H5026, GP_PT_ThumbnailEquipModel = &H5027, GP_PT_ThumbnailFormat = &H5012, GP_PT_ThumbnailHeight = &H5014, GP_PT_ThumbnailImageDescription = &H5025, GP_PT_ThumbnailImageHeight = &H5021, GP_PT_ThumbnailImageWidth = &H5020, GP_PT_ThumbnailOrientation = &H5029, GP_PT_ThumbnailPhotometricInterp = &H5024
    Private Const GP_PT_ThumbnailPlanarConfig = &H502F, GP_PT_ThumbnailPlanes = &H5016, GP_PT_ThumbnailPrimaryChromaticities = &H5036, GP_PT_ThumbnailRawBytes = &H5017, GP_PT_ThumbnailRefBlackWhite = &H503A, GP_PT_ThumbnailResolutionUnit = &H5030, GP_PT_ThumbnailResolutionX = &H502D, GP_PT_ThumbnailResolutionY = &H502E, GP_PT_ThumbnailRowsPerStrip = &H502B, GP_PT_ThumbnailSamplesPerPixel = &H502A, GP_PT_ThumbnailSize = &H5018, GP_PT_ThumbnailSoftwareUsed = &H5032, GP_PT_ThumbnailStripBytesCount = &H502C, GP_PT_ThumbnailStripOffsets = &H5028
    Private Const GP_PT_ThumbnailTransferFunction = &H5031, GP_PT_ThumbnailWhitePoint = &H5035, GP_PT_ThumbnailWidth = &H5013, GP_PT_ThumbnailYCbCrCoefficients = &H5037, GP_PT_ThumbnailYCbCrPositioning = &H5039, GP_PT_ThumbnailYCbCrSubsampling = &H5038, GP_PT_TileByteCounts = &H145, GP_PT_TileLength = &H143, GP_PT_TileOffset = &H144, GP_PT_TileWidth = &H142, GP_PT_TransferFunction = &H12D, GP_PT_TransferRange = &H156, GP_PT_WhitePoint = &H13E, GP_PT_XPosition = &H11E, GP_PT_XResolution = &H11A, GP_PT_YCbCrCoefficients = &H211
    Private Const GP_PT_YCbCrPositioning = &H213, GP_PT_YCbCrSubsampling = &H212, GP_PT_YPosition = &H11F, GP_PT_YResolution = &H11B
#End If

Private Enum GP_PropertyTagType
    GP_PTT_Byte = 1
    GP_PTT_ASCII = 2
    GP_PTT_Short = 3
    GP_PTT_Long = 4
    GP_PTT_Rational = 5
    GP_PTT_Undefined = 7
    GP_PTT_SLONG = 9
    GP_PTT_SRational = 10
End Enum

#If False Then
    Private Const GP_PTT_Byte = 1, GP_PTT_ASCII = 2, GP_PTT_Short = 3, GP_PTT_Long = 4, GP_PTT_Rational = 5, GP_PTT_Undefined = 7, GP_PTT_SLONG = 9, GP_PTT_SRational = 10
#End If

Public Enum GP_RotateFlip
    GP_RF_NoneFlipNone = 0
    GP_RF_90FlipNone = 1
    GP_RF_180FlipNone = 2
    GP_RF_270FlipNone = 3
    GP_RF_NoneFlipX = 4
    GP_RF_90FlipX = 5
    GP_RF_180FlipX = 6
    GP_RF_270FlipX = 7
    GP_RF_NoneFlipY = GP_RF_180FlipX
    GP_RF_90FlipY = GP_RF_270FlipX
    GP_RF_180FlipY = GP_RF_NoneFlipX
    GP_RF_270FlipY = GP_RF_90FlipX
    GP_RF_NoneFlipXY = GP_RF_180FlipNone
    GP_RF_90FlipXY = GP_RF_270FlipNone
    GP_RF_180FlipXY = GP_RF_NoneFlipNone
    GP_RF_270FlipXY = GP_RF_90FlipNone
End Enum

#If False Then
    Private Const GP_RF_NoneFlipNone = 0, GP_RF_90FlipNone = 1, GP_RF_180FlipNone = 2, GP_RF_270FlipNone = 3, GP_RF_NoneFlipX = 4, GP_RF_90FlipX = 5, GP_RF_180FlipX = 6, GP_RF_270FlipX = 7, GP_RF_NoneFlipY = GP_RF_180FlipX
    Private Const GP_RF_90FlipY = GP_RF_270FlipX, GP_RF_180FlipY = GP_RF_NoneFlipX, GP_RF_270FlipY = GP_RF_90FlipX, GP_RF_NoneFlipXY = GP_RF_180FlipNone, GP_RF_90FlipXY = GP_RF_270FlipNone, GP_RF_180FlipXY = GP_RF_NoneFlipNone, GP_RF_270FlipXY = GP_RF_90FlipNone
#End If

Public Enum GP_SmoothingMode
    GP_SM_Invalid = GP_QM_Invalid
    GP_SM_Default = GP_QM_Default
    GP_SM_HighSpeed = GP_QM_Low
    GP_SM_HighQuality = GP_QM_High
    GP_SM_None = 3&
    GP_SM_Antialias = 4&
End Enum

#If False Then
    Private Const GP_SM_Invalid = GP_QM_Invalid, GP_SM_Default = GP_QM_Default, GP_SM_HighSpeed = GP_QM_Low, GP_SM_HighQuality = GP_QM_High, GP_SM_None = 3, GP_SM_Antialias = 4
#End If

Public Enum GP_Unit
    GP_U_World = 0&
    GP_U_Display = 1&
    GP_U_Pixel = 2&
    GP_U_Point = 3&
    GP_U_Inch = 4&
    GP_U_Document = 5&
    GP_U_Millimeter = 6&
End Enum

#If False Then
    Private Const GP_U_World = 0, GP_U_Display = 1, GP_U_Pixel = 2, GP_U_Point = 3, GP_U_Inch = 4, GP_U_Document = 5, GP_U_Millimeter = 6
#End If

Public Enum GP_WrapMode
    GP_WM_Tile = 0
    GP_WM_TileFlipX = 1
    GP_WM_TileFlipY = 2
    GP_WM_TileFlipXY = 3
    GP_WM_Clamp = 4
End Enum

#If False Then
    Private Const GP_WM_Tile = 0, GP_WM_TileFlipX = 1, GP_WM_TileFlipY = 2, GP_WM_TileFlipXY = 3, GP_WM_Clamp = 4
#End If

Public Type PointFloat
    x As Single
    Y As Single
End Type

Private Type PointLong
    x As Long
    Y As Long
End Type

Private Type RECTF
    Left As Single
    Top As Single
    Width As Single
    Height As Single
End Type

Private Type RECTL
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type RectL_WH
    Left As Long
    Top As Long
    Width As Long
    Height As Long
End Type

'GDI+ uses a modified bitmap struct when performing things like raster format conversions
Public Type GP_BitmapData
    BD_Width As Long
    BD_Height As Long
    BD_Stride As Long
    BD_PixelFormat As GP_PixelFormat
    BD_Scan0 As Long
    BD_Reserved As Long
End Type

'GDI interop is made easier by declaring a few GDI-specific structs
Private Type BITMAPINFOHEADER
    Size As Long
    Width As Long
    Height As Long
    Planes As Integer
    BitCount As Integer
    Compression As Long
    ImageSize As Long
    XPelsPerMeter As Long
    YPelsPerMeter As Long
    Colorused As Long
    ColorImportant As Long
End Type

Private Type RGBQuad
    Blue As Byte
    Green As Byte
    Red As Byte
    Alpha As Byte
End Type

Private Type BITMAPINFO
    Header As BITMAPINFOHEADER
    Colors(0 To 255) As RGBQuad
End Type

'This (stupid) type is used so we can take advantage of LSet when performing some conversions
Private Type tmpLong
    lngResult As Long
End Type

'On GDI+ v1.1 or later, certain effects can be rendered via GDI+.  Note that these are buggy and *not* well-tested,
' so we avoid them in PD except for curiosity and testing purposes.
Private Type GP_BlurParams
    BP_Radius As Single
    BP_ExpandEdge As Long
End Type

'Exporting images via GDI+ is a big headache.  A number of convoluted structs are required if the user
' wants to custom-set any image properties.
Private Type GP_EncoderParameter
    EP_GUID(0 To 15) As Byte
    EP_NumOfValues As Long
    EP_ValueType As GP_EncoderValueType
    EP_ValuePtr As Long
End Type

Private Type GP_EncoderParameters
    EP_Count As Long
    EP_Parameter As GP_EncoderParameter
End Type

Private Type GP_ImageCodecInfo
    IC_ClassID(0 To 15) As Byte
    IC_FormatID(0 To 15) As Byte
    IC_CodecName As Long
    IC_DllName As Long
    IC_FormatDescription As Long
    IC_FilenameExtension As Long
    IC_MimeType As Long
    IC_Flags As Long
    IC_Version As Long
    IC_SigCount As Long
    IC_SigSize As Long
    IC_SigPattern As Long
    IC_SigMask As Long
End Type

'Helper structs for metafile headers.  IMPORTANT NOTE!  There are probably struct alignment issues with these structs,
' as they are legacy structs that intermix 16- and 32-bit datatypes.  I do not need these at present (I only need them
' as part of an unused union in a GDI+ metafile type), so I have not tested them thoroughly.  Use at your own risk.
Private Type GDI_SizeL
    cX As Long
    cY As Long
End Type

Private Type GDI_MetaHeader
    mtType As Integer
    mtHeaderSize As Integer
    mtVersion As Integer
    mtSize As Long
    mtNoObjects As Integer
    mtMaxRecord As Long
    mtNoParameters As Integer
End Type

Private Type GDIP_EnhMetaHeader3
    itype As Long
    nSize As Long
    rclBounds As RECTL
    rclFrame As RECTL
    dSignature As Long
    nVersion As Long
    nBytes As Long
    nRecords As Long
    nHandles As Integer
    sReserved As Integer
    nDescription As Long
    offDescription As Long
    nPalEntries As Long
    szlDevice As GDI_SizeL
    szlMillimeters As GDI_SizeL
End Type

Private Type GP_MetafileHeader_UNION
    muWmfHeader As GDI_MetaHeader
    muEmfHeader As GDIP_EnhMetaHeader3
End Type

'Want additional information on a metafile-type Image object?  This struct contains basic header data.
' IMPORTANT NOTE: please see the previous comment on struct alignment.  I can't guarantee that anything past
' the mfOrigHeader union is aligned correctly; use those at your own risk.
Private Type GP_MetafileHeader
    mfType As GP_MetafileType
    mfSize As Long
    mfVersion As Long
    mfEmfPlusFlags As Long
    mfDpiX As Single
    mfDpiY As Single
    mfBoundsX As Long
    mfBoundsY As Long
    mfBoundsWidth As Long
    mfBoundsHeight As Long
    mfOrigHeader As GP_MetafileHeader_UNION
    mfEmfPlusHeaderSize As Long
    mfLogicalDpiX As Long
    mfLogicalDpiY As Long
End Type

'GDI+ image properties
Public Type GP_PropertyItem
    propID As GP_PropertyTag    'Tag identifier
    propLength As Long          'Length of the property value, in bytes
    propType As Integer         'Type of tag value (one of GP_PropertyTagType)
    ignorePadding As Integer
    propValue As Long           'Property value or pointer to property value, contingent on propType, above
End Type

'GDI+ uses GUIDs to define image formats.  VB6 doesn't let us predeclare byte arrays (at least not easily),
' so we save ourselves the trouble and just use string versions.
Private Const GP_FF_GUID_Undefined = "{B96B3CA9-0728-11D3-9D7B-0000F81EF32E}"
Private Const GP_FF_GUID_MemoryBMP = "{B96B3CAA-0728-11D3-9D7B-0000F81EF32E}"
Private Const GP_FF_GUID_BMP = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}"
Private Const GP_FF_GUID_EMF = "{B96B3CAC-0728-11D3-9D7B-0000F81EF32E}"
Private Const GP_FF_GUID_WMF = "{B96B3CAD-0728-11D3-9D7B-0000F81EF32E}"
Private Const GP_FF_GUID_JPEG = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"
Private Const GP_FF_GUID_PNG = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"
Private Const GP_FF_GUID_GIF = "{B96B3CB0-0728-11D3-9D7B-0000F81EF32E}"
Private Const GP_FF_GUID_TIFF = "{B96B3CB1-0728-11D3-9D7B-0000F81EF32E}"
Private Const GP_FF_GUID_EXIF = "{B96B3CB2-0728-11D3-9D7B-0000F81EF32E}"
Private Const GP_FF_GUID_Icon = "{B96B3CB5-0728-11D3-9D7B-0000F81EF32E}"

'Like image formats, export encoder properties are also defined by GUID.  These values come from the Win 8.1
' version of gdiplusimaging.h.  Note that some are restricted to GDI+ v1.1.
Private Const GP_EP_Compression As String = "{E09D739D-CCD4-44EE-8EBA-3FBF8BE4FC58}"
Private Const GP_EP_ColorDepth As String = "{66087055-AD66-4C7C-9A18-38A2310B8337}"
Private Const GP_EP_ScanMethod As String = "{3A4E2661-3109-4E56-8536-42C156E7DCFA}"
Private Const GP_EP_Version As String = "{24D18C76-814A-41A4-BF53-1C219CCCF797}"
Private Const GP_EP_RenderMethod As String = "{6D42C53A-229A-4825-8BB7-5C99E2B9A8B8}"
Private Const GP_EP_Quality As String = "{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"
Private Const GP_EP_Transformation As String = "{8D0EB2D1-A58E-4EA8-AA14-108074B7B6F9}"
Private Const GP_EP_LuminanceTable As String = "{EDB33BCE-0266-4A77-B904-27216099E717}"
Private Const GP_EP_ChrominanceTable As String = "{F2E455DC-09B3-4316-8260-676ADA32481C}"
Private Const GP_EP_SaveFlag As String = "{292266FC-AC40-47BF-8CFC-A85B89A655DE}"

'THESE ENCODER PROPERTIES REQUIRE GDI+ v1.1 OR LATER!
Private Const GP_EP_ColorSpace As String = "{AE7A62A0-EE2C-49D8-9D07-1BA8A927596E}"
Private Const GP_EP_SaveAsCMYK As String = "{A219BBC9-0A9D-4005-A3EE-3A421B8BB06C}"

'Multi-frame (GIF) and multi-page (TIFF) files support retrieval of individual pages via something Microsoft
' confusingly calls "frame dimensions".  Frame retrieval functions require to specify which kind of frame
' you want to retrieve; these GUIDs control that.
Private Const GP_FD_Page As String = "{7462DC86-6180-4C7E-8E3F-EE7333A7A483}"
Private Const GP_FD_Resolution As String = "{84236F7B-3BD3-428F-8DAB-4EA1439CA315}"
Private Const GP_FD_Time As String = "{6AEDBD6D-3FB5-418A-83A6-7F45229DC872}"

'Core GDI+ functions:
Private Declare Function GdiplusStartup Lib "gdiplus" (ByRef gdipToken As Long, ByRef startupStruct As GDIPlusStartupInput, Optional ByVal OutputBuffer As Long = 0&) As GP_Result
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal gdipToken As Long) As GP_Result

'Object creation/destruction/property functions
Private Declare Function GdipAddPathRectangle Lib "gdiplus" (ByVal hPath As Long, ByVal X1 As Single, ByVal Y1 As Single, ByVal rectWidth As Single, ByVal rectHeight As Single) As GP_Result
Private Declare Function GdipAddPathRectangleI Lib "gdiplus" (ByVal hPath As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal rectWidth As Long, ByVal rectHeight As Long) As GP_Result
Private Declare Function GdipAddPathEllipse Lib "gdiplus" (ByVal hPath As Long, ByVal X1 As Single, ByVal Y1 As Single, ByVal rectWidth As Single, ByVal rectHeight As Single) As GP_Result
Private Declare Function GdipAddPathLine Lib "gdiplus" (ByVal hPath As Long, ByVal X1 As Single, ByVal Y1 As Single, ByVal x2 As Single, ByVal y2 As Single) As GP_Result
Private Declare Function GdipAddPathLineI Lib "gdiplus" (ByVal hPath As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As GP_Result
Private Declare Function GdipAddPathCurve2 Lib "gdiplus" (ByVal hPath As Long, ByVal ptrToFloatArray As Long, ByVal numOfPoints As Long, ByVal curveTension As Single) As GP_Result
Private Declare Function GdipAddPathCurve2I Lib "gdiplus" (ByVal hPath As Long, ByVal ptrToLongArray As Long, ByVal numOfPoints As Long, ByVal curveTension As Single) As GP_Result
Private Declare Function GdipAddPathClosedCurve2 Lib "gdiplus" (ByVal hPath As Long, ByVal ptrToFloatArray As Long, ByVal numOfPoints As Long, ByVal curveTension As Single) As GP_Result
Private Declare Function GdipAddPathClosedCurve2I Lib "gdiplus" (ByVal hPath As Long, ByVal ptrToLongArray As Long, ByVal numOfPoints As Long, ByVal curveTension As Single) As GP_Result
Private Declare Function GdipAddPathBezier Lib "gdiplus" (ByVal hPath As Long, ByVal X1 As Single, ByVal Y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal x3 As Single, ByVal y3 As Single, ByVal x4 As Single, ByVal y4 As Single) As GP_Result
Private Declare Function GdipAddPathLine2 Lib "gdiplus" (ByVal hPath As Long, ByVal ptrToFloatArray As Long, ByVal numOfPoints As Long) As GP_Result
Private Declare Function GdipAddPathLine2I Lib "gdiplus" (ByVal hPath As Long, ByVal ptrToIntArray As Long, ByVal numOfPoints As Long) As GP_Result
Private Declare Function GdipAddPathPolygon Lib "gdiplus" (ByVal hPath As Long, ByVal ptrToFloatArray As Long, ByVal numOfPoints As Long) As GP_Result
Private Declare Function GdipAddPathPolygonI Lib "gdiplus" (ByVal hPath As Long, ByVal ptrToLongArray As Long, ByVal numOfPoints As Long) As GP_Result
Private Declare Function GdipAddPathArc Lib "gdiplus" (ByVal hPath As Long, ByVal x As Single, ByVal Y As Single, ByVal arcWidth As Single, ByVal arcHeight As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As GP_Result
Private Declare Function GdipAddPathPath Lib "gdiplus" (ByVal hPath As Long, ByVal pathToAdd As Long, ByVal connectToPreviousPoint As Long) As GP_Result

Private Declare Function GdipBitmapLockBits Lib "gdiplus" (ByVal hImage As Long, ByRef srcRect As RECTL, ByVal lockMode As GP_BitmapLockMode, ByVal dstPixelFormat As GP_PixelFormat, ByRef srcBitmapData As GP_BitmapData) As GP_Result
Private Declare Function GdipBitmapUnlockBits Lib "gdiplus" (ByVal hImage As Long, ByRef srcBitmapData As GP_BitmapData) As GP_Result

Private Declare Function GdipCloneBitmapAreaI Lib "gdiplus" (ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal newPixelFormat As GP_PixelFormat, ByVal hSrcBitmap As Long, ByRef hDstBitmap As Long) As GP_Result
Private Declare Function GdipCloneMatrix Lib "gdiplus" (ByVal srcMatrix As Long, ByRef dstMatrix As Long) As GP_Result
Private Declare Function GdipClonePath Lib "gdiplus" (ByVal srcPath As Long, ByRef dstPath As Long) As GP_Result
Private Declare Function GdipCloneRegion Lib "gdiplus" (ByVal srcRegion As Long, ByRef dstRegion As Long) As GP_Result
Private Declare Function GdipClosePathFigure Lib "gdiplus" (ByVal hPath As Long) As GP_Result

Private Declare Function GdipCombineRegionRect Lib "gdiplus" (ByVal hRegion As Long, ByRef srcRectF As RECTF, ByVal dstCombineMode As GP_CombineMode) As GP_Result
Private Declare Function GdipCombineRegionRectI Lib "gdiplus" (ByVal hRegion As Long, ByRef srcRectL As RECTL, ByVal dstCombineMode As GP_CombineMode) As GP_Result
Private Declare Function GdipCombineRegionRegion Lib "gdiplus" (ByVal dstRegion As Long, ByVal srcRegion As Long, ByVal dstCombineMode As GP_CombineMode) As GP_Result
Private Declare Function GdipCombineRegionPath Lib "gdiplus" (ByVal dstRegion As Long, ByVal srcPath As Long, ByVal dstCombineMode As GP_CombineMode) As GP_Result

Private Declare Function GdipCreateBitmapFromGdiDib Lib "gdiplus" (ByRef origGDIBitmapInfo As BITMAPINFO, ByRef srcBitmapData As Any, ByRef dstGdipBitmap As Long) As GP_Result
Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal bmpWidth As Long, ByVal bmpHeight As Long, ByVal bmpStride As Long, ByVal bmpPixelFormat As GP_PixelFormat, ByRef Scan0 As Any, ByRef dstGdipBitmap As Long) As GP_Result
Private Declare Function GdipCreateCachedBitmap Lib "gdiplus" (ByVal hBitmap As Long, ByVal hGraphics As Long, ByRef dstCachedBitmap As Long) As GP_Result
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, ByRef dstGraphics As Long) As GP_Result
Private Declare Function GdipCreateHatchBrush Lib "gdiplus" (ByVal bHatchStyle As GP_PatternStyle, ByVal bForeColor As Long, ByVal bBackColor As Long, ByRef dstBrush As Long) As GP_Result
Private Declare Function GdipCreateImageAttributes Lib "gdiplus" (ByRef dstImageAttributes As Long) As GP_Result
Private Declare Function GdipCreateLineBrush Lib "gdiplus" (ByRef firstPoint As PointFloat, ByRef secondPoint As PointFloat, ByVal firstRGBA As Long, ByVal secondRGBA As Long, ByVal brushWrapMode As GP_WrapMode, ByRef dstBrush As Long) As GP_Result
Private Declare Function GdipCreateLineBrushFromRectWithAngle Lib "gdiplus" (ByRef srcRect As RECTF, ByVal firstRGBA As Long, ByVal secondRGBA As Long, ByVal gradAngle As Single, ByVal isAngleScalable As Long, ByVal gradientWrapMode As GP_WrapMode, ByRef dstLineGradientBrush As Long) As GP_Result
Private Declare Function GdipCreateMatrix Lib "gdiplus" (ByRef dstMatrix As Long) As GP_Result
Private Declare Function GdipCreateMatrix2 Lib "gdiplus" (ByVal mM11 As Single, ByVal mM12 As Single, ByVal mM21 As Single, ByVal mM22 As Single, ByVal mDx As Single, ByVal mDy As Single, ByRef dstMatrix As Long) As GP_Result
Private Declare Function GdipCreatePath Lib "gdiplus" (ByVal pathFillMode As GP_FillMode, ByRef dstPath As Long) As GP_Result
Private Declare Function GdipCreatePathGradientFromPath Lib "gdiplus" (ByVal ptrToSrcPath As Long, ByRef dstPathGradientBrush As Long) As GP_Result
Private Declare Function GdipCreatePen1 Lib "gdiplus" (ByVal srcColor As Long, ByVal srcWidth As Single, ByVal srcUnit As GP_Unit, ByRef dstPen As Long) As GP_Result
Private Declare Function GdipCreatePenFromBrush Lib "gdiplus" Alias "GdipCreatePen2" (ByVal srcBrush As Long, ByVal penWidth As Single, ByVal srcUnit As GP_Unit, ByRef dstPen As Long) As GP_Result
Private Declare Function GdipCreateRegion Lib "gdiplus" (ByRef dstRegion As Long) As GP_Result
Private Declare Function GdipCreateRegionPath Lib "gdiplus" (ByVal hPath As Long, ByRef hRegion As Long) As GP_Result
Private Declare Function GdipCreateRegionRect Lib "gdiplus" (ByRef srcRect As RECTF, ByRef hRegion As Long) As GP_Result
Private Declare Function GdipCreateRegionRgnData Lib "gdiplus" (ByVal ptrToRegionData As Long, ByVal sizeOfRegionData As Long, ByRef dstRegion As Long) As GP_Result
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal srcColor As Long, ByRef dstBrush As Long) As GP_Result
Private Declare Function GdipCreateTexture Lib "gdiplus" (ByVal hImage As Long, ByVal textureWrapMode As GP_WrapMode, ByRef dstTexture As Long) As GP_Result

Private Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal hBrush As Long) As GP_Result
Private Declare Function GdipDeleteCachedBitmap Lib "gdiplus" (ByVal hCachedBitmap As Long) As GP_Result

Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal hGraphics As Long) As GP_Result
Private Declare Function GdipDeleteMatrix Lib "gdiplus" (ByVal hMatrix As Long) As GP_Result
Private Declare Function GdipDeletePath Lib "gdiplus" (ByVal hPath As Long) As GP_Result
Private Declare Function GdipDeletePen Lib "gdiplus" (ByVal hPen As Long) As GP_Result
Private Declare Function GdipDeleteRegion Lib "gdiplus" (ByVal hRegion As Long) As GP_Result
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal hImage As Long) As GP_Result
Private Declare Function GdipDisposeImageAttributes Lib "gdiplus" (ByVal hImageAttributes As Long) As GP_Result

Private Declare Function GdipDrawArc Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal x As Single, ByVal Y As Single, ByVal nWidth As Single, ByVal nHeight As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As GP_Result
Private Declare Function GdipDrawArcI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal startAngle As Long, ByVal sweepAngle As Long) As GP_Result
Private Declare Function GdipDrawCachedBitmap Lib "gdiplus" (ByVal hGraphics As Long, ByVal hCachedBitmap As Long, ByVal x As Long, ByVal Y As Long) As GP_Result
Private Declare Function GdipDrawClosedCurve2 Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal ptrToPointFloats As Long, ByVal numOfPoints As Long, ByVal curveTension As Single) As GP_Result
Private Declare Function GdipDrawClosedCurve2I Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal ptrToPointLongs As Long, ByVal numOfPoints As Long, ByVal curveTension As Single) As GP_Result
Private Declare Function GdipDrawCurve2 Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal ptrToPointFloats As Long, ByVal numOfPoints As Long, ByVal curveTension As Single) As GP_Result
Private Declare Function GdipDrawCurve2I Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal ptrToPointLongs As Long, ByVal numOfPoints As Long, ByVal curveTension As Single) As GP_Result
Private Declare Function GdipDrawEllipse Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal x As Single, ByVal Y As Single, ByVal nWidth As Single, ByVal nHeight As Single) As GP_Result
Private Declare Function GdipDrawEllipseI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long) As GP_Result
Private Declare Function GdipDrawImage Lib "gdiplus" (ByVal dstGraphics As Long, ByVal srcImage As Long, ByVal x As Single, ByVal Y As Single) As GP_Result
Private Declare Function GdipDrawImageI Lib "gdiplus" (ByVal dstGraphics As Long, ByVal srcImage As Long, ByVal x As Long, ByVal Y As Long) As GP_Result
Private Declare Function GdipDrawImageRect Lib "gdiplus" (ByVal dstGraphics As Long, ByVal srcImage As Long, ByVal x As Single, ByVal Y As Single, ByVal dstWidth As Single, ByVal dstHeight As Single) As GP_Result
Private Declare Function GdipDrawImageRectI Lib "gdiplus" (ByVal dstGraphics As Long, ByVal srcImage As Long, ByVal x As Long, ByVal Y As Long, ByVal dstWidth As Long, ByVal dstHeight As Long) As GP_Result
Private Declare Function GdipDrawImageRectRect Lib "gdiplus" (ByVal dstGraphics As Long, ByVal srcImage As Long, ByVal dstX As Single, ByVal dstY As Single, ByVal dstWidth As Single, ByVal dstHeight As Single, ByVal srcX As Single, ByVal srcY As Single, ByVal srcWidth As Single, ByVal srcHeight As Single, ByVal srcUnit As GP_Unit, Optional ByVal newImgAttributes As Long = 0, Optional ByVal progCallbackFunction As Long = 0, Optional ByVal progCallbackData As Long = 0) As GP_Result
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal dstGraphics As Long, ByVal srcImage As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal srcUnit As GP_Unit, Optional ByVal newImgAttributes As Long = 0, Optional ByVal progCallbackFunction As Long = 0, Optional ByVal progCallbackData As Long = 0) As GP_Result
Private Declare Function GdipDrawImagePointsRect Lib "gdiplus" (ByVal dstGraphics As Long, ByVal srcImage As Long, ByVal ptrToPointFloats As Long, ByVal dstPtCount As Long, ByVal srcX As Single, ByVal srcY As Single, ByVal srcWidth As Single, ByVal srcHeight As Single, ByVal srcUnit As GP_Unit, Optional ByVal newImgAttributes As Long = 0, Optional ByVal progCallbackFunction As Long = 0, Optional ByVal progCallbackData As Long = 0) As GP_Result
Private Declare Function GdipDrawImagePointsRectI Lib "gdiplus" (ByVal dstGraphics As Long, ByVal srcImage As Long, ByVal ptrToPointInts As Long, ByVal dstPtCount As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal srcUnit As GP_Unit, Optional ByVal newImgAttributes As Long = 0, Optional ByVal progCallbackFunction As Long = 0, Optional ByVal progCallbackData As Long = 0) As GP_Result
Private Declare Function GdipDrawLine Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal X1 As Single, ByVal Y1 As Single, ByVal x2 As Single, ByVal y2 As Single) As GP_Result
Private Declare Function GdipDrawLineI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As GP_Result
Private Declare Function GdipDrawLines Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal ptrToPointFloats As Long, ByVal numOfPoints As Long) As GP_Result
Private Declare Function GdipDrawLinesI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal ptrToPointLongs As Long, ByVal numOfPoints As Long) As GP_Result
Private Declare Function GdipDrawPath Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal hPath As Long) As GP_Result
Private Declare Function GdipDrawPolygon Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal ptrToPointFloats As Long, ByVal numOfPoints As Long) As GP_Result
Private Declare Function GdipDrawPolygonI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal ptrToPointLongs As Long, ByVal numOfPoints As Long) As GP_Result
Private Declare Function GdipDrawRectangle Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal x As Single, ByVal Y As Single, ByVal nWidth As Single, ByVal nHeight As Single) As GP_Result
Private Declare Function GdipDrawRectangleI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long) As GP_Result

Private Declare Function GdipFillClosedCurve2 Lib "gdiplus" (ByVal hGraphics As Long, ByVal hBrush As Long, ByVal ptrToPointFloats As Long, ByVal numOfPoints As Long, ByVal curveTension As Single, ByVal fillMode As GP_FillMode) As GP_Result
Private Declare Function GdipFillClosedCurve2I Lib "gdiplus" (ByVal hGraphics As Long, ByVal hBrush As Long, ByVal ptrToPointLongs As Long, ByVal numOfPoints As Long, ByVal curveTension As Single, ByVal fillMode As GP_FillMode) As GP_Result
Private Declare Function GdipFillEllipse Lib "gdiplus" (ByVal hGraphics As Long, ByVal hBrush As Long, ByVal x As Single, ByVal Y As Single, ByVal nWidth As Single, ByVal nHeight As Single) As GP_Result
Private Declare Function GdipFillEllipseI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hBrush As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long) As GP_Result
Private Declare Function GdipFillPath Lib "gdiplus" (ByVal hGraphics As Long, ByVal hBrush As Long, ByVal hPath As Long) As GP_Result
Private Declare Function GdipFillPolygon Lib "gdiplus" (ByVal hGraphics As Long, ByVal hBrush As Long, ByVal ptrToPointFloats As Long, ByVal numOfPoints As Long, ByVal fillMode As GP_FillMode) As GP_Result
Private Declare Function GdipFillPolygonI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hBrush As Long, ByVal ptrToPointLongs As Long, ByVal numOfPoints As Long, ByVal fillMode As GP_FillMode) As GP_Result
Private Declare Function GdipFillRectangle Lib "gdiplus" (ByVal hGraphics As Long, ByVal hBrush As Long, ByVal x As Single, ByVal Y As Single, ByVal nWidth As Single, ByVal nHeight As Single) As GP_Result
Private Declare Function GdipFillRectangleI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hBrush As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long) As GP_Result
Private Declare Function GdipFlush Lib "gdiplus" (ByVal hGraphics As Long, ByVal intention As GP_FlushIntention) As Long

Private Declare Function GdipGetClip Lib "gdiplus" (ByVal hGraphics As Long, ByRef dstRegion As Long) As GP_Result
Private Declare Function GdipGetCompositingMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal dstCompositingMode As GP_CompositingMode) As GP_Result
Private Declare Function GdipGetCompositingQuality Lib "gdiplus" (ByVal hGraphics As Long, ByRef dstCompositingQuality As GP_CompositingQuality) As GP_Result
Private Declare Function GdipGetImageBounds Lib "gdiplus" (ByVal hImage As Long, ByRef dstRectF As RECTF, ByRef dstUnit As GP_Unit) As GP_Result
Private Declare Function GdipGetImageDimension Lib "gdiplus" (ByVal hImage As Long, ByRef dstWidth As Single, ByRef dstHeight As Single) As GP_Result
Private Declare Function GdipGetImageDecoders Lib "gdiplus" (ByVal numOfEncoders As Long, ByVal sizeOfEncoders As Long, ByVal ptrToDstEncoders As Long) As GP_Result
Private Declare Function GdipGetImageDecodersSize Lib "gdiplus" (ByRef numOfEncoders As Long, ByRef sizeOfEncoders As Long) As GP_Result
Private Declare Function GdipGetImageEncoders Lib "gdiplus" (ByVal numOfEncoders As Long, ByVal sizeOfEncoders As Long, ByVal ptrToDstEncoders As Long) As GP_Result
Private Declare Function GdipGetImageEncodersSize Lib "gdiplus" (ByRef numOfEncoders As Long, ByRef sizeOfEncoders As Long) As GP_Result
Private Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal hImage As Long, ByRef dstHeight As Long) As GP_Result
Private Declare Function GdipGetImageHorizontalResolution Lib "gdiplus" (ByVal hImage As Long, ByRef dstHResolution As Single) As GP_Result
Private Declare Function GdipGetImagePixelFormat Lib "gdiplus" (ByVal hImage As Long, ByRef dstPixelFormat As GP_PixelFormat) As GP_Result
Private Declare Function GdipGetImageRawFormat Lib "gdiplus" (ByVal hImage As Long, ByVal ptrToDstGuid As Long) As GP_Result
Private Declare Function GdipGetImageType Lib "gdiplus" (ByVal srcImage As Long, ByRef dstImageType As GP_ImageType) As GP_Result
Private Declare Function GdipGetImageVerticalResolution Lib "gdiplus" (ByVal hImage As Long, ByRef dstVResolution As Single) As GP_Result
Private Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal hImage As Long, ByRef dstWidth As Long) As GP_Result
Private Declare Function GdipGetInterpolationMode Lib "gdiplus" (ByVal hGraphics As Long, ByRef dstInterpolationMode As GP_InterpolationMode) As GP_Result
Private Declare Function GdipGetMetafileHeaderFromMetafile Lib "gdiplus" (ByVal hMetafile As Long, ByRef dstHeader As GP_MetafileHeader) As GP_Result
Private Declare Function GdipGetPathFillMode Lib "gdiplus" (ByVal hPath As Long, ByRef dstFillRule As GP_FillMode) As GP_Result
Private Declare Function GdipGetPathWorldBounds Lib "gdiplus" (ByVal hPath As Long, ByRef dstBounds As RECTF, ByVal tmpTransformMatrix As Long, ByVal tmpPenHandle As Long) As GP_Result
Private Declare Function GdipGetPathWorldBoundsI Lib "gdiplus" (ByVal hPath As Long, ByRef dstBounds As RECTL, ByVal tmpTransformMatrix As Long, ByVal tmpPenHandle As Long) As GP_Result
Private Declare Function GdipGetPenColor Lib "gdiplus" (ByVal hPen As Long, ByRef dstPARGBColor As Long) As GP_Result
Private Declare Function GdipGetPenDashCap Lib "gdiplus" Alias "GdipGetPenDashCap197819" (ByVal hPen As Long, ByRef dstCap As GP_DashCap) As GP_Result
Private Declare Function GdipGetPenDashOffset Lib "gdiplus" (ByVal hPen As Long, ByRef dstOffset As Single) As GP_Result
Private Declare Function GdipGetPenDashStyle Lib "gdiplus" (ByVal hPen As Long, ByRef dstDashStyle As GP_DashStyle) As GP_Result
Private Declare Function GdipGetPenEndCap Lib "gdiplus" (ByVal hPen As Long, ByRef dstLineCap As GP_LineCap) As GP_Result
Private Declare Function GdipGetPenStartCap Lib "gdiplus" (ByVal hPen As Long, ByRef dstLineCap As GP_LineCap) As GP_Result
Private Declare Function GdipGetPenLineJoin Lib "gdiplus" (ByVal hPen As Long, ByRef dstLineJoin As GP_LineJoin) As GP_Result
Private Declare Function GdipGetPenMiterLimit Lib "gdiplus" (ByVal hPen As Long, ByRef dstMiterLimit As Single) As GP_Result
Private Declare Function GdipGetPenMode Lib "gdiplus" (ByVal hPen As Long, ByRef dstPenMode As GP_PenAlignment) As GP_Result
Private Declare Function GdipGetPenWidth Lib "gdiplus" (ByVal hPen As Long, ByRef dstWidth As Single) As GP_Result
Private Declare Function GdipGetPixelOffsetMode Lib "gdiplus" (ByVal hGraphics As Long, ByRef dstMode As GP_PixelOffsetMode) As GP_Result
Private Declare Function GdipGetPropertyItem Lib "gdiplus" (ByVal hImage As Long, ByVal gpPropertyID As Long, ByVal srcPropertySize As Long, ByVal ptrToDstBuffer As Long) As GP_Result
Private Declare Function GdipGetPropertyItemSize Lib "gdiplus" (ByVal hImage As Long, ByVal gpPropertyID As GP_PropertyTag, ByRef dstPropertySize As Long) As GP_Result
Private Declare Function GdipGetRegionBounds Lib "gdiplus" (ByVal hRegion As Long, ByVal hGraphics As Long, ByRef dstRectF As RECTF) As GP_Result
Private Declare Function GdipGetRegionBoundsI Lib "gdiplus" (ByVal hRegion As Long, ByVal hGraphics As Long, ByRef dstRectL As RECTL) As GP_Result
Private Declare Function GdipGetRegionHRgn Lib "gdiplus" (ByVal hRegion As Long, ByVal hGraphics As Long, ByRef dstHRgn As Long) As GP_Result
Private Declare Function GdipGetRenderingOrigin Lib "gdiplus" (ByVal hGraphics As Long, ByRef dstX As Long, ByRef dstY As Long) As GP_Result
Private Declare Function GdipGetSmoothingMode Lib "gdiplus" (ByVal hGraphics As Long, ByRef dstMode As GP_SmoothingMode) As GP_Result
Private Declare Function GdipGetSolidFillColor Lib "gdiplus" (ByVal hBrush As Long, ByRef dstColor As Long) As GP_Result
Private Declare Function GdipGetTextureWrapMode Lib "gdiplus" (ByVal hBrush As Long, ByRef dstWrapMode As GP_WrapMode) As GP_Result

Private Declare Function GdipImageGetFrameCount Lib "gdiplus" (ByVal hImage As Long, ByVal ptrToDimensionGuid As Long, ByRef dstCount As Long) As GP_Result
Private Declare Function GdipImageRotateFlip Lib "gdiplus" (ByVal hImage As Long, ByVal rotateFlipType As GP_RotateFlip) As GP_Result
Private Declare Function GdipImageSelectActiveFrame Lib "gdiplus" (ByVal hImage As Long, ByVal ptrToDimensionGuid As Long, ByVal frameIndex As Long) As GP_Result

Private Declare Function GdipInvertMatrix Lib "gdiplus" (ByVal hMatrix As Long) As GP_Result

Private Declare Function GdipIsEmptyRegion Lib "gdiplus" (ByVal srcRegion As Long, ByVal srcGraphics As Long, ByRef dstResult As Long) As GP_Result
Private Declare Function GdipIsEqualRegion Lib "gdiplus" (ByVal srcRegion1 As Long, ByVal srcRegion2 As Long, ByVal srcGraphics As Long, ByRef dstResult As Long) As GP_Result
Private Declare Function GdipIsInfiniteRegion Lib "gdiplus" (ByVal srcRegion As Long, ByVal srcGraphics As Long, ByRef dstResult As Long) As GP_Result
Private Declare Function GdipIsMatrixInvertible Lib "gdiplus" (ByVal hMatrix As Long, ByRef dstResult As Long) As Long
Private Declare Function GdipIsOutlineVisiblePathPoint Lib "gdiplus" (ByVal hPath As Long, ByVal x As Single, ByVal Y As Single, ByVal hPen As Long, ByVal hGraphicsOptional As Long, ByRef dstResult As Long) As GP_Result
Private Declare Function GdipIsOutlineVisiblePathPointI Lib "gdiplus" (ByVal hPath As Long, ByVal x As Long, ByVal Y As Long, ByVal hPen As Long, ByVal hGraphicsOptional As Long, ByRef dstResult As Long) As GP_Result
Private Declare Function GdipIsVisiblePathPoint Lib "gdiplus" (ByVal hPath As Long, ByVal x As Single, ByVal Y As Single, ByVal hGraphicsOptional As Long, ByRef dstResult As Long) As GP_Result
Private Declare Function GdipIsVisiblePathPointI Lib "gdiplus" (ByVal hPath As Long, ByVal x As Long, ByVal Y As Long, ByVal hGraphicsOptional As Long, ByRef dstResult As Long) As GP_Result
Private Declare Function GdipIsVisibleRegionPoint Lib "gdiplus" (ByVal hRegion As Long, ByVal x As Single, ByVal Y As Single, ByVal hGraphicsOptional As Long, ByRef dstResult As Long) As GP_Result
Private Declare Function GdipIsVisibleRegionPointI Lib "gdiplus" (ByVal hRegion As Long, ByVal x As Long, ByVal Y As Long, ByVal hGraphicsOptional As Long, ByRef dstResult As Long) As GP_Result

Private Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal ptrSrcFilename As Long, ByRef dstGdipImage As Long) As GP_Result
Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal srcIStream As Long, ByRef dstGdipImage As Long) As GP_Result

Private Declare Function GdipResetClip Lib "gdiplus" (ByVal hGraphics As Long) As GP_Result
Private Declare Function GdipResetPath Lib "gdiplus" (ByVal hPath As Long) As GP_Result
Private Declare Function GdipRotateMatrix Lib "gdiplus" (ByVal hMatrix As Long, ByVal rotateAngle As Single, ByVal mOrder As GP_MatrixOrder) As GP_Result

Private Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal hImage As Long, ByVal ptrToFilename As Long, ByVal ptrToEncoderGUID As Long, ByVal ptrToEncoderParams As Long) As GP_Result
Private Declare Function GdipSaveImageToStream Lib "gdiplus" (ByVal hImage As Long, ByVal dstIStream As Long, ByVal ptrToEncoderGUID As Long, ByVal ptrToEncoderParams As Long) As GP_Result
Private Declare Function GdipScaleMatrix Lib "gdiplus" (ByVal hMatrix As Long, ByVal scaleX As Single, ByVal scaleY As Single, ByVal mOrder As GP_MatrixOrder) As GP_Result

Private Declare Function GdipSetClipRect Lib "gdiplus" (ByVal hGraphics As Long, ByVal x As Single, ByVal Y As Single, ByVal nWidth As Single, ByVal nHeight As Single, ByVal useCombineMode As GP_CombineMode) As GP_Result
Private Declare Function GdipSetClipRectI Lib "gdiplus" (ByVal hGraphics As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal useCombineMode As GP_CombineMode) As GP_Result
Private Declare Function GdipSetClipRegion Lib "gdiplus" (ByVal hGraphics As Long, ByVal hRegion As Long, ByVal useCombineMode As GP_CombineMode) As GP_Result
Private Declare Function GdipSetCompositingMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal newCompositingMode As GP_CompositingMode) As GP_Result
Private Declare Function GdipSetCompositingQuality Lib "gdiplus" (ByVal hGraphics As Long, ByVal newCompositingQuality As GP_CompositingQuality) As GP_Result
Private Declare Function GdipSetEmpty Lib "gdiplus" (ByVal hRegion As Long) As GP_Result
Private Declare Function GdipSetImageAttributesWrapMode Lib "gdiplus" (ByVal hImageAttributes As Long, ByVal newWrapMode As GP_WrapMode, ByVal argbOfClampMode As Long, ByVal bClampMustBeZero As Long) As GP_Result
Private Declare Function GdipSetImageAttributesColorMatrix Lib "gdiplus" (ByVal hImageAttributes As Long, ByVal typeOfAdjustment As GP_ColorAdjustType, ByVal enableSeparateAdjustmentFlag As Long, ByVal ptrToColorMatrix As Long, ByVal ptrToGrayscaleMatrix As Long, ByVal extraColorMatrixFlags As GP_ColorMatrixFlags) As GP_Result
Private Declare Function GdipSetInfinite Lib "gdiplus" (ByVal hRegion As Long) As GP_Result
Private Declare Function GdipSetInterpolationMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal newInterpolationMode As GP_InterpolationMode) As GP_Result
Private Declare Function GdipSetLinePresetBlend Lib "gdiplus" (ByVal hBrush As Long, ByVal ptrToFirstColor As Long, ByVal ptrToFirstPosition As Long, ByVal numOfPoints As Long) As GP_Result
Private Declare Function GdipSetMetafileDownLevelRasterizationLimit Lib "gdiplus" (ByVal hMetafile As Long, ByVal metafileRasterizationLimitDpi As Long) As GP_Result
Private Declare Function GdipSetPathGradientCenterPoint Lib "gdiplus" (ByVal hBrush As Long, ByRef newCenterPoints As PointFloat) As GP_Result
Private Declare Function GdipSetPathGradientPresetBlend Lib "gdiplus" (ByVal hBrush As Long, ByVal ptrToFirstColor As Long, ByVal ptrToFirstPosition As Long, ByVal numOfPoints As Long) As GP_Result
Private Declare Function GdipSetPathGradientWrapMode Lib "gdiplus" (ByVal hBrush As Long, ByVal newWrapMode As GP_WrapMode) As GP_Result
Private Declare Function GdipSetPathFillMode Lib "gdiplus" (ByVal hPath As Long, ByVal pathFillMode As GP_FillMode) As GP_Result
Private Declare Function GdipSetPenColor Lib "gdiplus" (ByVal hPen As Long, ByVal pARGBColor As Long) As GP_Result
Private Declare Function GdipSetPenDashArray Lib "gdiplus" (ByVal hPen As Long, ByVal ptrToDashArray As Long, ByVal numOfDashes As Long) As GP_Result
Private Declare Function GdipSetPenDashCap Lib "gdiplus" Alias "GdipSetPenDashCap197819" (ByVal hPen As Long, ByVal newCap As GP_DashCap) As GP_Result
Private Declare Function GdipSetPenDashOffset Lib "gdiplus" (ByVal hPen As Long, ByVal newPenOffset As Single) As GP_Result
Private Declare Function GdipSetPenDashStyle Lib "gdiplus" (ByVal hPen As Long, ByVal newDashStyle As GP_DashStyle) As GP_Result
Private Declare Function GdipSetPenEndCap Lib "gdiplus" (ByVal hPen As Long, ByVal endCap As GP_LineCap) As GP_Result
Private Declare Function GdipSetPenLineCap Lib "gdiplus" Alias "GdipSetPenLineCap197819" (ByVal hPen As Long, ByVal startCap As GP_LineCap, ByVal endCap As GP_LineCap, ByVal dashCap As GP_DashCap) As GP_Result
Private Declare Function GdipSetPenLineJoin Lib "gdiplus" (ByVal hPen As Long, ByVal newLineJoin As GP_LineJoin) As GP_Result
Private Declare Function GdipSetPenMiterLimit Lib "gdiplus" (ByVal hPen As Long, ByVal newMiterLimit As Single) As GP_Result
Private Declare Function GdipSetPenMode Lib "gdiplus" (ByVal hPen As Long, ByVal penMode As GP_PenAlignment) As GP_Result
Private Declare Function GdipSetPenStartCap Lib "gdiplus" (ByVal hPen As Long, ByVal startCap As GP_LineCap) As GP_Result
Private Declare Function GdipSetPenWidth Lib "gdiplus" (ByVal hPen As Long, ByVal penWidth As Single) As GP_Result
Private Declare Function GdipSetPixelOffsetMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal newMode As GP_PixelOffsetMode) As GP_Result
Private Declare Function GdipSetRenderingOrigin Lib "gdiplus" (ByVal hGraphics As Long, ByVal x As Long, ByVal Y As Long) As GP_Result
Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal newMode As GP_SmoothingMode) As GP_Result
Private Declare Function GdipSetSolidFillColor Lib "gdiplus" (ByVal hBrush As Long, ByVal newColor As Long) As GP_Result
Private Declare Function GdipSetTextureWrapMode Lib "gdiplus" (ByVal hBrush As Long, ByVal newWrapMode As GP_WrapMode) As GP_Result
Private Declare Function GdipSetWorldTransform Lib "gdiplus" (ByVal hGraphics As Long, ByVal hMatrix As Long) As GP_Result

Private Declare Function GdipShearMatrix Lib "gdiplus" (ByVal hMatrix As Long, ByVal shearX As Single, ByVal shearY As Single, ByVal mOrder As GP_MatrixOrder) As GP_Result
Private Declare Function GdipStartPathFigure Lib "gdiplus" (ByVal hPath As Long) As GP_Result

Private Declare Function GdipTranslateMatrix Lib "gdiplus" (ByVal hMatrix As Long, ByVal offsetX As Single, ByVal offsetY As Single, ByVal mOrder As GP_MatrixOrder) As GP_Result
Private Declare Function GdipTransformMatrixPoints Lib "gdiplus" (ByVal hMatrix As Long, ByVal ptrToFirstPointF As Long, ByVal numOfPoints As Long) As GP_Result
Private Declare Function GdipTransformPath Lib "gdiplus" (ByVal hPath As Long, ByVal hMatrix As Long) As GP_Result

Private Declare Function GdipWidenPath Lib "gdiplus" (ByVal hPath As Long, ByVal hPen As Long, ByVal hTransformMatrix As Long, ByVal allowableError As Single) As GP_Result
Private Declare Function GdipWindingModeOutline Lib "gdiplus" (ByVal hPath As Long, ByVal hTransformationMatrix As Long, ByVal allowableError As Single) As GP_Result

'Some GDI+ functions are *only* supported on GDI+ 1.1, which first shipped with Vista (but requires explicit activation
' via manifest, and as such, is unavailable to PD until Win 7).  Take care to confirm the availability of these functions
' before using them.
Private Declare Function GdipConvertToEmfPlus Lib "gdiplus" (ByVal hGraphics As Long, ByVal srcMetafile As Long, ByRef conversionSuccess As Long, ByVal typeOfEMF As GP_MetafileType, ByVal ptrToMetafileDescription As Long, ByRef dstMetafilePtr As Long) As GP_Result
Private Declare Function GdipConvertToEmfPlusToFile Lib "gdiplus" (ByVal hGraphics As Long, ByVal srcMetafile As Long, ByRef conversionSuccess As Long, ByVal filenamePointer As Long, ByVal typeOfEMF As GP_MetafileType, ByVal ptrToMetafileDescription As Long, ByRef dstMetafilePtr As Long) As GP_Result
Private Declare Function GdipCreateEffect Lib "gdiplus" (ByVal dwCid1 As Long, ByVal dwCid2 As Long, ByVal dwCid3 As Long, ByVal dwCid4 As Long, ByRef dstEffect As Long) As GP_Result
Private Declare Function GdipDeleteEffect Lib "gdiplus" (ByVal hEffect As Long) As GP_Result
Private Declare Function GdipDrawImageFX Lib "gdiplus" (ByVal hGraphics As Long, ByVal hImage As Long, ByRef drawRect As RECTF, ByVal hTransformMatrix As Long, ByVal hEffect As Long, ByVal hImageAttributes As Long, ByVal srcUnit As GP_Unit) As GP_Result
Private Declare Function GdipSetEffectParameters Lib "gdiplus" (ByVal hEffect As Long, ByRef srcParams As Any, ByVal srcParamSize As Long) As GP_Result

'Non-GDI+ helper functions:
Private Declare Function BitBlt Lib "gdi32" (ByVal hDstDC As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal hSrcDC As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal rastOp As Long) As Boolean
Private Declare Function CLSIDFromString Lib "ole32" (ByVal ptrToGuidString As Long, ByVal ptrToByteArray As Long) As Long
Private Declare Function CopyMemoryStrict Lib "kernel32" Alias "RtlMoveMemory" (ByVal ptrDst As Long, ByVal ptrSrc As Long, ByVal numOfBytes As Long) As Long
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ByVal ptrToDstStream As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetHGlobalFromStream Lib "ole32" (ByVal srcIStream As Long, ByRef dstHGlobal As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32" (ByVal oColor As OLE_COLOR, ByVal hPalette As Long, ByRef cColorRef As Long) As Long
Private Declare Function StringFromCLSID Lib "ole32" (ByVal ptrToGuid As Long, ByRef ptrToDstString As Long) As Long
Private Declare Function SysAllocStringByteLen Lib "oleaut32" (ByVal srcPtr As Long, ByVal strLength As Long) As String

'Used to quickly check if a file (or folder) exists
Private Const ERROR_SHARING_VIOLATION As Long = 32
Private Declare Function GetFileAttributesW Lib "kernel32" (ByVal lpFileName As Long) As Long

'If GDI+ is currently running, this will be set to a non-zero value.
Private m_GDIPlusToken As Long

Public Function GDIPlusLoadImage(ByVal srcFilename As String, ByRef dstDIB As pdDIB) As Boolean
' Using GDI+, load an image into a pdLayer class.  For brevity, things like embedded ICC profiles are ignored.

    ' Used to hold the return values of various GDI+ calls.  GDI+ functions return 0 if successful, and some error
    ' code if unsuccessful.  (See the Status enum at this link: http://www.jose.it-berater.org/gdiplus/reference/gdiplusenumerations.htm)
    Dim gdipReturn As GP_Result
    
    ' GDI+ requires a startup token for each application that accesses it.
    If (m_GDIPlusToken = 0) Then Exit Function
    
    ' Use GDI+ to load the image
    Dim hImage As Long
    gdipReturn = GdipLoadImageFromFile(StrPtr(srcFilename), hImage)
    
    If (gdipReturn = GP_OK) Then
        
        'Retrieve the image's size
        Dim imgWidth As Long, imgHeight As Long
        GdipGetImageWidth hImage, imgWidth
        GdipGetImageHeight hImage, imgHeight
        
        'Look for an alpha channel
        Dim imgHasAlpha As Boolean
        imgHasAlpha = False
        
        Dim imgPixelFormat As GP_PixelFormat
        GdipGetImagePixelFormat hImage, imgPixelFormat
        imgHasAlpha = ((imgPixelFormat And GP_PF_Alpha) <> 0)
        If (Not imgHasAlpha) Then imgHasAlpha = ((imgPixelFormat And GP_PF_PreMultAlpha) <> 0)
        
        'Create a blank layer with matching size and alpha channel
        If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
        If imgHasAlpha Then
            dstDIB.CreateBlank imgWidth, imgHeight, 32
        Else
            dstDIB.CreateBlank imgWidth, imgHeight, 24
        End If
        
        Dim copyBitmapData As GP_BitmapData, tmpRect As RECTL
        Dim hGraphics As Long
        
        'We now copy over image data in one of two ways.  If the image is 24bpp, our job is simple - use BitBlt and an hBitmap.
        ' 32bpp (including CMYK) images require a bit of extra work.
        If imgHasAlpha Then
        
            'Make sure the image is in 32bpp premultiplied ARGB format
            If (imgPixelFormat <> GP_PF_32bppPARGB) Then GdipCloneBitmapAreaI 0, 0, imgWidth, imgHeight, GP_PF_32bppPARGB, hImage, hImage
            
            'We are now going to copy the image's data directly into our destination DIB by using LockBits.  Very fast, and not much code!
            
            'Start by preparing a BitmapData variable with instructions on where GDI+ should paste the bitmap data
            With copyBitmapData
                .BD_Width = imgWidth
                .BD_Height = imgHeight
                .BD_PixelFormat = GP_PF_32bppPARGB
                .BD_Stride = dstDIB.GetDIBStride
                .BD_Scan0 = dstDIB.GetDIBPointer
            End With
            
            'Next, prepare a clipping rect
            With tmpRect
                .Left = 0
                .Top = 0
                .Right = imgWidth
                .Bottom = imgHeight
            End With
            
            'Use LockBits to perform the copy for us.
            GdipBitmapLockBits hImage, tmpRect, GP_BLM_UserInputBuf Or GP_BLM_Write Or GP_BLM_Read, GP_PF_32bppPARGB, copyBitmapData
            GdipBitmapUnlockBits hImage, copyBitmapData
            
            'Note alpha premultiplication state
            dstDIB.SetInitialAlphaPremultiplicationState True
        
        Else
        
            'Render the GDI+ image directly onto the newly created layer
            GdipCreateFromHDC dstDIB.GetDIBDC, hGraphics
            GdipDrawImageRect hGraphics, hImage, 0, 0, imgWidth, imgHeight
            GdipDeleteGraphics hGraphics
            
        End If
        
        GDIPlusLoadImage = True
        
    Else
        Debug.Print "WARNING!  GDIPlusLoadImage failed to load " & srcFilename & ".  GDI+ error code was # " & gdipReturn
        GDIPlusLoadImage = False
    End If
    
    'Release any remaining GDI+ handles and exit
    If (hImage <> 0) Then GdipDisposeImage hImage
    
End Function

Public Function GDIPlusSavePicture(ByRef srcDIB As pdDIB, ByVal dstFilename As String, ByVal imgFormat As GP_ImageFormat, Optional ByVal outputColorDepth As Long = 24, Optional ByVal JPEGQuality As Long = 85) As Boolean
' Save the contents of a pdLayer object to any format supported by GDI+.
' Additional save options are currently available for JPEGs (JPEG quality, range [1,100]).

    On Error GoTo GDIPlusSaveError

    ' GDI+ requires a startup token for each application that accesses it.
    If (m_GDIPlusToken <> 0) Then
    
        'If the output format is 24bpp (e.g. JPEG) but the input image is 32bpp, composite it against white
        If (srcDIB.GetDIBColorDepth <> 24) And imgFormat = GP_IF_JPEG Then srcDIB.ConvertTo24bpp
    
        'Begin by creating a generic bitmap header for the current DIB
        Dim imgHeader As BITMAPINFO
        With imgHeader.Header
            .Size = Len(imgHeader.Header)
            .Planes = 1
            .BitCount = srcDIB.GetDIBColorDepth
            .Width = srcDIB.GetDIBWidth
            .Height = -srcDIB.GetDIBHeight
        End With
    
        'Use GDI+ to create a GDI+-compatible bitmap
        Dim gdipReturn As GP_Result, hImage As Long
        
        'Debug.Print "Creating GDI+ image context..."
            
        'Different GDI+ calls are required for different color depths. GdipCreateBitmapFromGdiDib leads to a blank
        ' alpha channel for 32bpp images, so use GdipCreateBitmapFromScan0 in that case.
        If (srcDIB.GetDIBColorDepth = 32) Then
            gdipReturn = GdipCreateBitmapFromScan0(srcDIB.GetDIBWidth, srcDIB.GetDIBHeight, srcDIB.GetDIBWidth * 4, GP_PF_32bppPARGB, ByVal srcDIB.GetDIBPointer, hImage)
        Else
            gdipReturn = GdipCreateBitmapFromGdiDib(imgHeader, ByVal srcDIB.GetDIBPointer, hImage)
        End If
        
        If (gdipReturn = GP_OK) Then
            
            Dim gdipColorDepth As Long
            gdipColorDepth = outputColorDepth
            
            'Request an encoder from GDI+ based on the type passed to this routine
            Dim exportGuid(0 To 15) As Byte
            
            'GDI+ takes encoder parameters in a very particular sequential format:
            ' 4 byte long: number of encoder parameters
            ' (n) * LenB(GP_EncoderParameter): actual encoder parameters
            '
            'There's no easy way to create a variable-length struct like this in VB6, so instead, we create
            ' an array of encoder params, and as the final step before calling GDI+, we copy everything into
            ' a temporary byte array formatted per GDI+'s requirements.
            Dim numExportParams As Long
            Dim exportParams() As GP_EncoderParameter
            
            'Get the clsID for this encoder
            GetEncoderGUID imgFormat, VarPtr(exportGuid(0))
            
            Select Case imgFormat
                
                'BMP export
                Case GP_IF_BMP
                    
                    numExportParams = 1
                    ReDim exportParams(0 To numExportParams - 1) As GP_EncoderParameter
                    
                    With exportParams(0)
                        .EP_NumOfValues = 1
                        .EP_ValueType = GP_EVT_Long
                        CLSIDFromString StrPtr(GP_EP_ColorDepth), VarPtr(.EP_GUID(0))
                        .EP_ValuePtr = VarPtr(gdipColorDepth)
                    End With
            
                'GIF export
                Case GP_IF_GIF
                    
                    Dim gif_EncoderVersion As GP_EncoderValue
                    gif_EncoderVersion = GP_EV_VersionGif89
                    
                    numExportParams = 1
                    ReDim exportParams(0 To numExportParams - 1) As GP_EncoderParameter
                    
                    With exportParams(0)
                        .EP_NumOfValues = 1
                        .EP_ValueType = GP_EVT_Long
                        CLSIDFromString StrPtr(GP_EP_Version), VarPtr(.EP_GUID(0))
                        .EP_ValuePtr = VarPtr(gif_EncoderVersion)
                    End With
                    
                'JPEG export (requires extra work to specify a quality for the encode)
                Case GP_IF_JPEG
                    
                    numExportParams = 1
                    ReDim exportParams(0 To numExportParams - 1) As GP_EncoderParameter
                    
                    With exportParams(0)
                        .EP_NumOfValues = 1
                        .EP_ValueType = GP_EVT_Long
                        CLSIDFromString StrPtr(GP_EP_Quality), VarPtr(.EP_GUID(0))
                        .EP_ValuePtr = VarPtr(JPEGQuality)
                    End With
                
                'PNG export
                Case GP_IF_PNG
                            
                    numExportParams = 1
                    ReDim exportParams(0 To numExportParams - 1) As GP_EncoderParameter
                    
                    With exportParams(0)
                        .EP_NumOfValues = 1
                        .EP_ValueType = GP_EVT_Long
                        CLSIDFromString StrPtr(GP_EP_ColorDepth), VarPtr(.EP_GUID(0))
                        .EP_ValuePtr = VarPtr(gdipColorDepth)
                    End With
                    
                'TIFF export (requires extra work to specify compression and color depth for the encode)
                Case GP_IF_TIFF
                    
                    Dim TIFF_Compression As GP_EncoderValue
                    TIFF_Compression = GP_EV_CompressionLZW
                            
                    numExportParams = 2
                    ReDim exportParams(0 To numExportParams - 1) As GP_EncoderParameter
                    
                    With exportParams(0)
                        .EP_NumOfValues = 1
                        .EP_ValueType = GP_EVT_Long
                        CLSIDFromString StrPtr(GP_EP_Compression), VarPtr(.EP_GUID(0))
                        .EP_ValuePtr = VarPtr(TIFF_Compression)
                    End With
                    
                    With exportParams(1)
                        .EP_NumOfValues = 1
                        .EP_ValueType = GP_EVT_Long
                        CLSIDFromString StrPtr(GP_EP_ColorDepth), VarPtr(.EP_GUID(0))
                        .EP_ValuePtr = VarPtr(gdipColorDepth)
                    End With
            
            End Select
    
            'With our encoder prepared, we can finally continue with the save
            
            'Check to see if a file already exists at this location
            If FileExist(dstFilename) Then Kill dstFilename
            
            'Convert our list of params to a format GDI+ understands.
            Dim tmpEncodeParams() As Byte, tmpEncodeParamSize As Long
            If (numExportParams > 0) Then
                tmpEncodeParamSize = 4 + LenB(exportParams(0)) * numExportParams
            Else
                tmpEncodeParamSize = 4
            End If
            
            'First comes the number of parameters
            ReDim tmpEncodeParams(0 To tmpEncodeParamSize - 1) As Byte
            CopyMemoryStrict VarPtr(tmpEncodeParams(0)), VarPtr(numExportParams), 4&
            
            '...followed by each parameter in turn
            If (numExportParams > 0) Then
                Dim i As Long
                For i = 0 To numExportParams - 1
                    CopyMemoryStrict VarPtr(tmpEncodeParams(4)) + (i * LenB(exportParams(0))), VarPtr(exportParams(i)), LenB(exportParams(0))
                Next i
            End If
            
            'Pass all completed structs to GDI+ and let it handle everything from here
            gdipReturn = GdipSaveImageToFile(hImage, StrPtr(dstFilename), VarPtr(exportGuid(0)), VarPtr(tmpEncodeParams(0)))
            GDIPlusSavePicture = (gdipReturn = GP_OK)
            If (Not GDIPlusSavePicture) Then Debug.Print "WARNING!  GDI+ failed to save " & dstFilename & ".  Error # " & gdipReturn
            
        Else
            Debug.Print "GDI+ couldn't create temporary copy of source image; it may be too large."
            GDIPlusSavePicture = False
        End If
        
    Else
        Debug.Print "GDI+ couldn't be initialized."
        GDIPlusSavePicture = False
    End If
    
    If (hImage <> 0) Then GdipDisposeImage hImage
    Exit Function
    
GDIPlusSaveError:
    GDIPlusSavePicture = False
    
End Function

Private Function GetEncoderGUID(ByVal srcFormat As GP_ImageFormat, ByVal ptrToDstGuid As Long) As Boolean
' When exporting images, we need to find the unique GUID for a given exporter.  Matching via mimetype is a
' straightforward way to do this, and is the recommended solution from MSDN (see https://msdn.microsoft.com/en-us/library/ms533843(v=vs.85).aspx)
    
    GetEncoderGUID = False
    
    'Generate a matching mimetype for the given format
    Dim srcMimetype As String
    Select Case srcFormat
        Case GP_IF_BMP
            srcMimetype = "image/bmp"
        Case GP_IF_GIF
            srcMimetype = "image/gif"
        Case GP_IF_JPEG
            srcMimetype = "image/jpeg"
        Case GP_IF_PNG
            srcMimetype = "image/png"
        Case GP_IF_TIFF
            srcMimetype = "image/tiff"
        Case Else
            srcMimetype = vbNullString
    End Select
    
    If (LenB(srcMimetype) <> 0) Then
        
        'Start by retrieving the number of encoders, and the size of the full encoder list
        Dim numOfEncoders As Long, sizeOfEncoders As Long
        If (GdipGetImageEncodersSize(numOfEncoders, sizeOfEncoders) = GP_OK) Then
            
            If (numOfEncoders > 0) And (sizeOfEncoders > 0) Then
            
                Dim encoderBuffer() As Byte
                Dim tmpCodec As GP_ImageCodecInfo
                
                'Hypothetically, we could probably pull the encoder list directly into a GP_ImageCodecInfo() array,
                ' but I haven't tested to see if the byte values of the encoder sizes are exact.  To avoid any problems,
                ' let's just dump the return into a byte array, then parse out what we need as we go.
                ReDim encoderBuffer(0 To sizeOfEncoders - 1) As Byte
                If (GdipGetImageEncoders(numOfEncoders, sizeOfEncoders, VarPtr(encoderBuffer(0))) = GP_OK) Then
                
                    'Iterate through the encoder list, searching for a match
                    Dim i As Long
                    For i = 0 To numOfEncoders - 1
                    
                        'Extract this codec
                        CopyMemoryStrict VarPtr(tmpCodec), VarPtr(encoderBuffer(0)) + LenB(tmpCodec) * i, LenB(tmpCodec)
                        
                        'Compare mimetypes
                        If (StrComp(StringFromCharPtr(tmpCodec.IC_MimeType, True), srcMimetype, vbTextCompare) = 0) Then
                            GetEncoderGUID = True
                            CopyMemoryStrict ptrToDstGuid, VarPtr(tmpCodec.IC_ClassID(0)), 16&
                            Exit For
                        End If
                        
                    Next i
                
                End If
                
            End If
        End If
        
    End If

End Function

Public Sub GDIPlus_DrawCircle(ByVal dstDC As Long, ByVal centerx As Single, ByVal centery As Single, ByVal radius As Single, ByVal circleColor As Long, Optional ByVal circleLineWidth As Single = 1!, Optional ByVal useAntialiasing As Boolean = True, Optional ByVal circleOpacity As Single = 100!, Optional ByVal dashStyle As GP_DashStyle = GP_DS_Solid, Optional ByVal dashCapStyle As GP_DashCap = GP_DC_Round)
' Draw an antialiased line to an arbitrary hDC.  Coordinates *must* be in pixels (GDI+ doesn't understand twips).
    
    ' Wrap a GDI+ surface around the target DC
    Dim hGraphics As Long
    GdipCreateFromHDC dstDC, hGraphics
    
    ' Activate antialiasing on the surface
    If useAntialiasing Then GdipSetSmoothingMode hGraphics, GP_SM_Antialias Else GdipSetSmoothingMode hGraphics, GP_SM_None
    
    ' Create a pen (used to stroke the line)
    Dim hPen As Long
    If (GdipCreatePen1(FillQuadWithVBRGB(circleColor, circleOpacity * 2.55!), circleLineWidth, GP_U_Pixel, hPen) = GP_OK) Then
    
        'Set the dash style, if any
        GdipSetPenLineCap hPen, GP_LC_Round, GP_LC_Round, dashCapStyle
        GdipSetPenDashStyle hPen, dashStyle
        
        'Render the circle
        GdipDrawEllipse hGraphics, hPen, centerx - radius, centery - radius, radius * 2!, radius * 2!
        
    Else
        Debug.Print "WARNING! GDIPlus_Drawline failed to create a pen."
    End If
    
    ' Free all intermediary objects
    If (hPen <> 0) Then GdipDeletePen hPen
    If (hGraphics <> 0) Then GdipDeleteGraphics hGraphics
    
End Sub

Public Sub GDIPlus_DrawLine(ByVal dstDC As Long, ByVal X1 As Single, ByVal Y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal lineColor As Long, Optional ByVal lineWidth As Single = 1!, Optional ByVal useAntialiasing As Boolean = True, Optional ByVal lineOpacity As Single = 100!, Optional ByVal lineCap As GP_LineCap = GP_LC_Round, Optional ByVal dashStyle As GP_DashStyle = GP_DS_Solid, Optional ByVal dashCapStyle As GP_DashCap = GP_DC_Round)
' Draw an antialiased line to an arbitrary hDC.  Coordinates *must* be in pixels (GDI+ doesn't understand twips).
    
    'Wrap a GDI+ surface around the target DC
    Dim hGraphics As Long
    GdipCreateFromHDC dstDC, hGraphics
    
    'Activate antialiasing on the surface
    If useAntialiasing Then GdipSetSmoothingMode hGraphics, GP_SM_Antialias Else GdipSetSmoothingMode hGraphics, GP_SM_None
    
    'Create a pen (used to stroke the line)
    Dim hPen As Long
    If (GdipCreatePen1(FillQuadWithVBRGB(lineColor, lineOpacity * 2.55!), lineWidth, GP_U_Pixel, hPen) = GP_OK) Then
    
        'Set the line cap and dash style, if any
        GdipSetPenLineCap hPen, lineCap, lineCap, dashCapStyle
        GdipSetPenDashStyle hPen, dashStyle
        
        'Render the line
        GdipDrawLine hGraphics, hPen, X1, Y1, x2, y2
        
    Else
        Debug.Print "WARNING! GDIPlus_Drawline failed to create a pen."
    End If
    
    'Free all intermediary objects
    If (hPen <> 0) Then GdipDeletePen hPen
    If (hGraphics <> 0) Then GdipDeleteGraphics hGraphics

End Sub

Public Sub GDIPlus_PaintPicture(ByVal dstDC As Long, ByVal dstX As Single, ByVal dstY As Single, ByVal dstWidth As Single, ByVal dstHeight As Single, ByVal srcDC As Long, ByVal srcX As Single, ByVal srcY As Single, ByVal srcWidth As Single, ByVal srcHeight As Single, Optional ByVal resizeQuality As GP_InterpolationMode = GP_IM_HighQualityBicubic)
    
    ' Wrap a GDI+ surface around the target DC
    Dim hGraphics As Long
    GdipCreateFromHDC dstDC, hGraphics
    
    ' Wrap a GDI+ bitmap around the source DC, using an intermediary DIB as required
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    srcDIB.CreateBlank Int(srcWidth + 0.9999), Int(srcHeight + 0.9999)
    BitBlt srcDIB.GetDIBDC, 0, 0, srcWidth, srcHeight, srcDC, srcX, srcY, vbSrcCopy
    
    Dim imgHeader As BITMAPINFO
    With imgHeader.Header
        .Size = Len(imgHeader.Header)
        .Planes = 1
        .BitCount = srcDIB.GetDIBColorDepth
        .Width = srcDIB.GetDIBWidth
        .Height = -srcDIB.GetDIBHeight
    End With
    
    Dim gdipReturn As GP_Result, hImage As Long
    gdipReturn = GdipCreateBitmapFromGdiDib(imgHeader, ByVal srcDIB.GetDIBPointer, hImage)
    
    ' Request the smoothing mode we were passed
    GdipSetInterpolationMode hGraphics, resizeQuality
    
    'Perform the render
    gdipReturn = GdipDrawImageRectRect(hGraphics, hImage, dstX, dstY, dstWidth, dstHeight, 0, 0, srcWidth, srcHeight, GP_U_Pixel)
    
    GdipDisposeImage hImage
    GdipDeleteGraphics hGraphics
        
End Sub

Public Sub GDIPlus_PaintPictureRotated(ByVal dstDC As Long, ByVal dstCenterX As Single, ByVal dstCenterY As Single, ByVal dstAngle As Single, ByVal dstWidth As Single, ByVal dstHeight As Single, ByVal srcDC As Long, ByVal srcX As Single, ByVal srcY As Single, ByVal srcWidth As Single, ByVal srcHeight As Single, Optional ByVal resizeQuality As GP_InterpolationMode = GP_IM_HighQualityBicubic)
        
    ' Wrap a GDI+ surface around the target DC
    Dim hGraphics As Long
    GdipCreateFromHDC dstDC, hGraphics
    
    ' Wrap a GDI+ bitmap around the source DC, using an intermediary DIB as required
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    srcDIB.CreateBlank Int(srcWidth + 0.9999), Int(srcHeight + 0.9999)
    BitBlt srcDIB.GetDIBDC, 0, 0, srcWidth, srcHeight, srcDC, srcX, srcY, vbSrcCopy
    
    Dim imgHeader As BITMAPINFO
    With imgHeader.Header
        .Size = Len(imgHeader.Header)
        .Planes = 1
        .BitCount = srcDIB.GetDIBColorDepth
        .Width = srcDIB.GetDIBWidth
        .Height = -srcDIB.GetDIBHeight
    End With
    
    Dim gdipReturn As GP_Result, hImage As Long
    gdipReturn = GdipCreateBitmapFromGdiDib(imgHeader, ByVal srcDIB.GetDIBPointer, hImage)
    
    ' Request the smoothing mode we were passed
    GdipSetInterpolationMode hGraphics, resizeQuality
    
    ' Create a table of corner-points for the destination image
    Dim dstPoints() As PointFloat
    ReDim dstPoints(0 To 2) As PointFloat
    dstPoints(0).x = dstCenterX - dstWidth / 2!
    dstPoints(0).Y = dstCenterY - dstHeight / 2!
    dstPoints(1).x = dstCenterX + dstWidth / 2!
    dstPoints(1).Y = dstCenterY - dstHeight / 2!
    dstPoints(2).x = dstCenterX - dstWidth / 2!
    dstPoints(2).Y = dstCenterY + dstHeight / 2!
    
    Dim centerPoint As PointFloat
    centerPoint.x = dstCenterX
    centerPoint.Y = dstCenterY
    
    ' Rotate all points about the center
    Dim i As Long
    For i = 0 To 2
        dstPoints(i) = PictureSnapRotateVertex(dstPoints(i), centerPoint, dstAngle)
    Next i
    
    'Perform the render
    gdipReturn = GdipDrawImagePointsRect(hGraphics, hImage, VarPtr(dstPoints(0)), 3, 0!, 0!, srcWidth, srcHeight, GP_U_Pixel)
    
    GdipDisposeImage hImage
    GdipDeleteGraphics hGraphics
    
End Sub

Public Sub GDIPlus_PaintPictureFast(ByVal dstDC As Long, ByVal dstX As Single, ByVal dstY As Single, ByVal dstWidth As Single, ByVal dstHeight As Single, ByRef srcDIB As pdDIB, ByVal srcX As Single, ByVal srcY As Single, ByVal srcWidth As Single, ByVal srcHeight As Single, Optional ByVal resizeQuality As GP_InterpolationMode = GP_IM_HighQualityBicubic)
' If you are drawing from a DIB to a picture box, painting can be greatly accelerated as a temporary image copy
' is no longer required.  The "source" image parameter is the only change in this function.
    
    ' Wrap a GDI+ surface around the target DC
    Dim hGraphics As Long
    GdipCreateFromHDC dstDC, hGraphics
    
    ' Wrap a GDI+ bitmap around the source image using a shortcut function (and no intermediary memory!)
    Dim gdipReturn As GP_Result, hImage As Long
    GetGdipBitmapHandleFromDIB hImage, srcDIB
    
    ' Request the smoothing mode we were passed
    GdipSetInterpolationMode hGraphics, resizeQuality
    
    ' Perform the render
    gdipReturn = GdipDrawImageRectRect(hGraphics, hImage, dstX, dstY, dstWidth, dstHeight, srcX, srcY, srcWidth, srcHeight, GP_U_Pixel)
    
    GdipDisposeImage hImage
    GdipDeleteGraphics hGraphics
    
End Sub

Public Sub GDIPlus_PaintPictureFastRotated(ByVal dstDC As Long, ByVal dstCenterX As Single, ByVal dstCenterY As Single, ByVal dstAngle As Single, ByVal dstWidth As Single, ByVal dstHeight As Single, ByRef srcDIB As pdDIB, ByVal srcX As Single, ByVal srcY As Single, ByVal srcWidth As Single, ByVal srcHeight As Single, Optional ByVal resizeQuality As GP_InterpolationMode = GP_IM_HighQualityBicubic)
' If you are drawing from a DIB to a picture box, painting can be greatly accelerated as a temporary image copy
' is no longer required.  The "source" image parameter is the only change in this function.
    
    ' Wrap a GDI+ surface around the target DC
    Dim hGraphics As Long
    If (GdipCreateFromHDC(dstDC, hGraphics) = GP_OK) Then
    
        ' Wrap a GDI+ bitmap around the source DC, using an intermediary DIB as required
        Dim gdipReturn As GP_Result, hImage As Long
        If GetGdipBitmapHandleFromDIB(hImage, srcDIB) Then
        
            'Request the smoothing mode we were passed
            GdipSetInterpolationMode hGraphics, resizeQuality
            
            'Create a table of corner-points for the destination image
            Dim dstPoints() As PointFloat
            ReDim dstPoints(0 To 2) As PointFloat
            dstPoints(0).x = dstCenterX - dstWidth / 2!
            dstPoints(0).Y = dstCenterY - dstHeight / 2!
            dstPoints(1).x = dstCenterX + dstWidth / 2!
            dstPoints(1).Y = dstCenterY - dstHeight / 2!
            dstPoints(2).x = dstCenterX - dstWidth / 2!
            dstPoints(2).Y = dstCenterY + dstHeight / 2!
            
            Dim centerPoint As PointFloat
            centerPoint.x = dstCenterX
            centerPoint.Y = dstCenterY
            
            'Rotate all points about the center
            Dim i As Long
            For i = 0 To 2
                dstPoints(i) = PictureSnapRotateVertex(dstPoints(i), centerPoint, dstAngle)
            Next i
            
            'Perform the render
            gdipReturn = GdipDrawImagePointsRect(hGraphics, hImage, VarPtr(dstPoints(0)), 3, srcX, srcY, srcWidth, srcHeight, GP_U_Pixel)
            
            GdipDisposeImage hImage
            
        End If
        
        GdipDeleteGraphics hGraphics
        
    End If
    
End Sub

Private Function PictureSnapRotateVertex(vCorner As PointFloat, vOrigin As PointFloat, rotation As Single) As PointFloat
' Calculate the rotated corners of an unrotated rectangle (rotation is passed in degrees)
' NOTE: copy of John's corresponding function in CodePictureSnapPOVAnnotations; the only difference is using
' the standard GDI+ PointF type

ierror = False
On Error GoTo PictureSnapRotateVertexError

Dim arad As Single
    
arad! = rotation! * PI! / 180
    
PictureSnapRotateVertex.x = ((vCorner.x - vOrigin.x) * Cos(arad!) - (vCorner.Y - vOrigin.Y) * Sin(arad)) + vOrigin.x
PictureSnapRotateVertex.Y = ((vCorner.Y - vOrigin.Y) * Cos(arad!) + (vCorner.x - vOrigin.x) * Sin(arad)) + vOrigin.Y

Exit Function

' Errors
PictureSnapRotateVertexError:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapRotateVertex"
ierror = True
Exit Function

End Function

Public Function GetGdipBitmapHandleFromDIB(ByRef dstBitmapHandle As Long, ByRef srcDIB As pdDIB) As Boolean
' Simpler shorthand function for obtaining a GDI+ bitmap handle from a pdDIB object.  Note that 24/32bpp cases have to be
' handled separately because GDI+ is unpredictable at automatically detecting color depth with 32-bpp DIBs.  (This behavior
' is forgivable, given GDI's unreliable handling of alpha bytes.)
    
    If (srcDIB Is Nothing) Then Exit Function
    
    If (srcDIB.GetDIBColorDepth = 32) Then
        GetGdipBitmapHandleFromDIB = (GdipCreateBitmapFromScan0(srcDIB.GetDIBWidth, srcDIB.GetDIBHeight, srcDIB.GetDIBWidth * 4, GP_PF_32bppPARGB, ByVal srcDIB.GetDIBPointer, dstBitmapHandle) = GP_OK)
    Else
    
        'Use GdipCreateBitmapFromGdiDib for 24bpp DIBs
        Dim imgHeader As BITMAPINFO
        With imgHeader.Header
            .Size = Len(imgHeader.Header)
            .Planes = 1
            .BitCount = srcDIB.GetDIBColorDepth
            .Width = srcDIB.GetDIBWidth
            .Height = -srcDIB.GetDIBHeight
        End With
        GetGdipBitmapHandleFromDIB = (GdipCreateBitmapFromGdiDib(imgHeader, ByVal srcDIB.GetDIBPointer, dstBitmapHandle) = GP_OK)
        
    End If

End Function

Public Function StartGDIPlus() As Boolean
' Start GDI+.  For better performance, you could simply call this function once, when your app starts.
' (Just remember to call the shutdown function before your app ends.)
'
' RETURNS: TRUE if GDI+ is running; FALSE if startup failed
    
    If (m_GDIPlusToken = 0) Then
    
        Dim gdiCheck As GDIPlusStartupInput
        gdiCheck.GDIPlusVersion = 1
        
        'Initialize a GDI+ interface
        StartGDIPlus = (GdiplusStartup(m_GDIPlusToken, gdiCheck) = GP_OK)
        If (Not StartGDIPlus) Then Debug.Print "WARNING!  Could not start GDI+!"
        
    Else
        Debug.Print "WARNING!  GDI+ is already running - you don't need to call StartGDIPlus()!"
    End If
    
End Function

Public Sub StopGDIPlus()
' Shutdown GDI+.  See notes on StartGDIPlus, above, for additional details.
    If (m_GDIPlusToken <> 0) Then
        GdiplusShutdown m_GDIPlusToken
        m_GDIPlusToken = 0
    Else
        Debug.Print "WARNING!  GDI+ was never started - how can I shut it down?!"
    End If
End Sub

Public Function StringFromCharPtr(ByVal srcPointer As Long, Optional ByVal srcStringIsUnicode As Boolean = True, Optional ByVal maxLength As Long = -1) As String
' Given an arbitrary pointer to a null-terminated CHAR or WCHAR run, measure the resulting string and copy the results
' into a VB string.
'
' For security reasons, if an upper limit of the string's length is known in advance (e.g. MAX_PATH), pass that limit
' via the optional maxLength parameter to avoid a buffer overrun.  This function has a hard-coded limit of 65k chars,
' a limit you can easily lift but which makes sense for most software.  If a string exceeds the limit (whether passed
' or hard-coded), *a string will still be created and returned*, but it will be clamped to the max length.
    
    'Check string length
    Dim strLength As Long
    If srcStringIsUnicode Then strLength = lstrlenW(srcPointer) Else strLength = lstrlenA(srcPointer)
    
    'Make sure the length/pointer isn't null
    If (strLength <= 0) Then
        StringFromCharPtr = vbNullString
    Else
        
        'Make sure the string's length is valid.
        Dim maxAllowedLength As Long
        If (maxLength = -1) Then maxAllowedLength = 65535 Else maxAllowedLength = maxLength
        If (strLength > maxAllowedLength) Then strLength = maxAllowedLength
        
        'Create the target string and copy the bytes over
        If srcStringIsUnicode Then
            StringFromCharPtr = String$(strLength, 0)
            CopyMemoryStrict StrPtr(StringFromCharPtr), srcPointer, strLength * 2
        Else
            StringFromCharPtr = SysAllocStringByteLen(srcPointer, strLength)
        End If
    
    End If
    
End Function

Private Function FileExist(ByRef fName As String) As Boolean
' Returns a boolean as to whether or not a given file exists
    Select Case (GetFileAttributesW(StrPtr(fName)) And vbDirectory) = 0
        Case True: FileExist = True
        Case Else: FileExist = (Err.LastDllError = ERROR_SHARING_VIOLATION)
    End Select
End Function

Public Function FillQuadWithVBRGB(ByVal vbRGB As Long, ByVal alphaValue As Byte) As Long
' GDI+ requires RGBQUAD colors with alpha in the 4th byte.  This function returns an RGBQUAD (long-type) from
' a standard RGB() long and supplied alpha (on the range [0, 255]).
    
    ' The vbRGB constant may be an OLE color constant; if that happens, we want to convert it to a normal RGB quad.
    vbRGB = TranslateColor(vbRGB)
    
    Dim dstQuad As RGBQuad
    dstQuad.Red = ExtractRed(vbRGB)
    dstQuad.Green = ExtractGreen(vbRGB)
    dstQuad.Blue = ExtractBlue(vbRGB)
    dstQuad.Alpha = alphaValue
    
    Dim placeHolder As tmpLong
    LSet placeHolder = dstQuad
    
    FillQuadWithVBRGB = placeHolder.lngResult
    
End Function

Private Function TranslateColor(ByVal colorRef As Long) As Long
' Translate an OLE color to an RGB Long.  Note that the API function returns -1 on failure; if this happens, we return white.
    If OleTranslateColor(colorRef, 0, TranslateColor) Then TranslateColor = vbWhite
End Function

Public Function ExtractRed(ByVal srcColor As Long) As Integer
' Helper color functions for moving individual RGB components between RGB() Longs.  Note that these functions only
' return values in the range [0, 255], but declaring them as integers prevents overflow during intermediary math steps.
    ExtractRed = srcColor And 255
End Function

Public Function ExtractGreen(ByVal srcColor As Long) As Integer
    ExtractGreen = (srcColor \ 256) And 255
End Function

Public Function ExtractBlue(ByVal srcColor As Long) As Integer
    ExtractBlue = (srcColor \ 65536) And 255
End Function
