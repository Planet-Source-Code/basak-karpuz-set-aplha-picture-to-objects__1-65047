Attribute VB_Name = "Module1"
Option Explicit

Private Type PICTUREPROPERTIES
    pType As Long
    pWidth As Long
    pHeight As Long
    pWidthBytes As Long
    pPlanes As Integer
    pBitsPixel As Integer
    pBits As Long
End Type

Private Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type

Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdcDest&, ByVal nXOriginDest&, ByVal nYOriginDest&, ByVal nWidthDest&, ByVal nHeightDest&, ByVal hdcSrc&, ByVal nXOriginSrc&, ByVal nYOriginSrc&, ByVal nWidthSrc&, ByVal nHeightSrc&, ByVal lBlendFunction As Long) As Long
Private Declare Function CreateCompatibleDC& Lib "gdi32" (ByVal hDC&)
Private Declare Function DeleteDC& Lib "gdi32" (ByVal hDC&)
Private Declare Function DeleteObject& Lib "gdi32" (ByVal hObject&)
Private Declare Function GetObject& Lib "gdi32" Alias "GetObjectA" (ByVal hObject&, ByVal nCount&, lpObject As Any)
Private Declare Function SelectObject& Lib "gdi32" (ByVal hDC&, ByVal hObject&)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Public Sub PPic(rObj As Object, Picture_Path$, Alpha_Value%)

    Dim hSrcDC&, hBF&, BF As BLENDFUNCTION, PicProp As PICTUREPROPERTIES

    If Dir(Picture_Path) = vbNullString Then Exit Sub
    If 0 > Alpha_Value Or Alpha_Value > 255 Then Exit Sub

    rObj.AutoRedraw = True
    rObj.Cls

    Call GetObject(LoadPicture(Picture_Path).Handle, Len(PicProp), PicProp)
    hSrcDC = CreateCompatibleDC(rObj.hDC)
    Call SelectObject(hSrcDC, LoadPicture(Picture_Path).Handle)
    BF.SourceConstantAlpha = Alpha_Value
    Call CopyMemory(hBF, BF, 4)
    Call AlphaBlend(rObj.hDC, 0, 0, PicProp.pWidth, PicProp.pHeight, hSrcDC, 0, 0, PicProp.pWidth, PicProp.pHeight, hBF)
    Call DeleteDC(hSrcDC)

End Sub

