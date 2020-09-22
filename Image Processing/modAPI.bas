Attribute VB_Name = "modAPI"
'WINDOWS API DECLARATIONS
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Const conHwndTopmost = -1
Const conHwndNoTopmost = -2
Const conSwpNoActivate = &H10
Const conSwpShowWindow = &H40

'DECLARES PUBLIC VARIABLES
Public indeks As Integer
Public picforms(0 To 99) As New frmPicture
Public fpath As String
Public hBMPSour(0 To 99) As Long
Public hDCSour(0 To 99) As Long
Public hBMPDest(0 To 99) As Long
Public hDCDest(0 To 99) As Long
Public iCancel As Boolean
Public currDir As String

Function GetRed(cValue As Long) As Long 'A function that is used to get RED value
    GetRed = cValue Mod 256
End Function

Function GetGreen(cValue As Long) As Long   'A function that is used to get GREEN value
    GetGreen = Int((cValue / 256)) Mod 256
End Function

Function GetBlue(cValue As Long) As Long    'A function that is used to get BLUE value
    GetBlue = Int(cValue / 65536)
End Function

Sub Balancing(pfIndex As Integer)
    Dim pX As Long, pY As Long
    Dim x As Long, y As Long
    Dim colorval As Long
    Dim red As Long, green As Long, blue As Long
    Dim red2 As Long, green2 As Long, blue2 As Long
    
    pX = mdiImgProcess.ActiveForm.pcbPicture.ScaleWidth - 1
    pY = mdiImgProcess.ActiveForm.pcbPicture.ScaleHeight - 1
    For x = 0 To pX
        For y = 0 To pY
            colorval = GetPixel(hDCSour(pfIndex), x, y)
            red = GetRed(colorval)
            green = GetGreen(colorval)
            blue = GetBlue(colorval)
            
            red2 = red + Int(frmBalance.mRedVal / 100 * red)
            green2 = green + Int(frmBalance.mGreenVal / 100 * green)
            blue2 = blue + Int(frmBalance.mBlueVal / 100 * blue)
            
            If red2 >= 255 Then red2 = 255
            If green2 >= 255 Then green2 = 255
            If blue2 >= 255 Then blue2 = 255
            If red2 <= 0 Then red2 = 0
            If green2 <= 0 Then green2 = 0
            If blue2 <= 0 Then blue2 = 0
            
            SetPixel hDCDest(pfIndex), x, y, RGB(red2, green2, blue2)
        Next y
        If (x Mod (pX / 20)) = 0 Then frmBalance.ProgressBar.Value = x / pX * 100
    Next x
    BitBlt mdiImgProcess.ActiveForm.pcbPicture.hdc, 0, 0, mdiImgProcess.ActiveForm.pcbPicture.ScaleWidth, mdiImgProcess.ActiveForm.pcbPicture.ScaleHeight, hDCDest(pfIndex), 0, 0, vbSrcCopy
    mdiImgProcess.ActiveForm.pcbPicture.Refresh
    If frmBalance.Preview = False Then BitBlt hDCSour(pfIndex), 0, 0, mdiImgProcess.ActiveForm.pcbPicture.ScaleWidth, mdiImgProcess.ActiveForm.pcbPicture.ScaleHeight, mdiImgProcess.ActiveForm.pcbPicture.hdc, 0, 0, vbSrcCopy
End Sub

Sub Brightness(pfIndex As Integer)
    Dim pX As Long, pY As Long
    Dim x As Long, y As Long
    Dim colorval As Long
    Dim red As Long, green As Long, blue As Long
    Dim red2 As Long, green2 As Long, blue2 As Long
    
    pX = mdiImgProcess.ActiveForm.pcbPicture.ScaleWidth - 1
    pY = mdiImgProcess.ActiveForm.pcbPicture.ScaleHeight - 1
    For x = 0 To pX
        For y = 0 To pY
            colorval = GetPixel(hDCSour(pfIndex), x, y)
            red = GetRed(colorval)
            green = GetGreen(colorval)
            blue = GetBlue(colorval)
            
            red2 = red + frmBright.vBright
            green2 = green + frmBright.vBright
            blue2 = blue + frmBright.vBright
            
            If red2 >= 255 Then red2 = 255
            If green2 >= 255 Then green2 = 255
            If blue2 >= 255 Then blue2 = 255
            If red2 <= 0 Then red2 = 0
            If green2 <= 0 Then green2 = 0
            If blue2 <= 0 Then blue2 = 0
            
            SetPixel hDCDest(pfIndex), x, y, RGB(red2, green2, blue2)
        Next y
        If (x Mod (pX / 20)) = 0 Then frmBright.ProgressBar.Value = x / pX * 100
    Next x
    BitBlt mdiImgProcess.ActiveForm.pcbPicture.hdc, 0, 0, mdiImgProcess.ActiveForm.pcbPicture.ScaleWidth, mdiImgProcess.ActiveForm.pcbPicture.ScaleHeight, hDCDest(pfIndex), 0, 0, vbSrcCopy
    mdiImgProcess.ActiveForm.pcbPicture.Refresh
    If frmBright.Preview = False Then BitBlt hDCSour(pfIndex), 0, 0, mdiImgProcess.ActiveForm.pcbPicture.ScaleWidth, mdiImgProcess.ActiveForm.pcbPicture.ScaleHeight, mdiImgProcess.ActiveForm.pcbPicture.hdc, 0, 0, vbSrcCopy
End Sub

Sub Grayscaling(pfIndex As Integer)
    Dim pX As Long, pY As Long
    Dim x As Long, y As Long
    Dim colorval As Long
    Dim red As Long, green As Long, blue As Long
    Dim red2 As Long, green2 As Long, blue2 As Long
    
    pX = mdiImgProcess.ActiveForm.pcbPicture.ScaleWidth - 1
    pY = mdiImgProcess.ActiveForm.pcbPicture.ScaleHeight - 1
    Load frmProgress
    frmProgress.Caption = "Grayscaling..."
    frmProgress.Show
    DoEvents
    For x = 0 To pX
        For y = 0 To pY
            colorval = GetPixel(hDCSour(pfIndex), x, y)
            red = GetRed(colorval)
            green = GetGreen(colorval)
            blue = GetBlue(colorval)
            
            red2 = Int((red + green + blue) / 3)
            green2 = red2
            blue2 = red2
            
            If red2 >= 255 Then red2 = 255
            If green2 >= 255 Then green2 = 255
            If blue2 >= 255 Then blue2 = 255
            If red2 <= 0 Then red2 = 0
            If green2 <= 0 Then green2 = 0
            If blue2 <= 0 Then blue2 = 0
            
            SetPixel hDCDest(pfIndex), x, y, RGB(red2, green2, blue2)
        Next y
        If (x Mod (pX / 20)) = 0 Then frmProgress.ProgressBar.Value = x / pX * 100
    Next x
    BitBlt mdiImgProcess.ActiveForm.pcbPicture.hdc, 0, 0, mdiImgProcess.ActiveForm.pcbPicture.ScaleWidth, mdiImgProcess.ActiveForm.pcbPicture.ScaleHeight, hDCDest(pfIndex), 0, 0, vbSrcCopy
    mdiImgProcess.ActiveForm.pcbPicture.Refresh
    BitBlt hDCSour(pfIndex), 0, 0, mdiImgProcess.ActiveForm.pcbPicture.ScaleWidth, mdiImgProcess.ActiveForm.pcbPicture.ScaleHeight, mdiImgProcess.ActiveForm.pcbPicture.hdc, 0, 0, vbSrcCopy
    Unload frmProgress
End Sub

Sub Inverting(pfIndex As Integer)
    Dim pX As Long, pY As Long
    Dim x As Long, y As Long
    Dim colorval As Long
    Dim red As Long, green As Long, blue As Long
    Dim red2 As Long, green2 As Long, blue2 As Long
    
    pX = mdiImgProcess.ActiveForm.pcbPicture.ScaleWidth - 1
    pY = mdiImgProcess.ActiveForm.pcbPicture.ScaleHeight - 1
    Load frmProgress
    frmProgress.Caption = "Inverting..."
    frmProgress.Show
    DoEvents
    For x = 0 To pX
        For y = 0 To pY
            colorval = GetPixel(hDCSour(pfIndex), x, y)
            red = GetRed(colorval)
            green = GetGreen(colorval)
            blue = GetBlue(colorval)
            
            red2 = 255 - red
            green2 = 255 - green
            blue2 = 255 - blue
            
            If red2 >= 255 Then red2 = 255
            If green2 >= 255 Then green2 = 255
            If blue2 >= 255 Then blue2 = 255
            If red2 <= 0 Then red2 = 0
            If green2 <= 0 Then green2 = 0
            If blue2 <= 0 Then blue2 = 0
            
            SetPixel hDCDest(pfIndex), x, y, RGB(red2, green2, blue2)
        Next y
        If (x Mod (pX / 20)) = 0 Then frmProgress.ProgressBar.Value = x / pX * 100
    Next x
    BitBlt mdiImgProcess.ActiveForm.pcbPicture.hdc, 0, 0, mdiImgProcess.ActiveForm.pcbPicture.ScaleWidth, mdiImgProcess.ActiveForm.pcbPicture.ScaleHeight, hDCDest(pfIndex), 0, 0, vbSrcCopy
    mdiImgProcess.ActiveForm.pcbPicture.Refresh
    BitBlt hDCSour(pfIndex), 0, 0, mdiImgProcess.ActiveForm.pcbPicture.ScaleWidth, mdiImgProcess.ActiveForm.pcbPicture.ScaleHeight, mdiImgProcess.ActiveForm.pcbPicture.hdc, 0, 0, vbSrcCopy
    Unload frmProgress
End Sub

Sub Softening(pfIndex As Integer)
    Dim pX As Long, pY As Long
    Dim x As Long, y As Long
    Dim colorval(8) As Long
    Dim red(8) As Long, green(8) As Long, blue(8) As Long
    Dim red2 As Long, green2 As Long, blue2 As Long
    Dim i As Integer
    
    pX = mdiImgProcess.ActiveForm.pcbPicture.ScaleWidth - 1
    pY = mdiImgProcess.ActiveForm.pcbPicture.ScaleHeight - 1
    Load frmProgress
    frmProgress.Caption = "Softening..."
    frmProgress.Show
    DoEvents
    For x = 1 To pX
        For y = 1 To pY
            colorval(0) = GetPixel(hDCSour(pfIndex), x - 1, y - 1)
            colorval(1) = GetPixel(hDCSour(pfIndex), x - 1, y)
            colorval(2) = GetPixel(hDCSour(pfIndex), x - 1, y + 1)
            colorval(3) = GetPixel(hDCSour(pfIndex), x, y - 1)
            colorval(4) = GetPixel(hDCSour(pfIndex), x, y)
            colorval(5) = GetPixel(hDCSour(pfIndex), x, y + 1)
            colorval(6) = GetPixel(hDCSour(pfIndex), x + 1, y - 1)
            colorval(7) = GetPixel(hDCSour(pfIndex), x + 1, y)
            colorval(8) = GetPixel(hDCSour(pfIndex), x + 1, y + 1)
                'Get color value in 3x3 pixels box
            For i = 0 To 8
                red(i) = GetRed(colorval(i))
                green(i) = GetGreen(colorval(i))
                blue(i) = GetBlue(colorval(i))

                red2 = red2 + red(i)
                green2 = green2 + green(i)
                blue2 = blue2 + blue(i)
            Next i
            
            red2 = Int(red2 / 9)
            green2 = Int(green2 / 9)
            blue2 = Int(blue2 / 9)
            
            If red2 >= 255 Then red2 = 255
            If green2 >= 255 Then green2 = 255
            If blue2 >= 255 Then blue2 = 255
            If red2 <= 0 Then red2 = 0
            If green2 <= 0 Then green2 = 0
            If blue2 <= 0 Then blue2 = 0
            
            SetPixel hDCDest(pfIndex), x, y, RGB(red2, green2, blue2)
        Next y
        If (x Mod (pX / 20)) = 0 Then frmProgress.ProgressBar.Value = x / pX * 100
    Next x
    BitBlt mdiImgProcess.ActiveForm.pcbPicture.hdc, 0, 0, mdiImgProcess.ActiveForm.pcbPicture.ScaleWidth, mdiImgProcess.ActiveForm.pcbPicture.ScaleHeight, hDCDest(pfIndex), 0, 0, vbSrcCopy
    mdiImgProcess.ActiveForm.pcbPicture.Refresh
    BitBlt hDCSour(pfIndex), 0, 0, mdiImgProcess.ActiveForm.pcbPicture.ScaleWidth, mdiImgProcess.ActiveForm.pcbPicture.ScaleHeight, mdiImgProcess.ActiveForm.pcbPicture.hdc, 0, 0, vbSrcCopy
    Unload frmProgress
End Sub

Sub Lighting(pfIndex As Integer)
    Dim pX As Long, pY As Long
    Dim x As Long, y As Long
    Dim colorval As Long
    Dim red As Long, green As Long, blue As Long
    Dim red2 As Long, green2 As Long, blue2 As Long
    Dim rval As Long
    Dim intensity As Single
    
    pX = mdiImgProcess.ActiveForm.pcbPicture.ScaleWidth - 1
    pY = mdiImgProcess.ActiveForm.pcbPicture.ScaleHeight - 1
    Load frmProgress
    frmProgress.Caption = "Lighting..."
    frmProgress.Show
    DoEvents
    Randomize Timer
    For x = 0 To pX
        rval = pY + 1 - x
        i = 0
        For y = 0 To pY
            i = i + pX / 90
            rval = rval - 1 * Sin(i * 3.14 / 180)
            colorval = GetPixel(hDCSour(pfIndex), x, y)
            red = GetRed(colorval)
            green = GetGreen(colorval)
            blue = GetBlue(colorval)
            intensity = 0.25     'Increase this value for smaller picture
                                 'Decrease this value for larger picture
            red2 = Int(red / 2 + rval * intensity)
            green2 = Int(green / 2 + rval * intensity)
            blue2 = Int(blue / 2 + rval * intensity)
            
            If red2 >= 255 Then red2 = 255
            If green2 >= 255 Then green2 = 255
            If blue2 >= 255 Then blue2 = 255
            If red2 <= 0 Then red2 = 0
            If green2 <= 0 Then green2 = 0
            If blue2 <= 0 Then blue2 = 0
            
            SetPixel hDCDest(pfIndex), x, y, RGB(red2, green2, blue2)
            If rval <= 50 Then rval = 50
            If i >= 90 Then i = 90
        Next y
        If (x Mod (pX / 20)) = 0 Then frmProgress.ProgressBar.Value = x / pX * 100
    Next x
    BitBlt mdiImgProcess.ActiveForm.pcbPicture.hdc, 0, 0, mdiImgProcess.ActiveForm.pcbPicture.ScaleWidth, mdiImgProcess.ActiveForm.pcbPicture.ScaleHeight, hDCDest(pfIndex), 0, 0, vbSrcCopy
    mdiImgProcess.ActiveForm.pcbPicture.Refresh
    BitBlt hDCSour(pfIndex), 0, 0, mdiImgProcess.ActiveForm.pcbPicture.ScaleWidth, mdiImgProcess.ActiveForm.pcbPicture.ScaleHeight, mdiImgProcess.ActiveForm.pcbPicture.hdc, 0, 0, vbSrcCopy
    Unload frmProgress
End Sub

Sub Sharpening(pfIndex As Integer)
    Dim pX As Long, pY As Long
    Dim x As Long, y As Long
    Dim colorval As Long
    Dim red As Long, green As Long, blue As Long
    Dim red1 As Long, green1 As Long, blue1 As Long
    Dim red2 As Long, green2 As Long, blue2 As Long
    
    pX = mdiImgProcess.ActiveForm.pcbPicture.ScaleWidth - 1
    pY = mdiImgProcess.ActiveForm.pcbPicture.ScaleHeight - 1
    Load frmProgress
    frmProgress.Caption = "Sharpening..."
    frmProgress.Show
    DoEvents
    For x = 1 To pX - 2
        For y = 1 To pY - 2
            colorval = GetPixel(hDCSour(pfIndex), x, y)
            red = GetRed(colorval)
            green = GetGreen(colorval)
            blue = GetBlue(colorval)
            colorval = GetPixel(hDCSour(pfIndex), x - 1, y - 1)
            red1 = GetRed(colorval)
            green1 = GetGreen(colorval)
            blue1 = GetBlue(colorval)
            
            red2 = red + 0.5 * (red - red1)
            green2 = green + 0.5 * (green - green1)
            blue2 = blue + 0.5 * (blue - blue1)
            
            If red2 >= 255 Then red2 = 255
            If green2 >= 255 Then green2 = 255
            If blue2 >= 255 Then blue2 = 255
            If red2 <= 0 Then red2 = 0
            If green2 <= 0 Then green2 = 0
            If blue2 <= 0 Then blue2 = 0
            
            SetPixel hDCDest(pfIndex), x, y, RGB(red2, green2, blue2)
        Next y
        If (x Mod (pX / 20)) = 0 Then frmProgress.ProgressBar.Value = x / pX * 100
    Next x
    BitBlt mdiImgProcess.ActiveForm.pcbPicture.hdc, 0, 0, mdiImgProcess.ActiveForm.pcbPicture.ScaleWidth, mdiImgProcess.ActiveForm.pcbPicture.ScaleHeight, hDCDest(pfIndex), 0, 0, vbSrcCopy
    mdiImgProcess.ActiveForm.pcbPicture.Refresh
    BitBlt hDCSour(pfIndex), 0, 0, mdiImgProcess.ActiveForm.pcbPicture.ScaleWidth, mdiImgProcess.ActiveForm.pcbPicture.ScaleHeight, mdiImgProcess.ActiveForm.pcbPicture.hdc, 0, 0, vbSrcCopy
    Unload frmProgress
End Sub

Sub Solarizing(pfIndex As Integer)
    Dim pX As Long, pY As Long
    Dim x As Long, y As Long
    Dim colorval As Long
    Dim red As Long, green As Long, blue As Long
    Dim red2 As Long, green2 As Long, blue2 As Long
    
    pX = mdiImgProcess.ActiveForm.pcbPicture.ScaleWidth - 1
    pY = mdiImgProcess.ActiveForm.pcbPicture.ScaleHeight - 1
    Load frmProgress
    frmProgress.Caption = "Solarizing..."
    frmProgress.Show
    DoEvents
    For x = 0 To pX
        For y = 0 To pY
            colorval = GetPixel(hDCSour(pfIndex), x, y)
            red = GetRed(colorval)
            green = GetGreen(colorval)
            blue = GetBlue(colorval)
            
            If ((red < 128) Or (red > 255)) Then red2 = 255 - red
            If ((green < 128) Or (green > 255)) Then green2 = 255 - green
            If ((blue < 128) Or (blue > 128)) Then blue2 = 255 - blue
            
            If red2 >= 255 Then red2 = 255
            If green2 >= 255 Then green2 = 255
            If blue2 >= 255 Then blue2 = 255
            If red2 <= 0 Then red2 = 0
            If green2 <= 0 Then green2 = 0
            If blue2 <= 0 Then blue2 = 0
            
            SetPixel hDCDest(pfIndex), x, y, RGB(red2, green2, blue2)
        Next y
        If (x Mod (pX / 20)) = 0 Then frmProgress.ProgressBar.Value = x / pX * 100
    Next x
    BitBlt mdiImgProcess.ActiveForm.pcbPicture.hdc, 0, 0, mdiImgProcess.ActiveForm.pcbPicture.ScaleWidth, mdiImgProcess.ActiveForm.pcbPicture.ScaleHeight, hDCDest(pfIndex), 0, 0, vbSrcCopy
    mdiImgProcess.ActiveForm.pcbPicture.Refresh
    BitBlt hDCSour(pfIndex), 0, 0, mdiImgProcess.ActiveForm.pcbPicture.ScaleWidth, mdiImgProcess.ActiveForm.pcbPicture.ScaleHeight, mdiImgProcess.ActiveForm.pcbPicture.hdc, 0, 0, vbSrcCopy
    Unload frmProgress
End Sub

Sub Embossing(pfIndex As Integer)
    Dim pX As Long, pY As Long
    Dim x As Long, y As Long
    Dim colorval As Long
    Dim red As Long, green As Long, blue As Long
    Dim red1 As Long, green1 As Long, blue1 As Long
    Dim red2 As Long, green2 As Long, blue2 As Long
    
    pX = mdiImgProcess.ActiveForm.pcbPicture.ScaleWidth - 1
    pY = mdiImgProcess.ActiveForm.pcbPicture.ScaleHeight - 1
    Load frmProgress
    frmProgress.Caption = "Embossing..."
    frmProgress.Show
    DoEvents
    For x = 1 To pX - 2
        For y = 1 To pY - 2
            colorval = GetPixel(hDCSour(pfIndex), x, y)
            red = GetRed(colorval)
            green = GetGreen(colorval)
            blue = GetBlue(colorval)
            colorval = GetPixel(hDCSour(pfIndex), x + 1, y + 1)
            red1 = GetRed(colorval)
            green1 = GetGreen(colorval)
            blue1 = GetBlue(colorval)
            
            red2 = Abs(red - red1 + 128)
            green2 = Abs(green - green1 + 128)
            blue2 = Abs(blue - blue1 + 128)
            
            If red2 >= 255 Then red2 = 255
            If green2 >= 255 Then green2 = 255
            If blue2 >= 255 Then blue2 = 255
            If red2 <= 0 Then red2 = 0
            If green2 <= 0 Then green2 = 0
            If blue2 <= 0 Then blue2 = 0
            
            SetPixel hDCDest(pfIndex), x, y, RGB(red2, green2, blue2)
        Next y
        If (x Mod (pX / 20)) = 0 Then frmProgress.ProgressBar.Value = x / pX * 100
    Next x
    BitBlt mdiImgProcess.ActiveForm.pcbPicture.hdc, 0, 0, mdiImgProcess.ActiveForm.pcbPicture.ScaleWidth, mdiImgProcess.ActiveForm.pcbPicture.ScaleHeight, hDCDest(pfIndex), 0, 0, vbSrcCopy
    mdiImgProcess.ActiveForm.pcbPicture.Refresh
    BitBlt hDCSour(pfIndex), 0, 0, mdiImgProcess.ActiveForm.pcbPicture.ScaleWidth, mdiImgProcess.ActiveForm.pcbPicture.ScaleHeight, mdiImgProcess.ActiveForm.pcbPicture.hdc, 0, 0, vbSrcCopy
    Unload frmProgress
End Sub

Sub Diffusing(pfIndex As Integer)
    Dim pX As Long, pY As Long
    Dim x As Long, y As Long
    Dim colorval As Long
    Dim red As Long, green As Long, blue As Long
    Dim rX As Long, rY As Long
    
    pX = mdiImgProcess.ActiveForm.pcbPicture.ScaleWidth - 1
    pY = mdiImgProcess.ActiveForm.pcbPicture.ScaleHeight - 1
    Load frmProgress
    frmProgress.Caption = "Diffusing..."
    frmProgress.Show
    DoEvents
    Randomize Timer
    For x = 0 To pX
        For y = 0 To pY
            rX = Rnd() * 4 - 2
            rY = Rnd() * 4 - 2
            colorval = GetPixel(hDCSour(pfIndex), x + rX, y + rY)
            red = GetRed(colorval)
            green = GetGreen(colorval)
            blue = GetBlue(colorval)
            
            If red >= 255 Then red = 255
            If green >= 255 Then green = 255
            If blue >= 255 Then blue = 255
            If red <= 0 Then red = 0
            If green <= 0 Then green = 0
            If blue <= 0 Then blue = 0
            
            SetPixel hDCDest(pfIndex), x, y, RGB(red, green, blue)
        Next y
        If (x Mod (pX / 20)) = 0 Then frmProgress.ProgressBar.Value = x / pX * 100
    Next x
    BitBlt mdiImgProcess.ActiveForm.pcbPicture.hdc, 0, 0, mdiImgProcess.ActiveForm.pcbPicture.ScaleWidth, mdiImgProcess.ActiveForm.pcbPicture.ScaleHeight, hDCDest(pfIndex), 0, 0, vbSrcCopy
    mdiImgProcess.ActiveForm.pcbPicture.Refresh
    BitBlt hDCSour(pfIndex), 0, 0, mdiImgProcess.ActiveForm.pcbPicture.ScaleWidth, mdiImgProcess.ActiveForm.pcbPicture.ScaleHeight, mdiImgProcess.ActiveForm.pcbPicture.hdc, 0, 0, vbSrcCopy
    Unload frmProgress
End Sub

