Attribute VB_Name = "MatrixDisplay"
'The main code of the display is in this module.
'First i wanted to create an ActiveX control from it,
'but a module is much better for this purpose, 'cause
'this way i can provide direct access to the MatrixData
'and MatrixDisplay arrays, so you can write much cooler
'effects and other stuff for the display i can think of...
'The workflow of the MatrixDisplay 'engine' is fairly
'simple:
'First, you call InitMatrix(),then InitDisplay(),
'and you're ready to use the thing.
'If you don't want to mess with adding simple text,
'you can call LoadCharSet() after the init, and
'start adding text using AddTextToData().
'In the display loop, you call CopyDataToDisplay(), then
'UpdateDisplay. It's that simple.
'Remember, if your display picturebox has it's AutoRedraw
'property on, you need to call Refresh on it after
'UnpdateDisplay to see a thing.

'The following code is heavily commented, so it couldn't be
'a problem understanding and/or using it.
'Again, if you have a question, comment, etc., write me a
'mail, or write to this code's thread on Planet SourceCode.

Option Base 0
'For description in on this Enum, see AddTextToData().
Public Enum ColorMethodEnum
   CM_ConstantColor = 0
   CM_Rainbow = 1
   CM_Rainbow_Reverse = 2
   CM_Random = 3
   CM_CustomLoop = 4
   CM_CustomTimes = 5
End Enum

'This is used to return the result of a HitTest function.
Public Type MatrixCoord
   X As Long
   Y As Long
End Type

Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
Public Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public MatrixData() As Byte, DATAW As Long, DATAH As Long
Public MatrixDisplay() As Byte, DISPW As Long, DISPH As Long

Private LEDW As Long
Private NUMCOLORS As Byte

Private CharSet() As Byte, CHARW As Long, CHARH As Long
Private UnderLine(0 To 255) As Byte
Private UnderLineHeight As Byte
Private LastOffs As Long

'For HitTest_Data
Private Last_DorigX As Long, Last_DorigY As Long
Private Last_CopyWidth As Long, Last_CopyHeight As Long
Private Last_DisplayOffsetX As Long, Last_DisplayOffsetY As Long
'For HitTest_Data END

Private X As Long, Y As Long

'Inits the matrixes and sets internal variables:

'NewLEDW is the width(and height) of the picture
'you want to use for the matrix elements.

'NewNumColors specify the number of colors used
'in the display, including 'off' state. So if it uses only
'red 'LED'-s, this argument should be 2.

Public Sub InitMatrix(NewDataW As Long, NewDataH As Long, NewDispW As Long, NewDispH As Long, NewLEDW As Long, NewNumColors As Byte)
DATAW = NewDataW
DATAH = NewDataH

DISPW = NewDispW
DISPH = NewDispH

LEDW = NewLEDW
NUMCOLORS = NewNumColors - 1

Last_CopyWidth = -1
Last_CopyHeight = -1

ReDim MatrixData(DATAW - 1, DATAH - 1)
ReDim MatrixDisplay(DISPW - 1, DISPH - 1)
LastOffs = 0

End Sub

'Resizes your pictureboxes, so the MatrixDisplay can fit in.
Public Sub InitDisplay(DisplayPic As PictureBox, BackBufferPic As PictureBox)
DisplayPic.Width = DISPW * LEDW
DisplayPic.Height = DISPH * LEDW

BackBufferPic.Width = DISPW * LEDW
BackBufferPic.Height = DISPH * LEDW

End Sub

'It's simple: this loads a Charset from a .chr file.
'It's a simple file with an "MDCS" header, then come CharW
'and CharH, both in 2 bytes.
'After them, the real Charset array saved with 'Put',
'the 'UnderLine'(-able) array, also saved with 'Put'.
'The last variable saved in one byte is the Y-coordinate
'of the underline.
Public Function LoadChr(FileName As String) As Byte
Dim FreeFil As Integer, Temp As String, C As Byte
FreeFil = FreeFile
Open FileName For Binary As #FreeFil

Temp = Input(4, #FreeFil)
If Temp <> "MDCS" Then
   Close #FreeFil
   LoadChr = 0
   Exit Function
End If

CHARW = BackToDec(Input(2, #FreeFil))
CHARH = BackToDec(Input(2, #FreeFil))

ReDim CharSet(0 To 255, CHARW - 1, CHARH - 1)
Get #FreeFil, , CharSet

Get #FreeFil, , UnderLine

Temp = Input(1, #FreeFil)
UnderLineHeight = BackToDec(Temp)

Close #FreeFil
LoadChr = 1

End Function

'The following function is here to support the LoadChr()
'function.
Public Function BackToDec(ByVal Back As String) As Variant
Dim z As Integer, BackConTmp As Long

For z = Len(Back) - 1 To 0 Step -1
    BackConTmp = BackConTmp + Asc(Mid$(Back, Len(Back) - z, 1)) * (256 ^ z)
Next z

BackToDec = BackConTmp
End Function

'Adds text to the DataMatrix, positioned to XOrig and YOrig.
'if XOrig is -1, it appends the given text to the previous
'one.
'After the first three obvious arguments, comes the tricky
'part.
'ColorMethod specifies the way the program assignes color
'to the induvidual characters. Possible values:
'CM_ConstantColor : uses the AfterCustom argument to specify
'  a color for all characters in the string. You can skip
'  CustomPattern and CustomTimes.
'CM_Rainbow : Goes through the colors in forward order,
'  taking one step by character. You can skip CustomPattern,
'  AfterCustom and CustomTimes.
'CM_Rainbow_Reverse : Goes through the colors in reverse
'  order, taking one step by character. You can skip
'  CustomPattern, AfterCustom and CustomTimes.
'CM_Random : Gives a random color to every character.
'  You can skip CustomPattern, AfterCustom and CustomTimes
'  this time, too.
'CM_CustomLoop : This uses the CustomPattern string's
'  charcodes for the color of the corresponding character.
'  If this pattern is shorter than inText, the function
'  loops it as many times as needed.
'  If it wasn't too clear, look at this picture:
'    Pattern: &H01 &H02 &H03 &H02
'     inText:   A    B    C    D
'  Color No.:   1    2    3    2
'  If you use this ColorMethod, you can skip CustomTimes and
'  AfterCustom.
'CM_CustomTimes : Same as CM_CustomLoop, except this doesn't
'  loops the pattern to infinity, just 'CustomTimes' times.
'  After the pattern comes to the end, it uses AfterCustom
'  as a color for the remaining characters.

'If the UnderlineIt attribute is 1 ,then it adds underlined
'text to the MatrixData. It underlines characters only that
'are underlineable. This is determined by the creator of the
'Charset. (Note: You can create Charsets with the CharMaker
'(included).)
Public Sub AddTextToData(ByVal XOrig As Long, YOrig As Long, inText As String, Optional ColorMethod As ColorMethodEnum = CM_ConstantColor, Optional CustomPattern As String = "", Optional ByVal CustomTimes As Long = 1, Optional AfterCustom As Byte = 1, Optional UnderLineIt As Byte = 0)
If XOrig = -1 Then XOrig = LastOffs
LastOffs = LastOffs + Len(inText) * CHARW

If ColorMethod = CM_ConstantColor Then
   For p = 1 To Len(inText)
      AddCharToData XOrig + (p - 1) * CHARW, YOrig, Asc(Mid(inText, p, 1)), AfterCustom, UnderLineIt
   Next 'P
ElseIf ColorMethod = CM_Rainbow Then
   For p = 1 To Len(inText)
      AddCharToData XOrig + (p - 1) * CHARW, YOrig, Asc(Mid(inText, p, 1)), ((p - 1) Mod NUMCOLORS + 1), UnderLineIt
   Next 'P
ElseIf ColorMethod = CM_Rainbow_Reverse Then
   For p = 1 To Len(inText)
      AddCharToData XOrig + (p - 1) * CHARW, YOrig, Asc(Mid(inText, p, 1)), (NUMCOLORS + 1) - ((p - 1) Mod NUMCOLORS + 1), UnderLineIt
   Next 'P
ElseIf ColorMethod = CM_Random Then
   For p = 1 To Len(inText)
      AddCharToData XOrig + (p - 1) * CHARW, YOrig, Asc(Mid(inText, p, 1)), Int(Rnd * NUMCOLORS + 1), UnderLineIt
   Next 'P
ElseIf ColorMethod = CM_CustomLoop Then
   For p = 1 To Len(inText)
      AddCharToData XOrig + (p - 1) * CHARW, YOrig, Asc(Mid(inText, p, 1)), Asc(Mid(CustomPattern, ((p - 1) Mod Len(CustomPattern) + 1), 1)), UnderLineIt
   Next 'P
ElseIf ColorMethod = CM_CustomTimes Then
   For p = 1 To Len(inText)
      If CustomTimes < 1 Then
         AddCharToData XOrig + (p - 1) * CHARW, YOrig, Asc(Mid(inText, p, 1)), AfterCustom, UnderLineIt
      Else
         AddCharToData XOrig + (p - 1) * CHARW, YOrig, Asc(Mid(inText, p, 1)), Asc(Mid(CustomPattern, ((p - 1) Mod Len(CustomPattern) + 1), 1)), UnderLineIt
         If (p Mod Len(CustomPattern)) = 0 Then CustomTimes = CustomTimes - 1
      End If
   Next 'P
End If

End Sub

'Obvious: Copies a portion of the DataMatrix to
'the MatrixDisplay.
'If you do not specify CopyWidth and CopyHeight, these
'will be DisplayWidth and DisplayHeight.
'DisplayOffset arguments are useful for special effects
'like graphics dropping in lines to the bottom, etc.
'Data will be looped if it reaches the end.
'It loops data for X and for Y now, in both ways, but
'you better not try to loop it in both ways, 'cause it
'will produce strange errors.
Public Sub CopyDataToDisplay(DorigX As Long, DorigY As Long, Optional ByVal CopyWidth As Long = -1, Optional ByVal CopyHeight As Long = -1, Optional DisplayOffsetX As Long = 0, Optional DisplayOffsetY As Long = 0)
Dim ErrX As Long, ErrY As Long, ErrInit As Byte

'For MatrixData hittest
Last_DorigX = DorigX
Last_DorigY = DorigY
Last_CopyWidth = CopyWidth
Last_CopyHeight = CopyHeight
Last_DisplayOffsetX = DisplayOffsetX
Last_DisplayOffsetY = DisplayOffsetY
'For MatrixData hittest END

If (CopyWidth > -1) And (CopyHeight > -1) Then
   If CopyWidth + DisplayOffsetX > DISPW Then CopyWidth = DISPW - DisplayOffsetX
   If CopyHeight + DisplayOffsetY > DISPH Then CopyHeight = DISPH - DisplayOffsetY
Else
   CopyWidth = DISPW - 1
   CopyHeight = DISPH - 1
End If

For X = 0 To CopyWidth
   For Y = 0 To CopyHeight
      If (DorigX + X < DATAW) And (DorigY + Y < DATAH) Then
         MatrixDisplay(X + DisplayOffsetX, Y + DisplayOffsetY) = MatrixData(DorigX + X, DorigY + Y)
      Else
      'this is for data looping purposes
         If ErrInit = 0 Then
            ErrX = X
            ErrY = Y
            ErrInit = 1
         End If
         MatrixDisplay(X + DisplayOffsetX, Y + DisplayOffsetY) = MatrixData(X - ErrX, Y - ErrY)
      End If
   Next 'Y
Next 'X

End Sub

'perform a LED-hittest at MatrixDisplay level.
'X and Y are in pixels.
Public Function HitTest_Display(X As Long, Y As Long) As MatrixCoord
HitTest_Display.X = X \ 8
HitTest_Display.Y = Y \ 8

End Function

'perform a LED-hittest at MatrixData level.
'X and Y are in pixels.
'Well, i'm not sure this is the most efficient method for
'this, but this is the safest nonetheless.
'It works quite well, i tested it(imagine a simple paint
'program with a scrolling 'paper'...).
Public Function HitTest_Data(X As Long, Y As Long) As MatrixCoord
Dim ErrX As Long, ErrY As Long, ErrInit As Byte, TempMC As MatrixCoord

TempMC = HitTest_Display(X, Y)

If (Last_CopyWidth > -1) And (Last_CopyHeight > -1) Then
   If Last_CopyWidth + Last_DisplayOffsetX > DISPW Then Last_CopyWidth = DISPW - Last_DisplayOffsetX
   If Last_CopyHeight + Last_DisplayOffsetY > DISPH Then Last_CopyHeight = DISPH - Last_DisplayOffsetY
Else
   Last_CopyWidth = DISPW - 1
   Last_CopyHeight = DISPH - 1
End If

For Xx = 0 To Last_CopyWidth
   For Yy = 0 To Last_CopyHeight
      If (Last_DorigX + Xx < DATAW) And (Last_DorigY + Yy < DATAH) Then
         If (TempMC.X = Xx + Last_DisplayOffsetX) And (TempMC.Y = Yy + Last_DisplayOffsetY) Then
            HitTest_Data.X = Last_DorigX + Xx
            HitTest_Data.Y = Last_DorigY + Yy
            Exit Function
         End If
      Else
         If ErrInit = 0 Then
            ErrX = X
            ErrY = Y
            ErrInit = 1
         End If
         If (TempMC.X = Xx + Last_DisplayOffsetX) And (TempMC.Y = Yy + Last_DisplayOffsetY) Then
            HitTest_Data.X = Xx - ErrX
            HitTest_Data.Y = Yy - ErrY
            Exit Function
         End If
      End If
   Next 'Yy
Next 'Xx

End Function

'perform a LED-hittest at Char level.
'X and Y are in pixels.
'Here: MatrixCoord.X=Column of char
'Here: MatrixCoord.Y=Row of char
'It seems to be accurate, but if you want to do something
'serious with it, put the function in some real testing.
'Yeah, and don't change the Charset after adding some
'text with the previous, 'cause you will not be able
'to do HitTest_Char!
Public Function HitTest_Char(X As Long, Y As Long) As MatrixCoord
Dim TempDHTResult As MatrixCoord
TempDHTResult = HitTest_Data(X, Y)
HitTest_Char.X = TempDHTResult.X \ CHARW
HitTest_Char.Y = TempDHTResult.Y \ CHARH

End Function


'Blits the LEDPicures to BackBufferPic, then blits
'the BackBufferPic to the DisplayPic.
'This process is necessary for flickerless graphics.
'The first two arguments should be clear, but the third one
'is a little tricky. You need to pass to it an array of
'PictureBoxes, starting with an index of 0(Off state),
'ending with NUMCOLORS-1. These PictureBoxes should
'contain your pictures for the LEDs, properly colored in
'order. So if you want to use CM_Rainbow coloring mode for
'rainbow text, then the order should be red, orange, yellow,
'and so on.

'Colors in this demo:0=Off;1=Red;2=Yellow;3=Green;4=Blue
Public Sub UpdateDisplay(DisplayPic As PictureBox, BackBufferPic As PictureBox, LEDPicArray As Variant)
For Y = 0 To DISPH - 1
   For X = 0 To DISPW - 1
      BitBlt BackBufferPic.hDC, X * LEDW, Y * LEDW, LEDW, LEDW, LEDPicArray(MatrixDisplay(X, Y)).hDC, 0, 0, vbSrcCopy
   Next 'X
Next 'Y
BitBlt DisplayPic.hDC, 0, 0, BackBufferPic.Width, BackBufferPic.Height, BackBufferPic.hDC, 0, 0, vbSrcCopy

End Sub

'You don't need to use this directly;
'This function is only called by AddTextToData().
Private Sub AddCharToData(XOrig As Long, YOrig As Long, CharCode As Byte, Optional CharColor As Byte = 1, Optional UnderLineIt As Byte = 0)
For X = 0 To CHARW - 1
   For Y = 0 To CHARH - 1
      If CharSet(CharCode, X, Y) <> 0 Then
         MatrixData(X + XOrig, Y + YOrig) = CharColor
      Else
         MatrixData(X + XOrig, Y + YOrig) = 0
      End If
   Next 'Y
Next 'X

If (UnderLineIt > 0) And (UnderLine(CharCode) = 1) Then
   For X = 0 To CHARW - 1
      MatrixData(X + XOrig, UnderLineHeight + YOrig) = CharColor
   Next X
End If

End Sub
