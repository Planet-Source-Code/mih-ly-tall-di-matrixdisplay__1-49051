Attribute VB_Name = "Common"
Public CharSet() As Byte
Public CHARW As Long, CHARH As Long
Public UnderLine(0 To 255) As Byte
Public UnderLineHeight As Byte

Public Sub SaveChr(FileName As String)
Dim FreeFil As Integer, Temp As String, C As Byte
FreeFil = FreeFile
If Dir(FileName) <> "" Then Kill FileName
Open FileName For Binary As #FreeFil

Temp = "MDCS" & CompleteString(DecTo256(CHARW), 2) & CompleteString(DecTo256(CHARH), 2)
Put #FreeFil, , Temp

Put #FreeFil, , CharSet

Put #FreeFil, , UnderLine
Temp = DecTo256(UnderLineHeight)
Put #FreeFil, , Temp

Close #FreeFil

End Sub

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

Public Function DecTo256(ByVal Dec As Variant) As String
Dim ConTmp As String

Do
ConTmp = Chr$(Dec Mod 256) & ConTmp
If Dec <= 255 Then GoTo Done
Dec = Int(Dec / 256)
Loop

Done:
DecTo256 = ConTmp
End Function

Public Function BackToDec(ByVal Back As String) As Variant
Dim z As Integer, BackConTmp As Long

For z = Len(Back) - 1 To 0 Step -1
    BackConTmp = BackConTmp + Asc(Mid$(Back, Len(Back) - z, 1)) * (256 ^ z)
Next z

BackToDec = BackConTmp
End Function

Public Function CompleteString(InNum As String, ByVal OutLen As Byte) As String
OutLen = OutLen - Len(InNum)
If OutLen <= 0 Then
   CompleteString = InNum
   Exit Function
End If
CompleteString = String(OutLen, Chr(0)) & InNum
End Function

