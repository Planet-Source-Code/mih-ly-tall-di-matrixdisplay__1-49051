VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MatrixDisplay demo"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   160
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   603
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TestTim 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   3480
      Top             =   1920
   End
   Begin VB.Timer ScrollTim 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8640
      Top             =   1800
   End
   Begin VB.PictureBox BackBuffer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   240
      ScaleHeight     =   129
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   569
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   8535
   End
   Begin VB.PictureBox LEDPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   120
      Index           =   4
      Left            =   8880
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox LEDPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   120
      Index           =   3
      Left            =   8880
      Picture         =   "Form1.frx":0102
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox LEDPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   120
      Index           =   2
      Left            =   8880
      Picture         =   "Form1.frx":0204
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox LEDPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   120
      Index           =   1
      Left            =   8880
      Picture         =   "Form1.frx":0306
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox LEDPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   120
      Index           =   0
      Left            =   8880
      Picture         =   "Form1.frx":0408
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox Canvas 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   120
      ScaleHeight     =   129
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   569
      TabIndex        =   0
      Top             =   120
      Width           =   8535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tip: Click on underlined words!"
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   3615
   End
   Begin VB.Image MatrixImg 
      Height          =   240
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   240
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'It's just a little demo code, to show you how to use
'the MatrixDisplay(The main code is in the module).
'I bet you can write some other cool effects, too,
'if you take a look at the Timer's procedure.
'If you wrote something really cool, send it to me, if you
'wish(If you send it by mail, do it by sending text instead
'of attachments. Thanks.).

Dim X As Long, Y As Long, NowOrigX As Long, NowOrigY As Long, PrevTick As Long
Dim LoopTimes As Byte

Private Sub Canvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim TempCHTResult As MatrixCoord

If (LoopTimes = 10) And (Button = 1) Then
   TempCHTResult = HitTest_Char(CLng(X), CLng(Y))
   If (TempCHTResult.X > 28) And (TempCHTResult.X < 33) Then
      MsgBox "Yhea, voting is important, so go to PSC and VOTE!!" & vbNewLine & "(BTW, if you haven't noticed, this was a character-level hittest!)", vbExclamation
   End If
End If

End Sub

Private Sub Form_Load()
InitMatrix 1000, 80, 70, 15, 8, 5
InitDisplay Canvas, BackBuffer

If LoadChr(App.Path & "\ANSI_8x15.chr") = 0 Then
   MsgBox "An error occurred while loading the Charset." & vbNewLine & "The program will close.", vbCritical
   End
End If

ScrollTim.Enabled = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub ScrollTim_Timer()
'------======Init main text======------
AddTextToData 0, 0, "Knock-knock Neo...   The Matrix has you...  ", CM_ConstantColor, , , 3
AddTextToData -1, 0, "MatrixDisplay by Msi", CM_Rainbow
AddTextToData -1, 0, "  --== ", CM_ConstantColor, , , 2
AddTextToData -1, 0, "Comments are welcome: msist@freemail.hu & at PSC", CM_CustomTimes, String(22, Chr(4)) & String(17, Chr(1)) & String(6, Chr(4)), 1, 1
AddTextToData -1, 0, " ==-- ", CM_ConstantColor, , , 2

'------======Drop in======------
NowOrigY = 15
Do
   CopyDataToDisplay 0, NowOrigY
   UpdateDisplay Canvas, BackBuffer, LEDPic
   Canvas.Refresh
   DoEvents
   
   NowOrigY = NowOrigY - 1
   PrevTick = GetTickCount
   Do Until PrevTick + 50 < GetTickCount
   Loop
Loop Until NowOrigY = -1

'------======Pause======------
PrevTick = GetTickCount
Do Until PrevTick + 600 < GetTickCount
   DoEvents
Loop

'------======Scroll main text======------
Do
   CopyDataToDisplay NowOrigX, 0
   UpdateDisplay Canvas, BackBuffer, LEDPic
   Canvas.Refresh
   DoEvents
   
   NowOrigX = NowOrigX + 1
   If (NowOrigX > DATAW - DISPW - 2) Then
      PrevTick = GetTickCount
      Do Until PrevTick + 200 < GetTickCount
         DoEvents
      Loop
      NowOrigX = NowOrigX - 1
      LoopTimes = 1
   End If
   
   PrevTick = GetTickCount
   Do Until PrevTick + 20 < GetTickCount
   Loop
   
Loop Until LoopTimes = 1

'------======Pause======------
PrevTick = GetTickCount
Do Until PrevTick + 600 < GetTickCount
   DoEvents
Loop

'------======Fly up======------
NowOrigY = 0
Do
   CopyDataToDisplay NowOrigX, NowOrigY
   UpdateDisplay Canvas, BackBuffer, LEDPic
   Canvas.Refresh
   DoEvents
   
   NowOrigY = NowOrigY + 1
   PrevTick = GetTickCount
   Do Until PrevTick + 50 < GetTickCount
   Loop
Loop Until NowOrigY = 16

'------======Reset matrix & init end text======------
InitMatrix 1000, 80, 70, 15, 8, 5

AddTextToData 0, 0, "         And don't forget to", CM_ConstantColor, , , 2
AddTextToData -1, 0, " VOTE", CM_ConstantColor, , , 1, 1
AddTextToData -1, 0, "!", CM_ConstantColor, , , 3

LoopTimes = 10

'------======Scroll end text======------
NowOrigX = 0
Do
   CopyDataToDisplay NowOrigX, 0
   UpdateDisplay Canvas, BackBuffer, LEDPic
   Canvas.Refresh
   DoEvents
   
   NowOrigX = NowOrigX + 1
   PrevTick = GetTickCount
   Do Until PrevTick + 30 < GetTickCount
   Loop
Loop Until NowOrigX > 300

'------======Pause & end======------
PrevTick = GetTickCount
Do Until PrevTick + 1500 < GetTickCount
   DoEvents
Loop

ScrollTim.Enabled = False
End

End Sub

