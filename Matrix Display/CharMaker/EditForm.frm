VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form EditForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Charset editor"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   450
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   671
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton UnderLineAllButt 
      Caption         =   "Underline all"
      Height          =   375
      Left            =   6600
      TabIndex        =   22
      Top             =   6240
      Width           =   1095
   End
   Begin VB.TextBox ULHBox 
      Height          =   285
      Left            =   5040
      TabIndex        =   19
      Text            =   "12"
      Top             =   5640
      Width           =   735
   End
   Begin VB.CommandButton ExitButt 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   375
      Left            =   8880
      TabIndex        =   17
      Top             =   6240
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CmDlg 
      Left            =   9600
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "CharSet Files (*.chr)|*.chr"
   End
   Begin VB.CheckBox ULineChk 
      Alignment       =   1  'Right Justify
      Caption         =   "This character is underlineable"
      Height          =   255
      Left            =   2280
      TabIndex        =   15
      Top             =   5520
      Width           =   2535
   End
   Begin VB.CommandButton SaveButt 
      Caption         =   "&Save..."
      Height          =   375
      Left            =   8880
      TabIndex        =   14
      Top             =   5880
      Width           =   1095
   End
   Begin VB.PictureBox ANSISetPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3600
      Left            =   6240
      Picture         =   "EditForm.frx":0000
      ScaleHeight     =   240
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   13
      Top             =   600
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.PictureBox ASCIISetPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2880
      Left            =   7800
      Picture         =   "EditForm.frx":7C42
      ScaleHeight     =   192
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   12
      Top             =   480
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.PictureBox CurrCharPic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   4440
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   5
      Top             =   5040
      Width           =   375
   End
   Begin VB.CommandButton SetChangeButt 
      Caption         =   "&Showing User set"
      Default         =   -1  'True
      Height          =   375
      Left            =   6240
      TabIndex        =   4
      Top             =   5400
      Width           =   2295
   End
   Begin VB.PictureBox SetPic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4935
      Left            =   5040
      ScaleHeight     =   329
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   329
      TabIndex        =   2
      Top             =   360
      Width           =   4935
   End
   Begin VB.HScrollBar CharScroll 
      Height          =   255
      LargeChange     =   5
      Left            =   240
      Max             =   255
      TabIndex        =   1
      Top             =   5880
      Width           =   4575
   End
   Begin VB.PictureBox EditPic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   4575
      Left            =   240
      ScaleHeight     =   305
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   305
      TabIndex        =   0
      Top             =   360
      Width           =   4575
      Begin VB.Line ULLine 
         BorderColor     =   &H00FF00FF&
         Visible         =   0   'False
         X1              =   0
         X2              =   304
         Y1              =   264
         Y2              =   264
      End
   End
   Begin VB.CommandButton LoadButt 
      Caption         =   "&Load..."
      Height          =   375
      Left            =   8880
      TabIndex        =   16
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label CharNumLab 
      Alignment       =   2  'Center
      Caption         =   "#0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   21
      Top             =   5040
      Width           =   735
   End
   Begin VB.Line Line6 
      X1              =   576
      X2              =   664
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line5 
      X1              =   576
      X2              =   576
      Y1              =   392
      Y2              =   360
   End
   Begin VB.Line Line4 
      X1              =   412
      X2              =   576
      Y1              =   392
      Y2              =   392
   End
   Begin VB.Line Line3 
      X1              =   412
      X2              =   412
      Y1              =   392
      Y2              =   355
   End
   Begin VB.Line Line2 
      X1              =   328
      X2              =   412
      Y1              =   356
      Y2              =   356
   End
   Begin VB.Label Label7 
      Caption         =   "This is shown on underlineable characters as a magenta line."
      Height          =   495
      Left            =   5040
      TabIndex        =   20
      Top             =   6000
      Width           =   3015
   End
   Begin VB.Label Label5 
      Caption         =   "Underline height:"
      Height          =   255
      Left            =   4920
      TabIndex        =   18
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Browse characters using the scrollbar above, or using QuickJump"
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   6240
      Width           =   4575
   End
   Begin VB.Label ASCIILab 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   10
      Top             =   5400
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "Same ASCII character:"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label ANSILab 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   5040
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "Same ANSI character:"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Your character:"
      Height          =   255
      Left            =   3120
      TabIndex        =   6
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Quick jump(Click on a character):"
      Height          =   255
      Left            =   5040
      TabIndex        =   3
      Top             =   120
      Width           =   2415
   End
   Begin VB.Line Line1 
      X1              =   328
      X2              =   328
      Y1              =   356
      Y2              =   16
   End
End
Attribute VB_Name = "EditForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Sorry, but this is not commented as well as
'the MatrixDisplay(in fact, it is NOT commented AT ALL).
'There are 2(two) reasons for this:
'1-This code is not so hard to understand, even for
'  beginners.
'2-Commenting this would be *really* painful, and i don't
'  have *that* much time right now... -:-)
Dim X As Long, Y As Long, C As Long
Dim iX As Long, iY As Long
Dim QJMode As Byte

Private Sub ANSISetPic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
   C = Int(X \ 8 + (Y \ 15) * 16)
   If C > 255 Then C = 255
   CharScroll.Value = C
End If

End Sub

Private Sub ANSISetPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
   C = Int(X \ 8 + (Y \ 15) * 16)
   If C > 255 Then C = 255
   CharScroll.Value = C
End If

End Sub

Private Sub ASCIISetPic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
   C = Int(X \ 8 + (Y \ 12) * 16)
   If C > 255 Then C = 255
   CharScroll.Value = C
End If

End Sub

Private Sub ASCIISetPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
   C = Int(X \ 8 + (Y \ 12) * 16)
   If C > 255 Then C = 255
   CharScroll.Value = C
End If

End Sub

Private Sub CharScroll_Change()
RefreshEdit
End Sub

Private Sub CharScroll_Scroll()
RefreshEdit
End Sub

Private Sub UnderLineAllButt_Click()
For X = 0 To 255
   UnderLine(X) = 1
Next 'X
ULineChk.Value = 1
ULLine.Visible = True

End Sub

Private Sub EditPic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
   iX = X \ 15
   iY = Y \ 15
   If (iX > -1) And (iX < CHARW) And (iY > -1) And (iY < CHARH) Then
      EditPic.FillColor = 0
      EditPic.ForeColor = 0
      EditPic.Line (iX * 15, iY * 15)-((iX + 1) * 15, (iY + 1) * 15), , BF
      CurrCharPic.PSet (iX, iY), EditPic.ForeColor
      CharSet(CharScroll.Value, iX, iY) = 1
   End If
ElseIf Button = 2 Then
   iX = X \ 15
   iY = Y \ 15
   If (iX > -1) And (iX < CHARW) And (iY > -1) And (iY < CHARH) Then
      EditPic.FillColor = 2 ^ 24 - 1
      EditPic.ForeColor = 2 ^ 24 - 1
      EditPic.Line (iX * 15, iY * 15)-((iX + 1) * 15, (iY + 1) * 15), , BF
      CurrCharPic.PSet (iX, iY), EditPic.ForeColor
      CharSet(CharScroll.Value, iX, iY) = 0
   End If
End If

End Sub

Private Sub EditPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
   iX = X \ 15
   iY = Y \ 15
   If (iX > -1) And (iX < CHARW) And (iY > -1) And (iY < CHARH) Then
      EditPic.FillColor = 0
      EditPic.ForeColor = 0
      EditPic.Line (iX * 15, iY * 15)-((iX + 1) * 15, (iY + 1) * 15), , BF
      CurrCharPic.PSet (iX, iY), EditPic.ForeColor
      CharSet(CharScroll.Value, iX, iY) = 1
   End If
ElseIf Button = 2 Then
   iX = X \ 15
   iY = Y \ 15
   If (iX > -1) And (iX < CHARW) And (iY > -1) And (iY < CHARH) Then
      EditPic.FillColor = 2 ^ 24 - 1
      EditPic.ForeColor = 2 ^ 24 - 1
      EditPic.Line (iX * 15, iY * 15)-((iX + 1) * 15, (iY + 1) * 15), , BF
      CurrCharPic.PSet (iX, iY), EditPic.ForeColor
      CharSet(CharScroll.Value, iX, iY) = 0
   End If
End If

End Sub

Private Sub EditPic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (Button = 1) Or (Button = 2) Then
   RefreshSet CharScroll.Value, CharScroll.Value
End If
End Sub

Private Sub ExitButt_Click()
If MsgBox("Are you sure you want to quit?", vbYesNo + vbQuestion) = vbYes Then End
End Sub

Private Sub Form_Load()
CHARW = UBound(CharSet, 2) + 1
CHARH = UBound(CharSet, 3) + 1
EditPic.Width = CHARW * 15
EditPic.Height = CHARH * 15
CurrCharPic.Width = CHARW
CurrCharPic.Height = CHARH
SetPic.Width = 16 * CHARW
SetPic.Height = 16 * CHARH

RefreshEdit
RefreshSet 0, 255

ULHBox = UnderLineHeight
ULLine.Y1 = 15 * UnderLineHeight - 7.5
ULLine.Y2 = ULLine.Y1
ULLine.Visible = UnderLine(0)

End Sub

Private Sub RefreshSet(StartChar As Long, EndChar As Long)
For C = StartChar To EndChar
   For X = 0 To CHARW - 1
      For Y = 0 To CHARH - 1
         SetPic.PSet (X + (C Mod 16) * CHARW, Y + (C \ 16) * CHARH), (2 ^ 24 - 1) * (1 - CharSet(C, X, Y))
      Next 'Y
   Next 'X
Next 'C
SetPic.Refresh

End Sub

Private Sub RefreshEdit()
For X = 0 To CHARW - 1
   For Y = 0 To CHARH - 1
      EditPic.FillColor = (2 ^ 24 - 1) * (1 - CharSet(CharScroll.Value, X, Y))
      EditPic.ForeColor = (2 ^ 24 - 1) * (1 - CharSet(CharScroll.Value, X, Y))
      EditPic.Line (X * 15, Y * 15)-((X + 1) * 15, (Y + 1) * 15), , BF
      CurrCharPic.PSet (X, Y), EditPic.ForeColor
   Next 'Y
Next 'X
ULineChk.Value = UnderLine(CharScroll.Value)
ULLine.Visible = UnderLine(CharScroll.Value)
CharNumLab = "#" & CharScroll.Value
ANSILab = Chr(CharScroll.Value)
ASCIILab = Chr(CharScroll.Value)
End Sub

Private Sub LoadButt_Click()
On Error GoTo ee

CmDlg.Flags = cdlOFNLongNames Or cdlOFNFileMustExist
CmDlg.ShowOpen

If LoadChr(CmDlg.FileName) = 0 Then
   MsgBox "Cannot open Charset, file is corrupted.", vbCritical
   Exit Sub
End If

CHARW = UBound(CharSet, 2)
CHARH = UBound(CharSet, 3)
EditPic.Width = CHARW * 15
EditPic.Height = CHARH * 15
CurrCharPic.Width = CHARW
CurrCharPic.Height = CHARH
SetPic.Width = 16 * CHARW
SetPic.Height = 16 * CHARH
CharScroll.Value = 0
CharNumLab = "#0"

RefreshEdit
RefreshSet 0, 255

ULHBox = UnderLineHeight
ULLine.Y1 = 15 * UnderLineHeight - 7.5
ULLine.Y2 = ULLine.Y1
ULLine.Visible = UnderLine(0)
ee:

End Sub

Private Sub SaveButt_Click()
On Error GoTo ee

CmDlg.Flags = cdlOFNLongNames Or cdlOFNOverwritePrompt
CmDlg.ShowSave

SaveChr CmDlg.FileName

ee:
End Sub

'QJMode=
'0: Showing User set
'1: Showing ANSI set
'2: Showing ASCII set
Private Sub SetChangeButt_Click()
If QJMode = 0 Then
   SetChangeButt.Caption = "Showing ANSI set"
   QJMode = 1
   SetPic.Visible = False
   ANSISetPic.Left = SetPic.Left
   ANSISetPic.Top = SetPic.Top
   ANSISetPic.Visible = True
ElseIf QJMode = 1 Then
   SetChangeButt.Caption = "Showing ASCII set"
   QJMode = 2
   ANSISetPic.Visible = False
   ASCIISetPic.Left = SetPic.Left
   ASCIISetPic.Top = SetPic.Top
   ASCIISetPic.Visible = True
ElseIf QJMode = 2 Then
   SetChangeButt.Caption = "Showing User set"
   QJMode = 0
   ASCIISetPic.Visible = False
   SetPic.Visible = True
End If
End Sub

Private Sub SetPic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
   C = Int(X \ CHARW + (Y \ CHARH) * 16)
   If C > 255 Then C = 255
   CharScroll.Value = C
End If

End Sub

Private Sub SetPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
   C = Int(X \ CHARW + (Y \ CHARH) * 16)
   If C > 255 Then C = 255
   CharScroll.Value = C
End If

End Sub

Private Sub ULHBox_Change()
UnderLineHeight = Val(ULHBox)
ULLine.Y1 = 15 * UnderLineHeight - 7.5
ULLine.Y2 = ULLine.Y1
End Sub

Private Sub ULineChk_Click()
UnderLine(CharScroll.Value) = ULineChk.Value
ULLine.Visible = UnderLine(CharScroll.Value)

End Sub
