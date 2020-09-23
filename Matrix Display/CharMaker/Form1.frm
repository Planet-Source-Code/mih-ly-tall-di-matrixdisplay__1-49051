VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form SetupForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Charset creator"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   252
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   252
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton LoadCharSetButt 
      Caption         =   "&Load an existing charset"
      Height          =   615
      Left            =   120
      TabIndex        =   13
      Top             =   2640
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CmDlg 
      Left            =   1680
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "Bitmap Files (*.bmp)|*.bmp"
   End
   Begin VB.PictureBox TempCPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   2400
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   12
      Top             =   960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox CharsPerRowBox 
      Height          =   285
      Left            =   2280
      TabIndex        =   11
      Text            =   "16"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton ExitButt 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton BMPButt 
      Caption         =   "New charset from a &BMP"
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton ScratchButt 
      Caption         =   "&Create a new charset"
      Default         =   -1  'True
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox HeightBox 
      Height          =   285
      Left            =   840
      TabIndex        =   5
      Text            =   "12"
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox WidthBox 
      Height          =   285
      Left            =   840
      TabIndex        =   2
      Text            =   "8"
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "Note: you do not need to define character dimensions for this one."
      Height          =   615
      Left            =   1560
      TabIndex        =   14
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label Label6 
      Caption         =   "  Enter the number of characters   in a row:"
      Height          =   495
      Left            =   1440
      TabIndex        =   10
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label Label5 
      Caption         =   "pixels"
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Height:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "pixels"
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Width:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Please enter the width and height of a character:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "SetupForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim C As Long, X As Long, Y As Long

Private Sub CreateCharSet(CharSetPic As PictureBox, CharWidth As Long, CharHeight As Long, CharsInARow As Long)
For C = 0 To 255
   For X = 0 To CharWidth - 1
      For Y = 0 To CharHeight - 1
         If CharSetPic.Point(X + (C Mod CharsInARow) * CharWidth, Y + (C \ CharsInARow) * CharHeight) <> 0 Then
            CharSet(C, X, Y) = 0
         Else
            CharSet(C, X, Y) = 1
         End If
      Next 'Y
   Next 'X
Next 'C

End Sub

Private Sub BMPButt_Click()
On Error GoTo ee

CmDlg.Flags = cdlOFNLongNames Or cdlOFNFileMustExist
CmDlg.Filter = "Bitmap Files (*.bmp)|*.bmp"
CmDlg.ShowOpen

TempCPic.Picture = LoadPicture(CmDlg.FileName)

ReDim CharSet(0 To 255, WidthBox - 1, HeightBox - 1)
UnderLineHeight = HeightBox - 2

CreateCharSet TempCPic, WidthBox, HeightBox, CharsPerRowBox
Set TempCPic.Picture = Nothing

EditForm.Show
Unload Me

ee:
End Sub

Private Sub ExitButt_Click()
End
End Sub

Private Sub LoadCharSetButt_Click()
On Error GoTo ee

CmDlg.Flags = cdlOFNLongNames Or cdlOFNFileMustExist
CmDlg.Filter = "CharSet Files (*.chr)|*.chr"
CmDlg.ShowOpen

If LoadChr(CmDlg.FileName) = 0 Then
   MsgBox "Failed to load Charset, file is corrupted.", vbCritical
   Exit Sub
End If

EditForm.Show
Unload Me

ee:
End Sub

Private Sub ScratchButt_Click()
ReDim CharSet(0 To 255, WidthBox - 1, HeightBox - 1)
UnderLineHeight = HeightBox - 2
EditForm.Show
Unload Me

End Sub
