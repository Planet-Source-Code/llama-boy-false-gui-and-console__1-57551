VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7905
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   344
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   527
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Save Screen"
      Height          =   375
      Left            =   6600
      TabIndex        =   7
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton clsCons 
      Caption         =   "Clear Console"
      Height          =   375
      Left            =   6600
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.Frame frmData 
      Caption         =   "Data"
      Height          =   2415
      Left            =   3840
      TabIndex        =   3
      Top             =   2640
      Width           =   3975
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   1080
         Top             =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Char"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Line:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   345
      End
   End
   Begin VB.CommandButton cmdDebug 
      Caption         =   "Run Debug"
      Height          =   375
      Left            =   6600
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtCode 
      Height          =   2295
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   6375
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   2400
      Left            =   120
      ScaleHeight     =   160
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   240
      TabIndex        =   0
      Top             =   2640
      Width           =   3600
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub clsCons_Click()
Picture1.Cls
Picture1.BackColor = vbBlack

End Sub

Private Sub cmdDebug_Click()
CurChar = 0
ParseCode txtCode
SetBG RGB(0, 128, 128)
Picture1.FillColor = RGB(0, 128, 128)
ConsLine 0, 240, 147, 147, cHighlight
ConsLine 0, 240, 148, 148, cButton
ConsLine 0, 240, 149, 149, cButton
ConsLine 0, 240, 150, 150, cButton
ConsLine 0, 240, 151, 151, cButton
ConsLine 0, 240, 152, 152, cButton
ConsLine 0, 240, 153, 153, cButton
ConsLine 0, 240, 154, 154, cButton
ConsLine 0, 240, 155, 155, cButton
ConsLine 0, 240, 156, 156, cButton
ConsLine 0, 240, 157, 157, cButton
ConsLine 0, 240, 158, 158, cButton
ConsLine 0, 240, 159, 159, cButton
ConsLine 0, 240, 160, 160, cButton
mButton 2, 19, 149, 159, "kos"
mWindow "kaotic-os run", 20, 20, 70, 30, "win1"
mWindow "kos editz", 70, 10, 100, 82, "win2"
mControl 1, "win1", 48, 12, 10, 19, "run", "cmdrun"
mControl 2, "win2", 3, 12, 54, 93, "kaotic-os edit is " & vbCrLf & "tight. woah!", "cmdrun"
mControl 1, "win2", 3, 69, 10, 25, "Save", "cmdSave"
SetHotSpot 2, 19, 149, 159
End Sub

Private Sub Command1_Click()
SavePicture Picture1.Image, "C:\ScreenShot.gif"
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
ClickEven x, y, Button
End Sub

Private Sub Timer1_Timer()
Label1(0).Caption = "Line: " & CurLine
Label1(1).Caption = "Char: " & CurChar
End Sub

Private Sub ParseCode(code As String)
    For i = 1 To Len(code)
        If Mid(code, i, 5) = "PRNT " Then
            i = i + 5
            Picture1.Print Mid(code, i, InStr(i, code, ";") - i)
        End If
    Next i
End Sub
