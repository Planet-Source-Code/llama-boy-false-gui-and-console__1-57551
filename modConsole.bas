Attribute VB_Name = "modConsole"
Type HotSpot
    x1 As Single
    x2 As Single
    y1 As Single
    y2 As Single
End Type

Type window
    wTitle As String
    wName As String
    wHeight As Single
    wWidth As Single
    x As Single
    y As Single
    controls As String
    values As String
    layer As Byte
End Type

Public CurChar As Long
Public CurLine As Long
Const CharH = 14
Const CharW = 9
Const ConsW = 26
Const ConsH = 11
Public Const cText = &H0&
Public Const cBG = &H0&
Public Const cHighlight = &HFFFFFF
Public Const cButton = &HC0C0C0
Public Const cShadow = &H0&
Public Const cFG = &HFFFFFF
Public Const cTitleBar = &H6F6FFF

Dim HS(999) As HotSpot
Dim HSCount As Integer
Dim Wins(999) As window
Dim wCount As Integer
Dim arrLayers(999) As String

Function ConsDraw(x As Single, y As Single, color As Long)
    Form1.Picture1.PSet (x, y), color
End Function

Function ConsLine(x1 As Single, x2 As Single, y1 As Single, y2 As Single, color As Long)
    Form1.Picture1.Line (x1, y1)-(x2, y2), color
End Function

Function SetBG(color As Long)
    Form1.Picture1.BackColor = color
End Function

Function ClickEven(x As Single, y As Single, Button As Integer)
    'MsgBox "X: " & X & vbCrLf & "Y: " & Y
    If onWin(x, y) Then
        
    End If
End Function

Function onWin(x As Single, y As Single) As Boolean
    For i = 0 To wCount
        If Wins(i).x < x And Wins(i).y < y Then
            If Wins(i).wWidth + Wins(i).x > x And Wins(i).wHeight + Wins(i).y > y Then
                For x = 0 To wCount
                    arrLayers(Wins(x).layer) = Wins(x).wName
                Next
                RedrawWindow Wins(i).wName
            End If
        End If
    Next
End Function

Function SetHotSpot(x1 As Single, x2 As Single, y1 As Single, y2 As Single)
    i = HSCount
        HS(i).x1 = x1: HS(i).x2 = x2
        HS(i).y1 = y1: HS(i).y2 = y2
    HSCount = i + 1
End Function

Function mButton(x1 As Single, x2 As Single, y1 As Single, y2 As Single, Optional text As String)
    ConsLine x1 + 1, x2, y1, y1, cHighlight
    ConsLine x1, x1, y1, y2, cHighlight
    ConsLine x2, x2, y1, y2, cShadow
    ConsLine x1 + 1, x2 + 1, y2, y2, cShadow
    Dim i As Single
    For i = y1 + 1 To y2 - 1
        ConsLine x1 + 1, x2, i, i, cButton
    Next
    If text <> "" Then
        WriteXY x1 + 2, y1 + 2, text
    End If
End Function

Function mTextbox(x1 As Single, x2 As Single, y1 As Single, y2 As Single, Optional text As String)
    ConsLine x1 + 1, x2, y1, y1, cShadow
    ConsLine x1, x1, y1, y2, cShadow
    ConsLine x2, x2, y1, y2, cHighlight
    ConsLine x1 + 1, x2 + 1, y2, y2, cHighlight
    Dim i As Single
    For i = y1 + 1 To y2 - 1
        ConsLine x1 + 1, x2, i, i, cHighlight
    Next
    If text <> "" Then
        WriteXY x1 + 2, y1 + 2, text
    End If
End Function

Function WriteXY(x As Single, y As Single, text As String)
    Dim newx As Single
    Dim newy As Single
    newx = x + 1
    newy = y
    For i = 1 To Len(text)
        Select Case Mid(text, i, 1)
        Case "A", "a":
            ConsDraw newx + 1, newy, cText: ConsDraw newx + 2, newy, cText
            ConsDraw newx, newy + 1, cText: ConsDraw newx + 3, newy + 1, cText
            ConsDraw newx, newy + 2, cText: ConsDraw newx + 3, newy + 2, cText
            ConsDraw newx, newy + 3, cText: ConsDraw newx + 1, newy + 3, cText
            ConsDraw newx + 2, newy + 3, cText: ConsDraw newx + 3, newy + 3, cText
            ConsDraw newx, newy + 4, cText: ConsDraw newx + 3, newy + 4, cText
            ConsDraw newx, newy + 5, cText: ConsDraw newx + 3, newy + 5, cText
            ConsDraw newx, newy + 6, cText: ConsDraw newx + 3, newy + 6, cText
            newx = newx + 5
        Case "B", "b":
            ConsDraw newx, newy, cText: ConsDraw newx + 1, newy, cText
            ConsDraw newx + 2, newy, cText
            ConsDraw newx, newy + 1, cText: ConsDraw newx + 3, newy + 1, cText
            ConsDraw newx, newy + 2, cText: ConsDraw newx + 3, newy + 2, cText
            ConsDraw newx, newy + 3, cText: ConsDraw newx + 1, newy + 3, cText
            ConsDraw newx + 2, newy + 3, cText
            ConsDraw newx, newy + 4, cText: ConsDraw newx + 3, newy + 4, cText
            ConsDraw newx, newy + 5, cText: ConsDraw newx + 3, newy + 5, cText
            ConsDraw newx, newy + 6, cText: ConsDraw newx + 1, newy + 6, cText
            ConsDraw newx + 2, newy + 6, cText
            newx = newx + 5
        Case "C", "c":
            ConsDraw newx + 1, newy, cText: ConsDraw newx + 2, newy, cText
            ConsDraw newx, newy + 1, cText: ConsDraw newx + 3, newy + 1, cText
            ConsDraw newx, newy + 2, cText
            ConsDraw newx, newy + 3, cText
            ConsDraw newx, newy + 4, cText
            ConsDraw newx, newy + 5, cText: ConsDraw newx + 3, newy + 5, cText
            ConsDraw newx + 1, newy + 6, cText: ConsDraw newx + 2, newy + 6, cText
            newx = newx + 5
        Case "D", "d":
            ConsDraw newx, newy, cText: ConsDraw newx + 1, newy, cText: ConsDraw newx + 2, newy, cText
            ConsDraw newx, newy + 1, cText: ConsDraw newx + 3, newy + 1, cText
            ConsDraw newx, newy + 2, cText: ConsDraw newx + 3, newy + 2, cText
            ConsDraw newx, newy + 3, cText: ConsDraw newx + 3, newy + 3, cText
            ConsDraw newx, newy + 4, cText: ConsDraw newx + 3, newy + 4, cText
            ConsDraw newx, newy + 5, cText: ConsDraw newx + 3, newy + 5, cText
            ConsDraw newx, newy + 6, cText: ConsDraw newx + 1, newy + 6, cText
            ConsDraw newx + 2, newy + 6, cText
            newx = newx + 5
        Case "E", "e":
            ConsDraw newx, newy, cText: ConsDraw newx + 1, newy, cText: ConsDraw newx + 2, newy, cText
            ConsDraw newx + 3, newy, cText
            ConsDraw newx, newy + 1, cText
            ConsDraw newx, newy + 2, cText
            ConsDraw newx, newy + 3, cText: ConsDraw newx + 1, newy + 3, cText
            ConsDraw newx + 2, newy + 3, cText
            ConsDraw newx, newy + 4, cText
            ConsDraw newx, newy + 5, cText
            ConsDraw newx, newy + 6, cText: ConsDraw newx + 1, newy + 6, cText
            ConsDraw newx + 2, newy + 6, cText: ConsDraw newx + 3, newy + 6, cText
            newx = newx + 5
        Case "F", "f":
            ConsDraw newx, newy, cText: ConsDraw newx + 1, newy, cText: ConsDraw newx + 2, newy, cText
            ConsDraw newx + 3, newy, cText
            ConsDraw newx, newy + 1, cText
            ConsDraw newx, newy + 2, cText
            ConsDraw newx, newy + 3, cText: ConsDraw newx + 1, newy + 3, cText
            ConsDraw newx + 2, newy + 3, cText
            ConsDraw newx, newy + 4, cText
            ConsDraw newx, newy + 5, cText
            ConsDraw newx, newy + 6, cText
            newx = newx + 5
        Case "G", "g":
            ConsDraw newx + 1, newy, cText: ConsDraw newx + 2, newy, cText
            ConsDraw newx, newy + 1, cText: ConsDraw newx + 3, newy + 1, cText
            ConsDraw newx, newy + 2, cText
            ConsDraw newx, newy + 3, cText
            ConsDraw newx, newy + 4, cText: ConsDraw newx + 2, newy + 4, cText: ConsDraw newx + 3, newy + 4, cText
            ConsDraw newx, newy + 5, cText: ConsDraw newx + 3, newy + 5, cText
            ConsDraw newx + 1, newy + 6, cText: ConsDraw newx + 2, newy + 6, cText
            ConsDraw newx + 3, newy + 6, cText
            newx = newx + 5
        Case "H", "h":
            ConsDraw newx, newy, cText: ConsDraw newx + 3, newy, cText
            ConsDraw newx + 3, newy, cText: ConsDraw newx + 3, newy, cText
            ConsDraw newx, newy + 1, cText: ConsDraw newx + 3, newy + 1, cText
            ConsDraw newx, newy + 2, cText: ConsDraw newx + 3, newy + 2, cText
            ConsDraw newx, newy + 3, cText: ConsDraw newx + 3, newy + 3, cText
            ConsDraw newx + 1, newy + 3, cText: ConsDraw newx + 2, newy + 3, cText
            ConsDraw newx, newy + 4, cText: ConsDraw newx + 3, newy + 4, cText
            ConsDraw newx, newy + 5, cText: ConsDraw newx + 3, newy + 5, cText
            ConsDraw newx, newy + 6, cText: ConsDraw newx + 3, newy + 6, cText
            newx = newx + 5
        Case "I", "i":
            ConsDraw newx, newy, cText: ConsDraw newx + 1, newy, cText: ConsDraw newx + 2, newy, cText
            ConsDraw newx + 1, newy + 1, cText
            ConsDraw newx + 1, newy + 2, cText
            ConsDraw newx + 1, newy + 3, cText
            ConsDraw newx + 1, newy + 4, cText
            ConsDraw newx + 1, newy + 5, cText
            ConsDraw newx, newy + 6, cText: ConsDraw newx + 1, newy + 6, cText
            ConsDraw newx + 2, newy + 6, cText
            newx = newx + 5
        Case "J", "j":
            ConsDraw newx, newy, cText: ConsDraw newx + 1, newy, cText: ConsDraw newx + 2, newy, cText
            ConsDraw newx + 3, newy, cText
            ConsDraw newx + 2, newy + 1, cText
            ConsDraw newx + 2, newy + 2, cText
            ConsDraw newx + 2, newy + 3, cText
            ConsDraw newx + 2, newy + 4, cText
            ConsDraw newx, newy + 5, cText: ConsDraw newx + 2, newy + 5, cText
            ConsDraw newx + 1, newy + 6, cText
            newx = newx + 5
        Case "K", "k":
            ConsDraw newx, newy, cText: ConsDraw newx + 3, newy, cText
            ConsDraw newx, newy + 1, cText: ConsDraw newx + 2, newy + 1, cText
            ConsDraw newx, newy + 2, cText: ConsDraw newx + 1, newy + 2, cText
            ConsDraw newx, newy + 3, cText: ConsDraw newx + 1, newy + 3, cText
            ConsDraw newx, newy + 4, cText: ConsDraw newx + 2, newy + 4, cText
            ConsDraw newx, newy + 5, cText: ConsDraw newx + 3, newy + 5, cText
            ConsDraw newx, newy + 6, cText: ConsDraw newx + 3, newy + 6, cText
            newx = newx + 5
        Case "L", "l":
            ConsDraw newx, newy, cText
            ConsDraw newx, newy + 1, cText
            ConsDraw newx, newy + 2, cText
            ConsDraw newx, newy + 3, cText
            ConsDraw newx, newy + 4, cText
            ConsDraw newx, newy + 5, cText
            ConsDraw newx, newy + 6, cText: ConsDraw newx + 1, newy + 6, cText
            ConsDraw newx + 2, newy + 6, cText: ConsDraw newx + 3, newy + 6, cText
            newx = newx + 5
        Case "M", "m":
            ConsDraw newx, newy, cText: ConsDraw newx + 3, newy, cText
            ConsDraw newx + 3, newy, cText: ConsDraw newx + 3, newy, cText
            ConsDraw newx, newy + 1, cText: ConsDraw newx + 3, newy + 1, cText
            ConsDraw newx, newy + 3, cText: ConsDraw newx + 3, newy + 3, cText
            ConsDraw newx, newy + 2, cText: ConsDraw newx + 3, newy + 2, cText
            ConsDraw newx + 1, newy + 2, cText: ConsDraw newx + 2, newy + 2, cText
            ConsDraw newx, newy + 4, cText: ConsDraw newx + 3, newy + 4, cText
            ConsDraw newx, newy + 5, cText: ConsDraw newx + 3, newy + 5, cText
            ConsDraw newx, newy + 6, cText: ConsDraw newx + 3, newy + 6, cText
            newx = newx + 5
        Case "N", "n":
            ConsDraw newx, newy, cText: ConsDraw newx + 3, newy, cText
            ConsDraw newx + 3, newy, cText: ConsDraw newx + 3, newy, cText
            ConsDraw newx, newy + 1, cText: ConsDraw newx + 3, newy + 1, cText
            ConsDraw newx, newy + 3, cText: ConsDraw newx + 3, newy + 3, cText
            ConsDraw newx, newy + 2, cText: ConsDraw newx + 3, newy + 2, cText
            ConsDraw newx + 1, newy + 2, cText: ConsDraw newx + 2, newy + 3, cText
            ConsDraw newx, newy + 4, cText: ConsDraw newx + 3, newy + 4, cText
            ConsDraw newx, newy + 5, cText: ConsDraw newx + 3, newy + 5, cText
            ConsDraw newx, newy + 6, cText: ConsDraw newx + 3, newy + 6, cText
            newx = newx + 5
        Case "0", "o", "O":
            ConsDraw newx + 1, newy, cText: ConsDraw newx + 2, newy, cText
            ConsDraw newx, newy + 1, cText: ConsDraw newx + 3, newy + 1, cText
            ConsDraw newx, newy + 2, cText: ConsDraw newx + 3, newy + 2, cText
            ConsDraw newx, newy + 3, cText: ConsDraw newx + 3, newy + 3, cText
            ConsDraw newx, newy + 4, cText: ConsDraw newx + 3, newy + 4, cText
            ConsDraw newx, newy + 5, cText: ConsDraw newx + 3, newy + 5, cText
            ConsDraw newx + 1, newy + 6, cText: ConsDraw newx + 2, newy + 6, cText
            newx = newx + 5
        Case "P", "p":
            ConsDraw newx, newy, cText: ConsDraw newx + 1, newy, cText
            ConsDraw newx + 2, newy, cText
            ConsDraw newx, newy + 1, cText: ConsDraw newx + 3, newy + 1, cText
            ConsDraw newx, newy + 2, cText: ConsDraw newx + 3, newy + 2, cText
            ConsDraw newx, newy + 3, cText: ConsDraw newx + 1, newy + 3, cText
            ConsDraw newx + 2, newy + 3, cText
            ConsDraw newx, newy + 4, cText
            ConsDraw newx, newy + 5, cText
            ConsDraw newx, newy + 6, cText
            newx = newx + 5
        Case "Q", "q":
            ConsDraw newx + 1, newy, cText: ConsDraw newx + 2, newy, cText
            ConsDraw newx, newy + 1, cText: ConsDraw newx + 3, newy + 1, cText
            ConsDraw newx, newy + 2, cText: ConsDraw newx + 3, newy + 2, cText
            ConsDraw newx, newy + 3, cText: ConsDraw newx + 3, newy + 3, cText
            ConsDraw newx, newy + 4, cText: ConsDraw newx + 3, newy + 4, cText
            ConsDraw newx, newy + 5, cText: ConsDraw newx + 3, newy + 5, cText
            ConsDraw newx + 2, newy + 5, cText
            ConsDraw newx + 1, newy + 6, cText: ConsDraw newx + 2, newy + 6, cText
            ConsDraw newx + 3, newy + 6, cText
            newx = newx + 5
        Case "R", "r":
            ConsDraw newx, newy, cText: ConsDraw newx + 1, newy, cText
            ConsDraw newx + 2, newy, cText
            ConsDraw newx, newy + 1, cText: ConsDraw newx + 3, newy + 1, cText
            ConsDraw newx, newy + 2, cText: ConsDraw newx + 3, newy + 2, cText
            ConsDraw newx, newy + 3, cText: ConsDraw newx + 1, newy + 3, cText
            ConsDraw newx + 2, newy + 3, cText
            ConsDraw newx, newy + 4, cText: ConsDraw newx + 3, newy + 4, cText
            ConsDraw newx, newy + 5, cText: ConsDraw newx + 3, newy + 5, cText
            ConsDraw newx, newy + 6, cText: ConsDraw newx + 3, newy + 6, cText
            newx = newx + 5
        Case "S", "s":
            ConsDraw newx + 1, newy, cText: ConsDraw newx + 2, newy, cText
            ConsDraw newx, newy + 1, cText: ConsDraw newx + 3, newy + 1, cText
            ConsDraw newx, newy + 2, cText
            ConsDraw newx + 1, newy + 3, cText: ConsDraw newx + 2, newy + 3, cText
            ConsDraw newx + 3, newy + 4, cText
            ConsDraw newx, newy + 5, cText: ConsDraw newx + 3, newy + 5, cText
            ConsDraw newx + 1, newy + 6, cText: ConsDraw newx + 2, newy + 6, cText
            newx = newx + 5
        Case "T", "t":
            ConsDraw newx, newy, cText: ConsDraw newx + 1, newy, cText: ConsDraw newx + 2, newy, cText
            ConsDraw newx + 1, newy + 1, cText
            ConsDraw newx + 1, newy + 2, cText
            ConsDraw newx + 1, newy + 3, cText
            ConsDraw newx + 1, newy + 4, cText
            ConsDraw newx + 1, newy + 5, cText
            ConsDraw newx + 1, newy + 6, cText
            newx = newx + 5
        Case "U", "u":
            ConsDraw newx, newy, cText: ConsDraw newx + 3, newy, cText
            ConsDraw newx, newy + 1, cText: ConsDraw newx + 3, newy + 1, cText
            ConsDraw newx, newy + 2, cText: ConsDraw newx + 3, newy + 2, cText
            ConsDraw newx, newy + 3, cText: ConsDraw newx + 3, newy + 3, cText
            ConsDraw newx, newy + 4, cText: ConsDraw newx + 3, newy + 4, cText
            ConsDraw newx, newy + 5, cText: ConsDraw newx + 3, newy + 5, cText
            ConsDraw newx + 1, newy + 6, cText: ConsDraw newx + 2, newy + 6, cText
            newx = newx + 5
        Case "V", "v":
            ConsDraw newx, newy, cText: ConsDraw newx + 3, newy, cText
            ConsDraw newx, newy + 1, cText: ConsDraw newx + 3, newy + 1, cText
            ConsDraw newx, newy + 2, cText: ConsDraw newx + 3, newy + 2, cText
            ConsDraw newx, newy + 3, cText: ConsDraw newx + 3, newy + 3, cText
            ConsDraw newx, newy + 4, cText: ConsDraw newx + 3, newy + 4, cText
            ConsDraw newx + 1, newy + 5, cText: ConsDraw newx + 3, newy + 5, cText
            ConsDraw newx + 2, newy + 6, cText
            newx = newx + 5
        Case "W", "w":
            ConsDraw newx, newy, cText: ConsDraw newx + 3, newy, cText
            ConsDraw newx + 3, newy, cText: ConsDraw newx + 3, newy, cText
            ConsDraw newx, newy + 1, cText: ConsDraw newx + 3, newy + 1, cText
            ConsDraw newx, newy + 2, cText: ConsDraw newx + 3, newy + 2, cText
            ConsDraw newx, newy + 3, cText: ConsDraw newx + 3, newy + 3, cText
            ConsDraw newx, newy + 4, cText: ConsDraw newx + 3, newy + 4, cText
            ConsDraw newx, newy + 5, cText: ConsDraw newx + 3, newy + 5, cText
            ConsDraw newx + 1, newy + 5, cText: ConsDraw newx + 2, newy + 5, cText
            ConsDraw newx, newy + 6, cText: ConsDraw newx + 3, newy + 6, cText
            newx = newx + 5
        Case "X", "x":
            ConsDraw newx, newy, cText: ConsDraw newx + 3, newy, cText
            ConsDraw newx, newy + 1, cText: ConsDraw newx + 3, newy + 1, cText
            ConsDraw newx + 1, newy + 2, cText: ConsDraw newx + 2, newy + 2, cText
            ConsDraw newx + 1, newy + 3, cText: ConsDraw newx + 2, newy + 2, cText
            ConsDraw newx + 1, newy + 4, cText: ConsDraw newx + 2, newy + 3, cText
            ConsDraw newx, newy + 5, cText: ConsDraw newx + 3, newy + 4, cText
            ConsDraw newx, newy + 6, cText: ConsDraw newx + 3, newy + 5, cText
            newx = newx + 5
        Case "Y", "y":
            ConsDraw newx, newy, cText: ConsDraw newx + 3, newy, cText
            ConsDraw newx, newy + 1, cText: ConsDraw newx + 3, newy + 1, cText
            ConsDraw newx, newy + 2, cText: ConsDraw newx + 3, newy + 2, cText
            ConsDraw newx + 1, newy + 3, cText: ConsDraw newx + 2, newy + 3, cText
            ConsDraw newx + 3, newy + 3, cText
            ConsDraw newx + 3, newy + 4, cText
            ConsDraw newx + 3, newy + 5, cText
            ConsDraw newx + 3, newy + 6, cText
            newx = newx + 5
        Case "Z", "z":
            ConsDraw newx, newy, cText: ConsDraw newx + 1, newy, cText
            ConsDraw newx + 2, newy, cText: ConsDraw newx + 3, newy, cText
            ConsDraw newx + 3, newy + 1, cText: ConsDraw newx + 3, newy + 2, cText
            ConsDraw newx + 1, newy + 3, cText: ConsDraw newx + 2, newy + 3, cText
            ConsDraw newx, newy + 4, cText: ConsDraw newx, newy + 5, cText
            ConsDraw newx, newy + 6, cText: ConsDraw newx + 1, newy + 6, cText
            ConsDraw newx + 2, newy + 6, cText: ConsDraw newx + 3, newy + 6, cText
            newx = newx + 5
        Case "™":
            ConsDraw newx, newy + 2, cText: ConsDraw newx + 1, newy + 2, cText
            ConsDraw newx + 2, newy + 2, cText
            ConsDraw newx + 1, newy + 3, cText
            newx = newx + 5
        Case "-":
            ConsDraw newx, newy + 2, cText: ConsDraw newx + 1, newy + 2, cText
            ConsDraw newx + 2, newy + 2, cText
            newx = newx + 5
        Case ".":
            ConsDraw newx + 1, newy + 5, cText: ConsDraw newx + 2, newy + 5, cText
            ConsDraw newx + 1, newy + 6, cText: ConsDraw newx + 2, newy + 6, cText
            newx = newx + 5
        Case "!":
            ConsDraw newx + 1, newy, cText: ConsDraw newx + 1, newy + 1, cText
            ConsDraw newx + 1, newy + 2, cText: ConsDraw newx + 1, newy + 3, cText
            ConsDraw newx + 1, newy + 4, cText: ConsDraw newx + 1, newy + 6, cText
            newx = newx + 5
        Case " ":
            newx = newx + 5
        Case Chr(13):
            newy = newy + 9
            newx = x + 1
        End Select
    Next
End Function

Function mWindow(Title As String, x As Single, y As Single, width As Single, height As Single, wName As String)
    i = wCount
        Wins(i).wHeight = height
        Wins(i).wWidth = width
        Wins(i).x = x
        Wins(i).y = y
        Wins(i).wName = wName
        Wins(i).wTitle = Title
        Wins(i).layer = 0
    wCount = i + 1
    
    For a = 0 To wCount - 2
        Wins(a).layer = Wins(a).layer + 1
    Next
    
    mButton x, width + x, y, height + y
    Dim b As Single
    For b = y + 1 To y + 9
        ConsLine x + 1, width + x, b, b, cTitleBar
    Next
    WriteXY x + 2, y + 2, Title
End Function

Function mControl(conType As Byte, wParent As String, x As Single, y As Single, height As Single, width As Single, Value As String, conName As String)
    Select Case conType
    Case 1: 'Button
            For i = 0 To 255
                If Wins(i).wName = wParent Then
                    Wins(i).controls = Wins(i).controls & "btn;" & x & ";" & y & _
                                       ";" & height & ";" & width & ";" & conName & _
                                       ";" & Value & "|"
                    If isWinOnTop(wParent) Then
                        mButton Wins(i).x + x, Wins(i).x + x + width, Wins(i).y + y, Wins(i).y + y + height, Value
                    End If
                End If
            Next
    Case 2: 'Textbox
            For i = 0 To 255
                If Wins(i).wName = wParent Then
                    Wins(i).controls = Wins(i).controls & "txt;" & x & ";" & y & _
                                       ";" & height & ";" & width & ";" & conName & _
                                       ";" & Value & "|"
                    If isWinOnTop(wParent) Then
                        mTextbox Wins(i).x + x, Wins(i).x + x + width, Wins(i).y + y, Wins(i).y + y + height, Value
                    End If
                End If
            Next
    End Select
End Function

Function isWinOnTop(wName As String) As Boolean
    For i = 0 To 255
        If Wins(i).wName = wName Then
            If Wins(i).layer = 0 Then
                isWinOnTop = True
                Exit Function
            End If
        End If
    Next
    isWinOnTop = False
End Function

Function RedrawWindow(wName As String)
On Error Resume Next
    Dim temparr() As String
    Dim temparr2() As String
    For i = 0 To wCount
        With Wins(i)
            If .wName = wName Then
                
                mButton .x, .wWidth + .x, .y, .wHeight + .y
                Dim b As Single
                For b = .y + 1 To .y + 9
                    ConsLine .x + 1, .wWidth + .x, b, b, cTitleBar
                Next
                WriteXY .x + 2, .y + 2, .wTitle
                
                temparr = Split(.controls, "|")
                For x = 0 To UBound(temparr)
                    temparr2 = Split(temparr(x), ";")
                    If temparr2(0) = "txt" Then
                        mControl 2, .wName, Val(temparr2(1)), Val(temparr2(2)), _
                        Val(temparr2(3)), Val(temparr2(4)), temparr2(6), temparr2(5)
                    ElseIf temparr2(0) = "btn" Then
                        mControl 1, .wName, Val(temparr2(1)), Val(temparr2(2)), _
                        Val(temparr2(3)), Val(temparr2(4)), temparr2(6), temparr2(5)
                    End If
                Next
            End If
        End With
    Next
End Function



























