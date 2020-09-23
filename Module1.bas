Attribute VB_Name = "Module1"

Public Function PlotMap(lLat As Integer, lLong As Integer) As String
Dim x As Long ' long value
Dim y As Long ' long value
Dim lXHalf As Long ' half way point of the width of the map
Dim lYHalf As Long ' half way point of the height of the map
Dim l1XPoint As Long ' the value of what 1 deg on the x axis is equal to
Dim l1YPoint As Long ' the value of what 1 deg on the y axis is equal to
Dim lTemp As Long ' temporary value

' get the scale width and scale height values
x = Form1.Picture1.ScaleWidth
y = Form1.Picture1.ScaleHeight

' get the half way mark of the map in pixels
lXHalf = x / 2 ' get the 0 deg mark
lYHalf = y / 2 ' get the 0 deg mark

' determine what 1 deg on the map is equal to
l1XPoint = Form1.Picture1.ScaleWidth / 360
l1YPoint = Form1.Picture1.ScaleHeight / 180

' Latitude

If lLat < 0 Then ' negative value -180 to 0

' make the value positive - just easyier to work with
lLat = lLat * -1

    If lLat = 180 Then ' first point on map
        lLat = 1 ' just use 1 so the hair shows on the map
    Else
        lTemp = 180 - lLat
        
        lLat = lTemp * l1XPoint
    End If

ElseIf lLat = 0 Then
    ' half way point on map
    lLat = lXHalf
Else ' positive value 0 to 180
    
    lLat = (lLat * l1XPoint) + lXHalf
End If

' Longitude

If lLong < 0 Then ' negative value -90 to 0

' make the value positive
lLong = lLong * -1

    If lLong = 90 Then ' first point on map
        lLong = Form1.Picture1.ScaleHeight - 1 ' just use -1 so the hair shows on the map
    Else
        lTemp = 90 + lLong
        
        lLong = lTemp * l1YPoint
    End If

ElseIf lLong = 0 Then
    lLong = lYHalf
Else
    lLong = (90 - lLong) * l1YPoint
    ' lLong = (lLong - lYHalf) * l1YPoint ') - lYHalf
End If

' draw the lines
Form1.Line1.X1 = lLat
Form1.Line1.X2 = lLat

Form1.Line2.Y1 = lLong
Form1.Line2.Y2 = lLong

Form1.Shape1.Left = lLat - 2 ' crosshair center
Form1.Shape1.Top = lLong - 4

PlotMap = lLat & "," & lLong

End Function

Public Function RoundNum(Number As Double) As Integer
' round off the number
' this was taken from PSC, I forgot who uploaded
' this function, but thank you whoever you are

    If Int(Number + 0.5) > Int(Number) Then
        RoundNum = Int(Number) + 1
    Else
        RoundNum = Int(Number)
    End If
    
End Function

Public Function FindLatLong(x As Integer, y As Integer) As String
On Error GoTo errPlot
Dim lXHalf As Long
Dim lYHalf As Long
Dim xPlot As String
Dim yPlot As String
Dim l1XPoint As Long
Dim l1YPoint As Long

DoEvents

lXHalf = Form1.Picture1.ScaleWidth / 2
lYHalf = Form1.Picture1.ScaleHeight / 2

l1XPoint = Form1.Picture1.ScaleWidth / 360
l1YPoint = Form1.Picture1.ScaleHeight / 180

If x = lXHalf Then ' half way point
    xPlot = "0"
ElseIf x < lXHalf Then ' negative value
    xPlot = x - lXHalf
    
    xPlot = xPlot / l1XPoint
Else
    xPlot = (x - lXHalf) * l1XPoint
End If

If y = lYHalf Then ' half way point
    yPlot = "0"
ElseIf y < lYHalf Then ' negative value
    yPlot = y - 90
    
    yPlot = yPlot / l1YPoint
Else
    yPlot = (y / l1YPoint) - lYHalf
End If

FindLatLong = "Lat:" & xPlot & " - Long:" & yPlot

Exit Function
errPlot:
    FindLatLong = "Error..."
    Exit Function
End Function
