Attribute VB_Name = "RGB2HSV"
Public Function RGB2HUE(r As Double, g As Double, b As Double) As Double

    RGB2HUE = RGB2HSV(r, g, b)(0)

End Function

Public Function RGB2SATURATION(r As Double, g As Double, b As Double) As Double

    RGB2SATURATION = RGB2HSV(r, g, b)(1)

End Function

Public Function RGB2VALUE(r As Double, g As Double, b As Double) As Double

    RGB2VALUE = RGB2HSV(r, g, b)(2)

End Function

Private Function RGB2HSV(r As Double, g As Double, b As Double) As Double()

    Dim result(0 To 2) As Double
    
    r = ShapeRGB(r)
    g = ShapeRGB(g)
    b = ShapeRGB(b)
    
    Dim max As Double: max = WorksheetFunction.max(r, g, b)
    Dim min As Double: min = WorksheetFunction.min(r, g, b)
    
    If r = g And g = b And b = r Then
        result(0) = 0
    ElseIf r = max Then
        result(0) = (g - b) / (max - min) * 60
    ElseIf g = max Then
        result(0) = (b - r) / (max - min) * 60 + 120
    ElseIf b = max Then
        result(0) = (r - g) / (max - min) * 60 + 240
    End If
    
    If result(0) < 0 Then
        result(0) = result(0) + 360
    End If
    
    result(0) = Round(result(0), 0)
    result(1) = Round((max - min) / max * 100, 0)
    result(2) = Round(max / 255 * 100, 0)
    
    RGB2HSV = result

End Function

Private Function ShapeRGB(value As Double) As Double
    
    If value < 0 Then
        ShapeRGB = 0
    ElseIf value > 255 Then
        ShapeRGB = 255
    Else
        ShapeRGB = value
    End If

End Function
