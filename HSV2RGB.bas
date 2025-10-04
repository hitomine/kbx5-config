Attribute VB_Name = "HSV2RGB"
Public Function HSV2RED(h As Double, s As Double, v As Double) As Double

    HSV2RED = HSV2RGB(h, s, v)(0)
    
End Function

Public Function HSV2GREEN(h As Double, s As Double, v As Double) As Double

    HSV2GREEN = HSV2RGB(h, s, v)(1)
    
End Function

Public Function HSV2BLUE(h As Double, s As Double, v As Double) As Double

    HSV2BLUE = HSV2RGB(h, s, v)(2)
    
End Function

Private Function HSV2RGB(h As Double, s As Double, v As Double) As Double()

    Dim result(0 To 2) As Double

    h = ShapeHue(h)
    s = ShapeSaturation(s)
    v = ShapeValue(v)
    
    Dim max As Double: max = v * 255
    Dim min As Double: min = max * (1 - s)
    
    If h < 60 Then
        result(0) = max
        result(1) = (h / 60) * (max - min) + min
        result(2) = min
    ElseIf h < 120 Then
        result(0) = ((120 - h) / 60) * (max - min) + min
        result(1) = max
        result(2) = min
    ElseIf h < 180 Then
        result(0) = min
        result(1) = max
        result(2) = ((h - 120) / 60) * (max - min) + min
    ElseIf h < 240 Then
        result(0) = min
        result(1) = ((240 - h) / 60) * (max - min) + min
        result(2) = max
    ElseIf h < 300 Then
        result(0) = ((h - 240) / 60) * (max - min) + min
        result(1) = min
        result(2) = max
    Else
        result(0) = max
        result(1) = min
        result(2) = ((360 - h) / 60) * (max - min) + min
    End If
    
    result(0) = Round(result(0), 0)
    result(1) = Round(result(1), 0)
    result(2) = Round(result(2), 0)
    
    HSV2RGB = result

End Function

Private Function ShapeHue(h As Double) As Double
    
    h = h Mod 360
    If h < 0 Then
        ShapeHue = 360 - h
    Else
        ShapeHue = h
    End If
    
End Function

Private Function ShapeSaturation(s As Double) As Double
   
    If s < 0 Then
        ShapeSaturation = 0
    ElseIf s > 100 Then
        ShapeSaturation = 1
    Else
        ShapeSaturation = s / 100
    End If

End Function

Private Function ShapeValue(v As Double) As Double
    
    If v < 0 Then
        ShapeValue = 0
    ElseIf v > 100 Then
        ShapeValue = 1
    Else
        ShapeValue = v / 100
    End If

End Function
