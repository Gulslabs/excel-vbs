Sub GetFillColorRGB()
    Dim cellColor As Long
    Dim redVal As Long, greenVal As Long, blueVal As Long
    cellColor = ActiveCell.Interior.Color
    redVal = cellColor Mod 256
    greenVal = (cellColor \ 256) Mod 256
    blueVal = (cellColor \ 65536) Mod 256
    MsgBox "RGB: (" & redVal & ", " & greenVal & ", " & blueVal & ")"
End Sub