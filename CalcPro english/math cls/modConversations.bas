Attribute VB_Name = "modFuncConvertAng"
Public Function Deg2Rad(ByVal deg As Double) As Double

    Deg2Rad = deg / 180 * PI
    
End Function


Public Function Rad2Deg(ByVal rad As Double) As Double

    Rad2Deg = rad / PI * 180
    
End Function
