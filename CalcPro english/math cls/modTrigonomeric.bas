Attribute VB_Name = "modFuncTrigonomeric"
Option Explicit

Public Function CCos(ByRef z As Complex) As Complex

    If z.Im = 0 Then
    
        CCos.Im = 0
        CCos.Re = Cos(z.Re)
        
    Else
    
        Dim z1 As Complex, z2 As Complex, zi As Complex
        'e^iz
        zi.Im = 1
        zi.Re = 0
        z1 = CMul(z, zi)
        z1 = CExp(z1)
        'e^-iz
        zi.Im = -1
        z2 = CMul(z, zi)
        z2 = CExp(z2)
        'e^iz+e^-iz
        CCos = CAdd(z1, z2)
        z1.Re = 2
        z1.Im = 0
        '(e^iz+e^-iz)/2
        CCos.Re = CCos.Re / 2
        CCos.Im = CCos.Im / 2
        
    End If
End Function


Public Function CSin(ByRef z As Complex) As Complex

    If z.Im = 0 Then
    
        CSin.Im = 0
        CSin.Re = Sin(z.Re)
        
    Else
    
        Dim z1 As Complex, z2 As Complex, zi As Complex
        'e^iz
        zi.Im = 1
        zi.Re = 0
        z1 = CMul(z, zi)
        z1 = CExp(z1)
        'e^-iz
        zi.Im = -1
        z2 = CMul(z, zi)
        z2 = CExp(z2)
        'e^iz-e^-iz
        CSin = CSub(z1, z2)
        z1.Re = 0
        z1.Im = 2
        '(e^iz-e^-iz)/2i
        CSin = CDiv(CSin, z1)
        
    End If
End Function


Public Function CTan(ByRef z As Complex) As Complex

    CTan = CDiv(CSin(z), CCos(z))
    
End Function

Public Function CAtan(ByRef z As Complex) As Complex

    If z.Im = 0 Then
    
        CAtan.Im = 0
        CAtan.Re = Atn(z.Re)
        
    Else
    
        Dim z1 As Complex, z2 As Complex, zi As Complex, zo As Complex
        
        zi.Im = 1
        zi.Re = 0
        
        zo.Im = 0
        zo.Re = 1
        '
        z1 = CMul(z, zi)    '
        z1 = CLn(CSub(zo, z1))   '
        '
        z2 = CMul(z, zi)
        z2 = CAdd(zo, z2)
        z2 = CLn(z2)    '
        '
        z1 = CSub(z1, z2)   '
        z1 = CMul(z1, zi)
        
        z2.Im = 0
        z2.Re = 2
        CAtan = CDiv(z1, z2)
        
    End If
    
End Function

'ovo sve valja
'sec(z) = 1/cos(z)
'Public Function CSec(ByRef z As Complex) As Complex
'    CSec = CDiv(createComplex(1, 0), CCos(z))
'End Function
''cosec(z) = 1/sin(z)
'Public Function CCosec(ByRef z As Complex) As Complex
'    CSec = CDiv(createComplex(1, 0), CSin(z))
'End Function
''cotan(z) = 1/tan(z)
'Public Function Ccotan(ByRef z As Complex) As Complex
'    CSec = CDiv(createComplex(1, 0), CTan(z))
'End Function
''arcsin(z) = atan(z / sqr(-z * z + 1));
'Public Function CArcsin(ByRef z As Complex) As Complex
'    Dim z1 As Complex
'    '
'    z1 = CMul(z, createComplex(-1, 0))
'    z1 = CMul(z, z1)
'    z1.Re = z1.Re + 1
'    z1 = CSqr(z1)
'    CArcsin = CAtan(CDiv(z, z1))
'End Function
''arccos(z) = atan(-$x / sqr(-$x * $x + 1)) + 2 * atan(1);
'Public Function CArccos(ByRef z As Complex) As Complex
'    Dim z1 As Complex
'    '
'    z1 = CMul(z, createComplex(-1, 0))
'    z1 = CMul(z, z1)
'    z1.Re = z1.Re + 1
'    z1 = CSqr(z1)
'    CArccos = CAtan(CDiv(z, z1))
'    CArccos.Re = CArccos.Re + 2 * Atn(1)
'End Function

'arcsec($X)= atan($x / sqr($x * $x- 1)) + sgn(($x) - 1) * (2 * atan(1));
'Public Function arcsec(ByRef z As Complex) As Complex
'
'    Dim z1 As Complex
'    '
'    z1 = CMul(z, createComplex(-1, 0))
'    z1 = CMul(z, z1)
'    z1.Re = z1.Re + 1
'    z1 = CSqr(z1)
'    CArccos = CAtan(CDiv(z, z1))
'
'    '? sgn za cplx
'
'End Function

