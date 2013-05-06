Option Explicit
' ************************************************************************************************
' DIMENSIONLESS NUMBERS
' This module constains functions to compute dimensionless numbers
'*************************************************************************************************
'

Public Function archimedesNumber(particleDiameter As Double, particleDensity As Double, gasDensity As Double, gasDynamicViscosity As Double)
    '*****************************************************
    ' Purpose:  Computes the Archimedes Number
    ' Inputs:
    '           particleDiameter in m
    '           particleDensity in kg/m^3,
    '           gasDensity in kg/m^3
    '           gasDynamicViscosity in kg m /s
    ' Returns:   the archimedes number (dimensionless)
    '*****************************************************
    'Revised in 01 May 2013

    archimedesNumber = (particleDiameter ^ 3 * gasDensity * (particleDensity - gasDensity) * gravity) / gasDynamicViscosity ^ 2

End Function

Public Function reynoldsNumber(length As Double, density As Double, velocity As Double, _
viscosity As Double)
    '*****************************************************
    ' Purpose:  Computes the (standard) Reynolds Number
    ' Inputs:
    '           length in m
    '           density in kg/m^3,
    '           velocity in m/s
    '           dynamic viscosity in kg / m s
    ' Returns:   Reynolds number (dimensionless)
    '*****************************************************
    ' Revised in 01 May 2013
    
    reynoldsNumber = (length * density * velocity) / viscosity

End Function

