Attribute VB_Name = "Módulo3"
Option Explicit
' ************************************************************************************************
' PHYSICAL PROPERTIES FUNCTIONS AND CORRELATIONS
' This module constains functions and correlations to compute physical properties
'*************************************************************************************************

Public Function idealGasDensity(molarMass As Double, pressure As Double, temperature As Double)
    '*****************************************************
    ' Purpose:   computes the density of a gas following the ideal gas law
    ' Inputs:
    '           molarMass in g/mol
    '           pressure in pascal,
    '           temperature in celsius
    ' Returns:   the gas density in kg/m^3
    '*****************************************************

    idealGasDensity = ((molarMass / 1000) * pressure) / (idealGasConstant * (temperature + 273.15))

End Function

Public Function airViscosity(temperature As Double)
    '*****************************************************
    ' Purpose:  Computes the atmospheric air viscosity following Sutherland's formula
    ' Inputs:
    '           temperature in celsius
    ' Returns:   air viscosity in kg m/s
    '*****************************************************

    Const sutherCons = 120
    Const refTemp = 524.07
    Const refVisc = 0.01827

    Dim auxVar1 As Double

    auxVar1 = temperature * 1.8 + 32 + 459.67

    airViscosity = 0.000001 * refVisc * ((0.555 * refTemp + sutherCons) / (0.555 * auxVar1 + sutherCons)) _
    * (auxVar1 / refTemp) ^ (3 / 2)

End Function


