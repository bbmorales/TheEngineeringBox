Attribute VB_Name = "Módulo1"
Option Explicit

' ************************************************************************************************
' UNIT CONVERSION MODULE
' This module constains unit conversion functions
'*************************************************************************************************

'**********************************************************
'VOLUME UNIT CONVERSION FUNCTIONS

Public Function m3ToNm3(volume As Double, temperature As Double, pressure As Double)
    '*****************************************************
    ' Purpose: converts gas volume in m^3 at a specifiec temperature and pressure to Nm^3
    ' (volume in Normal Conditions, - 273.15 K and 101325 Pa)
    ' Inputs:
    '           volume in m^3
    '           temperature in Kelvin
    '           pressure in Pa
    ' Returns:   volume in Nm^3
    '*****************************************************

    m3ToNm3 = volume * (pressure / temperature) * (273.15 / 101325#)

End Function

Public Function Nm3ToM3(volume As Double, temperature As Double, pressure As Double)
    '*****************************************************
    ' Purpose:  converts from volume in Normal Conditions (273.15K, 101325 Pa) Nm^3  to
    ' volume in the user defined temperature (K) and presure (Pa)
    ' Inputs:
    '           volume in Nm^3
    '           temperature in Kelvin
    '           pressure in Pa
    ' Returns:   volume in m^3
    '*****************************************************

    Nm3ToM3 = volume * (temperature / 273.15) * (101325 / pressure)

End Function

Public Function NLToM3(volume As Double, temperature As Double, pressure As Double)
    '*****************************************************
    ' Purpose:  ' function which converts from volume in Normal liter (liters in Normal
    ' Conditions)(273.15K, 101325 Pa) NL  to volume in cubic meters in the user specified
    'temperature (K) and presure (Pa)
    ' Inputs:
    '           volume in Nm^3
    '           temperature in Kelvin
    '           pressure in Pa
    ' Returns:   volume in m^3
    '*****************************************************

    NLToM3 = volume * (temperature / 273.15) * (101325 / pressure) * (1 / (60000#))

End Function

Public Function KgToNm3(massKg As Double, molarMass As Double)
    '*****************************************************
    ' Purpose:  converts a gas mass in kg to Nm3
    '(volume in Normal conditions - 273.15K, 101325 Pa)
    ' Inputs:
    '           massKg in kg
    '           molarMass in g/mol
    ' Returns:   volume in Nm^3
    '*****************************************************

    KgToNm3 = massKg * (1344600# / molarMass) * 0.06 / (3600#)

End Function

'**********************************************************
'PRESSURE UNIT CONVERSION FUNCTIONS

Public Function PaToMmH20(pressure As Double)
    '*****************************************************
    ' Purpose: converts from Pa to mmH20
    ' Inputs:
    '           pressure in Pa
    ' Returns:   pressure in mmH2O
    '*****************************************************

    PaToMmH20 = 0.101974 * pressure

End Function

Public Function mmH20ToPa(pressure As Double)
    '*****************************************************
    ' Purpose: converts from mmH2O to Pa
    ' Inputs:
    '           pressure in mmH2O
    ' Returns:   pressure in Pa
    '*****************************************************

    mmH20ToPa = 9.80642 * pressure

End Function

'**********************************************************
'TEMPERATURE UNIT CONVERSION FUNCTIONS

Public Function KelvinToCelsius(temperature As Double)
    '*****************************************************
    ' Purpose: converts from Kelvin to Celsius
    ' Inputs:
    '           temperature in Kelvin
    ' Returns:   temperature in Celsius
    '*****************************************************

    KelvinToCelsius = temperature - 273.15

End Function

Public Function CelsiusToKelvin(temperature As Double)
    '*****************************************************
    ' Purpose: converts from Kelvin to Celsius
    ' Inputs:
    '           temperature in Celsius
    ' Returns:   temperature in Kelvin
    '*****************************************************

    CelsiusToKelvin = temperature + 273.15

End Function

Public Function RankineToCelsius(temperature As Double)
    '*****************************************************
    ' Purpose: converts from Rankine to Celsius
    ' Inputs:
    '           temperature in Rankine
    ' Returns:   temperature in Celsius
    '*****************************************************

    RankineToCelsius = (temperature - 491.67) * (5 / 9)

End Function

Public Function KelvinToRankine(temperature As Double)
    '*****************************************************
    ' Purpose: converts from Kelvin to Rankine
    ' Inputs:
    '           temperature in Rankine
    ' Returns:   temperature in Celsius
    '*****************************************************

    KelvinToRankine = temperature * (9 / 5)

End Function

'**********************************************************
'ENERGY UNIT CONVERSION FUNCTIONS

Public Function JouleTocal(energy As Double)
    '*****************************************************
    ' Purpose: converts energy from Joule to Calories
    ' Inputs:
    '           energy in joules
    ' Returns:   energy in Calories
    '*****************************************************

    JouleTocal = (1.987 / 8.314) * energy

End Function


