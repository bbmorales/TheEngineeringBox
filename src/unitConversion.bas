Option Explicit

' ************************************************************************************************
' UNIT CONVERSION MODULE
' This module constains unit conversion functions
'*************************************************************************************************

'**********************************************************
'LENGTH UNIT CONVERSION FUNCTIONS

Public Function feetToMeter(length As Double)
    '*****************************************************
    ' Purpose: convert from feet (US) to meters (SI)
    ' Inputs:
    '           length in feet
    ' Returns:   length in meters
    '*****************************************************
    ' REVISED IN 23 Mar 2013 - OK - Reference: Perry, 1999

    feetToMeter = 3.048 * 10 ^ -1 * length

End Function

Public Function meterToFeet(length As Double)
    '*****************************************************
    ' Purpose: convert from meters (SI) to feet (US)
    ' Inputs:
    '           length in meters
    ' Returns:   length in feet
    '*****************************************************
    ' REVISED IN 23 Mar 2013 - OK - Reference: Perry, 1999

    meterToFeet = 3.2808 * length

End Function

Public Function inchToMeter(length As Double)
    '*****************************************************
    ' Purpose: convert from inches to meters (SI)
    ' Inputs:
    '           length in inches
    ' Returns:   length in meters
    '*****************************************************
    ' REVISED IN 23 Mar 2013 - OK - Reference: Perry, 1999

    inchToMeter = (2.54 * 10 ^ -2) * length

End Function

Public Function meterToInch(length As Double)
    '*****************************************************
    ' Purpose: convert from inches to meters (SI)
    ' Inputs:
    '           length in meters
    ' Returns:   length in inches
    '*****************************************************
    ' REVISED IN 23 Mar 2013 - OK - Reference: Perry, 1999

    meterToInch = ((2.54 * 10 ^ -2) ^ -1) * length

End Function

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
    'Revised on 01 May 2013
    
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

Public Function paToMbar(pressure As Double)
    '*****************************************************
    ' Purpose: converts from Pa to mbar
    ' Inputs:
    '           pressure in Pa
    ' Returns:   pressure in mbar
    '*****************************************************

    paToMbar = 10 ^ -2 * pressure

End Function

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

Public Function paToAtm(pressure As Double)
    '*****************************************************
    ' Purpose: converts from Pa to Atm
    ' Inputs:
    '           pressure in Pa
    ' Returns:   pressure in Atm
    '*****************************************************

    paToAtm = (1.0135 * 10 ^ 5) ^ -1 * pressure

End Function

Public Function atmToPa(pressure As Double)
    '*****************************************************
    ' Purpose: converts from Atm to Pa
    ' Inputs:
    '           pressure in Atm
    ' Returns:   pressure in Pa
    '*****************************************************

    atmToPa = (1.0135 * 10 ^ 5) * pressure

End Function

Public Function paToBar(pressure As Double)
    '*****************************************************
    ' Purpose: converts from Pa to Bar
    ' Inputs:
    '           pressure in Pa
    ' Returns:   pressure in Bar
    '*****************************************************

    paToBar = (1# * 10 ^ 5) ^ -1 * pressure

End Function

Public Function barToPa(pressure As Double)
    '*****************************************************
    ' Purpose: converts from Bar to Pa
    ' Inputs:
    '           pressure in Bar
    ' Returns:   pressure in Pa
    '*****************************************************

    barToPa = (1# * 10 ^ 5) * pressure

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
    ' REVISED IN 23 Mar 2013 - OK - Reference: Perry, 1999

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

Public Function FahrenheitToKelvin(temperature As Double)
    '*****************************************************
    ' Purpose: converts from Fahrenheit to Kelvin
    ' Inputs:
    '           temperature in Fahrenheit
    ' Returns:   temperature in Kelvin
    '*****************************************************
    ' REVISED IN 23 Mar 2013 - OK - Reference: Perry, 1999

    FahrenheitToKelvin = (5 / 9) * (temperature + 459.67)

End Function

Public Function FahrenheitToCelsius(temperature As Double)
    '*****************************************************
    ' Purpose: converts from Fahrenheit to Celsius
    ' Inputs:
    '           temperature in Fahrenheit
    ' Returns:   temperature in Celsius
    '*****************************************************
    ' REVISED IN 23 Mar 2013 - OK - Reference: Perry, 1999

    FahrenheitToCelsius = (5 / 9) * (temperature - 32)

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

Public Function CalToJoule(energy As Double)
    '*****************************************************
    ' Purpose: converts energy from Calories to Joules
    ' Inputs:
    '           energy in Calories
    ' Returns:   energy in Joules
    '*****************************************************
    ' REVISED IN 23 Mar 2013 - OK - Reference: Perry, 1999

    CalToJoule = 4.184 * energy

End Function

Public Function KwhToCal(energy As Double)
    '*****************************************************
    ' Purpose: converts energy from kWh to Calories
    ' Inputs:
    '           energy in kWh
    ' Returns:   energy in Calories
    '*****************************************************

    KwhToCal = 3.6 * 10 ^ 3 * (1.987 / 8.314) * energy

End Function

Public Function KwhToJoule(energy As Double)
    '*****************************************************
    ' Purpose: converts energy from kWh to Joules
    ' Inputs:
    '           energy in kWh
    ' Returns:   energy in Joule
    '*****************************************************

    KwhToJoule = 3.6 * 10 ^ 3 * energy

End Function

Public Function JouleToKwh(energy As Double)
    '*****************************************************
    ' Purpose: converts energy from Joule to kWh
    ' Inputs:
    '           energy in Joule
    ' Returns:   energy in kWh
    '*****************************************************

    JouleToKwh = (3.6 * 10 ^ 3) ^ -1 * energy

End Function

Public Function CalToKwh(energy As Double)
    '*****************************************************
    ' Purpose: converts energy from Cal to kWh
    ' Inputs:
    '           energy in Calories
    ' Returns:   energy in kWh
    '*****************************************************

    CalToKwh = 4.184 * 3.6 * 10 ^ 3 * energy

End Function

