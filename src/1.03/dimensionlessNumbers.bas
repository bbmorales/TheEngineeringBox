' Copyright (c) 2013 Bayard Beling Morales

'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:

'The above copyright notice and this permission notice shall be included in
'all copies or substantial portions of the Software.

'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
'THE SOFTWARE.

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

