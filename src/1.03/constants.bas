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
' PHYSICAL AND AUXILIARY CONSTANTS
' This module constains fundamental physical constanst and auxiliary constants

' DECLARING CONSTANTS
' The declaration of constants shall follow the example below:

' Description of the constant
' Public Const nameOf TheConstant as double 0.00000000

' SOURCES:
' NIST Reference on constants, units and uncertainity - http://physics.nist.gov/cuu/index.html

'*************************************************************************************************

' This section declares constants to be used WITHIN the TEB code
' Pi - [dimensionless]
Public Const Pi As Double = 3.14159265

' Standard gravity - SI [m s^-2]
Public Const gravityConst As Double = 9.80665

' Ideal gas constant [J mol^-1 K^-1]
Public Const idealGasConst As Double = 8.31446

' Stefan-boltzmann constant [W m^-2 K^-4]
Public Const stefanBoltzmannConst As Double = 0.000000056704

'This section declares constants to be used WITHIN Spreadsheets

Public Function gravity()
    '*****************************************************
    ' Purpose: returns the value of Standard acceleration of gravity
    ' Inputs:
    '           none
    ' Returns:   gravity in m s-2
    '*****************************************************
    ' Revised on 01 May 2013

    gravity = 9.80665

End Function

Public Function molarGasConstant()
    '*****************************************************
    ' Purpose: returns the value of molar gas constant R - also know as ideal gas constant
    ' Inputs:
    '           none
    ' Returns:   molarGasConstant in J mol-1 K-1
    '*****************************************************
    ' Revised on 01 May 2013

    molarGasConstant = 8.3144621

End Function
Public Function stefanBoltzmann()
    '*****************************************************
    ' Purpose: returns the value of Stefan-Boltzmann constant
    ' Inputs:
    '           none
    ' Returns:   stefanBoltzmann in W m-2 K-4
    '*****************************************************
    ' Revised on 01 May 2013

    stefanBoltzmann = 5.670373 * 10 ^ -8

End Function

