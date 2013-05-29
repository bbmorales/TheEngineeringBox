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

