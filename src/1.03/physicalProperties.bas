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
    'Revised on 01 May 2013

    idealGasDensity = ((molarMass / 1000) * pressure) / (idealGasConst * (temperature + 273.15))

End Function

Public Function airViscosity(temperature As Double)
    '*****************************************************
    ' Purpose:  Computes the atmospheric air viscosity following Sutherland's formula
    ' Inputs:
    '           temperature in celsius
    ' Returns:   air viscosity in kg m/s
    '*****************************************************
    ' Revised in 01 May 2013

    Const sutherCons = 120
    Const refTemp = 524.07
    Const refVisc = 0.01827

    Dim auxVar1 As Double

    auxVar1 = temperature * 1.8 + 32 + 459.67

    airViscosity = 0.001 * refVisc * ((0.555 * refTemp + sutherCons) / ((0.555 * auxVar1 + sutherCons)) _
    * (auxVar1 / refTemp) ^ (3 / 2))

End Function

Public Function specificSurfaceArea(particleDiameter As Double, voidFraction As Double)
    '*****************************************************
    ' Purpose:  Computes the specific surface area
    ' Inputs:
    '           particleDiameter in m
    '           voidFraction (dimensionless) 0 < voidFraction < 1
    ' Returns:   specific Surface area in m -1 (m2 m-3)
    '*****************************************************
    'Revised in 23 May 2013 - Ref. Handbook of Fluidization - Yang

    specificSurfaceArea = 6 * (1 - voidFraction) / particleDiameter

End Function

Public Function PressureLossErgun(particleDiameter As Double, sphericityParticle As Double, _
voidFraction As Double, gasMassFlux As Double, gasDensity As Double, _
gasDynamicViscosity As Double, bedHeight As Double)
    '*****************************************************
    ' Purpose:  Computes the pressure drop over a packed bed following Ergun Equation with known gas flow
    '           considering the gas as incompressible
    ' Inputs:
    '           particleDiameter in m
    '           sphericityParticle (dimensionless) 0 < sphericityParticle < 1
    '           voidFraction (dimensionless) 0 < voidFraction < 1
    '           gasMassFlux in kg m^-2 s^-1
    '           gasDensity in kg m-3
    '           gasDynamicViscosity in kg m-1 s-1
    '           bedHeight in m
    ' Returns:   pressure loss in Pa
    '*****************************************************
    ' Revised in 29 May 2013
    
    Dim b1, b2 As Double

    b1 = (1.75 * (1 - voidFraction)) / _
    (sphericityParticle * particleDiameter * (voidFraction ^ 3) * gasDensity)
    
    b2 = (150 * (1 - voidFraction) ^ 2 * gasDynamicViscosity) / _
    (sphericityParticle ^ 2 * particleDiameter ^ 2 * voidFraction ^ 3 * gasDensity)
    
    PressureLossErgun = (b2 * gasMassFlux + b1 * gasMassFlux ^ 2) * bedHeight

End Function

Public Function arrheniusEquation(preExponentialFactor As Double, activationEnergy As Double, temperature As Double)
    '*****************************************************
    ' Purpose:  Computes the reaction rate constant following Arrhenius equation
    ' Inputs:
    '           preExponentialFactor in s-1
    '           activationEnergy in kJ mol-1
    '           temperature in K
    ' Returns:   reaction rate constant in s-1
    '*****************************************************
    ' Revised in 29 May 2013

    arrheniusEquation = preExponentialFactor * Exp(-activationEnergy * 1000 / (idealGasConst * temperature))

End Function

Public Function fluidMassFlow(density As Double, velocity As Double, area As Double)
    '*****************************************************
    ' Purpose:  Computes the fluid mass flow given density, velocity and
    '           cross-sectional area, assuming uniform velocity (plug flow) and incompressibility
    ' Inputs:
    '           density in kg m-3
    '           velocity in s-1
    '           area in m-2
    ' Returns:   mass flow in kg s-1
    '*****************************************************
    ' Revised in 29 May 2013
    
    fluidMassFlow = density * velocity * area

End Function

Public Function fluidVelocity(density As Double, massFlow As Double, area As Double)
    '*****************************************************
    ' Purpose:  Computes fluid velocity given density, massFlow and
    '           cross-sectional area, assuming uniform velocity (plug flow) and incompressibility
    ' Inputs:
    '           density in kg m-3
    '           mass flow in kg s-1
    '           area in m-2
    ' Returns:  velocity in s-1
    '*****************************************************
    ' Revised in 29 May 2013
    
    fluidVelocity = massFlow / (density * area)

End Function

Public Function minimumFluidisationVelocity(particleDiameter As Double, particleDensity As Double, _
fluidDensity As Double, fluidDynamicViscosity As Double)
    '*****************************************************
    ' Purpose:   computes the minimum fluidisation velocity for a fluidised bed
    ' Inputs:
    '           particleDiameter in m
    '           particleDensity in kg/m^3,
    '           fluidDensity in kg/m^3
    '           fluidDynamicViscosity in kg m /s
    ' Returns:   minimumFluidisationVelocity in m s-1
    '*****************************************************
    'Revised on 21 Jun 2013
    
    ' Crowe, C., 2006. Multiphase flow handbook, Boca Raton: Taylor & Francis.
    'Different values of the constants, C1 and C2, are available in the literature, with the most popular being
    'C1 ? 33.7 and C2 ? 0.0408, as recommended by Wen and Yu (1966). Typically this equation predicts the minimum
    'fluidization velocity no better than within approximately ? 25%, so it is best to measure it experimentally
    'whenever possible.
    
    Dim c1 As Double
    Dim c2 As Double
    Dim Ar As Double
    
    ' constants values - Wen and Yu (1966)
    c1 = 33.7
    c2 = 0.0408
    ' computes Archimedes Number
    Ar = archimedesNumber(particleDiameter, particleDensity, fluidDensity, fluidDynamicViscosity)
    
    minimumFluidisationVelocity = (fluidDynamicViscosity / (particleDiameter * fluidDensity)) * _
    (((c1 ^ 2 + c2 * Ar) ^ 0.5) - c1)
    
End Function

Public Function terminalVelocity(particleDiameter As Double, particleDensity As Double, fluidDensity As Double, _
fluidDynamicViscosity As Double)
    '*****************************************************
    ' Purpose:  Computes the terminal velocity of a free spherical particle inside a fluid (unhindered)
    '           following Khan and Richardson correlation.
    ' Inputs:
    '           particleDiameter in m
    '           particleDensity in kg m-3
    '           fluidDensity in kg m-3
    '           fluidDynamicViscosity in Pa s
    ' Returns:  velocity in s-1
    '*****************************************************
    ' Revised in 08 Jul 13
    
    Dim var1 As Double
    
    'computes the archimedes number
    var1 = archimedesNumber(particleDiameter, particleDensity, fluidDensity, fluidDynamicViscosity)
    
    terminalVelocity = (fluidDynamicViscosity / (particleDiameter * fluidDensity)) * _
                        (2.33 * var1 ^ 0.018 - 1.53 * var1 ^ -0.016) ^ 13.3

End Function

Public Function geldartGroup(particleDiameter As Double, particleDensity As Double, fluidDensity As Double)
    '*****************************************************
    ' Purpose: returns the Geldart particle group
    ' Inputs:
    '           particleDiameter in m
    '           particleDensity in kg m-3
    '           fluidDensity in kg m-3
    ' Returns:   geldart group [string] - allowed values: A, B, C, D
    '*****************************************************
    ' REVISED IN 20 JUN 13
    
    Dim densityDifference As Double
    Dim diameter As Double
    Dim curvePQ As Double
    Dim equation6 As Double
    Dim equation8 As Double
    
    'converts particle diameter to micrometers
    diameter = particleDiameter * 1000000#
    
    'computes the density difference in g cm-3
    densityDifference = (particleDensity - fluidDensity) / 1000
    
    ' compute values for curves taken from Geldart diagram
    curvePQ = 162.83 * diameter ^ (-1.529)
    equation6 = 228.35 * diameter ^ (-1.003)
    equation8 = 1000000# * diameter ^ (-2.012)
    
    ' Define Geldart group
    If (diameter < 20) Then
        geldartGroup = "C"
    End If
    
    If (diameter >= 20 And diameter <= 50) Then
        If (densityDifference < curvePQ) Then
            geldartGroup = "C"
        Else
            geldartGroup = "A"
        End If
    End If
    
    If (diameter > 50 And diameter < 140) Then
        If (densityDifference < curvePQ) Then
            geldartGroup = "C"
        ElseIf (densityDifference > equation6) Then
            geldartGroup = "B"
        Else
            geldartGroup = "A"
        End If
    End If

    If (diameter >= 140 And diameter < 400) Then
        If (densityDifference < equation6) Then
            geldartGroup = "A"
        Else
            geldartGroup = "B"
        End If
    End If
    
    If (diameter >= 400) Then
        If (densityDifference < equation6) Then
            geldartGroup = "A"
        ElseIf (densityDifference > equation8) Then
            geldartGroup = "D"
        Else
            geldartGroup = "B"
        End If
    End If
    
End Function

Public Function porousCoefficientAlfa(particleDiameter As Double, porosity As Double)
    '*****************************************************
    ' Purpose: computes the porous coefficient alfa (viscous) from Ergun's equation
    ' Inputs:
    '           particleDiameter in m
    '           porosity (dimensionless) - 0 < porosity < 1
    ' Returns:   porousCoefficientAlfa in m2
    '*****************************************************
    ' REVISED IN 21 JUN 13
    
    'Dullien, F., 1991. Porous media: fluid transport and pore structure, San Diego: Academic Press.
    
    porousCoefficientAlfa = 150 * (1 - porosity) ^ 2 / (particleDiameter ^ 2 * porosity ^ 3)
    
End Function

Public Function porousCoefficientC2(particleDiameter As Double, porosity As Double)
    '*****************************************************
    ' Purpose: computes the porous coefficient C2 (inertial) from Ergun's equation
    ' Inputs:
    '           particleDiameter in m
    '           porosity (dimensionless) - 0 < porosity < 1
    ' Returns:   porousCoefficientC2 in m-1
    '*****************************************************
    ' REVISED IN 21 JUN 13
    
    'Dullien, F., 1991. Porous media: fluid transport and pore structure, San Diego: Academic Press.
    
    porousCoefficientC2 = 3.5 * (1 - porosity) / (particleDiameter * porosity ^ 3)
    
End Function

