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
