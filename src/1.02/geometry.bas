Option Explicit
' ************************************************************************************************
' GEOMETRY FUNCTIONS
' This module constains functions to compute geometrical relations
'*************************************************************************************************
'

Public Function circleArea(diameter As Double)
    '*****************************************************
    ' Purpose: compute the area of a circle
    ' Inputs:
    '           diameter in meters
    ' Returns:   area in square meters
    '*****************************************************
    ' REVISED IN 27 May 2013

    circleArea = Pi * (diameter / 2) ^ 2

End Function

Public Function circlePerimeter(diameter As Double)
    '*****************************************************
    ' Purpose: compute the perimeter of a circle
    ' Inputs:
    '           diameter in meters
    ' Returns:   perimeter in meters
    '*****************************************************
    ' REVISED IN 27 May 2013

    circleArea = 2 * Pi * (diameter / 2)

End Function


Public Function cylinderVolume(diameter As Double, height As Double)
    '*****************************************************
    ' Purpose: compute the volume of a regular cylinder
    ' Inputs:
    '           diameter in meters
    '           height in meters
    ' Returns:   volume in cubic meters
    '*****************************************************
    ' REVISED IN 23 May 2013

    cylinderVolume = height * Pi * (diameter / 2) ^ 2

End Function

Public Function cylinderArea(diameter As Double, height As Double)
    '*****************************************************
    ' Purpose: compute the external area of a regular cylinder
    ' Inputs:
    '           diameter in meters
    '           height in meters
    ' Returns:   area in square meters
    '*****************************************************
    ' REVISED IN 27 May 2013

    cylinderArea = 2 * Pi * (diameter / 2) ^ 2 + 2 * Pi * height * (diameter / 2)

End Function

Public Function sphereVolume(diameter As Double)
    '*****************************************************
    ' Purpose: compute the volume of a sphere
    ' Inputs:
    '           diameter in meters
    ' Returns:   volume in cubic meters
    '*****************************************************
    ' REVISED IN 27 May 2013

    sphereVolume = (4 / 3) * Pi * (diameter / 2) ^ 3

End Function

Public Function sphereArea(diameter As Double)
    '*****************************************************
    ' Purpose: compute the external area of a sphere
    ' Inputs:
    '           diameter in meters
    ' Returns:   area in square meters
    '*****************************************************
    ' REVISED IN 27 May 2013

    sphereArea = 4 * Pi * (diameter / 2) ^ 2

End Function

Public Function ellipseArea(majorSemiAxis As Double, minorSemiAxis As Double)
    '*****************************************************
    ' Purpose: compute the area of a ellipse
    ' Inputs:
    '           diameter in meters
    ' Returns:   area in square meters
    '*****************************************************
    ' REVISED IN 27 May 2013

    ellipsePerimeter = Pi * majorSemiAxis * minorSemiAxis

End Function

Public Function ellipsePerimeter(majorSemiAxis As Double, minorSemiAxis As Double)
    '*****************************************************
    ' Purpose: compute the perimeter (circunference) of a ellipse following Ramanujan approximation
    ' Inputs:
    '           diameter in meters
    ' Returns:   perimeter in meters
    '*****************************************************
    ' REVISED IN 27 May 2013

    ellipsePerimeter = Pi * (3 * (majorSemiAxis + minorSemiAxis) - (10 * majorSemiAxis * minorSemiAxis + _
    3 * (majorSemiAxis ^ 2 + minorSemiAxis ^ 2)) ^ (1 / 2))

End Function
