Attribute VB_Name = "METALS"
Option Explicit

'Declare type properties of materials:
Type material
    name   As String * 20   'Material common name
    v      As Double        'Poisson's ratio: unitless
    E      As Double        'Youngs Modulus of elasticity: psi
    ys     As Double        'Yield strength (cmprssn/tnsn): psi
End Type

'Declare an array of some common materials:
Global matArray(0 To 8) As material

Sub SetProperties()
'The exact value of these properties vary widely!
'So design with a large safety factor.
    'Material Array index counter:
    Dim n As Integer
    
    n = 0  'Starting index 'change to n=1 or n = 0 as necessary.
    'Aluminum (2024-T4)
    matArray(n).name = "Aluminum (2024-T4)"
    matArray(n).v = 0.32           'Poisson's ratio: unitless
    matArray(n).E = 10600000#      'Youngs Modulus of elasticity: psi
    matArray(n).ys = 48000         'Yield strength (cmprssn/tnsn): psi
    n = n + 1
    'Aluminum (7075-T6)
    matArray(n).name = "Aluminum (7075-T6)"
    matArray(n).v = 0.32           'Poisson's ratio: unitless
    matArray(n).E = 10400000#      'Youngs Modulus of elasticity: psi
    matArray(n).ys = 72000         'Yield strength (cmprssn/tnsn): psi
    n = n + 1
    'Red Brass, soft
    matArray(n).name = "Red Brass, soft"
    matArray(n).v = 0.33           'Poisson's ratio: unitless
    matArray(n).E = 15000000#      'Youngs Modulus of elasticity: psi
    matArray(n).ys = 15000         'Yield strength (cmprssn/tnsn): psi
    n = n + 1
    'Polycarbonate
    matArray(n).name = "Polycarbonate"
    matArray(n).v = 0.285          'Poisson's ratio: unitless
    matArray(n).E = 340000#        'Youngs Modulus of elasticity: psi
    matArray(n).ys = 8700          'Yield strength (cmprssn/tnsn): psi
    n = n + 1
    'Steel-1018
    matArray(n).name = "Steel-1018"
    matArray(n).v = 0.285          'Poisson's ratio: unitless
    matArray(n).E = 30000000#      'Youngs Modulus of elasticity: psi
    matArray(n).ys = 32000         'Yield strength (cmprssn/tnsn): psi
    n = n + 1
    'Steel-4130/4340
    matArray(n).name = "Steel-4130/4340"
    matArray(n).v = 0.285          'Poisson's ratio: unitless
    matArray(n).E = 30000000#      'Youngs Modulus of elasticity: psi
    matArray(n).ys = 45000         'Yield strength (cmprssn/tnsn): psi
    n = n + 1
    'Steel-304 SS
    matArray(n).name = "Steel-304 SS"
    matArray(n).v = 0.25           'Poisson's ratio: unitless
    matArray(n).E = 28000000#      'Youngs Modulus of elasticity: psi
    matArray(n).ys = 35000         'Yield strength (cmprssn/tnsn): psi
    n = n + 1
    'Steel-410 SS
    matArray(n).name = "Steel-410 SS"
    matArray(n).v = 0.28           'Poisson's ratio: unitless
    matArray(n).E = 29000000#      'Youngs Modulus of elasticity: psi
    matArray(n).ys = 40000         'Yield strength (cmprssn/tnsn): psi
    n = n + 1
    'Titanium alloy
    matArray(n).name = "Titanium alloy"
    matArray(n).v = 0.34           'Poisson's ratio: unitless
    matArray(n).E = 14000000#      'Youngs Modulus of elasticity: psi
    matArray(n).ys = 135000        'Yield strength (cmprssn/tnsn): psi
End Sub

