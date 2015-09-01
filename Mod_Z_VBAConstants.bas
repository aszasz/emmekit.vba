Attribute VB_Name = "Mod_Z_VBAConstants"
' Generic global constants

Option Explicit

''''    VBA STANDARD DATA TYPES MAX & MIN    ''''
' Integer
Public Const INTEGER_MAX As Integer = 32767 ' 2 bytes / two's complement
Public Const INTEGER_MIN As Integer = -32768 ' 2 bytes / two's complement
' Long
Public Const LONG_MAX As Long = 2147483647 ' 4 bytes / two's complement
Public Const LONG_MIN As Long = -2147483648# ' 4 bytes / two's complement
' Single
Public Const SINGLE_POS_MIN As Single = 1.401298E-45
Public Const SINGLE_POS_MAX As Single = 3.402823E+38
Public Const SINGLE_NEG_MIN As Single = -3.402823E+38
Public Const SINGLE_NEG_MAX As Single = -1.401298E-45
' Double
Public Const DOUBLE_POS_MIN As Double = 4.94065645841247E-324
Public Const DOUBLE_POS_MAX As Double = 1.79769313486231E+308
Public Const DOUBLE_NEG_MIN As Double = -1.79769313486231E+308
Public Const DOUBLE_NEG_MAX As Double = -4.94065645841247E-324
' Byte
Public Const BYTE_MAX As Byte = 255
Public Const BYTE_MIN As Byte = 0

''''    FLOAT COMPARISSON STANDARD TOLERANCE    ''''
Public Const SINGLE_COMPARISSON_EPSILON As Single = 0.000001 ' 1e-6
Public Const DOUBLE_COMPARISSON_EPSILON As Double = 0.00000000000001 ' 1e-14

''''    FLOAT EMPTY    ''''
Public Const SINGLE_EMPTY As Single = -1.7E+38
Public Const SINGLE_EMPTY_THRESHOLD As Single = -1E+38
Public Const DOUBLE_EMPTY As Double = -1.5E+308
Public Const DOUBLE_EMPTY_THRESHOLD As Double = -1E+308
Public Const SHAPEFILE_EMPTY As Double = -1.7E+38
Public Const SHAPEFILE_EMPTY_THRESHOLD As Double = -1E+38
' shapefiles work with the -1e38 threshold for an IEEE 64-bit double-precision floating-point number


