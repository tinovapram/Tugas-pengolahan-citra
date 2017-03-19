Attribute VB_Name = "Module1"
Option Explicit
' Deklarasi Jenis type Data RGB, untuk keperluan Image Processing
Public Type tRGB24
B As Byte
G As Byte
R As Byte
End Type
Public Declare Function GetPixel Lib "gdi32" ( _
ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Global vImage(0 To 319, 0 To 239) As tRGB24

