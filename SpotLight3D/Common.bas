Attribute VB_Name = "BasMaths"
Option Explicit

Public Const Pi = 3.141549265
Public Const Deg = (Pi / 180)

Public Declare Function Polygon Lib "Gdi32.dll" (ByVal hDC As Long, lpPoint As Point2D, ByVal nCount As Long) As Long

Public Type Point2D
 X As Long
 Y As Long
End Type

Public Type Vector
 X As Single
 Y As Single
 Z As Single
End Type

Public Type ColRGB
 R As Integer
 G As Integer
 B As Integer
End Type

'############################

Public Type Face
 A As Integer
 B As Integer
 C As Integer
 Normal As Vector
 Center As Vector
 Col As ColRGB
 TCol As ColRGB
 Visible As Boolean
End Type

Public Type SpotLight3D
 Origin As Vector
 Direction As Vector
 Color As ColRGB
 Falloff As Single
 Hotspot As Single
 BrightRange As Single
 DarkRange As Single
 Enabled As Boolean
End Type
Function ColorInterpolate(A As ColRGB, B As ColRGB, Alpha As Single) As ColRGB
 ColorInterpolate.R = A.R + Alpha * (B.R - A.R)
 ColorInterpolate.G = A.G + Alpha * (B.G - A.G)
 ColorInterpolate.B = A.B + Alpha * (B.B - A.B)
End Function
Function ColorAdd(A As ColRGB, B As ColRGB) As ColRGB
 ColorAdd.R = A.R + B.R
 ColorAdd.G = A.G + B.G
 ColorAdd.B = A.B + B.B
End Function
Function ColorScale(A As ColRGB, B As Single) As ColRGB
 ColorScale.R = A.R * B
 ColorScale.G = A.G * B
 ColorScale.B = A.B * B
End Function
Function CrossProduct(V1 As Vector, V2 As Vector) As Vector

 CrossProduct.X = (V1.Y * V2.Z) - (V1.Z * V2.Y)
 CrossProduct.Y = (V1.Z * V2.X) - (V1.X * V2.Z)
 CrossProduct.Z = (V1.X * V2.Y) - (V1.Y * V2.X)

End Function
Function DotProduct(V1 As Vector, V2 As Vector) As Single

 DotProduct = (V1.X * V2.X) + (V1.Y * V2.Y) + (V1.Z * V2.Z)

End Function
Function GetCenter(V1 As Vector, V2 As Vector, V3 As Vector) As Vector

 GetCenter.X = (V1.X + V2.X + V3.X) / 3
 GetCenter.Y = (V1.Y + V2.Y + V3.Y) / 3
 GetCenter.Z = (V1.Z + V2.Z + V3.Z) / 3

End Function
Function GetNormal(V1 As Vector, V2 As Vector, V3 As Vector) As Vector

 GetNormal = CrossProduct(VectorSubtract(V1, V2), VectorSubtract(V3, V2))

End Function
Function Normalize(V As Vector) As Vector

 Dim L As Single

 L = VectorLength(V)

 If L <> 0 Then
  Normalize.X = (V.X / L)
  Normalize.Y = (V.Y / L)
  Normalize.Z = (V.Z / L)
 End If

End Function
Function Rotate(SrcVec As Vector, Axis As Byte, Angle As Single) As Vector

 Select Case Axis
  Case 0:
   Rotate.X = SrcVec.X
   Rotate.Y = (Cos(Angle) * SrcVec.Y) - (Sin(Angle) * SrcVec.Z)
   Rotate.Z = (Sin(Angle) * SrcVec.Y) + (Cos(Angle) * SrcVec.Z)
  Case 1:
   Rotate.X = (Cos(Angle) * SrcVec.X) + (Sin(Angle) * SrcVec.Z)
   Rotate.Y = SrcVec.Y
   Rotate.Z = -(Sin(Angle) * SrcVec.X) + (Cos(Angle) * SrcVec.Z)
  Case 2:
   Rotate.X = (Cos(Angle) * SrcVec.X) - (Sin(Angle) * SrcVec.Y)
   Rotate.Y = (Sin(Angle) * SrcVec.X) + (Cos(Angle) * SrcVec.Y)
   Rotate.Z = SrcVec.Z
 End Select

End Function
Function VectorAdd(VecA As Vector, VecB As Vector) As Vector

 VectorAdd.X = VecA.X + VecB.X
 VectorAdd.Y = VecA.Y + VecB.Y
 VectorAdd.Z = VecA.Z + VecB.Z

End Function
Function VectorAngle(VecA As Vector, VecB As Vector) As Single

 If VectorCompare(VecA, VectorNull) = False And VectorCompare(VecB, VectorNull) = False Then
  VectorAngle = DotProduct(Normalize(VecA), Normalize(VecB))
 End If

End Function
Function VectorCompare(VecA As Vector, VecB As Vector) As Boolean

 If VecA.X = VecB.X And VecA.Y = VecB.Y And VecA.Z = VecB.Z Then VectorCompare = True

End Function
Function VectorDistance(VecA As Vector, VecB As Vector) As Single

 VectorDistance = VectorLength(VectorSubtract(VecB, VecA))

End Function
Function VectorGetXPitch(V1 As Vector, V2 As Vector) As Single

 If VectorCompare(V1, VectorNull) = False And VectorCompare(V2, VectorNull) = False Then
  VectorGetXPitch = VectorAngle(VectorInput(V1.X, 0, V1.Z), VectorInput(V2.X, 0, V2.Z))
 End If

End Function
Function VectorGetYYaw(V1 As Vector, V2 As Vector) As Single

 If VectorCompare(V1, VectorNull) = False And VectorCompare(V2, VectorNull) = False Then
  VectorGetYYaw = VectorAngle(VectorInput(0, V1.Y, V1.Z), VectorInput(0, V2.Y, V2.Z))
 End If

End Function
Function VectorInput(X!, Y!, Z!) As Vector

 VectorInput.X = X
 VectorInput.Y = Y
 VectorInput.Z = Z

End Function
Function VectorLength(V As Vector) As Single

 VectorLength = Sqr((V.X * V.X) + (V.Y * V.Y) + (V.Z * V.Z))

End Function
Function VectorNull() As Vector

End Function
Function VectorReflect(VecA As Vector, VecB As Vector) As Vector

 If VectorAngle(VecA, VecB) < 0 Then
  VectorReflect = VectorAdd(VecA, VectorScale(VectorScale(Normalize(VecB), DotProduct(VecA, Normalize(VecB))), -2))
 End If

End Function
Function VectorScale(Vec As Vector, S As Single) As Vector

 VectorScale.X = Vec.X * S
 VectorScale.Y = Vec.Y * S
 VectorScale.Z = Vec.Z * S

End Function
Function VectorSubtract(V1 As Vector, V2 As Vector) As Vector

 VectorSubtract.X = V1.X - V2.X
 VectorSubtract.Y = V1.Y - V2.Y
 VectorSubtract.Z = V1.Z - V2.Z

End Function
