VERSION 5.00
Begin VB.Form Main 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   DrawStyle       =   5  'Transparent
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const CNear = 50                  'Near and Far Z value of camera
Const CFar = 0

Dim File3DName As String          'Path and name for a 3d model (inclued)

Dim Cam As Vector                 'Camera vector
Dim Spot As SpotLight3D           'Spotlight
Dim Amb As ColRGB                 'Ambient color

Dim Vecs() As Vector              'Origin vectors array
Dim Tmps() As Vector              'Transformed vectors array
Dim Faces() As Face               '1 Face = 3 connected vectors

Dim OptSpecRef As Boolean         'Specular On/Off
Dim OptAttenuation As Boolean     'Light attenuation
Dim OptFog As Boolean             'Fog = Camera attenuation

Dim PP(2) As Point2D              'Rasterization
Dim XAng!, YAng!, ZAng!, I&, Dp!
Sub Calc()

 'Rotate vectors on 3 axis:
 For I = LBound(Vecs) To UBound(Vecs)
  Tmps(I) = Rotate(Vecs(I), 0, XAng)
  Tmps(I) = Rotate(Tmps(I), 1, YAng)
  Tmps(I) = Rotate(Tmps(I), 2, ZAng)
 Next I

 'Get the normal and the center, then check the
 ' visibility from the camera.
 For I = LBound(Faces) To UBound(Faces)
  Faces(I).Normal = GetNormal(Tmps(Faces(I).A), Tmps(Faces(I).B), Tmps(Faces(I).C))
  Faces(I).Center = GetCenter(Tmps(Faces(I).A), Tmps(Faces(I).B), Tmps(Faces(I).C))

  Dp = VectorAngle(Cam, Faces(I).Normal)
  If Dp < 0 Then
   Faces(I).Visible = True
   Shade I
  Else
   Faces(I).Visible = False
  End If

 Next I

 'Perspective distortion, for a realistic projection,
 ' so it's look Ok !
 For I = LBound(Vecs) To UBound(Vecs)
  Tmps(I).X = (Tmps(I).X / (Tmps(I).Z + 100)) * 250
  Tmps(I).Y = (Tmps(I).Y / (Tmps(I).Z + 100)) * 250
 Next I

End Sub
Sub Draw()

 'Rasterization (2D):
 ' Draw the faces only visible with 'TCol' (Transformed color),
 ' with an origine(0, 0, 0) as (320, 240, 0) (center screen)
 For I = LBound(Faces) To UBound(Faces)
  If Faces(I).Visible = True Then
   If Spot.Enabled = True Then
    FillColor = RGB(Faces(I).TCol.R, Faces(I).TCol.G, Faces(I).TCol.B)
   Else
    FillColor = RGB(Faces(I).Col.R, Faces(I).Col.G, Faces(I).Col.B)
   End If
   PP(0).X = 320 + Tmps(Faces(I).A).X: PP(0).Y = 240 + Tmps(Faces(I).A).Y
   PP(1).X = 320 + Tmps(Faces(I).B).X: PP(1).Y = 240 + Tmps(Faces(I).B).Y
   PP(2).X = 320 + Tmps(Faces(I).C).X: PP(2).Y = 240 + Tmps(Faces(I).C).Y
   Polygon Me.hDC, PP(0), 3
  End If
 Next I

End Sub
Sub LoadScene()

 'Load 3d model (you can change the file name)

 File3DName = App.Path & "\Primatives\Grid.dex"

 Dim Buff As Long

 Open File3DName For Binary As 1

  Get #1, , Buff
  ReDim Vecs(CLng(Buff))
  ReDim Tmps(CLng(Buff))
  Get #1, , Buff
  ReDim Faces(CLng(Buff))

  For I = LBound(Vecs) To UBound(Vecs)
   Get #1, , Vecs(I).X
   Get #1, , Vecs(I).Y
   Get #1, , Vecs(I).Z
  Next I

  For I = LBound(Faces) To UBound(Faces)
   Get #1, , Faces(I).A
   Get #1, , Faces(I).B
   Get #1, , Faces(I).C
   Faces(I).Col.G = 100
  Next I

 Close 1

 '###################################

 'Do somme stuffs (setup the camera and the spotlight)

 Cam = VectorInput(0, 0, -1)
 Amb.R = 10: Amb.G = 10: Amb.B = 10

 OptSpecRef = True: OptAttenuation = True: OptFog = True

 Spot.Origin = VectorInput(10, 10, 10)
 Spot.Direction = VectorInput(0, 0, 0)
 Spot.DarkRange = CFar: Spot.BrightRange = CNear
 Spot.Falloff = 1: Spot.Hotspot = 0.6
 Spot.Color.R = 100: Spot.Color.G = 100: Spot.Color.B = 100
 Spot.Enabled = True

End Sub
Sub Shade(FaceIndex As Long)

 'Shade faces:

 Dim ColorSum As ColRGB
 Dim Alpha!, Beta!, Gamma!, Delta!, Epsilon!
 '  Diffuse  Spec   Attenu   Fog     Spot

 If Spot.Enabled = True Then     ' Only if the spot is turned On

  '################################################

  'Note: For the PointLight algorithm (Sphere of light),
  '       you change simply the Epsilon value,
  '        (I have not find), if you find this, mail me a copy !

  'use spotlight filter
  Epsilon = 1
  If Spot.Falloff > 0 Then
   Epsilon = VectorAngle(Spot.Direction, VectorSubtract(Faces(FaceIndex).Center, Spot.Origin))
   If Epsilon < 0 Then Epsilon = 0
   If Spot.Falloff <> Spot.Hotspot Then
    Epsilon = (Spot.Falloff - Epsilon) / (Spot.Falloff - Spot.Hotspot)
    If Epsilon < 0 Then Epsilon = 0 Else If Epsilon > 1 Then Epsilon = 1
   Else
    Exit Sub
   End If
  End If

  '################################################
  'perform incident ray shading (Diffusion)
  Alpha = VectorAngle(VectorSubtract(Spot.Origin, Faces(FaceIndex).Center), Faces(FaceIndex).Normal)
  If Alpha < 0 Then Alpha = 0

  '################################################
  'perform reflected ray shading (Specular)
  If OptSpecRef = True Then
   Beta = VectorAngle(VectorReflect(VectorSubtract(Faces(FaceIndex).Center, Spot.Origin), Faces(FaceIndex).Normal), VectorSubtract(Cam, Faces(FaceIndex).Center))
   If Beta < 0 Then Beta = 0
  End If

  '################################################
  'apply light distance decay (Attenuation)
  Gamma = 1
  If OptFog = True Then
   If Spot.DarkRange <> Spot.BrightRange Then
    Gamma = (Spot.DarkRange - VectorDistance(Faces(FaceIndex).Center, Spot.Origin)) / (Spot.DarkRange - Spot.BrightRange)
   If Gamma < 0 Then Gamma = 0 Else If Gamma > 1 Then Gamma = 1
   Else
    Exit Sub
   End If
  End If

  '################################################

  'Note this: Always: Alpha + Beta, in other word: Diffusion + Sepcular
  ColorSum = ColorScale(Spot.Color, ((Alpha + Beta) * Gamma) * Epsilon)
  ColorSum = ColorAdd(ColorSum, Amb) 'Add the ambience

  'Set limitations:
  If ColorSum.R > 255 Then ColorSum.R = 255
  If ColorSum.G > 255 Then ColorSum.G = 255
  If ColorSum.B > 255 Then ColorSum.B = 255

  'Here, Lighting is Ok, the follow operation
  'is simply 'Scale' the color to apply the Fog,
  'with the CNear and CFar properties of the camera.

  '################################################
  'apply camera distance decay (Fog)
  Delta = 1
  If OptFog = True Then
   If CFar <> CNear Then
    Delta = (CFar - VectorDistance(Faces(FaceIndex).Center, Cam)) / (CFar - CNear)
    If Delta < 0 Then Delta = 0 Else If Delta > 1 Then Delta = 1
    Else
    Exit Sub
   End If
  End If

  '################################################
  Faces(FaceIndex).TCol = ColorScale(ColorSum, Delta)

 End If

End Sub
Private Sub Form_Activate()

 'Main loop
 '=========

 Do
  Cls

  Calc
  Draw

  XAng = XAng + Deg
  YAng = YAng + Deg

  DoEvents
 Loop

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

 If KeyCode = vbKeyEscape Then Unload Me: End
 If KeyCode = vbKeyNumpad0 Then OptSpecRef = Not OptSpecRef
 If KeyCode = vbKeyNumpad1 Then OptAttenuation = Not OptAttenuation
 If KeyCode = vbKeyNumpad2 Then OptFog = Not OptFog

End Sub
Private Sub Form_Load()

 Move 0, 0, (640 * 15), (480 * 15)
 ScaleMode = vbPixels

 LoadScene

End Sub
