Attribute VB_Name = "materialColors"
' MIT License
'
'Copyright (c) 2017 inDeev.eu Petr Katerinak
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'SOFTWARE.
'
'
' Version of materialColors: 1.0

Option Explicit

Public matColorNames As Variant

Public Function matColor(Name As String, Optional intensity As Integer = 500) As Long
  'Color Names:
  'Red, Pink, Purple, Deep Purple, Indigo, Blue, Light Blue, Cyan, Teal, Green,
  'Light Green, Lime, Yellow, Amber, Orange, Deep Orange, Brown, Grey, Blue Grey
  '
  'Color names are case insensitive and can be writen with or without spaces
  '
  'Color to be obtained from code by e.g. (r = range/selection, white font on dark blue background)
  ' r.Interior.color = matColor("blue", 800)
  '
  'Original color intensities are 50, 100, 200, 300, 400, 500, 600, 700, 800, 900,
  'but function can interpolate any integer value between 0 (white) and 1000 (black)
  '
  'White and Black can be obtained as intensity value 0 or 1000 of any color name but also as color name "white" or "black"
  ' r.Font.color = matColor("white")
  '

  Dim colorValue As Long
  Name = LCase(Replace(Name, " ", ""))            'Allows case insensitive names w/ or w/o spaces e.g. "lIght BlUe"

  If Name = "white" Or intensity <= 0 Then matColor = RGB(255, 255, 255): Exit Function
  If Name = "black" Or intensity >= 1000 Then matColor = RGB(0, 0, 0): Exit Function

  If intensity Mod 100 = 0 Or intensity = 50 Then 'intensity equals to 50, 100, 200, 300, ...
    Select Case Name
      Case "red"
        Select Case intensity
          Case 50: colorValue = RGB(255, 235, 238)
          Case 100: colorValue = RGB(255, 205, 210)
          Case 200: colorValue = RGB(239, 154, 154)
          Case 300: colorValue = RGB(229, 115, 115)
          Case 400: colorValue = RGB(239, 83, 80)
          Case 500: colorValue = RGB(244, 67, 54)
          Case 600: colorValue = RGB(229, 57, 53)
          Case 700: colorValue = RGB(211, 47, 47)
          Case 800: colorValue = RGB(198, 40, 40)
          Case 900: colorValue = RGB(183, 28, 28)
        End Select
      Case "pink"
        Select Case intensity
          Case 50: colorValue = RGB(252, 228, 236)
          Case 100: colorValue = RGB(248, 187, 208)
          Case 200: colorValue = RGB(244, 143, 177)
          Case 300: colorValue = RGB(240, 98, 146)
          Case 400: colorValue = RGB(236, 64, 122)
          Case 500: colorValue = RGB(233, 30, 99)
          Case 600: colorValue = RGB(216, 27, 96)
          Case 700: colorValue = RGB(194, 24, 91)
          Case 800: colorValue = RGB(173, 20, 87)
          Case 900: colorValue = RGB(136, 14, 79)
        End Select
      Case "purple"
        Select Case intensity
          Case 50: colorValue = RGB(243, 229, 245)
          Case 100: colorValue = RGB(225, 190, 231)
          Case 200: colorValue = RGB(206, 147, 216)
          Case 300: colorValue = RGB(186, 104, 200)
          Case 400: colorValue = RGB(171, 71, 188)
          Case 500: colorValue = RGB(156, 39, 176)
          Case 600: colorValue = RGB(142, 36, 170)
          Case 700: colorValue = RGB(123, 31, 162)
          Case 800: colorValue = RGB(106, 27, 154)
          Case 900: colorValue = RGB(74, 20, 140)
        End Select
      Case "deeppurple"
        Select Case intensity
          Case 50: colorValue = RGB(237, 231, 246)
          Case 100: colorValue = RGB(209, 196, 233)
          Case 200: colorValue = RGB(179, 157, 219)
          Case 300: colorValue = RGB(149, 117, 205)
          Case 400: colorValue = RGB(126, 87, 194)
          Case 500: colorValue = RGB(103, 58, 183)
          Case 600: colorValue = RGB(94, 53, 177)
          Case 700: colorValue = RGB(81, 45, 168)
          Case 800: colorValue = RGB(69, 39, 160)
          Case 900: colorValue = RGB(49, 27, 146)
        End Select
      Case "indigo"
        Select Case intensity
          Case 50: colorValue = RGB(232, 234, 246)
          Case 100: colorValue = RGB(197, 202, 233)
          Case 200: colorValue = RGB(159, 168, 218)
          Case 300: colorValue = RGB(121, 134, 203)
          Case 400: colorValue = RGB(92, 107, 192)
          Case 500: colorValue = RGB(63, 81, 181)
          Case 600: colorValue = RGB(57, 73, 171)
          Case 700: colorValue = RGB(48, 63, 159)
          Case 800: colorValue = RGB(40, 53, 147)
          Case 900: colorValue = RGB(26, 35, 126)
        End Select
      Case "blue"
        Select Case intensity
          Case 50: colorValue = RGB(227, 242, 253)
          Case 100: colorValue = RGB(187, 222, 251)
          Case 200: colorValue = RGB(144, 202, 249)
          Case 300: colorValue = RGB(100, 181, 246)
          Case 400: colorValue = RGB(66, 165, 245)
          Case 500: colorValue = RGB(33, 150, 243)
          Case 600: colorValue = RGB(30, 136, 229)
          Case 700: colorValue = RGB(25, 118, 210)
          Case 800: colorValue = RGB(21, 101, 192)
          Case 900: colorValue = RGB(13, 71, 161)
        End Select
      Case "lightblue"
        Select Case intensity
          Case 50: colorValue = RGB(225, 245, 254)
          Case 100: colorValue = RGB(179, 229, 252)
          Case 200: colorValue = RGB(129, 212, 250)
          Case 300: colorValue = RGB(79, 195, 247)
          Case 400: colorValue = RGB(41, 182, 246)
          Case 500: colorValue = RGB(3, 169, 244)
          Case 600: colorValue = RGB(3, 155, 229)
          Case 700: colorValue = RGB(2, 136, 209)
          Case 800: colorValue = RGB(2, 119, 189)
          Case 900: colorValue = RGB(1, 87, 155)
        End Select
      Case "cyan"
        Select Case intensity
          Case 50: colorValue = RGB(224, 247, 250)
          Case 100: colorValue = RGB(178, 235, 242)
          Case 200: colorValue = RGB(128, 222, 234)
          Case 300: colorValue = RGB(77, 208, 225)
          Case 400: colorValue = RGB(38, 198, 218)
          Case 500: colorValue = RGB(0, 188, 212)
          Case 600: colorValue = RGB(0, 172, 193)
          Case 700: colorValue = RGB(0, 151, 167)
          Case 800: colorValue = RGB(0, 131, 143)
          Case 900: colorValue = RGB(0, 96, 100)
        End Select
      Case "teal"
        Select Case intensity
          Case 50: colorValue = RGB(224, 242, 241)
          Case 100: colorValue = RGB(178, 223, 219)
          Case 200: colorValue = RGB(128, 203, 196)
          Case 300: colorValue = RGB(77, 182, 172)
          Case 400: colorValue = RGB(38, 166, 154)
          Case 500: colorValue = RGB(0, 150, 136)
          Case 600: colorValue = RGB(0, 137, 123)
          Case 700: colorValue = RGB(0, 121, 107)
          Case 800: colorValue = RGB(0, 105, 92)
          Case 900: colorValue = RGB(0, 77, 64)
        End Select
      Case "green"
        Select Case intensity
          Case 50: colorValue = RGB(232, 245, 233)
          Case 100: colorValue = RGB(200, 230, 201)
          Case 200: colorValue = RGB(165, 214, 167)
          Case 300: colorValue = RGB(129, 199, 132)
          Case 400: colorValue = RGB(102, 187, 106)
          Case 500: colorValue = RGB(76, 175, 80)
          Case 600: colorValue = RGB(67, 160, 71)
          Case 700: colorValue = RGB(56, 142, 60)
          Case 800: colorValue = RGB(46, 125, 50)
          Case 900: colorValue = RGB(27, 94, 32)
        End Select
      Case "lightgreen"
        Select Case intensity
          Case 50: colorValue = RGB(241, 248, 233)
          Case 100: colorValue = RGB(220, 237, 200)
          Case 200: colorValue = RGB(197, 225, 165)
          Case 300: colorValue = RGB(174, 213, 129)
          Case 400: colorValue = RGB(156, 204, 101)
          Case 500: colorValue = RGB(139, 195, 74)
          Case 600: colorValue = RGB(124, 179, 66)
          Case 700: colorValue = RGB(104, 159, 56)
          Case 800: colorValue = RGB(85, 139, 47)
          Case 900: colorValue = RGB(51, 105, 30)
        End Select
      Case "lime"
        Select Case intensity
          Case 50: colorValue = RGB(249, 251, 231)
          Case 100: colorValue = RGB(240, 244, 195)
          Case 200: colorValue = RGB(230, 238, 156)
          Case 300: colorValue = RGB(220, 231, 117)
          Case 400: colorValue = RGB(212, 225, 87)
          Case 500: colorValue = RGB(205, 220, 57)
          Case 600: colorValue = RGB(192, 202, 51)
          Case 700: colorValue = RGB(175, 180, 43)
          Case 800: colorValue = RGB(158, 157, 36)
          Case 900: colorValue = RGB(130, 119, 23)
        End Select
      Case "yellow"
        Select Case intensity
          Case 50: colorValue = RGB(255, 253, 231)
          Case 100: colorValue = RGB(255, 249, 196)
          Case 200: colorValue = RGB(255, 245, 157)
          Case 300: colorValue = RGB(255, 241, 118)
          Case 400: colorValue = RGB(255, 238, 88)
          Case 500: colorValue = RGB(255, 235, 59)
          Case 600: colorValue = RGB(253, 216, 53)
          Case 700: colorValue = RGB(251, 192, 45)
          Case 800: colorValue = RGB(249, 168, 37)
          Case 900: colorValue = RGB(245, 127, 23)
        End Select
      Case "amber"
        Select Case intensity
          Case 50: colorValue = RGB(255, 248, 225)
          Case 100: colorValue = RGB(255, 236, 179)
          Case 200: colorValue = RGB(255, 224, 130)
          Case 300: colorValue = RGB(255, 213, 79)
          Case 400: colorValue = RGB(255, 202, 40)
          Case 500: colorValue = RGB(255, 193, 7)
          Case 600: colorValue = RGB(255, 179, 0)
          Case 700: colorValue = RGB(255, 160, 0)
          Case 800: colorValue = RGB(255, 143, 0)
          Case 900: colorValue = RGB(255, 111, 0)
        End Select
      Case "orange"
        Select Case intensity
          Case 50: colorValue = RGB(255, 243, 224)
          Case 100: colorValue = RGB(255, 224, 178)
          Case 200: colorValue = RGB(255, 204, 128)
          Case 300: colorValue = RGB(255, 183, 77)
          Case 400: colorValue = RGB(255, 167, 38)
          Case 500: colorValue = RGB(255, 152, 0)
          Case 600: colorValue = RGB(251, 140, 0)
          Case 700: colorValue = RGB(245, 124, 0)
          Case 800: colorValue = RGB(239, 108, 0)
          Case 900: colorValue = RGB(230, 81, 0)
        End Select
      Case "deeporange"
        Select Case intensity
          Case 50: colorValue = RGB(251, 233, 231)
          Case 100: colorValue = RGB(255, 204, 188)
          Case 200: colorValue = RGB(255, 171, 145)
          Case 300: colorValue = RGB(255, 138, 101)
          Case 400: colorValue = RGB(255, 112, 67)
          Case 500: colorValue = RGB(255, 87, 34)
          Case 600: colorValue = RGB(244, 81, 30)
          Case 700: colorValue = RGB(230, 74, 25)
          Case 800: colorValue = RGB(216, 67, 21)
          Case 900: colorValue = RGB(191, 54, 12)
        End Select
      Case "brown"
        Select Case intensity
          Case 50: colorValue = RGB(239, 235, 233)
          Case 100: colorValue = RGB(215, 204, 200)
          Case 200: colorValue = RGB(188, 170, 164)
          Case 300: colorValue = RGB(161, 136, 127)
          Case 400: colorValue = RGB(141, 110, 99)
          Case 500: colorValue = RGB(121, 85, 72)
          Case 600: colorValue = RGB(109, 76, 65)
          Case 700: colorValue = RGB(93, 64, 55)
          Case 800: colorValue = RGB(78, 52, 46)
          Case 900: colorValue = RGB(62, 39, 35)
        End Select
      Case "gray", "grey"
        Select Case intensity
          Case 50: colorValue = RGB(250, 250, 250)
          Case 100: colorValue = RGB(245, 245, 245)
          Case 200: colorValue = RGB(238, 238, 238)
          Case 300: colorValue = RGB(224, 224, 224)
          Case 400: colorValue = RGB(189, 189, 189)
          Case 500: colorValue = RGB(158, 158, 158)
          Case 600: colorValue = RGB(117, 117, 117)
          Case 700: colorValue = RGB(97, 97, 97)
          Case 800: colorValue = RGB(66, 66, 66)
          Case 900: colorValue = RGB(33, 33, 33)
        End Select
      Case "bluegray", "bluegrey"
        Select Case intensity
          Case 50: colorValue = RGB(236, 239, 241)
          Case 100: colorValue = RGB(207, 216, 220)
          Case 200: colorValue = RGB(176, 190, 197)
          Case 300: colorValue = RGB(144, 164, 174)
          Case 400: colorValue = RGB(120, 144, 156)
          Case 500: colorValue = RGB(96, 125, 139)
          Case 600: colorValue = RGB(84, 110, 122)
          Case 700: colorValue = RGB(69, 90, 100)
          Case 800: colorValue = RGB(55, 71, 79)
          Case 900: colorValue = RGB(38, 50, 56)
        End Select
    End Select
  Else                                           ' intensity is not covered by cases 50, 100, 200, 300,...
    Dim lighter, darker As Long
    Dim factor As Integer

    If intensity < 50 Then                       ' 0-50
      lighter = matColor(Name, 0)
      darker = matColor(Name, 50)
      factor = intensity * 2
    ElseIf intensity < 100 Then                  ' 50-100
      lighter = matColor(Name, 50)
      darker = matColor(Name, 100)
      factor = (intensity - 50) * 2
    Else                                         ' 100-1000
      lighter = matColor(Name, Floor(intensity, 100))
      darker = matColor(Name, Ceiling(intensity, 100))
      factor = intensity - Floor(intensity, 100)
    End If

    colorValue = interpolateColor(lighter, darker, factor)

  End If
  matColor = colorValue
End Function

Public Function Ceiling(ByVal X As Double, Optional ByVal factor As Double = 1) As Double
  Ceiling = (Int(X / factor) - (X / factor - Int(X / factor) > 0)) * factor
End Function

Public Function Floor(ByVal X As Double, Optional ByVal factor As Double = 1) As Double
  Floor = Int(X / factor) * factor
End Function

Public Function interpolateColor(lighter As Variant, darker As Long, factor As Integer) As Long

  Dim r1, g1, b1 As Integer
  r1 = lighter Mod 256
  g1 = (lighter \ 256) Mod 256
  b1 = (lighter \ 256 \ 256) Mod 256

  Dim r2, g2, b2 As Integer
  r2 = darker Mod 256
  g2 = (darker \ 256) Mod 256
  b2 = (darker \ 256 \ 256) Mod 256

  interpolateColor = RGB(Int(r1 - ((r1 - r2) * factor / 100)), _
                         Int(g1 - ((g1 - g2) * factor / 100)), _
                         Int(b1 - ((b1 - b2) * factor / 100)))
End Function

Public Function isMatColor(color As Long, matColorName As String, Optional intensity As Integer = 0) As Boolean

  If intensity > 0 Then
    If color = matColor(matColorName, intensity) Then
      isMatColor = True: Exit Function
    Else
      isMatColor = False: Exit Function
    End If
  End If

  If color = matColor(matColorName, 50) Then isMatColor = True: Exit Function
  Dim o, p As Integer
  For o = 100 To 900 Step 100
    If color = matColor(matColorName, CInt(o)) Then isMatColor = True: Exit Function
  Next o

  For p = 1 To 999 Step 1
    If color = matColor(matColorName, CInt(p)) Then isMatColor = True: Exit Function
  Next p

  isMatColor = False
End Function

Public Function isMatColorN(color As Long) As String
  matColorNames = Array("Red", "Pink", "Purple", "DeepPurple", "Indigo", "Blue", "LightBlue", "Cyan", "Teal", "Green", "LightGreen", "Lime", "Yellow", "Amber", "Orange", "DeepOrange", "Brown", "Grey", "BlueGrey")
  Dim o, p As Integer
  Dim matColorName As Variant
  For Each matColorName In matColorNames
    If color = matColor(CStr(matColorName), 50) Then isMatColorN = CStr(matColorName): Exit Function

    For o = 100 To 900 Step 100
      If color = matColor(CStr(matColorName), CInt(o)) Then isMatColorN = CStr(matColorName): Exit Function
    Next o

    For p = 1 To 999 Step 1
      If color = matColor(CStr(matColorName), CInt(p)) Then isMatColorN = CStr(matColorName): Exit Function
    Next p
  Next
  isMatColorN = vbNullString
End Function


