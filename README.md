# Material-Colors-in-Excel-VBA
Allows to use material colors in MS Excel VBA projects

![Material Colors](https://github.com/inDeev/Material-Colors-in-Excel-VBA/blob/master/VBAMaterialColorPaletteBG.png)

Color Names:
Red, Pink, Purple, Deep Purple, Indigo, Blue, Light Blue, Cyan, Teal, Green,
Light Green, Lime, Yellow, Amber, Orange, Deep Orange, Brown, Grey, Blue Grey

more details in attached PDF file

Color names are case insensitive and can be writen with or without spaces

Color to be obtained from code by e.g. (r = range/selection, white font on dark blue background)
r.Interior.color = matColor("blue", 800)

Original color intensities are 50, 100, 200, 300, 400, 500, 600, 700, 800, 900,
but function can interpolate any integer value between 0 (white) and 1000 (black)

White and Black can be obtained as intensity value 0 or 1000 of any color name but also as color name "white" or "black"
r.Font.color = matColor("white")
