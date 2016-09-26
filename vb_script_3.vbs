' Two Dimentional Arrays

option explicit
'On Error Resume Next

'Declare variables
'Declare an array named A with 10 elements (1, 4 ?)
'Column, Row

Dim arrFavFruit(1,4)

Const SITE_TITLE = "www.mamedbekov.com"

arrFavFruit(0,0) = "John"
arrFavFruit(0,1) = "Peter"
arrFavFruit(0,2) = "Mike"
arrFavFruit(0,3) = "Sam"
arrFavFruit(0,4) = "Jimmy"

arrFavFruit(1,0) = "Apples"
arrFavFruit(1,1) = "Bananas"
arrFavFruit(1,2) = "Grapes"
arrFavFruit(1,3) = "Oranges"
arrFavFruit(1,4) = "Mangoes"

MsgBox arrFavFruit(0,0) & " likes " & arrFavFruit(1,0) & " ! ! ! "
MsgBox arrFavFruit(0,0) & " likes " & arrFavFruit(1,0) & " ! ! ! ", 0, SITE_TITLE

MsgBox arrFavFruit(0,3) & " likes " & arrFavFruit(1,3) & " ! ! ! ", 0, SITE_TITLE