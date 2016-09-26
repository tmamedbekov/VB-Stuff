'=========
' Arrays
'=========

option explicit
'On Error Resume Next

'Declare Variables

Dim arrNums (4)

Const SITE_TITLE = "www.mamedbekov.com"

'Assigning value to the elements in A
arrNums(0) = 100
arrNums(1) = 200
arrNums(2) = 300
arrNums(3) = 400
arrNums(4) = 500

Dim num1

num1 = arrNums(1)

MsgBox "Value of num1 is " & num1, 0, SITE_TITLE

MsgBox "Value of arrNums(1) is "& arrNums(1), 0, SITE_TITLE
MsgBox "Value of arrNums(3) is "& arrNums(3), 0, SITE_TITLE