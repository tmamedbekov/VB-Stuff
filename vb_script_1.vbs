option explicit
On Error Resume Next

Dim num1
Dim num2
Dim sum
Dim name, city

Const SITE_TITLE = "www.mamedbekov.com"

num1 = 10
num2 = 20

sum = num1 + num2

MsgBox sum

MsgBox "The sum of " & num1 & " and " & num2 & " is " & sum, 0, SITE_TITLE

MsgBox "** The sum of " & num1 & " and " & num2 & " is " & sum, 0, SITE_TITLE

name = "Tony"
city = "Houston"

MsgBox name & " lives in " & city & "!!! ", 0, SITE_TITLE

