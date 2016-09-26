'==============
'This is a very simple VB Script to display data
'==============

option explicit 
On Error Resume Next

' Declare Variables
Dim num1, num2
Dim sum
Dim name, city

' Make sure to have constants in all Caps
Const SITE_TITLE = "www.mamedbekov.com"
Const FIRST_NAME = "Togrul"

num1 = 10
num2 = 20

sum = num1 + num2

MsgBox sum

MsgBox "The sum of " & num1 & " and " & num2 & " is " & sum, 0, SITE_TITLE

MsgBox "** The sum of " & num1 & " and " & num2 &_
       " is " & sum, 0, SITE_TITLE

name = "Tony"
city = "Houston"

MsgBox fname & " lives in " & city & "!!! ", 0, SITE_TITLE

