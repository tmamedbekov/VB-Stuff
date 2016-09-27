'Arrays
'ReDim - Preserves the data in an existing array when you change the size of the last dimension.

option explicit
On Error Resume Next

'Declare variables
Const SITE_TITLE = "www.mamedbekov.com"

'Declare an array named A with 10 elements
Dim arrFriends()
ReDim arrFriends(4)
    arrFriends(0) = "John"
    arrFriends(1) = "Peter"
    arrFriends(2) = "Jimmy"
    arrFriends(3) = "Mike"
    arrFriends(4) = "Kenny"
ReDim Preserve arrFriends(9)

arrFriends(8) = "Tom"

MsgBox arrFriends(4), 0, SITE_TITLE
MsgBox arrFriends(8), 0, SITE_TITLE