Attribute VB_Name = "Module1"
'This is the array that the user names are input to and used
'for logging in.
Public usernameArray(1 To 100) As String
'This is the array that the passwords are input to and used
'for logging in.
Public passwordArray(1 To 100) As String
'This is the size of the price and name arrays used for
'various loops.
Public Size As Integer
'This is the array that the prices of all the items are input
'to.
Public priceArray(1 To 1000) As Single
'This is the array that the names of all the items are input
'to.
Public nameArray(1 To 1000) As String
'This is the total price of all of the items that have been
'rung up.
Public Sum As Single
'This is used for logging in so that the user only gets 3
'attempts to log in.
Public LoginCounter As Integer
'This is the timer set to give the user only 20 seconds to
'enter their user name and password. It is globally declared
'so that when you log out you only get 20 seconds to re-enter.
Public Timer1
'This is used to count the number of times an item has been
'put into the pic box.
Public ArrayCounter As Integer
