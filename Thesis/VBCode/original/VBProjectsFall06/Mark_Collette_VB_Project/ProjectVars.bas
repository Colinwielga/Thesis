Attribute VB_Name = "ProjectVars"
Public Items(1 To 100) As String, Prices(1 To 100) As Single
Public SItems As String, SPrices As Single, SStartInv As Single
Public Pos As Integer, Counter As Integer, Sum As Integer, Size As Integer, Pass As Integer
Public Found As Boolean
Public StartInv(1 To 100) As Single, QOH As Single
Public QSold(1 To 100) As Single, Sales As Single, EndInv(1 To 100) As Single
Public TempItems As String, TempPrices As Single, TempStartInv As Single

Public Subtotal As Single, Tax As Single, Total As Single, Discount As Single

Public picItemsSold
Public picPricesSold
Public picTotal

