Attribute VB_Name = "NataliesProjectModule"
Public Savings As Single    'variable for user's savings value
Public Paycheck As Single   'variable for user's paycheck value
Public PayPeriods As Integer    'variable for user's number of pay periods remaining before big bill is due
Public BigBill As Single    'variable for user's expected big bill value
Public Sum As Single        'variable for sum of user's resources (savings and earnings)
Public Surplus As Single    'variable for the value that the user has 'extra'
Public Bill As Single       'variable for user's monthly bill values
Public BillTotal As Single  'variable for sum of monthly bills
Public Item(1 To 100) As String 'Item array for grocery store
Public Price(1 To 100) As Single    'price array for grocery store
Public Category(1 To 100) As String 'category array for grocery store
Public I As Integer, J As Integer   'counters
Public TempPrice As Single      'variable used as temporary storage of the price value when sorting the price array
Public TempItem As String       'variable used as temporary storage of the item value when sorting the item array
Public TempCategory As String   'variable used as temporary storage of the category value when sorting the category array
Public Pass As Integer, K As Integer, L As Integer  'counters


