VERSION 5.00
Begin VB.Form frmSkiStore 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Ski Store"
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11970
   BeginProperty Font 
      Name            =   "Papyrus"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   Picture         =   "frmSkiStore.frx":0000
   ScaleHeight     =   8040
   ScaleWidth      =   11970
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNavitgate 
      Caption         =   "Navigate our Site"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   11
      Top             =   7080
      Width           =   2295
   End
   Begin VB.CommandButton cmdCheckOUt 
      Caption         =   "Proceed To Checkout"
      Height          =   615
      Left            =   7440
      TabIndex        =   10
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CommandButton cmdSearchSki 
      Caption         =   "Search For Specific Ski"
      Height          =   615
      Left            =   2640
      TabIndex        =   9
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton cmdSortStyle 
      Caption         =   "Sort According to Style"
      Height          =   615
      Left            =   600
      TabIndex        =   8
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton cmdSortWeight 
      Caption         =   "Sort According to Weight"
      Height          =   615
      Left            =   600
      TabIndex        =   7
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton cmdSortPrice 
      Caption         =   "Sort According to Price"
      Height          =   615
      Left            =   600
      TabIndex        =   6
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Calculator"
      Height          =   615
      Left            =   9240
      TabIndex        =   4
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate Total"
      Height          =   615
      Left            =   7440
      TabIndex        =   3
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton cmdViewSkis 
      Caption         =   "View Ski Inventory"
      Height          =   615
      Left            =   2640
      TabIndex        =   2
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton cmdSkWizard 
      Caption         =   "Begin Ski Wizard"
      Height          =   615
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1440
      Width           =   1695
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FF8080&
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   4575
      Left            =   4680
      Negotiate       =   -1  'True
      ScaleHeight     =   4515
      ScaleWidth      =   6795
      TabIndex        =   0
      Top             =   480
      Width           =   6855
   End
   Begin VB.Label lblSkiWelcome 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSkiStore.frx":10E67
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   3735
      Left            =   4440
      MousePointer    =   1  'Arrow
      TabIndex        =   5
      Top             =   1080
      Width           =   7335
   End
End
Attribute VB_Name = "frmSkiStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dave and JB's ski and skate store
'frmFront
'Dave Lambert and John Brine
'Thursday March 23
'This form is used to search, by name or given types, the skates that are in the inventory at that moment.  It then can calculate all of the totals and also show the inventory.


    Dim FoundSki As Integer
    ' establishes variable as global for use with multipul commands

Private Sub cmdCalculate_Click()
    'establishes all required variables for command
    Dim subtotal As Double
    Dim tax As Double
    Dim total As Double
    'clears pic box
    picResults.Cls
    
        'if user found ski in other commands then proceeds
        'code figures totals
        If FoundSki > 0 Then
            'prints header
            picResults.Print "Item", , , "Price", "Type", "Weight Range"
            picResults.Print "*******************************************************************************************************************************************************"
            picResults.Print
            picResults.Print
            'prints the found ski
            picResults.Print SkiName(FoundSki), , FormatCurrency(SkiPrice(FoundSki)), SkiType(FoundSki), SkiWeight(FoundSki)
            picResults.Print
            picResults.Print
            'shows the subtotal
            subtotal = SkiPrice(FoundSki)
            'shows found subtotal
            picResults.Print "Subtotal", , , FormatCurrency(subtotal)
            'finds tax amount
            tax = subtotal * 0.07
            'shows shipping and handling title and amount
            picResults.Print "Shipping and Handling", , "$25.00"
            'shows tax title and amount
            picResults.Print "Tax", , , FormatCurrency(tax)
            picResults.Print "*******************************************************************************************************************************************************"
            'finds to total for selected ski
            total = 25 + subtotal * 1.07
            'displays total and title
            picResults.Print "Total", , , FormatCurrency(total)
        End If
        'displays error message for when no ski is found
        If FoundSki = 0 Then
            MsgBox "You have not selected anything!", , ":("
        End If
        'shows checkout button
        cmdCheckout.Visible = True
        'changes values in public variables for use when printing reciept
        skitotal = total
        SubTotalSki = subtotal + SubTotalSki
End Sub

Private Sub cmdCheckout_Click()
    'shows checkout form and moves from ski store form
    frmCheckOut.Visible = True
    frmSkiStore.Visible = False
End Sub

Private Sub cmdClear_Click()
    picResults.Cls 'clears the pic box
End Sub
Private Sub cmdNavitgate_Click()
    'moves from ski store form to navigate form
    frmNavigate.Visible = True
    frmSkiStore.Visible = False
End Sub

Private Sub cmdSearchSki_Click()
    'establishes all needed variables
    Dim search As Single
    Dim Counter As Integer
    Dim found As Boolean
    Dim pos2 As Integer
    Dim I As Integer
    Dim SearchSki As String
    
    'takes away the welcome label
    lblSkiWelcome.Visible = False
    
    'establishes variables values
    found = False
    pos2 = 0
    
    'message box to get the search sting for the search command
    SearchSki = InputBox("Input Desired Ski.", "Search Ski")
    
    'shows and clears the pic box
    picResults.Visible = True
    picResults.Cls
    
        'prints header in pic box
        picResults.Print "Item", , , "Price", "Type", "Weight Range"
        picResults.Print "*******************************************************************************************************************************************************"
        picResults.Print
        'begins the search throught the ski array
        For I = 1 To SizeSki
            search = InStr(SkiName(I), SearchSki)
            pos2 = pos2 + 1
            'changes global variable to true so can be used later
            If search > 0 Then
                  found = True
            'prints the found ski
            picResults.Print
            picResults.Print SkiName(I), , FormatCurrency(SkiPrice(I)), SkiType(I), SkiWeight(I)
            'changes the number of the found ski from array to the FoundSKi variable (public variable)
            FoundSki = I
            End If
        Next I
        'when found = false (no matches) shows an error message in both message box and pic box
        If found = False Then
            picResults.Print "Sorry no matches in inventory right now."
            picResults.Print "Click the Inventory Button to see what we have in stock!"
            MsgBox "Sorry No matches in inventory right now.", , ":("
        End If
        'shows calculate and clear buttons
        cmdCalculate.Visible = True
        cmdClear.Visible = True
        
End Sub

Private Sub cmdSkWizard_Click()

    'establishes needed variables for command
    Dim Wizard As Single
    Dim Match As String
    Dim SkierType, SkierPrice, SkierWeight As Integer
    Dim found As Boolean
    Dim pos1 As Integer
    
    'shows and clears pic box
    picResults.Visible = True
    picResults.Cls
    
    'prints header
    picResults.Print "Item", , , "Price", "Type", "Weight Range"
    picResults.Print "*******************************************************************************************************************************************************"
    picResults.Print
    
    'shows input box and sets the gained information to the three global variables
    SkierType = Val(InputBox("What type of Skier are you? (1 = Green, 2 = Blue, 3 = Black, 4 = Park)", "Skier Type"))
    SkierPrice = Val(InputBox("How much do you want to spend? (1 = under $200, 2 = $200 - $350, 3 = $350 - $500, 4 = Over $500)", "Skier Price"))
    SkierWeight = Val(InputBox("How much do you weigh? (1 = under 150 lbs, 2 = 150 - 175, 3 = 175 - 200, 4 = over 200)", "Skier Weight"))
    
    'shows calcualte and clear buttons while taking away the intro label
    lblSkiWelcome.Visible = False
    cmdCalculate.Visible = True
    cmdClear.Visible = True
    
    'establishes the variables to desired values
    found = False
    pos1 = 0
    
    'takes the input from the user and searches it with info from ski array
    Do While found = False And pos1 < SizeSki
        pos1 = pos1 + 1
        If SkiPriceNum(pos1) = SkierPrice And SkiTypeNum(pos1) = SkierPrice And SkiWeightNum(pos1) = SkierWeight _
            Or (SkiPriceNum(pos1) = SkierPrice) And (SkiTypeNum(pos1) = SkierPrice) _
            Or (SkiTypeNum(pos1) = SkierPrice) And (SkiWeightNum(pos1) = SkierWeight) _
            Or (SkiWeightNum(pos1) = SkierWeight) And (SkiPriceNum(pos1) = SkierPrice) Then
            found = True
        End If
    Loop
    
    'if ski is found then prints the result
    If found = True Then
        picResults.Print SkiName(pos1), SkiPrice(pos1), SkiType(pos1), SkiWeight(pos1)
    Else
        'if no ski is found then prints error message
        picResults.Print "Sorry No Matches!"
        
        'input box to see if user wants to find ski that is the closest match to their respective inputs
        Match = InputBox("Would You like to see the closest match that we have in stock? (Yes or No)", "No Matches", , 25, 25)
        
        'establishes variable for this function of sub routine
        Dim pos As Integer
        
        'runs search for yes statement
        If InStr(Match, "es") > 0 Then
            'equates the wiziard
            'wiziard is the total of all user inputs
            Wizard = SkierType + SkierPrice + SkierWeight
            'begins actual search for wizard number for each ski
            Do While found = False And pos < SizeSki
                pos = pos + 1
                'equates wizard with ski array number
                If Wizard = SkiOverall(pos) Then
                'prints the found ski
                    picResults.Print
                    picResults.Print SkiName(pos), SkiPrice(pos), SkiType(pos), SkiWeight(pos)
                    found = True
                
                'displays error message for when no skis are found
                Else
                    MsgBox "No matches found! Check inventory to see what we have in stock right now!", , ":("
                End If
            Loop
        End If
    End If
    'changes value of gobal variable to that of position of found ski
    FoundSki = pos1
End Sub

Private Sub cmdSortPrice_Click()
    'establishes variables needed for sub routine
    Dim pass As Integer
    Dim pos As Integer
    Dim tempWeight, tempStyle, tempName As String
    Dim tempPrice As Single
    Dim I As Integer
    
    'clears pic box
    picResults.Cls
    
        'begins sort according to price of items
        For pass = 1 To (SizeSki - 1)
            For pos = 1 To (SizeSki - pass)
                If SkiPrice(pos) > SkiPrice(pos + 1) Then
                    
                    'swaps all other items in array so that all relative information is printed correctly
                    tempName = SkiName(pos)
                    SkiName(pos) = SkiName(pos + 1)
                    SkiName(pos + 1) = tempName
                    
                    tempWeight = SkiWeight(pos)
                    SkiWeight(pos) = SkiWeight(pos + 1)
                    SkiWeight(pos + 1) = tempWeight
                    
                    tempStyle = SkiType(pos)
                    SkiType(pos) = SkiType(pos + 1)
                    SkiType(pos + 1) = tempStyle
                    
                    tempPrice = SkiPrice(pos)
                    SkiPrice(pos) = SkiPrice(pos + 1)
                    SkiPrice(pos + 1) = tempPrice
                End If
        Next pos
    Next pass
    
 'prints header
    picResults.Print "Item", , , "Price", "Type", "Weight"
    picResults.Print "*******************************************************************************************************************************************************"
    picResults.Print
    
    'prints all of the sorted info in the pic box
        For I = 1 To SizeSki
            picResults.Print SkiName(I), , SkiPrice(I), SkiType(I), SkiWeight(I)
        Next I
                
End Sub

Private Sub cmdSortStyle_Click()
    'establishes all needed variables for sub routine
    Dim pass As Integer
    Dim pos As Integer
    Dim tempWeight, tempStyle, tempName As String
    Dim tempPrice As Single
    Dim I As Integer
    
    'clears pic box
    picResults.Cls
    
        'sorts the info from the array according to type of product
        For pass = 1 To (SizeSki - 1)
            For pos = 1 To (SizeSki - pass)
                If SkiType(pos) > SkiType(pos + 1) Then 'find the actual position of items in array
                    
                'swaps all of the information so that it is in its respective order
                    tempName = SkiName(pos)
                    SkiName(pos) = SkiName(pos + 1)
                    SkiName(pos + 1) = tempName
                    
                    tempWeight = SkiWeight(pos)
                    SkiWeight(pos) = SkiWeight(pos + 1)
                    SkiWeight(pos + 1) = tempWeight
                    
                    tempStyle = SkiType(pos)
                    SkiType(pos) = SkiType(pos + 1)
                    SkiType(pos + 1) = tempStyle
                    
                    tempPrice = SkiPrice(pos)
                    SkiPrice(pos) = SkiPrice(pos + 1)
                    SkiPrice(pos + 1) = tempPrice
                End If
        Next pos
    Next pass
    
    'prints header
    picResults.Print "Item", , , "Price", "Type", "Weight"
    picResults.Print "*******************************************************************************************************************************************************"
    picResults.Print
    
        'prints the items from affay in new order
        For I = 1 To SizeSki
            picResults.Print SkiName(I), , SkiPrice(I), SkiType(I), SkiWeight(I)
        Next I
End Sub

Private Sub cmdSortWeight_Click()
    'establishes all needed variables for sub routine
    Dim pass As Integer
    Dim pos As Integer
    Dim tempWeight, tempStyle, tempName As String
    Dim tempPrice As Single
    Dim I As Integer
    
    'clears pic box
    picResults.Cls
    
        'sorts the items from the ski array according to recomended weight
        For pass = 1 To (SizeSki - 1)
            For pos = 1 To (SizeSki - pass)
                If SkiWeight(pos) > SkiWeight(pos + 1) Then 'changes the order of the items
                    
                'swaps all of the other items from the array into their correct respectively
                    tempName = SkiName(pos)
                    SkiName(pos) = SkiName(pos + 1)
                    SkiName(pos + 1) = tempName
                    
                    tempWeight = SkiWeight(pos)
                    SkiWeight(pos) = SkiWeight(pos + 1)
                    SkiWeight(pos + 1) = tempWeight
                    
                    tempStyle = SkiType(pos)
                    SkiType(pos) = SkiType(pos + 1)
                    SkiType(pos + 1) = tempStyle
                    
                    tempPrice = SkiPrice(pos)
                    SkiPrice(pos) = SkiPrice(pos + 1)
                    SkiPrice(pos + 1) = tempPrice
                End If
        Next pos
    Next pass
    
    'prints header
    picResults.Print "Item", , , "Price", "Type", "Weight"
    picResults.Print "*******************************************************************************************************************************************************"
    picResults.Print
        'prints the newly rearranged items in pic box
        For I = 1 To SizeSki
            picResults.Print SkiName(I), , SkiPrice(I), SkiType(I), SkiWeight(I)
        
        Next I
End Sub

Private Sub cmdStoreFront_Click(Index As Integer)
'shows the store front and moves away from the ski store
    frmSkiStore.Hide
    frmFront.Show
End Sub

Private Sub cmdViewSkis_Click()
    'takes away and displays all neccessary buttons, labels and pic boxes
    cmdClear.Visible = True
    lblSkiWelcome.Visible = False
    picResults.Visible = True
    cmdSortPrice.Visible = True
    cmdSortWeight.Visible = True
    cmdSortStyle.Visible = True
    
    'establishes needed variable for sub routine
    Dim I As Integer
    
    'clears pic box
    picResults.Cls
    
    'prints header
    picResults.Print "Item", , , "Price", "Type", "Weight"
    picResults.Print "*******************************************************************************************************************************************************"
    picResults.Print
        'prints all of selected items from array in order in pic box
        For I = 1 To SizeSki
            picResults.Print SkiName(I), , SkiPrice(I), SkiType(I), SkiWeight(I)
        Next I
End Sub

Private Sub Form_Load()
    'hides all of the unwanted buttons and pic boxes when form loads
    cmdCheckout.Visible = False
    picResults.Visible = False
    cmdClear.Visible = False
    cmdSortPrice.Visible = False
    cmdSortWeight.Visible = False
    cmdSortStyle.Visible = False
    cmdCalculate.Visible = False
    
    'establishes the needed variables
    Dim pos As Integer
    
        'loads the array for use with all other commands
        Open App.Path & "\ski.txt" For Input As #1
        Do Until EOF(1)
        pos = pos + 1
            SizeSki = SizeSki + 1
            Input #1, SkiName(pos), SkiPrice(pos), SkiPriceNum(pos), SkiType(pos), SkiTypeNum(pos), SkiWeight(pos), SkiWeightNum(pos), SkiOverall(pos)
        Loop
        Close #1
    'changes the value of the global variable to the found size of the array
    SizeSki = pos
End Sub
