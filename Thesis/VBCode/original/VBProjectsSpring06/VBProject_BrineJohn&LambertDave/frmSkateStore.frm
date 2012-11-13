VERSION 5.00
Begin VB.Form frmSkateStore 
   Caption         =   "Skate Store!"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9030
   FillStyle       =   0  'Solid
   FontTransparent =   0   'False
   LinkTopic       =   "Form2"
   Picture         =   "frmSkateStore.frx":0000
   ScaleHeight     =   7920
   ScaleWidth      =   9030
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
      Left            =   360
      TabIndex        =   11
      Top             =   6960
      Width           =   2295
   End
   Begin VB.CommandButton cmdSortStyle 
      Caption         =   "Sort According to Style"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   10
      Top             =   3960
      Width           =   2055
   End
   Begin VB.CommandButton cmdSortFrequency 
      Caption         =   "Sort According to Frequency "
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   9
      Top             =   2880
      Width           =   2055
   End
   Begin VB.CommandButton cmdSortPrice 
      Caption         =   "Sort According to Price"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   8
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search Inventory"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2520
      TabIndex        =   7
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton cmdCheckout 
      Caption         =   "Proceed to Checkout"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3960
      TabIndex        =   6
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Calculator"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   4
      Top             =   5040
      Width           =   1575
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate Total"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   3
      Top             =   5040
      Width           =   1575
   End
   Begin VB.CommandButton cmdViewSkates 
      Caption         =   "Click to View Skates in Stock"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton cmdSkateWizard 
      BackColor       =   &H00FF0000&
      Caption         =   "Begin Skate Wizard"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4800
      MaskColor       =   &H00FF0000&
      Picture         =   "frmSkateStore.frx":CDAB
      TabIndex        =   1
      Top             =   720
      Width           =   2055
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00808080&
      Height          =   2895
      Left            =   2640
      ScaleHeight     =   2835
      ScaleWidth      =   6195
      TabIndex        =   0
      Top             =   1800
      Width           =   6255
   End
   Begin VB.Label lblSkateWelcome 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSkateStore.frx":105F6
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3855
      Left            =   1080
      TabIndex        =   5
      Top             =   2400
      Width           =   6735
   End
End
Attribute VB_Name = "frmSkateStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dave and JB's ski and skate store
'frmFront
'Dave Lambert and John Brine
'Thursday March 23
'This form is used to search, by name or given types, the skis that are in the inventory at that moment.  It then can calculate all of the totals and also show the inventory.

    Dim SizeSkate As Integer
    Dim FoundSkate As Integer
    
Private Sub cmdBack_Click()
    frmSkateStore.Hide ' Shows desired form and hides undesired forms
    frmFront.Show
End Sub

Private Sub cmdCalculate_Click()
'establishes variables needed for sub routine
    picResults.Print
    Dim subtotal As Double
    Dim tax As Double
    Dim total As Double
    
    'clears the pic box
    picResults.Cls
    
    'calculates the subtotal, tax, s&h and total for a found skate
        If FoundSkate > 0 Then
            'prints header
            picResults.Print "Item", "Frequency", "Type", "Price"
            picResults.Print "*******************************************************************************************************************************************************"
            picResults.Print
            picResults.Print
            'prints the found skate
            picResults.Print SkateName(FoundSkate), SkateFrequency(FoundSkate), SkateType(FoundSkate), FormatCurrency(SkatePrice(FoundSkate))
            picResults.Print
            picResults.Print
            subtotal = SkatePrice(FoundSkate) 'establishes the found skate as the subtotal
            picResults.Print "Subtotal", , , FormatCurrency(subtotal) 'displats the subtotal and the amount
            tax = subtotal * 0.07 'figures tax
            picResults.Print "Shipping and Handling", , "$25.00" 'shows the amount and name of shipping and handling
            picResults.Print "Tax", , , FormatCurrency(tax) 'displays tax as money and shows title
            picResults.Print "*******************************************************************************************************************************************************"
            total = 25 + subtotal * 1.07 'figures total
            picResults.Print "Total", FormatCurrency(total) 'displays actual total and title
        End If
        'if statement to display error
        If FoundSkate = 0 Then
            MsgBox "You have not selected anything!", , ":("
        End If
        'changes the value of the public variables for use when printing reciept
        SubTotalSkate = SubTotalSkate + subtotal
        SkateTotal = total + SkateTotal
        cmdCheckout.Visible = True
        

End Sub

Private Sub cmdCheckout_Click()
    frmCheckOut.Visible = True ' Shows desired comands and hides comand forms
    frmSkateStore.Visible = False
End Sub

Private Sub cmdClear_Click()
    'clears the picture box
    picResults.Cls
End Sub

Private Sub cmdExit_Click()
    'ends program and displays error message thanking user for stopping bye
    MsgBox "Thanks for stopping come back for all of your Skate needs!", , "Come back soon!"
End
End Sub

Private Sub cmdNavitgate_Click()
    frmNavigate.Visible = True ' Shows desired form and hides undesired forms
    frmSkateStore.Visible = False
End Sub

Private Sub cmdSearch_Click()
    'establishes all variables to be used in sub routine
    Dim search As Single
    Dim Counter As Integer
    Dim found As Boolean
    Dim pos2 As Integer
    Dim I As Integer
    Dim SearchSkate As String
    
    'takes welcome message away from form
    lblSkateWelcome.Visible = False
    
    'establishes all of the dimed variables as desired values
    found = False
    pos2 = 0
    
    'shows message box to get the searched name
    SearchSkate = InputBox("Input Desired Skate.", "Search Skate")
    
    'begins the actual search for the input string
    picResults.Visible = True 'shows pic box
    picResults.Cls 'clears pic box
    'prints header
        picResults.Print "Item", , "Price", "Type", "Frequency of Use"
        picResults.Print "*******************************************************************************************************************************************************"
        picResults.Print
        
        'acutal search takes place
        For I = 1 To SizeSkate
            search = InStr(SkateName(I), SearchSkate)
            pos2 = pos2 + 1
            If search > 0 Then
                  found = True 'changes found variable to true when found
            'prints the found ski according to previous search
            picResults.Print
            picResults.Print SkateName(I), , FormatCurrency(SkatePrice(I)), SkateType(I), SkateFrequency(I)
            FoundSkate = I 'changes the found skate to the variable found skate to use globally
            End If
        Next I
        
        'for when no match is found displays error message in picture box
        If found = False Then
            picResults.Print "Sorry no matches in inventory right now."
            picResults.Print "Click the Inventory Button to see what we have in stock!"
            MsgBox "Sorry No matches in inventory right now.", , ":("
        End If
        
        'displays calculate and clear buttons for when the user found a ski
        If found = True Then
            cmdCalculate.Visible = True
            cmdClear.Visible = True
        End If
End Sub

Private Sub cmdSkateWizard_Click()
    'establishes all of the variables
    Dim Wizard As Single
    Dim Match As String
    Dim SkierType, SkierPrice, SkierFrequency As Integer
    Dim found As Boolean
    Dim pos1 As Integer
    
   
    ' gains the users input to use for finding approptiate product
    SkierType = Val(InputBox("What kind of Skater are you? (1 = Recreational, 2 = Figure, 3 = Hockey, 4 = Speed)", "Skater Type"))
    SkierPrice = Val(InputBox("How much do you want to spend on a Skate? (1 = under $100, 2 = $100 - 150, 3 = $150 - 250, 4 = Over $250)", "Cost?"))
    SkierFrequency = Val(InputBox("How often do you skate? (1 = Yearly, 2 = Monthly, 3 = Weekly, 4 = Daily)", "Frequency"))
    
    lblSkateWelcome.Visible = False ' Shows desired comands and picture boxes and hides ones
    cmdCalculate.Visible = True
    cmdClear.Visible = True
    picResults.Visible = True
    
    picResults.Cls 'clears picture box
    picResults.Print "Item", "Price", "Type", "Frequency of Use" 'prints header on picture box
    picResults.Print "*******************************************************************************************************************************************************"
    picResults.Print
    
    'establishes variables as desrired values
    found = False
    pos1 = 0
    'begins the actual search comparing the text file with the variables from the user
    Do While found = False And pos1 < SizeSkate
        pos1 = pos1 + 1
        If SkatePriceNum(pos1) = SkierPrice And SkateTypeNum(pos1) = SkierPrice And SkateFrequencyNum(pos1) = SkierFrequency _
            Or (SkatePriceNum(pos1) = SkierPrice) And (SkateTypeNum(pos1) = SkierPrice) _
            Or (SkateTypeNum(pos1) = SkierPrice) And (SkateFrequencyNum(pos1) = SkierFrequency) _
            Or (SkateFrequencyNum(pos1) = SkierFrequency) And (SkatePriceNum(pos1) = SkierPrice) Then
            found = True
        End If
    Loop
    'displays the result if product is found
    If found = True Then
        picResults.Print SkateName(pos1), SkatePrice(pos1), SkateType(pos1), SkateFrequency(pos1)
    Else
        'displays error when no match is found
        picResults.Print "Sorry No Matches!"
        'begins second search for cleses possible match for variables
        Match = InputBox("Would You like to see the closest match that we have in stock? (Yes or No)", "No Matches", , 25, 25)
        Dim pos As Integer
        found = False
        pos = 0
        If InStr(Match, "es") > 0 Then
            'creates new criteria for search by adding all variables together
            Wizard = SkierType + SkierPrice + SkierFrequency
            'begins the search for close product
            Do While found = False And pos < SizeSkate
                pos = pos + 1
                If Wizard = SkateOverall(pos) Then
                'displays the found product
                    picResults.Cls
                    picResults.Print "Item", , "Price", "Type", "Frequency of Use"
                    picResults.Print "*******************************************************************************************************************************************************"
                    picResults.Print
                    picResults.Print
                    picResults.Print SkateName(pos), , SkatePrice(pos), SkateType(pos), SkateFrequency(pos)
                    found = True
                End If
            Loop
        End If
    End If
    'creates a number value for the ski that has been found - used to calculate the total later on.
    FoundSkate = pos1
End Sub

Private Sub cmdSkiStore_Click()
    frmSkateStore.Hide ' Shows desired form and hides undesired forms
    frmSkiStore.Show
End Sub

Private Sub cmdSortFrequency_Click()
   'this sub routine sorts the product according to the frequency
    Dim pass As Integer
    Dim pos As Integer
    Dim TempFrequency, tempStyle, tempName As String
    Dim tempPrice As Single
    Dim I As Integer
    picResults.Cls
        For pass = 1 To (SizeSkate - 1) ' does sort for only the size of the array
            For pos = 1 To (SizeSkate - pass)
                If SkateFrequency(pos) > SkateFrequency(pos + 1) Then
                'keeps all items in correct order for all of the arrays
                    tempName = SkateName(pos)
                    SkateName(pos) = SkateName(pos + 1)
                    SkateName(pos + 1) = tempName
                    
                    TempFrequency = SkateFrequency(pos)
                    SkateFrequency(pos) = SkateFrequency(pos + 1)
                    SkateFrequency(pos + 1) = TempFrequency
                    
                    tempStyle = SkateType(pos)
                    SkateType(pos) = SkateType(pos + 1)
                    SkateType(pos + 1) = tempStyle
                    
                    tempPrice = SkatePrice(pos)
                    SkatePrice(pos) = SkatePrice(pos + 1)
                    SkatePrice(pos + 1) = tempPrice
                End If
        Next pos
    Next pass
    
 
    picResults.Print "Item", , "Price", "Type", "Frequency" ' prints the header on the picturebox
    picResults.Print "*******************************************************************************************************************************************************"
    picResults.Print 'prints the array in the new sorted way
        For I = 1 To SizeSkate
            picResults.Print SkateName(I), , SkatePrice(I), SkateType(I), SkateFrequency(I)
        
        Next I
End Sub

Private Sub cmdSortPrice_Click()
     'establishes desired variables necessary for sub routine
    Dim pass As Integer
    Dim pos As Integer
    Dim TempFrequency, tempStyle, tempName As String
    Dim tempPrice As Single
    Dim I As Integer
    'clears the picture box
    picResults.Cls
    
    'sorts the array according to skate price
        For pass = 1 To (SizeSkate - 1)
            For pos = 1 To (SizeSkate - pass)
                If SkatePrice(pos) > SkatePrice(pos + 1) Then
                    
                'swaps all of the information so it stays with in the correct _
                order for the new sorted style
                    tempName = SkateName(pos)
                    SkateName(pos) = SkateName(pos + 1)
                    SkateName(pos + 1) = tempName
                    
                    TempFrequency = SkateFrequency(pos)
                    SkateFrequency(pos) = SkateFrequency(pos + 1)
                    SkateFrequency(pos + 1) = TempFrequency
                    
                    tempStyle = SkateType(pos)
                    SkateType(pos) = SkateType(pos + 1)
                    SkateType(pos + 1) = tempStyle
                    
                    tempPrice = SkatePrice(pos)
                    SkatePrice(pos) = SkatePrice(pos + 1)
                    SkatePrice(pos + 1) = tempPrice
                End If
        Next pos
    Next pass
    
 'prints header in picture box
    picResults.Print "Item", , "Price", "Type", "Frequency"
    picResults.Print "*******************************************************************************************************************************************************"
    picResults.Print
    'prints the new sorted array in pic box
        For I = 1 To SizeSkate
            picResults.Print SkateName(I), , SkatePrice(I), SkateType(I), SkateFrequency(I)
        Next I
                
End Sub

Private Sub cmdSortStyle_Click()
    'establishes desired variables necessary for sub routine
    Dim pass As Integer
    Dim pos As Integer
    Dim TempFrequency, tempStyle, tempName As String
    Dim tempPrice As Single
    Dim I As Integer
    
    picResults.Cls 'clears the picturebox
    
    'sorts the array according to skate type
        For pass = 1 To (SizeSkate - 1)
            For pos = 1 To (SizeSkate - pass)
                If SkateType(pos) > SkateType(pos + 1) Then
                    
                'swaps all of the information in the array to keep all of the _
                factors in the correct array order
                    tempName = SkateName(pos)
                    SkateName(pos) = SkateName(pos + 1)
                    SkateName(pos + 1) = tempName
                    
                    TempFrequency = SkateFrequency(pos)
                    SkateFrequency(pos) = SkateFrequency(pos + 1)
                    SkateFrequency(pos + 1) = TempFrequency
                    
                    tempStyle = SkateType(pos)
                    SkateType(pos) = SkateType(pos + 1)
                    SkateType(pos + 1) = tempStyle
                    
                    tempPrice = SkatePrice(pos)
                    SkatePrice(pos) = SkatePrice(pos + 1)
                    SkatePrice(pos + 1) = tempPrice
                End If
        Next pos
    Next pass
    
 'prints the header in the picture box
    picResults.Print "Item", , "Price", "Type", "Frequency of Use"
    picResults.Print "*******************************************************************************************************************************************************"
    picResults.Print
    
    'prints all of the information from the array in the new sorted order
        For I = 1 To SizeSkate
            picResults.Print SkateName(I), , SkatePrice(I), SkateType(I), SkateFrequency(I)
        
        Next I
End Sub

Private Sub cmdViewSkates_Click()
'shows the desired buttons that display the are used to view inventory
    lblSkateWelcome.Visible = False
    cmdSortPrice.Visible = True
    cmdSortStyle.Visible = True
    cmdSortFrequency.Visible = True
    cmdClear.Visible = True
    picResults.Visible = True
    picResults.Cls
    'pulls all of the ski items from the established _
    array and then displays them within the picture box
    Dim I As Integer
    picResults.Print "Item", , "Price", "Type", "Frequency"
    picResults.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    picResults.Print
        For I = 1 To SizeSkate
            picResults.Print SkateName(I), , SkatePrice(I), SkateType(I), SkateFrequency(I)
        Next I
End Sub


Private Sub Form_Load()
'when form is loaded then all buttons and picture boxes are not visible
    cmdSortStyle.Visible = False
    cmdCheckout.Visible = False
    cmdSortFrequency.Visible = False
    cmdSortPrice.Visible = False
    picResults.Visible = False
    cmdClear.Visible = False
    cmdCalculate.Visible = False
    'loads the ski array for use in all of the form
    Dim pos As Integer
    Open App.Path & "\skate.txt" For Input As #2
    Do Until EOF(2)
        pos = pos + 1
        SizeSkate = SizeSkate + 1
        Input #2, SkateName(pos), SkatePrice(pos), SkatePriceNum(pos), SkateType(pos), SkateTypeNum(pos), SkateFrequency(pos), SkateFrequencyNum(pos), SkateOverall(pos)
    Loop
    Close #2
            
End Sub

Private Sub picResults_Click()
    'displays error message when you click on picture box
    MsgBox "Click the Calculate Button to Figure item total.", , "Error!"
End Sub
