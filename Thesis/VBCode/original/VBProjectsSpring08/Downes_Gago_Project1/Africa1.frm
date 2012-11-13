VERSION 5.00
Begin VB.Form Africa1 
   BackColor       =   &H80000008&
   Caption         =   "Africa"
   ClientHeight    =   10890
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12285
   LinkTopic       =   "Form1"
   ScaleHeight     =   45.375
   ScaleMode       =   0  'User
   ScaleTop        =   10
   ScaleWidth      =   102.375
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Africa"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      TabIndex        =   7
      Top             =   7440
      Width           =   1935
   End
   Begin VB.CommandButton cmdSize 
      Caption         =   "List countries from the most sizeable ones"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   6
      Top             =   2400
      Width           =   2415
   End
   Begin VB.CommandButton cmdMore 
      Caption         =   "List the First Country With Higher Population than Average"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      TabIndex        =   5
      Top             =   6000
      Width           =   2415
   End
   Begin VB.CommandButton cmdAverage 
      Caption         =   "What's the average population in Africa?"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   4
      Top             =   4680
      Width           =   2415
   End
   Begin VB.CommandButton cmdPopulation 
      Caption         =   "List countries according to the highest population"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   3
      Top             =   3480
      Width           =   2415
   End
   Begin VB.CommandButton cmdABC 
      Caption         =   "List countries in ABC order"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   2415
   End
   Begin VB.CommandButton cmdList 
      Caption         =   "Countries of Africa"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10815
      Left            =   3600
      LinkItem        =   "VScroll"
      ScaleHeight     =   10755
      ScaleWidth      =   8475
      TabIndex        =   0
      Top             =   0
      Width           =   8535
   End
End
Attribute VB_Name = "Africa1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: The Globe Trotter Experience
'Form name: Africa1.frm
'Author: Marta Gago & Brian Downes
'Date Written: Thursday March 27th, 2008
'Objective of form:  This form allows the user to open a data file of information on Africa into an Array
'The user can then click on the various buttons to bubble sort the list into alphabetical order
'Listing the Countries in order according to population, the calculation of the total countries average
'Print the first country that is higher than the average
'The user can use a Back button to return to the previous form (main form)

Option Explicit
Dim country(1 To 400000) As String
Dim area(1 To 400000) As Single
Dim population(1 To 400000) As Single
Dim pos As Integer
Dim Pass As Integer
Dim TempCountry As String
Dim TempArea As Single
Dim TempPopulation As Single
Dim avg As Single                   'Dim the all the variables for the entire form

'Sorting Countries alphabetically
Private Sub cmdABC_Click()

Pass = ctr - 1

For Pass = 1 To ctr - 1
    For pos = 1 To ctr - Pass
        If country(pos) > country(pos + 1) Then
        
            TempCountry = country(pos)          'TempCountry acts as another space to make arrangement possible
            country(pos) = country(pos + 1)
            country(pos + 1) = TempCountry
            TempArea = area(pos)
            area(pos) = area(pos + 1)
            area(pos + 1) = TempArea
            TempPopulation = population(pos)
            population(pos) = population(pos + 1)
            population(pos + 1) = TempPopulation
                        
        End If
    Next pos
Next Pass                       'Bubble Sort for alphabetical order

picResults.Cls                  'Clear Print Box
picResults.Print "No", "Country", Tab(50), "Area", Tab(75), "Population"
picResults.Print "****************************************************************************************************************"
                                'Print labels
For j = 1 To ctr
    picResults.Print j, country(j), Tab(50), area(j), Tab(75), FormatNumber(population(j), 0)
Next j
                                'Print data from Array to print box and tab inbetween so data lined up correctly
End Sub

Private Sub cmdAverage_Click()
Dim avg As Single
Dim sum As Single           'Dim variables for this subprogram
sum = 0                     'Start Sum as 0

picResults.Cls              'Clear Print Box

For j = 1 To ctr                     'j continues until the counter
    sum = sum + population(j)        'Calculating the Sum by adding itself to the population number
Next j
avg = sum / ctr                     'Calculating the average by dividing the sum by the counter

MsgBox "The average of the population in the countries in Africa is" & FormatNumber((avg), 0)


End Sub

Private Sub cmdBack_Click()
Africa1.Hide                'Make form disapear
Africa.Show                 'Make Africa Form appear
End Sub
'Loads Data into an array
Private Sub cmdList_Click()

Open App.Path & "\AfricaCountries.txt" For Input As #1  'Opening a data file into an array
ctr = 0
    'Loading the data into an array with a Do while loop
    Do While Not EOF(1)
        ctr = ctr + 1
        Input #1, country(ctr), area(ctr), population(ctr)
    Loop

picResults.Cls

picResults.Print "No", "Country", Tab(50); "Area"; Tab(70); "Population"
picResults.Print "_____________________________________________________________________________________________________________________________________________"
    'Printing results with a For/Next loop
For j = 1 To ctr
    picResults.Print j, country(j), Tab(50); area(j); Tab(70); FormatNumber(population(j), 0)
Next
Close       'Close the array

End Sub
'Find the first country that is above population average
Private Sub cmdMore_Click()
Found = False
j = 0
picResults.Print "Countries with population higher than the average are:"
picResults.Print "________________________________________________________________________________________________________________"
picResults.Print ""
    'A match and stop search to find the first country that is above average in population
Do While Not Found And j < ctr
    j = j + 1
        If population(j) > avg Then
            Found = True
        End If
Loop
picResults.Cls
picResults.Print country(j), FormatNumber(population(j), 0)     'Number without decimal places, and adds commas for larger numbers
 
End Sub
'Sorting Countries by population
Private Sub cmdPopulation_Click()
Pass = ctr - 1
    'Bubble Sort to arrange countries in order according to population
For Pass = 1 To ctr - 1
    For pos = 1 To ctr - Pass
        If population(pos) < population(pos + 1) Then
        
            TempCountry = country(pos)
            country(pos) = country(pos + 1)
            country(pos + 1) = TempCountry
            TempArea = area(pos)
            area(pos) = area(pos + 1)
            area(pos + 1) = TempArea
            TempPopulation = population(pos)
            population(pos) = population(pos + 1)
            population(pos + 1) = TempPopulation
                        
        End If
    Next pos
Next Pass

picResults.Cls
picResults.Print "No", "Country", Tab(50); "Area"; Tab(70); "Population"
picResults.Print "************************************************************************************************************************"
    'Using a For/Next loop to print each line of data in the print box
For j = 1 To ctr
    picResults.Print j, country(j), Tab(50); area(j); Tab(70); FormatNumber(population(j), 0)
Next j
End Sub
'Sorting Countries by Size
Private Sub cmdSize_Click()

Pass = ctr - 1
'Bubblesort to arrange countries by size
For Pass = 1 To ctr - 1
    For pos = 1 To ctr - Pass
        If area(pos) < area(pos + 1) Then
        
            TempCountry = country(pos)          'TempCountry acts as another space to make arrangement possible
            country(pos) = country(pos + 1)
            country(pos + 1) = TempCountry
            TempArea = area(pos)
            area(pos) = area(pos + 1)
            area(pos + 1) = TempArea            'TempArea acts the same as TempCountry
            TempPopulation = population(pos)
            population(pos) = population(pos + 1)
            population(pos + 1) = TempPopulation
                        
        End If
    Next pos
Next Pass

picResults.Cls
picResults.Print "No", "Country", Tab(50); "Area", Tab(70); "Population"
picResults.Print "*********************************************************************************************************"
'Using the Exhaustive Function, the data is printed
For j = 1 To ctr
    picResults.Print j, country(j), Tab(50); area(j); Tab(70); FormatNumber(population(j), 0)
Next j
End Sub

