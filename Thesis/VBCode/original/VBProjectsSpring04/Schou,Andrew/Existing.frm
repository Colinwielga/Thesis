VERSION 5.00
Begin VB.Form Existing 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Corporate Bonds"
   ClientHeight    =   4485
   ClientLeft      =   2520
   ClientTop       =   1365
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   ScaleHeight     =   4485
   ScaleWidth      =   10500
   Begin VB.CommandButton cmdHeadings 
      BackColor       =   &H00E0E0E0&
      Caption         =   "What do the headings mean?"
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox txtsort 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9240
      TabIndex        =   6
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmdEXquit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Quit"
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton cmdEXcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancel"
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton cmdSort 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Sort"
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton cmdopen 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Get List of Corporate Bonds"
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.PictureBox picResultsEX 
      BackColor       =   &H00E0E0E0&
      Height          =   3375
      Left            =   1560
      ScaleHeight     =   3315
      ScaleWidth      =   8715
      TabIndex        =   0
      Top             =   120
      Width           =   8775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   $"Existing.frx":0000
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1800
      TabIndex        =   5
      Top             =   3600
      Width           =   7335
   End
End
Attribute VB_Name = "Existing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Investments(project1.vbp)
'Existing(Existing.frm)\
'Author- Andrew Schou
'3/14/04
'The purpose of this form is to display and sort the data that was loaded from a file.
'The first button is to recieve the data, second to soret the data.  The last button is
'used to tell the user, if he or she does not know, what each caption means.

Dim interestrate(1 To 15) As Single
Dim maturity(1 To 15) As String
Dim hiprice(1 To 15) As Single, lowprice(1 To 15) As Single, hiyield(1 To 15) As Single, lowyield(1 To 15) As Single
Dim total As Single
Dim CTR As Integer
Dim Company(1 To 15) As String
Private Sub cmdEXcancel_Click()
'go back to main page
Existing.Hide
Introduction.Show
End Sub

Private Sub cmdEXquit_Click()
End
End Sub
'show the discriptions of the headings
Private Sub cmdHeadings_Click()
Headings.Show
End Sub

Private Sub cmdopen_Click()
'open the file
picResultsEX.Cls
Open "N:\CS130\handin\Schou, Andrew\corporatebonds.txt" For Input As #1
CTR = O
'print the headings at the top of the page
picResultsEX.Print "Company"; Tab(25); "Interest Rate"; Tab(40); "Date of maturity"; Tab(60); "High Price"; Tab(75); "Low Price"; Tab(90); "YTM(high)"; Tab(105); "YTM(low)"
picResultsEX.Print "______________________________________________________________________________________________"
'fill the array and print the data
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, Company(CTR), interestrate(CTR), maturity(CTR), hiprice(CTR), lowprice(CTR), hiyield(CTR), lowyield(CTR)
    picResultsEX.Print Company(CTR); Tab(25); FormatNumber(interestrate(CTR), 3); Tab(40); maturity(CTR); Tab(60); FormatNumber(hiprice(CTR), 3); Tab(75); FormatNumber(lowprice(CTR), 3); Tab(90); FormatNumber(hiyield(CTR), 3); Tab(105); FormatNumber(lowyield(CTR), 3)
Loop
Close #1

End Sub


Private Sub cmdSort_Click()
'this button will do all the sorting in the program
'each section of this code will sort the data in an order depending on the category choosen

Dim pass As Integer
Dim tempN As String
Dim temp As Single
Dim j As Integer
picResultsEX.Cls
cat = txtsort.Text
picResultsEX.Print "Company"; Tab(25); "Interest Rate"; Tab(40); "Date of maturity"; Tab(60); "High Price"; Tab(75); "Low Price"; Tab(90); "YTM(high)"; Tab(105); "YTM(low)"
picResultsEX.Print "______________________________________________________________________________________________"
'sort by company
If cat = 1 Then
    For pass = 1 To CTR - 1
        For j = 1 To CTR - pass
            If Company(j) > Company(j + 1) Then
                tempN = Company(j)
                Company(j) = Company(j + 1)
                Company(j + 1) = tempN
                temp = interestrate(j)
                interestrate(j) = interestrate(j + 1)
                interestrate(j + 1) = temp
                tempN = maturity(j)
                maturity(j) = maturity(j + 1)
                maturity(j + 1) = tempN
                temp = hiprice(j)
                hiprice(j) = hiprice(j + 1)
                hiprice(j + 1) = temp
                temp = lowprice(j)
                lowprice(j) = lowprice(j + 1)
                lowprice(j + 1) = temp
                temp = hiyield(j)
                hiyield(j) = hiyield(j + 1)
                hiyield(j + 1) = temp
                temp = lowyield(j)
                lowyield(j) = lowyield(j + 1)
                lowyield(j + 1) = temp
            End If
        Next j
    Next pass

'sort by interest rate
ElseIf cat = 2 Then
    For pass = 1 To CTR - 1
        For j = 1 To CTR - pass
            If interestrate(j) > interestrate(j + 1) Then
                temp = interestrate(j)
                interestrate(j) = interestrate(j + 1)
                interestrate(j + 1) = temp
                tempN = Company(j)
                Company(j) = Company(j + 1)
                Company(j + 1) = tempN
                tempN = maturity(j)
                maturity(j) = maturity(j + 1)
                maturity(j + 1) = tempN
                temp = hiprice(j)
                hiprice(j) = hiprice(j + 1)
                hiprice(j + 1) = temp
                temp = lowprice(j)
                lowprice(j) = lowprice(j + 1)
                lowprice(j + 1) = temp
                temp = hiyield(j)
                hiyield(j) = hiyield(j + 1)
                hiyield(j + 1) = temp
                temp = lowyield(j)
                lowyield(j) = lowyield(j + 1)
                lowyield(j + 1) = temp
            End If
        Next j
    Next pass

'sort by maturity date
ElseIf cat = 3 Then
    For pass = 1 To CTR - 1
        For j = 1 To CTR - pass
            If Right(maturity(j), 4) > Right(maturity(j + 1), 4) Then
                tempN = maturity(j)
                maturity(j) = maturity(j + 1)
                maturity(j + 1) = tempN
                temp = interestrate(j)
                interestrate(j) = interestrate(j + 1)
                interestrate(j + 1) = temp
                tempN = Company(j)
                Company(j) = Company(j + 1)
                Company(j + 1) = tempN
                temp = hiprice(j)
                hiprice(j) = hiprice(j + 1)
                hiprice(j + 1) = temp
                temp = lowprice(j)
                lowprice(j) = lowprice(j + 1)
                lowprice(j + 1) = temp
                temp = hiyield(j)
                hiyield(j) = hiyield(j + 1)
                hiyield(j + 1) = temp
                temp = lowyield(j)
                lowyield(j) = lowyield(j + 1)
                lowyield(j + 1) = temp
            End If
        Next j
    Next pass

'sort by highest price
ElseIf cat = 4 Then
    For pass = 1 To CTR - 1
        For j = 1 To CTR - pass
            If hiprice(j) < hiprice(j + 1) Then
                temp = interestrate(j)
                interestrate(j) = interestrate(j + 1)
                interestrate(j + 1) = temp
                tempN = Company(j)
                Company(j) = Company(j + 1)
                Company(j + 1) = tempN
                tempN = maturity(j)
                maturity(j) = maturity(j + 1)
                maturity(j + 1) = tempN
                temp = hiprice(j)
                hiprice(j) = hiprice(j + 1)
                hiprice(j + 1) = temp
                temp = lowprice(j)
                lowprice(j) = lowprice(j + 1)
                lowprice(j + 1) = temp
                temp = hiyield(j)
                hiyield(j) = hiyield(j + 1)
                hiyield(j + 1) = temp
                temp = lowyield(j)
                lowyield(j) = lowyield(j + 1)
                lowyield(j + 1) = temp
            End If
        Next j
    Next pass

'sort by lowest price
ElseIf cat = 5 Then
    For pass = 1 To CTR - 1
        For j = 1 To CTR - pass
            If lowprice(j) > lowprice(j + 1) Then
                temp = interestrate(j)
                interestrate(j) = interestrate(j + 1)
                interestrate(j + 1) = temp
                tempN = Company(j)
                Company(j) = Company(j + 1)
                Company(j + 1) = tempN
                tempN = maturity(j)
                maturity(j) = maturity(j + 1)
                maturity(j + 1) = tempN
                temp = hiprice(j)
                hiprice(j) = hiprice(j + 1)
                hiprice(j + 1) = temp
                temp = lowprice(j)
                lowprice(j) = lowprice(j + 1)
                lowprice(j + 1) = temp
                temp = hiyield(j)
                hiyield(j) = hiyield(j + 1)
                hiyield(j + 1) = temp
                temp = lowyield(j)
                lowyield(j) = lowyield(j + 1)
                lowyield(j + 1) = temp
            End If
        Next j
    Next pass

'sort by highest yield
ElseIf cat = 6 Then
    For pass = 1 To CTR - 1
        For j = 1 To CTR - pass
            If hiyield(j) < hiyield(j + 1) Then
                temp = interestrate(j)
                interestrate(j) = interestrate(j + 1)
                interestrate(j + 1) = temp
                tempN = Company(j)
                Company(j) = Company(j + 1)
                Company(j + 1) = tempN
                tempN = maturity(j)
                maturity(j) = maturity(j + 1)
                maturity(j + 1) = tempN
                temp = hiprice(j)
                hiprice(j) = hiprice(j + 1)
                hiprice(j + 1) = temp
                temp = lowprice(j)
                lowprice(j) = lowprice(j + 1)
                lowprice(j + 1) = temp
                temp = hiyield(j)
                hiyield(j) = hiyield(j + 1)
                hiyield(j + 1) = temp
                temp = lowyield(j)
                lowyield(j) = lowyield(j + 1)
                lowyield(j + 1) = temp
            End If
        Next j
    Next pass

'sort by lowest yield
ElseIf cat = 7 Then
    For pass = 1 To CTR - 1
        For j = 1 To CTR - pass
            If lowyield(j) > lowyield(j + 1) Then
                temp = interestrate(j)
                interestrate(j) = interestrate(j + 1)
                interestrate(j + 1) = temp
                tempN = Company(j)
                Company(j) = Company(j + 1)
                Company(j + 1) = tempN
                tempN = maturity(j)
                maturity(j) = maturity(j + 1)
                maturity(j + 1) = tempN
                temp = hiprice(j)
                hiprice(j) = hiprice(j + 1)
                hiprice(j + 1) = temp
                temp = lowprice(j)
                lowprice(j) = lowprice(j + 1)
                lowprice(j + 1) = temp
                temp = hiyield(j)
                hiyield(j) = hiyield(j + 1)
                hiyield(j + 1) = temp
                temp = lowyield(j)
                lowyield(j) = lowyield(j + 1)
                lowyield(j + 1) = temp
            End If
        Next j
    Next pass

Else
'in case the user does enter 1 -7 this meeage box will pop up
    MsgBox "Sorry but you must enter a number from 1 to 7", , "Error"
End If
'print the sorted information
For j = 1 To CTR
    picResultsEX.Print Company(j); Tab(25); FormatNumber(interestrate(j), 3); Tab(40); maturity(j); Tab(60); FormatNumber(hiprice(j), 3); Tab(75); FormatNumber(lowprice(j), 3); Tab(90); FormatNumber(hiyield(j), 3); Tab(105); FormatNumber(lowyield(j), 3)
Next j
                
End Sub
