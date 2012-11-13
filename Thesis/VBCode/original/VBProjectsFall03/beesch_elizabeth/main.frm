VERSION 5.00
Begin VB.Form first 
   BackColor       =   &H0000C000&
   Caption         =   " Basic Information"
   ClientHeight    =   9015
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12075
   DrawMode        =   7  'Invert
   BeginProperty Font 
      Name            =   "MS UI Gothic"
      Size            =   11.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9015
   ScaleWidth      =   12075
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton advance 
      BackColor       =   &H00008000&
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.CommandButton Figure 
      BackColor       =   &H0000FF00&
      Caption         =   "Begin Order"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      MaskColor       =   &H0000C000&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00008000&
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2295
      Left            =   1440
      ScaleHeight     =   2235
      ScaleWidth      =   3555
      TabIndex        =   0
      Top             =   600
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
      Caption         =   "BASIC JOB INFORMATION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "first"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Screen Printing(Main1.vpb)
'Form Name : first(main.frm)
'Author: Elizabeth Beesch
'Date Written: October 27, 2003
'Purpose of Form: To have the user input basic information on the order
'Purpose of Program: The purpose of this program is to produce an
    'easy to read order form. The form will show the costs of having
    'a piece of clothing screen printed with a logo/design/name/number
    'or anything else you may choose printed on it
'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.
Option Explicit
Private Sub Form_Load() ' when the user clicks on this button the following will happen
Figure.Enabled = True 'allows use of button
advance.Enabled = False 'does not allow use of button
End Sub ' ends the commands of the button
Private Sub Figure_Click() ' when the user clicks on this button the following will happen
Dim CTR1 As Integer ' declares the variable
Dim reorder As Integer ' declares the variable
Dim screens As Integer ' declares the variable
advance.Enabled = False ' allows user not to click to the next step until all steps have been completed
Figure.Enabled = True ' allows user to click on the begin order button
CTR = 0 'controls the numebr in the the counter
CTR1 = 0 ' controls the number in the second counter
contact = InputBox("Enter Person to Contact:", "Contact") ' asks user for person to contact
picResults.Print "Contact Person:   "; contact 'print contact information
phonenumber = InputBox("Enter Phone Number XXX-XXX-XXXX:", "Phone Number") ' asks user for phone number
picResults.Print "Phone Number:   "; phonenumber 'print phone number information
joborder = InputBox("Enter Job Title:", "Job Title") ' asks user for job title
picResults.Print "Job Title:   "; joborder ' print job order information
dateneeded = InputBox("Enter Date Needed By MM/DD/YYYY:", "Date") ' asks user for date needed by
picResults.Print "Date Needed By:   "; dateneeded ' print date needed by
numberofItems = InputBox("Enter the Total Number of Items:", "Total Number Of Items") ' ask user for total number of items
picResults.Print "Total Number of Items:"; numberofItems ' prints the total number of items
reorder = InputBox("Is this a Re-Order? Enter 1-Yes / 2- No", "Re-Order") ' asks user if its a reorder
If reorder = 1 Then ' if 1 is entered in the re-order box then Re-Order will be printed
    totalreorder = 2
    picResults.Print ("Re-order") ' prints out the total for reorder cost
ElseIf reorder = 2 Then ' if 2 is entered in the re-order box then no cost will be added
    totalreorder = 0
    picResults.Print ("Not a Re-order") ' prints out the fact there is no re-order cost
ElseIf reorder < 2 Then ' if it is unknown then no cost is added,
    totalreorder = 0
    picResults.Print ("Unknown Re-Order") ' prints out that it is unknown if it is a reorder
End If
screens = InputBox("Enter Total Number of Screens:", "Total Number of Screens") 'asks user for the number of screens needed
Do Until screens = 0  'adds $10.00 to the total screen cost for  each screen neded
    screens = screens - 1 ' subtracts 1 from the total number of screens inorder to add 10 for each screen
    CTR = CTR + 1 ' keeps trach of the number of times 10 is addeded, which is the number of screens
Loop ' repeats steps
totalscreens = CTR * 10#   ' number of screens times 10 to get the total cost for the screens
picResults.Print "Screen Cost:  "; FormatCurrency(totalscreens, 2) ' prints total screens cost
Figure.Enabled = False 'disables the "Begin Order Button"
advance.Enabled = True 'enables the "Next Button"
End Sub ' ends the commands of the button
Private Sub advance_Click() ' advance to the next screen
apparel.Show ' shows the next screen
first.Hide ' hides the current screen
End Sub ' ends the commands of the button
