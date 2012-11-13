VERSION 5.00
Begin VB.Form Employment 
   BackColor       =   &H00C00000&
   Caption         =   "Form1"
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   ScaleHeight     =   7470
   ScaleWidth      =   10935
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPhone3 
      Height          =   375
      Left            =   2280
      TabIndex        =   28
      Top             =   5160
      Width           =   975
   End
   Begin VB.TextBox txtPhone2 
      Height          =   375
      Left            =   1200
      TabIndex        =   27
      Top             =   5160
      Width           =   735
   End
   Begin VB.TextBox txtPhone1 
      Height          =   375
      Left            =   120
      TabIndex        =   26
      Top             =   5160
      Width           =   735
   End
   Begin VB.CommandButton cmdMinnesota 
      BackColor       =   &H0000C0C0&
      Caption         =   "Return to Minnesota Home Page"
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   6600
      Width           =   1575
   End
   Begin VB.CommandButton cmdSubmit 
      BackColor       =   &H0000C0C0&
      Caption         =   "Submit Your Application to the Schwan's Food Company"
      Height          =   1575
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton cmdMarshall 
      BackColor       =   &H0000C0C0&
      Caption         =   "Return to Marshall Home Page"
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5760
      Width           =   1575
   End
   Begin VB.TextBox txtDescription 
      Height          =   2055
      Left            =   3720
      TabIndex        =   22
      Top             =   5280
      Width           =   7095
   End
   Begin VB.TextBox txtCity 
      Height          =   285
      Left            =   9360
      TabIndex        =   19
      Top             =   4080
      Width           =   1455
   End
   Begin VB.ComboBox comboState 
      Height          =   315
      ItemData        =   "Employment.frx":0000
      Left            =   7800
      List            =   "Employment.frx":009D
      TabIndex        =   17
      Text            =   "---------------------"
      Top             =   4080
      Width           =   1335
   End
   Begin VB.ComboBox comboCountry 
      Height          =   315
      ItemData        =   "Employment.frx":02DF
      Left            =   5760
      List            =   "Employment.frx":0316
      TabIndex        =   15
      Text            =   "--------------------------------"
      Top             =   4080
      Width           =   1815
   End
   Begin VB.TextBox txtAddress2 
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   4200
      Width           =   5295
   End
   Begin VB.TextBox txtAddress1 
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   3720
      Width           =   5295
   End
   Begin VB.ComboBox comboYear 
      Height          =   315
      ItemData        =   "Employment.frx":03B0
      Left            =   9480
      List            =   "Employment.frx":03D5
      TabIndex        =   10
      Text            =   "Year"
      Top             =   2880
      Width           =   1335
   End
   Begin VB.ComboBox comboDay 
      Height          =   315
      ItemData        =   "Employment.frx":041B
      Left            =   8520
      List            =   "Employment.frx":047C
      TabIndex        =   9
      Text            =   "Day"
      Top             =   2880
      Width           =   975
   End
   Begin VB.ComboBox comboMonth 
      Height          =   315
      ItemData        =   "Employment.frx":04F5
      Left            =   7200
      List            =   "Employment.frx":051D
      TabIndex        =   8
      Text            =   "Month"
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox txtEmail 
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   2880
      Width           =   3135
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   2775
   End
   Begin VB.PictureBox Picture2 
      Height          =   2295
      Left            =   2160
      Picture         =   "Employment.frx":0584
      ScaleHeight     =   2235
      ScaleWidth      =   8715
      TabIndex        =   1
      Top             =   0
      Width           =   8775
   End
   Begin VB.PictureBox Picture1 
      Height          =   2295
      Left            =   0
      Picture         =   "Employment.frx":EB23
      ScaleHeight     =   2235
      ScaleWidth      =   2115
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      Index           =   1
      X1              =   2160
      X2              =   2040
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      Index           =   0
      X1              =   960
      X2              =   1080
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H0000C0C0&
      Caption         =   "Tell Us About Yourself! (Driving Experience,  Education, Salary, Talents, Previous Occupations):"
      Height          =   375
      Left            =   3720
      TabIndex        =   21
      Top             =   4800
      Width           =   7095
   End
   Begin VB.Label lblPhone 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      Caption         =   "                        Phone Number:                   # # #  -  # # #  -  # # # #"
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   4800
      Width           =   3135
   End
   Begin VB.Label lblCity 
      BackColor       =   &H0000C0C0&
      Caption         =   "City:"
      Height          =   375
      Left            =   9360
      TabIndex        =   18
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label lblState 
      BackColor       =   &H0000C0C0&
      Caption         =   "State:"
      Height          =   375
      Left            =   7800
      TabIndex        =   16
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label lblCountry 
      BackColor       =   &H0000C0C0&
      Caption         =   "Country:"
      Height          =   375
      Left            =   5760
      TabIndex        =   14
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label lblAddress 
      BackColor       =   &H0000C0C0&
      Caption         =   "Address:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3480
      Width           =   2775
   End
   Begin VB.Label lblBirthday 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      Caption         =   "Date of Birth:"
      Height          =   255
      Left            =   7200
      TabIndex        =   7
      Top             =   2640
      Width           =   3615
   End
   Begin VB.Label lblEmail 
      BackColor       =   &H0000C0C0&
      Caption         =   "E-Mail Address:"
      Height          =   255
      Left            =   3720
      TabIndex        =   5
      Top             =   2640
      Width           =   3135
   End
   Begin VB.Label lblName 
      BackColor       =   &H0000C0C0&
      Caption         =   "First Name,  Middle Initial, Last Name:"
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   2760
   End
   Begin VB.Label lblContact 
      BackColor       =   &H00000000&
      Caption         =   "Contact Profile"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   2400
      Width           =   10935
   End
End
Attribute VB_Name = "Employment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Minnesoooota
'Form Name: Employment
'Author: Danielle Johnson and Tony Blum
'Date Written: March 26th 2008
'The purpose of this form is to simulate the process of filling out an online application form to work at Schwan's.
Private Sub cmdMarshall_Click()
'Hides current form and returns to the Marshall page
Employment.Hide
Marshall.Show
End Sub

Private Sub cmdMinnesota_Click()
'Hides current form and returns to the Minnesota page
Employment.Hide
Minnesota.Show
End Sub

Private Sub cmdSubmit_Click()
'This saves all of the information entered into this form under variablers which are globalized under the module form
'This allows you to print results on the application form without having to enter the data twice
YourName = txtName.Text
EMail = txtEmail.Text
Address1 = txtAddress1.Text
Address2 = txtAddress2.Text
Month = comboMonth.Text
Day = comboDay.Text
Year = comboYear.Text
Country = comboCountry.Text
State = comboState.Text
City = txtCity.Text
Number1 = txtPhone1.Text
Number2 = txtPhone2.Text
Number3 = txtPhone3.Text
Description = txtDescription.Text
'Hides the current form and brings you to the Application form
Employment.Hide
Application.Show
End Sub

