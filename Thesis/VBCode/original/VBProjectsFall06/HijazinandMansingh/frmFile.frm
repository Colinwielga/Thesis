VERSION 5.00
Begin VB.Form frmEmployee 
   BackColor       =   &H008080FF&
   Caption         =   "Employees"
   ClientHeight    =   6450
   ClientLeft      =   4065
   ClientTop       =   2640
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   ScaleHeight     =   6450
   ScaleWidth      =   6030
   Begin VB.CommandButton show 
      BackColor       =   &H000000C0&
      Caption         =   "Display Data"
      Height          =   375
      Left            =   240
      MaskColor       =   &H000000C0&
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6000
      Width           =   1695
   End
   Begin VB.TextBox txtEmail 
      Height          =   375
      Left            =   2160
      TabIndex        =   17
      Top             =   5520
      Width           =   3015
   End
   Begin VB.TextBox txtPhone 
      Height          =   375
      Left            =   2160
      TabIndex        =   16
      Top             =   4920
      Width           =   2295
   End
   Begin VB.TextBox txtZip 
      Height          =   375
      Left            =   2160
      TabIndex        =   15
      Top             =   4320
      Width           =   975
   End
   Begin VB.TextBox txtState 
      Height          =   375
      Left            =   2160
      TabIndex        =   14
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox txtCity 
      Height          =   375
      Left            =   2160
      TabIndex        =   13
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox txtStreet 
      Height          =   375
      Left            =   2160
      TabIndex        =   12
      Top             =   2520
      Width           =   3135
   End
   Begin VB.TextBox txtFirstName 
      Height          =   375
      Left            =   2160
      TabIndex        =   11
      Top             =   1920
      Width           =   3135
   End
   Begin VB.TextBox txtLastName 
      Height          =   375
      Left            =   2160
      TabIndex        =   10
      Top             =   1320
      Width           =   3135
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H000000C0&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4080
      MaskColor       =   &H000000C0&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6000
      Width           =   1935
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H000000C0&
      Caption         =   "OK"
      Height          =   375
      Left            =   2040
      MaskColor       =   &H000000C0&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6000
      Width           =   1935
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FF8080&
      Caption         =   "Email"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FF8080&
      Caption         =   "Phone"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FF8080&
      Caption         =   "Zip"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FF8080&
      Caption         =   "State"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF8080&
      Caption         =   "City"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF8080&
      Caption         =   "Street"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      Caption         =   "First Name"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "Last Name"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mEmployees As CEmployees
Private mstrKey As String 'Key property of form
Private mstrAction As String 'Action property of form


Private Sub cmdCancel_Click()
 'Return to main form with no action
 
 frmEmployee.Hide

End Sub

Private Sub cmdOK_Click()
 'Choose action depending upon the selected command
 
 Dim strName As String
 Dim strKey As String
 Dim strMsg As String
 Dim intResponse As Integer
 Dim lngKey As Long
 
 Select Case mstrAction
   Case "A"  'Add
    'Add to listbox and set ItemData
    strName = txtLastName.Text & ", " & txtFirstName.Text
    With frmMain.lstEmployee
      .AddItem strName
      lngKey = Val(mEmployees.HighestKey) + 1
      .ItemData(.NewIndex) = lngKey
    End With
    'Add to collection and file
    strKey = Trim(Str(lngKey))
    mEmployees.Add txtLastName.Text, txtFirstName.Text, txtStreet.Text, txtCity.Text, txtState.Text, txtZip.Text, txtPhone.Text, txtEmail.Text, strKey
    Me.Hide
    
   Case "R" 'Remove
    'Remove from list box
    DisplayData
    With frmMain.lstEmployee
       .RemoveItem .ListIndex
    End With
    'Remove from collection and file
    mEmployees.Remove mstrKey
    Me.Hide
    
   Case "D" 'Display
     Me.Hide
         
          
   Case "E" 'Edit
    
    'Update list box
        
    With frmMain.lstEmployee
      .RemoveItem .ListIndex
      strName = Trim(txtLastName.Text) & "," & Trim(txtFirstName.Text)
      .AddItem strName
      .ItemData(.NewIndex) = Val(mstrKey)
    End With
    'Update object in collection and file
    With mEmployees(mstrKey)
        .LastName = txtLastName.Text
        .FirstName = txtFirstName.Text
        .Street = txtStreet.Text
        .City = txtCity.Text
        .State = txtState.Text
        .ZipCode = txtZip.Text
        .Phone = txtPhone.Text
        .Email = txtEmail.Text
    End With
    mEmployees.SaveRecord mstrKey 'Save changes in file
    Me.Hide
    
    Case "B" 'Browse
      'Display next employee
      With frmMain.lstEmployee
        If .ListIndex < .ListCount - 1 Then
        .ListIndex = .ListIndex + 1
        mstrKey = Trim(.ItemData(.ListIndex))
        DisplayData
      Else
        strMsg = "No more employee records to display."
        MsgBox strMsg, vbInformation, "Browse Employee Information"
        Me.Hide
      End If
    End With
 End Select

End Sub

Private Sub Form_Active()
 'Set up the form for the selected action
 
  Select Case mstrAction
   Case "A"  'Add
      lblCommand.Caption = "Add New Employee Information"
      UnlockTheControls
      ClearTextBoxes
      txtLastName.SetFocus
  Case "R" 'Remove
      LockTheControls
      DisplayData
      lblCommand.Caption = "Remove this record?"
  Case "D"   'Display
      LockTheControls
      DisplayData
      lblCommand.Caption = "Display Employee Information"
  Case "E"   'Edit
      UnlockTheControls
      DisplayData
      lblCommand.Caption = "Edit Employee Information"
  Case "B"   'Browse
      LockTheControls
      DisplayData
      lblCommand.Caption = "Browse Employee Information"
  End Select
         
End Sub

Private Sub Form_Load()
 'Create the employee collection object
 
 Dim Employee As CEmployee
 Dim strName As String
 Dim intResponse As Integer
 Dim strMsg  As String
 Set mEmployees = New CEmployees
 
 If mEmployees.FileOpened Then
   For Each Employee In mEmployees
       strName = Trim(Employee.LastName) & ", " & Employee.FirstName
       With frmMain.lstEmployee
         .AddItem strName
         .ItemData(.NewIndex) = (Val(Employee.EmployeeCode))
       End With
     Next
 Else
   strMsg = "File does not exist. Create new File?"
   intResponse = MsgBox(strMsg, vbQuestion + vbYesNo, "Employee File")
   If intResponse = vbYes Then
      mEmployees.OpenNewFile
   Else
      Set mEmployees = Nothing
      Set Employee = Nothing
      End
   End If
End If
       
End Sub

Private Sub Form_Unload(Cancel As Integer)
 'Remove the object from Memory
 
 Set mEmployees = Nothing
End Sub

Private Sub DisplayData()
 'Transfer from the collection to text fields
 
 With mEmployees(mstrKey)
    txtLastName.Text = .LastName
    txtFirstName.Text = .FirstName
    txtStreet.Text = .Street
    txtCity.Text = .City
    txtState.Text = .State
    txtZip.Text = .ZipCode
    txtPhone.Text = .Phone
    txtEmail.Text = .Email
 End With
End Sub

Private Sub ClearTextBoxes()
 'Clear all text boxes
 
 txtLastName.Text = " "
 txtFirstName.Text = " "
 txtStreet.Text = " "
 txtCity.Text = " "
 txtState.Text = " "
 txtZip.Text = " "
 txtPhone.Text = " "
 txtEmail.Text = " "
End Sub

Private Sub LockTheControls()
'Do not allow changes

 txtLastName.Locked = True
 txtFirstName.Locked = True
 txtStreet.Locked = True
 txtCity.Locked = True
 txtState.Locked = True
 txtZip.Locked = True
 txtPhone.Locked = True
 txtEmail.Locked = True
 
End Sub

Private Sub UnlockTheControls()
 'Do all changes
 txtLastName.Locked = False
 txtFirstName.Locked = False
 txtStreet.Locked = False
 txtCity.Locked = False
 txtState.Locked = False
 txtZip.Locked = False
 txtPhone.Locked = False
 txtEmail.Locked = False
 
End Sub

Public Property Let Key(ByVal strKey As String)
 'Write-only property to pass selected key value to form
 
 mstrKey = strKey
End Property

Public Property Let Action(ByVal strAction As String)
 'Write-only property to pass action to form
 
 mstrAction = strAction
End Property

Private Sub mnuAbout_Click()
Load About
About.show vbModal
End Sub

Private Sub show_Click()
 With frmMain.lstEmployee
        If .ListIndex <> -1 Then
          DisplayData
        Else
         MsgBox "No Selection was made", vbExclamation, "No Selection"
         Me.Hide
        End If
 End With
         
End Sub
