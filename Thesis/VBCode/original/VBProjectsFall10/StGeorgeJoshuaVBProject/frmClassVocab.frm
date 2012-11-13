VERSION 5.00
Begin VB.Form frmClassVocab 
   BackColor       =   &H00000080&
   Caption         =   "Lingua Vivens - Admin Options - Class Vocabulary"
   ClientHeight    =   10095
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15120
   LinkTopic       =   "Form1"
   ScaleHeight     =   10095
   ScaleWidth      =   15120
   Begin VB.CommandButton cmdEditDeleteVerb 
      BackColor       =   &H00000080&
      Caption         =   "Edit Selected VERB"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3360
      Width           =   2175
   End
   Begin VB.CommandButton cmdEditDeleteNoun 
      BackColor       =   &H00000080&
      Caption         =   "Edit Selected NOUN"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2640
      Width           =   2175
   End
   Begin VB.ListBox lstVerbs 
      Height          =   7665
      ItemData        =   "frmClassVocab.frx":0000
      Left            =   6600
      List            =   "frmClassVocab.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   13
      Top             =   1680
      Width           =   3975
   End
   Begin VB.ListBox lstNouns 
      Height          =   7665
      ItemData        =   "frmClassVocab.frx":0004
      Left            =   3120
      List            =   "frmClassVocab.frx":0006
      Sorted          =   -1  'True
      TabIndex        =   12
      Top             =   1680
      Width           =   3375
   End
   Begin VB.PictureBox picLevel 
      Height          =   375
      Left            =   13800
      ScaleHeight     =   315
      ScaleWidth      =   1035
      TabIndex        =   7
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdMove 
      BackColor       =   &H00000080&
      Caption         =   "Move VerbForm to Different Class"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4080
      Width           =   2175
   End
   Begin VB.ListBox lstVerbForms 
      Height          =   7665
      ItemData        =   "frmClassVocab.frx":0008
      Left            =   10680
      List            =   "frmClassVocab.frx":000A
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   1680
      Width           =   3015
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00808080&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8040
      Width           =   2175
   End
   Begin VB.CommandButton cmdLogOut 
      BackColor       =   &H00808080&
      Caption         =   "LogOut"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7200
      Width           =   2175
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00808080&
      Caption         =   "Return to Administration Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6360
      Width           =   2175
   End
   Begin VB.ComboBox cboClasses 
      Height          =   315
      Left            =   600
      TabIndex        =   1
      Text            =   "Select a Class ..."
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton cmdDisplayClass 
      BackColor       =   &H00000080&
      Caption         =   "Show Vocab List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label lblVerbTitle 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Verb Principle Parts"
      Height          =   375
      Left            =   6600
      TabIndex        =   15
      Top             =   1320
      Width           =   3975
   End
   Begin VB.Label lblNounsTitle 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nouns (nomS,genS)"
      Height          =   375
      Left            =   3120
      TabIndex        =   14
      Top             =   1320
      Width           =   3375
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "View and Modify Class Vocabulary"
      BeginProperty Font 
         Name            =   "Roman"
         Size            =   27.75
         Charset         =   255
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   1560
      TabIndex        =   11
      Top             =   480
      Width           =   10335
   End
   Begin VB.Label lblDirections 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmClassVocab.frx":000C
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   600
      TabIndex        =   10
      Top             =   4800
      Width           =   2175
   End
   Begin VB.Label lblLevel 
      BackColor       =   &H00FF0000&
      Caption         =   "Level of Currently Selected"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   13800
      TabIndex        =   9
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label lblForms 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Verb Forms"
      Height          =   375
      Left            =   10680
      TabIndex        =   8
      Top             =   1320
      Width           =   3015
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   9375
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   15375
   End
End
Attribute VB_Name = "frmClassVocab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FormItemsAdded As Integer
Dim nounItemsAdded As Integer
Dim verbItemsAdded As Integer
Private Sub cmdDisplayClass_Click()
    'Displays a particular or all classes selected by the user in a combobox
    'Declares variables
    Dim class As String
    Dim pos As Integer
    Dim ctr As Integer
    Dim Found As Boolean
    'Initializes varaibles
    Found = False
    ctr = 0
    class = cboClasses.Text
    'Searches for a match with the selection made by the user, saves the positions of this class at variable ctr
    Do While Not Found And ctr < classCtr
        ctr = ctr + 1
        If class = classList(ctr) Then
            Found = True
        End If
    Loop
    'Descision to display selected class, display all classes, or return an error if the user did not select a class
    'Clears the lists
    lstVerbForms.Clear
    lstNouns.Clear
    lstVerbs.Clear
    
    If Found Then
        
        
        'Loops cross the arrays to determine which nouns fir the user input criteria (exhaustive search)
        nounItemsAdded = 0
        For pos = 1 To NounCtr
            'Displays noun if the nounDifficulty of search position is less than or equal to the classLevel of the ctr from above
            If NounDifficulty(pos) <= classLevel(ctr) Then
                nounItemsAdded = nounItemsAdded + 1
                lstNouns.AddItem NomSNoun(pos) & ", " & GenSNoun(pos)
            End If
        Next pos
        'Does the same as above but for the verb
        verbItemsAdded = 0
        For pos = 1 To verbCtr
            If VerbDifficulty(pos) <= classLevel(ctr) Then
                verbItemsAdded = verbItemsAdded + 1
                'Displays the verb Principle parts
                lstVerbs.AddItem VerbPrincipleParts(pos)
            End If
        Next pos
        'Lists out the verbForms in lstverbForms
        FormItemsAdded = 0
        For pos = 1 To verbFormctr
            If formClassLevel(pos) <= classLevel(ctr) Then
                FormItemsAdded = FormItemsAdded + 1
                lstVerbForms.AddItem verbFormLevel(pos)
            End If
        Next pos
        
    ElseIf class = "All Classes" Then
        'Same as above only prints all classes regardless of class level
        
        nounItemsAdded = 0
        For pos = 1 To NounCtr
            'Displays noun if the nounDifficulty of search position is less than or equal to the classLevel of the ctr from above
            If NounDifficulty(pos) <= classLevel(ctr) Then
                nounItemsAdded = nounItemsAdded + 1
                lstNouns.AddItem NomSNoun(pos) & ", " & GenSNoun(pos)
            End If
        Next pos
        
        verbItemsAdded = 0
        For pos = 1 To verbCtr
            If VerbDifficulty(pos) <= classLevel(ctr) Then
                verbItemsAdded = verbItemsAdded + 1
                'Displays the verb Principle parts
                lstVerbs.AddItem VerbPrincipleParts(pos)
            End If
        Next pos
        
        FormItemsAdded = 0
        For pos = 1 To verbFormctr
            FormItemsAdded = verbFormctr
            lstVerbForms.AddItem verbFormLevel(pos)
        Next pos
        
    Else 'error handling for if user does not select a class
        MsgBox "Please select a class to view"
    End If
    
    
End Sub

Private Sub cmdEditDeleteNoun_Click()
    'Used to move to the EDIT/Delete form used to edit or delete a noun
    'usefult Varaibles
    Dim noun As String
    Dim Search As String
    Dim listPos As Integer
    Dim arrayPos As Integer
    Dim pos As Integer
    Dim Found As Boolean
    Dim selected As Boolean
    Dim className As String
    'initializes varaibles (listpos = -1 because the list starts at 0)
    selected = False
    listPos = -1
    'Match and stop to find the selected item
    Do Until selected Or listPos = nounItemsAdded
        listPos = listPos + 1
        If lstNouns.selected(listPos) = True Then
            selected = True
        End If
    Loop
    'if not selected returns a error(although the user cannot use the button without first selecting something)
    If Not selected Then
        MsgBox "Please select a Noun to be edited/deleted"
        Exit Sub
    End If
    'Gives search the string value fo the selected list item
    Search = lstNouns.List(listPos)
    'gives the noun varaible the nom s form of the noun selected
    noun = Left(Search, (InStr(Search, ",")) - 1)
    'initializes more varaibles
    Found = False
    arrayPos = 0
    'Another match and stop in order to search the array for a matching position for the selected noun
    Do Until Found Or arrayPos = NounCtr
        arrayPos = arrayPos + 1
        If noun = NomSNoun(arrayPos) Then
            Found = True
        End If
    Loop
    'Condition when found
    If Found Then
        'Sets the public varaible nounPosition = to the found array pos
        NounPosition = arrayPos
    Else
        'manages the rare/impossible case that the noun selected does not match any in the array
        MsgBox "Noun not found?!"
        Exit Sub
    End If
     
    'Shows the edit form and hides the class vocab form
    frmEditDeleteNoun.Show
    frmClassVocab.Hide
    'Clears the lists
    lstNouns.Clear
    lstVerbs.Clear
    lstVerbForms.Clear
    'disables the buttons
    cmdEditDeleteVerb.Enabled = False
    cmdEditDeleteNoun.Enabled = False
    'Sets the text boxes of the next form as their array counterparts of the selected noun
    frmEditDeleteNoun.txtNom.Text = NomSNoun(NounPosition)
    frmEditDeleteNoun.txtGen.Text = GenSNoun(NounPosition)
    frmEditDeleteNoun.txtStem.Text = stemNoun(NounPosition)
    frmEditDeleteNoun.txtDefinition.Text = definitionNoun(NounPosition)
    'Does the same as above only with optionbuttons, conditionally set by whatever declension, gender or difficulty is for the noun
    Select Case DeclensionNoun(NounPosition)
        Case 1
            frmEditDeleteNoun.optFirst.Value = True
        Case 2
            frmEditDeleteNoun.optSecond.Value = True
        Case 3
            frmEditDeleteNoun.optThird.Value = True
        Case 4
            frmEditDeleteNoun.optFourth.Value = True
        Case 5
            frmEditDeleteNoun.optFifth.Value = True
    End Select
    
    Select Case GenderNoun(NounPosition)
        Case 1
            frmEditDeleteNoun.optFeminine.Value = True
        Case 2
            frmEditDeleteNoun.optMasculine.Value = True
        Case 3
            frmEditDeleteNoun.optNeuter.Value = True
    End Select
    
    Select Case NounDifficulty(NounPosition)
        Case 1
            frmEditDeleteNoun.optEasy.Value = True
        Case 2
            frmEditDeleteNoun.optIntermediate.Value = True
        Case 3
            frmEditDeleteNoun.optDifficult.Value = True
        Case 4
            frmEditDeleteNoun.optCollegeLevel.Value = True
    End Select
    
End Sub

Private Sub cmdEditDeleteVerb_Click()
    'Shows the edit delete verb form for a selected verb from the lstVerbs
    'useful Varaibles
    Dim VerbParts As String
    Dim Search As String
    Dim listPos As Integer
    Dim arrayPos As Integer
    Dim pos As Integer
    Dim Found As Boolean
    Dim selected As Boolean
    Dim className As String
    'Initializes soem (list pos is -1 because the first item index = 0
    selected = False
    listPos = -1
    'Searches for whichever item is selected
    Do Until selected Or listPos = verbItemsAdded
        listPos = listPos + 1
        If lstVerbs.selected(listPos) = True Then
            selected = True
        End If
    Loop
    'deals with the impossible case that the user did not first select an item
    If Not selected Then
        MsgBox "Please select a Verb to be edited/deleted"
        Exit Sub
    End If
    'Gives the verbparts varaible the string value of the list item selected
    VerbParts = lstVerbs.List(listPos)
    'Initializes varaibles
    Found = False
    arrayPos = 0
    'match and stop to determine the position in the arrays of the verb which was selected in the list
    Do Until Found Or arrayPos = verbCtr
        arrayPos = arrayPos + 1
        If VerbParts = VerbPrincipleParts(arrayPos) Then
            Found = True
        End If
    Loop
    'Condition if found
    If Found Then
        'gives the public varaible verbposition the array position of the selected verb
        VerbPosition = arrayPos
    Else
        MsgBox "Verb not found?!"
        Exit Sub
    End If
    'Changes/updates the textboxes with the information of the selected verb
    frmEditDeleteVerb.txtPresentStem.Text = VerbPresStem(VerbPosition)
    frmEditDeleteVerb.txtInfinitive.Text = VerbInfinitive(VerbPosition)
    frmEditDeleteVerb.txtPerfectStem.Text = VerbPerfStem(VerbPosition)
    frmEditDeleteVerb.txtParticipleStem.Text = VerbPartStem(VerbPosition)
    frmEditDeleteVerb.txtDefinition.Text = VerbDefinition(VerbPosition)
    frmEditDeleteVerb.txtPrincipleParts.Text = VerbPrincipleParts(VerbPosition)
    'Does the same as above only conditionally for the option buttons
    Select Case VerbConjugation(VerbPosition)
        Case 1
            frmEditDeleteVerb.optFirst.Value = True
        Case 2
            frmEditDeleteVerb.optSecond.Value = True
        Case 3
            frmEditDeleteVerb.optThird.Value = True
        Case 4
            frmEditDeleteVerb.optThirdIO.Value = True
        Case 5
            frmEditDeleteVerb.optFourth.Value = True
    End Select
    
    Select Case VerbDifficulty(VerbPosition)
        Case 1
            frmEditDeleteVerb.optEasy.Value = True
        Case 2
            frmEditDeleteVerb.optIntermediate.Value = True
        Case 3
            frmEditDeleteVerb.optDifficult.Value = True
        Case 4
            frmEditDeleteVerb.optCollegeLevel.Value = True
    End Select
    
    Select Case VerbClass(VerbPosition)
        Case 1
            frmEditDeleteVerb.optRegular.Value = True
        Case 2
            frmEditDeleteVerb.optDeponent.Value = True
        Case 3
            frmEditDeleteVerb.optSemiDeponent.Value = True
        Case 4
            frmEditDeleteVerb.optDefective.Value = True
    End Select
    'Shows the form and hides the current one
    frmEditDeleteVerb.Show
    frmClassVocab.Hide
    'clears the list for further use
    lstVerbs.Clear
    lstNouns.Clear
    lstVerbForms.Clear
    'disables the buttons
    cmdEditDeleteVerb.Enabled = False
    cmdEditDeleteNoun.Enabled = False
    
End Sub

Private Sub cmdLogOut_Click()
    'Logs out using logout subroutine (cf. mdlPublicSubs)
    frmClassVocab.Hide
    lstNouns.Clear
    lstVerbs.Clear
    lstVerbForms.Clear
    cmdEditDeleteVerb.Enabled = False
    cmdEditDeleteNoun.Enabled = False
    Call LogOut
End Sub

Private Sub cmdMove_Click()
    'Used to shift about the verb form to other classes
    'Usefull varaibles
    Dim verbForm As String
    Dim listPos As Integer
    Dim arrayPos As Integer
    Dim pos As Integer
    Dim Found As Boolean
    Dim selected As Boolean
    Dim className As String
    Dim newLevel As String
    'initialization of variables, lispo = -1 because the first listIndex is 0
    selected = False
    listPos = -1
    'Searches via match and stop for whichever verbform is selected
    Do Until selected Or listPos = FormItemsAdded
        listPos = listPos + 1
        If lstVerbForms.selected(listPos) = True Then
            selected = True
        End If
    Loop
    'Deals witht eh impossible case in which the user does not select a verbform
    If Not selected Then
        MsgBox "Please select a verbForm to be modified"
        Exit Sub
    End If
    'Gives the verbform varaible the string value of the verbform selectd from the list
    verbForm = lstVerbForms.List(listPos)
    'initializes varaibles
    Found = False
    arrayPos = 0
    'Match and stop for the verb form in its respective array
    Do Until Found Or arrayPos = verbFormctr
        arrayPos = arrayPos + 1
        If verbForm = verbFormLevel(arrayPos) Then
            Found = True
        End If
    Loop
    'Gives the class name varaible the string value of the class name for the difficulty of the verbForm
    className = classList(formClassLevel(arrayPos))
    'initializes more varaibles (asking the user for what new level the verb form will be
    pos = 0
    newLevel = InputBox("Please enter a new class level for " & UCase(verbFormLevel(arrayPos)) & " for the class " & UCase(className) & " (1 for Latin I, 2 for Latin II, 3 for latin III, and 4 for Latin IV)(enter -1 to exit)")
    If newLevel = 1 Or newLevel = 2 Or newLevel = 3 Or newLevel = 4 Then ' if the user enters a valid integer
        If Found = True Then
            'gives the array value of the class difficulty level the new level and writes the entirety of the array to a blank textfile
            formClassLevel(arrayPos) = newLevel
            Open App.Path & "\Data\verbFormsbyClass.txt" For Output As #1
                For pos = 1 To verbFormctr
                    Write #1, verbFormLevel(pos), formClassLevel(pos), formTense(pos), formMood(pos), formVoice(pos), formClass(pos)
                Next pos
            Close #1
            MsgBox UCase(verbFormLevel(arrayPos)) & "'s class level has been changed to " & classList(newLevel)
        Else
            MsgBox "Form not found?!"
        End If
    ElseIf newLevel = -1 Then
        'Exits sub if the user enters the the exit flag of -1
        Exit Sub
    Else
        MsgBox "Please enter a valid number for new class level"
    End If
    
        
    
End Sub

Private Sub cmdQuit_Click()
    'Ends the program
    End
End Sub

Private Sub cmdReturn_Click()
    'Returns to the admin pane
    frmClassVocab.Hide
    frmAdmin.Show
    'reinitializes the form
    lstNouns.Clear
    lstVerbs.Clear
    lstVerbForms.Clear
    cmdEditDeleteVerb.Enabled = False
    cmdEditDeleteNoun.Enabled = False
End Sub

Private Sub Form_Load()
    'Upon form load this code reads the classes from the Class list text file and reads them into the combo box
    'defines the variables
    Dim pos As Integer
    Dim Pass As Integer
    'Calls the Read Classes subroutine (reference mdlPublicSubs for more info)
    Call ReadClasses
    'Reinitializes cbo box
    cboClasses.Clear
    cboClasses.Text = "Select a Class ..."
    cboClasses.AddItem "All Classes"
    'Adds each item into the combo box
    For pos = 1 To classCtr
        cboClasses.AddItem classList(pos)
    Next pos
    
    Open App.Path & "\Data\VerbFormsByClass.txt" For Input As #1
        verbFormctr = 0
        Do Until EOF(1)
            verbFormctr = verbFormctr + 1
            Input #1, verbFormLevel(verbFormctr), formClassLevel(verbFormctr), formTense(verbFormctr), formMood(verbFormctr), formVoice(verbFormctr), formClass(verbFormctr)
        Loop
    Close #1
    
End Sub

Private Sub lstNouns_Click()
    'Deals witht he selection of a noun for the lstNouns
    'Enables the cmdEditDeleteNoun button
    cmdEditDeleteNoun.Enabled = True
    'Declares sueful varaibles
    Dim pos As Integer
    Dim selected As Boolean
    Dim noun As String
    Dim Search As String
    Dim Found As Boolean
    Dim arrayPos As Integer
    'initializes varaibles
    selected = False
    pos = -1
    arrayPos = 0
    Found = False
    'Searches the list to find which one is selected
    Do Until selected Or pos = nounItemsAdded
        pos = pos + 1
        If lstNouns.selected(pos) = True Then
            selected = True
        End If
    Loop
    'Clears tthe picLevel
    picLevel.Cls
    'gives search the string avalue of the noun selected
    Search = lstNouns.List(pos)
    'gives the noun varaible the string value of the nomS of the noun selected
    noun = Left(Search, InStr(Search, ",") - 1)
    
    'Searches until it finds a corresponding NomSNoun, and prints the class label for the corresponding class difficulty
    Do Until Found Or arrayPos = NounCtr
        arrayPos = arrayPos + 1
        If noun = NomSNoun(arrayPos) Then
            Found = True
            picLevel.Print classList(NounDifficulty(arrayPos))
        End If
    Loop
    
   
End Sub

Private Sub lstVerbForms_Click()
    'Deals with the event when a verbform from the verbform list is clicked
    'Enables the button to edit the verbForms difficulty
    cmdMove.Enabled = True
    'Declares useful Varaibles
    Dim pos As Integer
    Dim selected As Boolean
    Dim verbForm As String
    Dim Found As Boolean
    Dim arrayPos As Integer
    'initializes varaibles
    selected = False
    pos = -1
    arrayPos = 0
    Found = False
    'Searches for the selected verbform
    Do Until selected Or pos = FormItemsAdded
        pos = pos + 1
        If lstVerbForms.selected(pos) = True Then
            selected = True
        End If
    Loop
    'prepares the picLevel
    picLevel.Cls
    'Gives the verbform varaibles the string value of the verbform selected
    verbForm = lstVerbForms.List(pos)
    'loops until is finds the corresponding verbform and prints the class label for the difficulty of the verbForm
    Do Until Found Or arrayPos = verbFormctr
        arrayPos = arrayPos + 1
        If verbForm = verbFormLevel(arrayPos) Then
            Found = True
            picLevel.Print classList(formClassLevel(arrayPos))
        End If
    Loop
    
    
End Sub


Private Sub lstVerbs_Click()
    'Deals witht he event of the selection of one of the verbs from the verb list
    'enables the editing button for the verbs
    cmdEditDeleteVerb.Enabled = True
    'Declares sueful varaibles
    Dim pos As Integer
    Dim selected As Boolean
    Dim VerbParts As String
    Dim Found As Boolean
    Dim arrayPos As Integer
    'initializes varaibles
    selected = False
    pos = -1
    arrayPos = 0
    Found = False
    'Loops to find the selected verb
    Do Until selected Or pos = verbItemsAdded
        pos = pos + 1
        If lstVerbs.selected(pos) = True Then
            selected = True
        End If
    Loop
    'prepares piclevel
    picLevel.Cls
    'Gives verbparts the string value of the item selected
    VerbParts = lstVerbs.List(pos)
    'loops and compares the selected verbparts with the principle parts in the verb arrays and prints the corresponding class label for the verb duifficulty.
    Do Until Found Or arrayPos = verbCtr
        arrayPos = arrayPos + 1
        If VerbParts = VerbPrincipleParts(arrayPos) Then
            Found = True
            picLevel.Print classList(VerbDifficulty(arrayPos))
        End If
    Loop
End Sub
