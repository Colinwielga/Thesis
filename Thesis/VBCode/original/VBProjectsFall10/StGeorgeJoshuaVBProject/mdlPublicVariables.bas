Attribute VB_Name = "mdlPublicVariables"
Option Explicit
'Defines program level varaibles for use in multiple forms

'Login Variables
Public firstName(1 To 200) As String
Public lastName(1 To 200) As String
Public userName(1 To 200) As String
Public password(1 To 200) As String
Public ClassEnrolled(1 To 200) As String
Public loginCtr As Integer
Public firstTime As Boolean
Public loginVerify(1 To 200) As String
Public verifyCtr As Integer
Public administrator As Boolean

'StudentGrade Data (concurrent with login data above)
Public studentGradeName(1 To 200) As String
Public StudentGrade(1 To 200) As Single
Public studentCorrect(1 To 200) As Integer
Public studentWrong(1 To 200) As Integer
Public StudentAttempted(1 To 200) As Integer

'Noun Parts Variables (for testable verbs)
Public NomSNoun(1 To 200) As String
Public GenSNoun(1 To 200) As String
Public stemNoun(1 To 200) As String
Public definitionNoun(1 To 200) As String
Public GenderNoun(1 To 200) As Integer
Public DeclensionNoun(1 To 200) As Integer
Public NounDifficulty(1 To 200) As Integer
Public NounCtr As Integer

'Class List Varaibles
Public classList(1 To 30) As String
Public classLevel(1 To 30) As Integer
Public classCtr As Integer

'Name and Pos variables for logged in student
Public StudentName As String
Public StudentPosition As Integer
Public StudentLevel As Integer

'Verb Parts Variables
Public VerbPresStem(1 To 200) As String
Public VerbInfinitive(1 To 200) As String
Public VerbPerfStem(1 To 200) As String
Public VerbPartStem(1 To 200) As String
Public VerbDefinition(1 To 200) As String
Public VerbConjugation(1 To 200) As Integer
Public VerbDifficulty(1 To 200) As Integer
Public VerbClass(1 To 200) As Integer
Public VerbPrincipleParts(1 To 200) As String
Public verbCtr As Integer

'Flash Card Variables
Public LatinFlash(1 To 500) As String
Public EnglishFlash(1 To 500) As String
Public partSpeechFlash(1 To 500) As String
Public flashCtr As Integer

'Student Usage Varaibles (to test whether or not a user has used a given function)
Public addedFlashVocab As Boolean

'Noun Declension Pattern Arrays and Ctrs
Public formName(1 To 12) As String
Public First(1 To 12) As String
Public SecondM(1 To 12) As String
Public SecondN(1 To 12) As String
Public ThirdMandF(1 To 12) As String
Public ThirdN(1 To 12) As String
Public FourthM(1 To 12) As String
Public FourthN(1 To 12) As String
Public Fifth(1 To 12) As String

'Verb Indicatvie Endings
Public IVerbFormName(1 To 200) As String
Public IFirstS(1 To 200) As String
Public ISecondS(1 To 200) As String
Public IThirdS(1 To 200) As String
Public IFirstP(1 To 200) As String
Public ISecondP(1 To 200) As String
Public IThirdP(1 To 200) As String
Public IendingCtr As Integer

'Verb Subjungtive Endings
Public SVerbFormName(1 To 200) As String
Public SFirstS(1 To 200) As String
Public SSecondS(1 To 200) As String
Public SThirdS(1 To 200) As String
Public SFirstP(1 To 200) As String
Public SSecondP(1 To 200) As String
Public SThirdP(1 To 200) As String
Public SendingCtr As Integer

'Verb Thematic Vowel Arrays (oraganized by conjugation 1st to 3rd)
Public VowelConjugation(1 To 5) As String
Public VowelIndicative(1 To 5) As String
Public VowelSubjungtive(1 To 5) As String

'The variable for the verbs storing the data about which verb forms are appropriate for each class
Public formClassLevel(1 To 50) As Integer
Public verbFormLevel(1 To 50) As String
Public formTense(1 To 50) As Integer
Public formMood(1 To 50) As Integer
Public formVoice(1 To 50) As Integer
Public formClass(1 To 50) As Integer
Public verbFormctr As Integer

'Edit/Delete Varaibles
Public NounPosition As Integer
Public VerbPosition As Integer
