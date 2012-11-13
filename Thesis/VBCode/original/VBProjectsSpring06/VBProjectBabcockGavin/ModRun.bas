Attribute VB_Name = "ModRun"
'The purpose of this module is to declare all our public variables, therefore allowing us to use them in all of our vast forms
Public Place(1 To 100) As Integer
Public Names(1 To 100) As String
Public Year(1 To 100) As Integer
Public School(1 To 100) As String
Public Minutes(1 To 100) As Integer
Public Seconds(1 To 100) As Single
Public Size As Integer
Public ArraySize As Integer

'To play the .wav file - found at "http://members.aol.com/danp600/vbfaq.html#wave"

Global Const SND_SYNC = &H0
Global Const SND_ASYNC = &H1
Global Const SND_NODEFAULT = &H2
Global Const SND_LOOP = &H8
Global Const SND_NOSTOP = &H10

Public Declare Function sndPlaySound% Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long)

Public Sub PlaySound(SoundName As String)
On Error Resume Next 'If can't find wav don't stall
Dim wFlags%, X%
SoundName = App.Path & "\" & SoundName
wFlags% = SND_ASYNC Or SND_NODEFAULT
X% = sndPlaySound(SoundName, wFlags%)
End Sub



