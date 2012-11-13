Attribute VB_Name = "Module1"
Global Variable 'these variables are public because theoretically they would need to be accessed in multiple forms.

    Public A As String
    Public SoundPath As String 'SoundFile must be accessible both within the SoundMenu form and also the DisplayedAlarm form.
    Public PicPath As String
    Public AlarmHour As String 'because it begins with a 0?
    Public AlarmMinute As Integer
    Public AlarmSecond As Integer
    
    
    
