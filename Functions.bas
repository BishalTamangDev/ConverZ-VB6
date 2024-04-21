Attribute VB_Name = "Module1"
Dim count As Integer
Dim sapi As Object
Dim MaleSpeaker As String * 6
Dim FemaleSpeaker As String * 6

Type info
    speech As String * 100
    character As String * 6
End Type

Public Function TrySpeech(location As String, character As String)
    Set sapi = CreateObject("sapi.spvoice")
    
    If character = "David" Then
        Set sapi.Voice = sapi.GetVoices.Item(0)
        sapi.Speak location
    Else
        Set sapi.Voice = sapi.GetVoices.Item(1)
        sapi.Speak location
    End If
End Function


Public Function Conversation(dialog As String, speaker As String)
    Set sapi = CreateObject("sapi.spvoice")
    
    MaleSpeaker = "David"
    FemaleSpeaker = "Zira"
    
    If speaker = MaleSpeaker Then
        Set sapi.Voice = sapi.GetVoices.Item(0)
        sapi.Speak dialog
    Else
        Set sapi.Voice = sapi.GetVoices.Item(1)
        sapi.Speak dialog
    End If
End Function


Public Function speech_count(filename As String, received As info) As Integer
    count = 1
    On Error GoTo errorhandler
    Open filename For Random As #1 Len = 106
        While EOF(1) = False
           Get #1, count, received
           count = count + 1
        Wend
    Close #1
    count = count - 2
    speech_count = count
    Exit Function
    
errorhandler:
    Close #1
    count = 0
    speech_count = count
End Function


Public Function filename_redundency(file_to_check As String) As Boolean
    On Error GoTo errorhandler
    Open file_to_check For Input As #2
    
    Close #2
    filename_redundency = True
    Exit Function
    
errorhandler:
    Close #2
    filename_redundency = False
End Function
