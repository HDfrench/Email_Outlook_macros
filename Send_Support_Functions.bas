Attribute VB_Name = "Send_Support_Functions"
Public Function StrContent(MyContent As String, Number As Integer, Language As Long) As String

    '**********************************************************************************************
    ' For a list of language codes, please refer to https://support.microsoft.com/en-us/kb/221435 *
    '**********************************************************************************************
    
        'Creates a list based on the language identified
        Select Case Language
            'The numbers for the cases correspond to the values Microsoft gave to each language
            Case 1036, 11276, 3084, 12300, 5132, 13324, 6156, 8204, 10252, 7180, 9228, 4108, 2060 'French
                StrContent = "Pièces jointes : " & Number & vbNewLine & MyContent
            'These are all the English version numbers. For instance 1033 means English-US and 2057 English-UK
            'The series of this case is redundant as we are using English as a default language
            Case 3081, 10249, 4105, 9925, 6153, 8201, 5129, 13321, 7177, 11273, 2057, 1033, 12297 'English
                StrContent = "Files attached: " & Number & vbNewLine & MyContent
            Case 1034, 11274, 16394, 13322, 9226, 5130, 7178, 12298, 17418, 4106, 18442, 19466, 6154, 15370, 10250, 20490, 3082, 14346, 8202 'Spanish
                StrContent = "Archivos adjuntos: " & Number & vbNewLine & MyContent
            Case 2055, 1031, 3079, 5127, 4103 'German
                If Number = 1 Then
                    StrContent = "Anhang: " & Number & " Datei" & vbNewLine & MyContent
                Else
                    StrContent = "Anhang: " & Number & " Dateien" & vbNewLine & MyContent
                End If
            Case Else 'Other languages, uses English by default
                StrContent = "Files attached: " & Number & vbNewLine & Content
        End Select

End Function

Public Function AddSalutations(Language As Long) As String

    Select Case Language
        Case 1036, 11276, 3084, 12300, 5132, 13324, 6156, 8204, 10252, 7180, 9228, 4108, 2060 'French
            AddSalutations = "Cordialement,"
        Case 3081, 10249, 4105, 9925, 6153, 8201, 5129, 13321, 7177, 11273, 2057, 1033, 12297 'English
            AddSalutations = "Kind regards,"
        Case 1034, 11274, 16394, 13322, 9226, 5130, 7178, 12298, 17418, 4106, 18442, 19466, 6154, 15370, 10250, 20490, 3082, 14346, 8202 'Spanish
            AddSalutations = "Saludos,"
        Case 2055, 1031, 3079, 5127, 4103 'German
            AddSalutations = "Mit freundlichen Grüßen,"
        Case Else 'Other languages, uses English by default
            AddSalutations = "Kind regards,"
    End Select
    
    'In this line, you will replace "Your name" with the relevant name to be added at the end of your signature
    AddSalutations = AddSalutations & vbNewLine & vbNewLine & "Hacène" & vbNewLine & vbNewLine
    
End Function

Public Function WhichSignature(Language As Long, MyFormat As Integer) As String

    'This section locates the signature, based on the language previously identified
    'For each language, it looks at the format of the e-mail to associate the relevant signature
    'If you are sending e-mail only in 1 language, you do not need the upper Select Case MyLanguage structure
        Select Case Language
            Case 1036, 11276, 3084, 12300, 5132, 13324, 6156, 8204, 10252, 7180, 9228, 4108, 2060 'French
                'You need to change the last section between \ and the end to put your own signature file
                'This select case structure concerns the format of the e-mail
                Select Case MyFormat
                    Case olFormatHTML
                        WhichSignature = Environ("appdata") & "\Microsoft\Signatures\HD-fr.htm"
                    Case olFormatRichText
                        WhichSignature = Environ("appdata") & "\Microsoft\Signatures\HD-fr.rtf"
                    Case olFormatPlain
                        WhichSignature = Environ("appdata") & "\Microsoft\Signatures\HD-fr.txt"
                End Select
            'This case about English language is redundant since we are using English as our default language (see Case Else)
            Case 3081, 10249, 4105, 9925, 6153, 8201, 5129, 13321, 7177, 11273, 2057, 1033, 12297 'English
                Select Case MyFormat
                    Case olFormatHTML
                        WhichSignature = Environ("appdata") & "\Microsoft\Signatures\HD-en.htm"
                    Case olFormatRichText
                        WhichSignature = Environ("appdata") & "\Microsoft\Signatures\HD-en.rtf"
                    Case olFormatPlain
                        WhichSignature = Environ("appdata") & "\Microsoft\Signatures\HD-en.txt"
                End Select
            Case 1034, 11274, 16394, 13322, 9226, 5130, 7178, 12298, 17418, 4106, 18442, 19466, 6154, 15370, 10250, 20490, 3082, 14346, 8202 'Spanish
                Select Case MyFormat
                    Case olFormatHTML
                        WhichSignature = Environ("appdata") & "\Microsoft\Signatures\HD-es.htm"
                    Case olFormatRichText
                        WhichSignature = Environ("appdata") & "\Microsoft\Signatures\HD-es.rtf"
                    Case olFormatPlain
                        WhichSignature = Environ("appdata") & "\Microsoft\Signatures\HD-es.txt"
                End Select
            Case 2055, 1031, 3079, 5127, 4103 'German
                Select Case MyFormat
                    Case olFormatHTML
                        WhichSignature = Environ("appdata") & "\Microsoft\Signatures\HD-de.htm"
                    Case olFormatRichText
                        WhichSignature = Environ("appdata") & "\Microsoft\Signatures\HD-de.rtf"
                    Case olFormatPlain
                        WhichSignature = Environ("appdata") & "\Microsoft\Signatures\HD-de.txt"
                End Select
            Case Else 'Other languages, uses English by default
                Select Case MyFormat
                    Case olFormatHTML
                        WhichSignature = Environ("appdata") & "\Microsoft\Signatures\HD-en.htm"
                    Case olFormatRichText
                        WhichSignature = Environ("appdata") & "\Microsoft\Signatures\HD-en.rtf"
                    Case olFormatPlain
                        WhichSignature = Environ("appdata") & "\Microsoft\Signatures\HD-en.txt"
                End Select
        End Select

End Function

