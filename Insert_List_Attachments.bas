Attribute VB_Name = "Insert_List_Attachments"
'This option is to force the coder to define all variables and objects used
'If a variable has not been defined before using it, this will generate an error
Option Explicit

'Since we will constantly refer to Em (which is our e-mail)
'Instead of writing a "Dim" statement for it in every single Sub,
'We define it as Public at the top of the module
Public Em As Outlook.MailItem

Public Sub Attachment_List()

    'This defines a chain of characters to inform the user
    Dim MyMessage As String
    'This defines MyAnswer as a small number
    Dim MyAnswer   As Integer
    
    'This section is not necessary since we are putting our button to call the macro on our e-mail ribbon
    'However, if you call the macro without composing an e-mail (by pressing Alt-F8 and selecting it from your inbox)
    'This will generate an error
    On Error Resume Next
    'Defines Em as an e-mail
    Set Em = Application.ActiveInspector.CurrentItem
    'If there is no message, this generates an error
    'The error generates a number
    'If this number is different from 0 then do not run the macro
    If Err <> 0 Then
        Err.Clear   'clear the error
        Exit Sub    'leaves the macro
    End If
    
    'In our e-mail to be sent, we are checking if there is "To"
    'If your cursor is still in the "To" line, it will tell you that the line is empty.
    'Make sure the cursor is in the body of the e-mail
    If Em.To = Empty Then
        MyMessage = "The 'To' line is empty."
    End If
    
    'In our e-mail to be sent, we are checking if there is "Subject"
    If Em.Subject = Empty Then
        MyMessage = MyMessage & vbNewLine & "There is no subject to that e-mail."
    End If
    
    'If there is an error message (due to missing elements), then displays the message and leaves macro
    If MyMessage <> vbNullString Then
        MsgBox MyMessage
        Exit Sub
    End If
    
    'In our e-mail to be sent, we are checking whether there are attachments
    If Em.Attachments.Count <> 0 Then
        'If there are attachments, we show the form to manage them
        With Fm_AttachmentList
            .UserForm_Initialize
            .Show
        End With
    Else
        'If there are no attachments, we give a message to confirm
        'MyAnswer is the value from clicking the "yes" or "no" button
        MyAnswer = MsgBox("This e-mail does not have any attachment." & vbNewLine & "Send anyway?", vbYesNo + vbQuestion, "No files attached to this e-mail")
        Select Case MyAnswer
            Case 6 'Yes
                'If the answer is "yes", then we send the e-mail after adding the salutations and signature
                Call InsertMyList(False)
            Case 7 'No
                'If the answer is "no", we go back to the e-mail
                Exit Sub
        End Select
    End If
    
End Sub

Public Sub InsertMyList(ByVal AddList As Boolean, Optional ListContent As String, Optional NumberAttachment As Integer)

    Dim OurEmail As Word.Document
    Dim CurPos As Long
    Dim i As Long
    Dim MyParCount As Long
    Dim MyTextRange As Range
    Dim MyLanguage As Long
    
    Set Em = Application.ActiveInspector.CurrentItem
    Em.Display
    Set OurEmail = Outlook.Application.ActiveInspector.WordEditor
    
    MyParCount = OurEmail.Paragraphs.Count
    
    'This section aims at checking if there already is a signature
    'If not checks if this is a reply or a forward message
    'In all cases, it assesses the number of paragraphs used in composing your message
    If OurEmail.Bookmarks.Exists("_MailAutoSig") = True Then
        CurPos = OurEmail.Bookmarks("_MailAutoSig").Start
        Set MyTextRange = OurEmail.Range(Start:=0, End:=CurPos)
        MyParCount = MyTextRange.Paragraphs.Count
    Else
        For i = 1 To MyParCount
            If InStr(1, OurEmail.Paragraphs(i), "From:") = 1 Or InStr(1, OurEmail.Paragraphs(i), "-----Original Message-----") = 1 Then
                'If we find such a marker, we change the value of MyParCount to 1 line before the previous exchange
                MyParCount = i - 1
                Exit For
            End If
        Next
    End If

    'This section removes the extra empty lines (if any) at the end of your text
    For i = MyParCount To 2 Step -1
        If Len(OurEmail.Paragraphs(i).Range.Text) = 1 Then
            MyParCount = MyParCount - 1
        Else
            If MyParCount = i Then
                'or adds an extra line if there is no empty line after your text and the beginning of the previous communication
                CurPos = OurEmail.Paragraphs(MyParCount).Range.End
                Set MyTextRange = OurEmail.Range(Start:=CurPos - 1, End:=CurPos - 1)
                MyTextRange = MyTextRange & vbNewLine
                MyParCount = MyParCount + 1
            End If
            Exit For
        End If
    Next

    'Now, we are defining our current composition as our range so that we can identify the language used
    Set MyTextRange = OurEmail.Paragraphs(1).Range
    MyTextRange.SetRange Start:=MyTextRange.Start, End:=OurEmail.Paragraphs(MyParCount).Range.End
    MyLanguage = MyTextRange.LanguageID
    
    'For a list of language codes, please refer to https://support.microsoft.com/en-us/kb/221435
    'If the option to prepare the list of attachments is selected, adds a localised line
    'In all cases adds salutations after
    If AddList = True And NumberAttachment <> 0 Then
        ListContent = StrContent(ListContent, NumberAttachment, MyLanguage) & vbNewLine & vbNewLine & AddSalutations(MyLanguage)
    Else
        ListContent = AddSalutations(MyLanguage)
    End If
    
    'Locates the cursor at the end of the composition and insert the additional text
    CurPos = OurEmail.Paragraphs(MyParCount).Range.End
    Set MyTextRange = OurEmail.Range(Start:=CurPos - 1, End:=CurPos - 1)
    MyTextRange = MyTextRange & vbNewLine & ListContent
    
    'This section opens the signature in Word to copy and paste it after your salutations
    'These definitions deal with the Word document to be opened
    If OurEmail.Bookmarks.Exists("_MailAutoSig") = False Then
    
        'After insertion, we locate the cursor at the end of the inserted text
        'We only need to do this if we want to insert a signature
        CurPos = CurPos + Len(MyTextRange) - 1
        Set MyTextRange = OurEmail.Range(Start:=CurPos, End:=CurPos)
    
        Dim appWord As Object
        Dim opWord As Boolean
        Dim MySigFile As Word.Document

        On Error Resume Next
        'Checks if Word is opened
        Set appWord = GetObject("Word.Application")
           opWord = False
        'If not, opens Word
        If appWord Is Nothing Then
           Set appWord = CreateObject("Word.Application")
           opWord = True
           Err.Clear
        End If
        appWord.Visible = False
        
        'Defines the signature file as MySigFile and opens it
        Set MySigFile = appWord.Documents.Open(WhichSignature(MyLanguage, Em.BodyFormat))
        'We select the entire content of the file and we copy it
        With MySigFile
            .Activate
            .Range.WholeStory
            .Range.Copy
        End With

        'We return to the e-mail
        Em.Display

        'We position the cursor at the end of our text (included what we previously added) and we paste the signature
        MyTextRange.Paste

        'Empty the clipboard and closes MySigFile without saving
        With MySigFile
            .Clipboard.Clear
            .Close
        End With

        'If Word is not already open, quit the application
        If opWord = True Then
            appWord.Quit
        End If

        appWord.Visible = True
        
        'We clear the memory
        Set appWord = Nothing
    End If
    
    Em.Display
    
    Em.Send

End Sub
