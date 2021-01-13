VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Fm_AttachmentList 
   Caption         =   "Check the file(s) attached to this email"
   ClientHeight    =   3705
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10260
   OleObjectBlob   =   "Fm_AttachmentList.frx":0000
End
Attribute VB_Name = "Fm_AttachmentList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************
' This section prepares the userform *
'*************************************
Public Sub UserForm_Initialize()

    'Positions the userform
    With Me
        .Left = 0
        .Top = 0
    End With

    'Refreshes the content of the list box named lb_Attachments
    Me.lb_Attachments.Clear
    Call UpdateAttachmentList
    
    'Sets the value to include the list of attachments in the email to true
    Me.Ch_AddListToEmail.value = True
    
    
End Sub

'******************************************
' This routine updates the list displayed *
'******************************************
Private Sub UpdateAttachmentList()


    Dim i As Integer
    
    'Defines Em as the current mail item selected
    Set Em = Application.ActiveInspector.CurrentItem
    
    'Creates the list to be contained in the list box using the name of all the files attached
    With Me
        For i = 1 To Em.Attachments.Count
            .lb_Attachments.AddItem Em.Attachments.Item(i).FileName
        Next
        
        'This updates the label caption for the number of attached files
        .Lbl_Result_Attach.Caption = Em.Attachments.Count
        'And the label for the number of items included in the list
        .Lbl_Result_List.Caption = Em.Attachments.Count
    End With
    
End Sub

'**********************************************************
' This routine deals with adding attachment to the e-mail *
'**********************************************************
Private Sub AddFiles_Click()

    'This is to check whether the selected file is already attached
    Dim AlreadyAttached As Boolean
    'This is to open a blank Word file
    Dim appWord As Object
    'This is to check whether Word is already opened
    Dim opWord As Boolean
    'This is to define a module to select files
    Dim FileSelect As FileDialog
    'This is to define which files are selected
    Dim ItemSelected As Variant
    'This is to count how many files have been selected
    Dim i As Long
    
    Set Em = Application.ActiveInspector.CurrentItem
    
    On Error Resume Next
    'Checks whether Word is opened
    Set appWord = GetObject(, "Word.Application")
       opWord = False
    'If not, opens Word
    If appWord Is Nothing Then
       Set appWord = CreateObject("Word.Application")
       opWord = True
       Err.Clear
    End If
    
    'This is to make the instance of Word get the focus so that you can select the files you want
    With appWord
        .Activate
        .Visible = True
    End With
    
    'Empties the list box
    Me.lb_Attachments.Clear
    
    'Opens a file dialog to pick files to add as attachments in Word as no file dialog is available in Outlook
    Word.ActiveWindow.SetFocus
    Set FileSelect = appWord.FileDialog(msoFileDialogFilePicker)
    With FileSelect
        'Gives a title to the opened dialog box
        .Title = "Select the files you want to attach to this email"
        'Accepts the selection of several files
        .AllowMultiSelect = True
        'Defines a starting point to look for files
        'Change the path to your preferred location
        .InitialFileName = "C:\"
        'Checks whether the user selected at least 1 file
        If .Show = True Then
            'For each selected file
            For Each ItemSelected In .SelectedItems
                AlreadyAttached = False
                'Checks whether this file is already attached
                For i = 1 To Em.Attachments.Count
                    If InStr(1, ItemSelected, "\" & Em.Attachments.Item(i).FileName) <> 0 Then
                        AlreadyAttached = True
                        Exit For
                    End If
                Next
                'If not then adds the file as attachment
                If AlreadyAttached = False Then
                    Em.Attachments.Add ItemSelected
                End If
            Next ItemSelected
        End If
    End With
    
    appWord.WindowState = wdWindowStateMinimize

    'Close the application just used
    If opWord = True Then
        appWord.Quit
    End If
    
    'Clear the name of all selected files in the file dialog process
    Set FileSelect = Nothing
    'Remove the Word file called
    Set appWord = Nothing

    'Update the list of attachments
    Call UpdateAttachmentList

End Sub

'************************************************************************
' This routine removes files from the list AND the selected attachments *
'************************************************************************
Private Sub RemSel_Click()


    Dim LC As Long, IL As Long
    
    Set Em = Application.ActiveInspector.CurrentItem

    'Goes through the list box and check which items are selected
    With Me.lb_Attachments
        For LC = .ListCount To 1 Step -1
            IL = LC - 1
            If .Selected(IL) = True Then
                'Removes the associated attachment
                Em.Attachments.Item(LC).Delete
                'Removes the selected line from the list
                .RemoveItem (IL)
            End If
        Next
        'Empties the list box
        .Clear
    End With
    

    'Updates the list of attachments
    Call UpdateAttachmentList
    
End Sub

'***************************************************************
' This routine removes files from the list, not the attachments *
'***************************************************************
Private Sub RemList_Click()

    Dim i As Long
    
    With Me.lb_Attachments
        'Loops through the list and remove selected items
        For i = .ListCount - 1 To 0 Step -1
            If .Selected(i) = True Then
                .RemoveItem (i)
            End If
        Next
        'Updates the labels for the number of items included in the list
        Me.Lbl_Result_List = .ListCount
    End With

End Sub

'*****************************************************************
' This section cancels sending and adding the list to the e-mail *
'*****************************************************************
Private Sub Btn_Cancel_Click()

    'hides the userform and returns to the e-mail
    Me.Hide

End Sub

'*************************************************************
' This routine sends the email based on the selected options *
'*************************************************************
Private Sub Btn_Send_Click()

    Dim MyList As String
    Dim i As Integer
    
    'Hiding the form
    Me.Hide
    
    'Emptying the list
    MyList = vbNullString
    
    'This section checks whether the option "Add to the list" was selected
    'If so, add each selected file name in the list in a new line
    'Then call the last routine to generate the end of the email
    With Me
        If .Ch_AddListToEmail.value = True Then
            For i = 1 To .lb_Attachments.ListCount
                MyList = MyList & "  -  " & .lb_Attachments.List(i - 1) & vbNewLine
            Next
            Call InsertMyList(True, MyList, Me.Lbl_Result_List.Caption)
        Else
            Call InsertMyList(False)
        End If
    End With

End Sub
