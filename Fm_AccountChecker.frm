VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Fm_AccountChecker 
   Caption         =   "Select the account to send your email"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3390
   OleObjectBlob   =   "Fm_AccountChecker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Fm_AccountChecker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()

    Dim Acc As Outlook.accounts
    Dim i As Integer
    
    'Positions the userform
    With Me
        .Left = 0
        .Top = 0
    End With

    Set Acc = Outlook.Session.accounts
    
    'identifies the list of avaialble accounts in the Outlook session
    For i = 1 To Acc.Count
        Me.Lst_Accounts.AddItem Acc.Item(i).DisplayName
    Next
    
End Sub

Private Sub Use_Account_Click()

    Dim Acc As Outlook.accounts
    Dim i As Integer
    Dim AccSelected As Boolean
    
    Dim MyAddress As AddressEntry
    
    Dim Em As Outlook.MailItem
    Set Acc = Outlook.Session.accounts
    Set Em = Application.ActiveInspector.CurrentItem
    
    AccSelected = False
    'Checks if an account was selected
    'If not ask the user to provide the data
    With Me
        For i = 1 To .Lst_Accounts.ListCount
            If .Lst_Accounts.Selected(i - 1) = True Then
                Set MyAddress = Session.accounts.Item(i).CurrentUser.AddressEntry
                Em.Sender = MyAddress
                AccSelected = True
                Exit For
            End If
        Next
        If AccSelected = True Then
            Me.Hide
        Else
            MsgBox "No account selected"
        End If
    End With
    
End Sub


