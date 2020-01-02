Imports System.Collections.Specialized
Imports System.ComponentModel

Public Enum Books1
    Book_1
    Book_2
    Book_3
    Book_4
    Book_5
    Book_6
    Book_7
    Book_8
    Book_9
    Book_10
    Book_11
    Book_12
    Book_13
    Book_14
    Book_15
    Book_16
    Book_17
    Book_18
    Book_19
    Book_20
End Enum

Public Enum Books2
    Book_1
    Book_2
    Book_3
    Book_4
    Book_5
    Book_6
    Book_7
    Book_8
    Book_9
    Book_10
    Book_11
    Book_12
    Book_13
    Book_14
    Book_15
    Book_16
    Book_17
    Book_18
    Book_19
    Book_20
End Enum

Public Enum InputTrigger
    Enter
    Shift
    Space
    None
End Enum

Public Class ASTextBox
    Inherits TextBox
    Public savedList As New StringCollection()
    Public readedList As New StringCollection()
    Public SaveToBook As Boolean = True
    Public keyboardkey As Keys = Keys.Enter
    Public BookRead As Books1
    Public BookSave As Books2

    Public Sub New()
        With Me
            .AutoCompleteMode = AutoCompleteMode.Suggest
            .AutoCompleteSource = AutoCompleteSource.CustomSource
        End With
        CheckTrigger()
        CheckBookSave()
        CheckBookRead()
        InputTrigger = InputTrigger.Enter
        RefreshAutoComplete()
    End Sub

    Public Shared Sub saveSetts()
        My.Settings.Save()
    End Sub

    Private Sub CheckBookSave()
        'checks which book the control is set to. because this control's selected book can also be changed programatically.
        If BookSave = Books2.Book_1 Then
            savedList = My.Settings.ASTBcollection1
        ElseIf BookSave = Books2.Book_2 Then
            savedList = My.Settings.ASTBcollection2
        ElseIf BookSave = Books2.Book_3 Then
            savedList = My.Settings.ASTBcollection3
        ElseIf BookSave = Books2.Book_4 Then
            savedList = My.Settings.ASTBcollection4
        ElseIf BookSave = Books2.Book_5 Then
            savedList = My.Settings.ASTBcollection5
        ElseIf BookSave = Books2.Book_6 Then
            savedList = My.Settings.ASTBcollection6
        ElseIf BookSave = Books2.Book_7 Then
            savedList = My.Settings.ASTBcollection7
        ElseIf BookSave = Books2.Book_8 Then
            savedList = My.Settings.ASTBcollection8
        ElseIf BookSave = Books2.Book_9 Then
            savedList = My.Settings.ASTBcollection9
        ElseIf BookSave = Books2.Book_10 Then
            savedList = My.Settings.ASTBcollection10
        ElseIf BookSave = Books2.Book_11 Then
            savedList = My.Settings.ASTBcollection11
        ElseIf BookSave = Books2.Book_12 Then
            savedList = My.Settings.ASTBcollection12
        ElseIf BookSave = Books2.Book_13 Then
            savedList = My.Settings.ASTBcollection13
        ElseIf BookSave = Books2.Book_14 Then
            savedList = My.Settings.ASTBcollection14
        ElseIf BookSave = Books2.Book_15 Then
            savedList = My.Settings.ASTBcollection15
        ElseIf BookSave = Books2.Book_16 Then
            savedList = My.Settings.ASTBcollection16
        ElseIf BookSave = Books2.Book_17 Then
            savedList = My.Settings.ASTBcollection17
        ElseIf BookSave = Books2.Book_18 Then
            savedList = My.Settings.ASTBcollection18
        ElseIf BookSave = Books2.Book_19 Then
            savedList = My.Settings.ASTBcollection19
        ElseIf BookSave = Books2.Book_20 Then
            savedList = My.Settings.ASTBcollection20
        End If
    End Sub

    Private Sub CheckBookRead()
        'checks which book the control is set to. because this control's selected book can also be changed programatically.
        If BookRead = Books1.Book_1 Then
            readedList = My.Settings.ASTBcollection1
        ElseIf BookRead = Books1.Book_2 Then
            readedList = My.Settings.ASTBcollection2
        ElseIf BookRead = Books1.Book_3 Then
            readedList = My.Settings.ASTBcollection3
        ElseIf BookRead = Books1.Book_4 Then
            readedList = My.Settings.ASTBcollection4
        ElseIf BookRead = Books1.Book_5 Then
            readedList = My.Settings.ASTBcollection5
        ElseIf BookRead = Books1.Book_6 Then
            readedList = My.Settings.ASTBcollection6
        ElseIf BookRead = Books1.Book_7 Then
            readedList = My.Settings.ASTBcollection7
        ElseIf BookRead = Books1.Book_8 Then
            readedList = My.Settings.ASTBcollection8
        ElseIf BookRead = Books1.Book_9 Then
            readedList = My.Settings.ASTBcollection9
        ElseIf BookRead = Books1.Book_10 Then
            readedList = My.Settings.ASTBcollection10
        ElseIf BookRead = Books1.Book_11 Then
            readedList = My.Settings.ASTBcollection11
        ElseIf BookRead = Books1.Book_12 Then
            readedList = My.Settings.ASTBcollection12
        ElseIf BookRead = Books1.Book_13 Then
            readedList = My.Settings.ASTBcollection13
        ElseIf BookRead = Books1.Book_14 Then
            readedList = My.Settings.ASTBcollection14
        ElseIf BookRead = Books1.Book_15 Then
            readedList = My.Settings.ASTBcollection15
        ElseIf BookRead = Books1.Book_16 Then
            readedList = My.Settings.ASTBcollection16
        ElseIf BookRead = Books1.Book_17 Then
            readedList = My.Settings.ASTBcollection17
        ElseIf BookRead = Books1.Book_18 Then
            readedList = My.Settings.ASTBcollection18
        ElseIf BookRead = Books1.Book_19 Then
            readedList = My.Settings.ASTBcollection19
        ElseIf BookRead = Books1.Book_20 Then
            readedList = My.Settings.ASTBcollection20
        End If
    End Sub

    Private Sub CheckTrigger()
        If InputTrigger = InputTrigger.Enter Then
            keyboardkey = Keys.Enter
        ElseIf InputTrigger = InputTrigger.Shift Then
            keyboardkey = Keys.ShiftKey
        ElseIf InputTrigger = InputTrigger.Space Then
            keyboardkey = Keys.Space
        ElseIf InputTrigger = InputTrigger.None Then
            keyboardkey = Keys.Tab
        End If
    End Sub

    Private Sub RefreshAutoComplete()
        Me.AutoCompleteCustomSource.Clear()
        For Each item In readedList
            Me.AutoCompleteCustomSource.Add(item)
        Next
    End Sub

    Private Sub SaveToSettings()
        CheckTrigger()
        CheckBookSave()
        CheckBookRead()
        If savedList.Contains(Me.Text) Then
        Else
            savedList.Add(Me.Text)
            My.Settings.Save()
        End If
        RefreshAutoComplete()
    End Sub

    Private Sub AutoSavingTextBox_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If Me.Text = Nothing Then
        Else
            If SaveToBook = True Then
                If e.KeyCode = keyboardkey Then
                    SaveToSettings()
                End If
            End If
        End If

    End Sub

    Private Sub AutoSavingTextBox_GotFocus(sender As Object, e As EventArgs) Handles Me.GotFocus
        CheckTrigger()
        CheckBookSave()
        CheckBookRead()
        RefreshAutoComplete()
    End Sub

    <Category("Auto-Saving TextBox")>
    <DisplayName("SaveToBook")>
    <Description("The selected book this textbox will write to.")>
    Public Property SaveTo() As Books2
        Get
            Return BookSave
        End Get
        Set(ByVal value As Books2)
            BookSave = value
        End Set
    End Property

    <Category("Auto-Saving TextBox")>
    <DisplayName("ReadFromBook")>
    <Description("The selected book this textbox will read from.")>
    Public Property ReadFrom() As Books1
        Get
            Return BookRead
        End Get
        Set(ByVal value As Books1)
            BookRead = value
        End Set
    End Property

    <Category("Auto-Saving TextBox")>
    <DisplayName("Save")>
    <Description("When set to true, the textbox will save all entries to the selected book. When set to false, it will not. However, it can still read from the selected book and auto suggest from it.")>
    Public Property Save_toBook() As Boolean
        Get
            Return SaveToBook
        End Get
        Set(ByVal value As Boolean)
            SaveToBook = value
        End Set
    End Property

    <Category("Auto-Saving TextBox")>
    <DisplayName("InputTrigger")>
    <Description("The key in which will save the entry.")>
    Public Property InputTrigger() As InputTrigger
        Get
            Return keyboardkey
        End Get
        Set(ByVal value As InputTrigger)
            keyboardkey = value
        End Set
    End Property

End Class

