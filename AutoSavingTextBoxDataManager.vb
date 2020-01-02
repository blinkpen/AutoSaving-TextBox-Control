Imports System.Collections.Specialized

Public Class AutoSavingTextBoxDataManager
    Dim SelectedList As New StringCollection()

    Private Sub AutoSavingTextBoxDataManager_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        For x = 1 To My.Settings.Properties.Count
            ComboBox1.Items.Add("Book_" & x)
        Next
        ComboBox1.SelectedIndex = 0
    End Sub

    Private Sub Command()
        ListBox1.Items.Clear()
        If ComboBox1.SelectedIndex = 0 Then
            SelectedList = My.Settings.ASTBcollection1
        ElseIf ComboBox1.SelectedIndex = 1 Then
            SelectedList = My.Settings.ASTBcollection2
        ElseIf ComboBox1.SelectedIndex = 2 Then
            SelectedList = My.Settings.ASTBcollection3
        ElseIf ComboBox1.SelectedIndex = 3 Then
            SelectedList = My.Settings.ASTBcollection4
        ElseIf ComboBox1.SelectedIndex = 4 Then
            SelectedList = My.Settings.ASTBcollection5
        ElseIf ComboBox1.SelectedIndex = 5 Then
            SelectedList = My.Settings.ASTBcollection6
        ElseIf ComboBox1.SelectedIndex = 6 Then
            SelectedList = My.Settings.ASTBcollection7
        ElseIf ComboBox1.SelectedIndex = 7 Then
            SelectedList = My.Settings.ASTBcollection8
        ElseIf ComboBox1.SelectedIndex = 8 Then
            SelectedList = My.Settings.ASTBcollection9
        ElseIf ComboBox1.SelectedIndex = 9 Then
            SelectedList = My.Settings.ASTBcollection10
        ElseIf ComboBox1.SelectedIndex = 10 Then
            SelectedList = My.Settings.ASTBcollection11
        ElseIf ComboBox1.SelectedIndex = 11 Then
            SelectedList = My.Settings.ASTBcollection12
        ElseIf ComboBox1.SelectedIndex = 12 Then
            SelectedList = My.Settings.ASTBcollection13
        ElseIf ComboBox1.SelectedIndex = 13 Then
            SelectedList = My.Settings.ASTBcollection14
        ElseIf ComboBox1.SelectedIndex = 14 Then
            SelectedList = My.Settings.ASTBcollection15
        ElseIf ComboBox1.SelectedIndex = 15 Then
            SelectedList = My.Settings.ASTBcollection16
        ElseIf ComboBox1.SelectedIndex = 16 Then
            SelectedList = My.Settings.ASTBcollection17
        ElseIf ComboBox1.SelectedIndex = 17 Then
            SelectedList = My.Settings.ASTBcollection18
        ElseIf ComboBox1.SelectedIndex = 18 Then
            SelectedList = My.Settings.ASTBcollection19
        ElseIf ComboBox1.SelectedIndex = 19 Then
            SelectedList = My.Settings.ASTBcollection20
        End If

        For Each item In SelectedList
            ListBox1.Items.Add(item)
        Next
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Command()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim whereOn As Integer = ComboBox1.SelectedIndex
        Dim result As DialogResult = MessageBox.Show("Are you sure you want to erase this book?", "AutoSaving TextBox Data Manager", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
        If result = DialogResult.Cancel Then
        ElseIf result = DialogResult.No Then
        ElseIf result = DialogResult.Yes Then
            SelectedList.Clear()
            My.Settings.Save()

            For x = 0 To ComboBox1.Items.Count - 1
                ComboBox1.SelectedIndex = x
            Next
            ComboBox1.SelectedIndex = whereOn

        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim result As DialogResult = MessageBox.Show("Are you sure you want to erase all books?", "AutoSaving TextBox Data Manager", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
        If result = DialogResult.Cancel Then
        ElseIf result = DialogResult.No Then
        ElseIf result = DialogResult.Yes Then
            Me.Enabled = False
            ComboBox1.SelectedIndex = 0
            For x = 0 To ComboBox1.Items.Count - 1
                ComboBox1.SelectedIndex = x
                SelectedList.Clear()
                My.Settings.Save()
            Next
            Me.Enabled = True
            ComboBox1.SelectedIndex = 0
        End If
    End Sub



    Private Sub AutoSavingTextBoxDataManager_MouseEnter(sender As Object, e As EventArgs) Handles Me.MouseEnter
        Dim selecteditem As Integer = ComboBox1.SelectedIndex
        For x = 0 To ComboBox1.Items.Count - 1
            ComboBox1.SelectedIndex = x
        Next
        ComboBox1.SelectedIndex = selecteditem
    End Sub
End Class
