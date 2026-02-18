Imports System.IO

Public Class EditMessage
    Private templatePath As String

    Public Sub New(templateFilePath As String)
        ' This call is required by the designer.
        InitializeComponent()

        ' Store the template path for use in this form
        templatePath = templateFilePath
    End Sub

    Private Sub EditMessage_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Load the template text into the TextBox
        Try
            If File.Exists(templatePath) Then
                txtMessage.Text = File.ReadAllText(templatePath)
            Else
                ' Create a default template if the file doesn't exist
                txtMessage.Text = "Your OTP verification code is: |otp-code|"
                File.WriteAllText(templatePath, txtMessage.Text)
            End If
        Catch ex As Exception
            MessageBox.Show($"Failed to load the message template: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        ' Check if the template includes |otp-code|
        If Not txtMessage.Text.Contains("|otp-code|") Then
            MessageBox.Show("The message template must include '|otp-code|' to insert the OTP code. Please update the template to include this placeholder.", "Invalid Message Template", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return ' Prevent saving if the placeholder is missing
        End If

        ' Save the template to the file
        Try
            File.WriteAllText(templatePath, txtMessage.Text)
            MessageBox.Show("Message template saved successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.DialogResult = DialogResult.OK
        Catch ex As Exception
            MessageBox.Show($"Failed to save the message template: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        ' Close the editor without saving
        Me.DialogResult = DialogResult.Cancel
    End Sub

    Private Sub txtMessage_TextChanged(sender As Object, e As EventArgs) Handles txtMessage.TextChanged

    End Sub
End Class
