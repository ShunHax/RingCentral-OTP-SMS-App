Public Class CustomPrompt
    Private _userResponse As DialogResult = DialogResult.None
    Public Property UserResponse As DialogResult
        Get
            Return _userResponse
        End Get
        Set(value As DialogResult) ' Changed from Private to Public
            _userResponse = value
        End Set
    End Property

    Public Sub New(message As String, title As String, iconImage As Image)
        InitializeComponent()

        ' Set the form title
        Me.Text = title

        ' Set the message and icon
        lblMessage.Text = message
        pbIcon.Image = iconImage
    End Sub

    Private Sub btnYes_Click(sender As Object, e As EventArgs) Handles btnYes.Click
        UserResponse = DialogResult.Yes
        Me.Close()
    End Sub

    Private Sub btnNo_Click(sender As Object, e As EventArgs) Handles btnNo.Click
        UserResponse = DialogResult.No
        Me.Close()
    End Sub

    Private Sub lblMessage_Click(sender As Object, e As EventArgs) Handles lblMessage.Click

    End Sub
End Class
