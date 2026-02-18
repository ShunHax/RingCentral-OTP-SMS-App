<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class EditMessage
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(EditMessage))
        lblEdtMsgDescription = New Label()
        GroupBox2 = New GroupBox()
        btnSave = New Button()
        btnCancel = New Button()
        txtMessage = New TextBox()
        Label1 = New Label()
        GroupBox2.SuspendLayout()
        SuspendLayout()
        ' 
        ' lblEdtMsgDescription
        ' 
        lblEdtMsgDescription.AccessibleName = ""
        lblEdtMsgDescription.BackColor = Color.Transparent
        lblEdtMsgDescription.ForeColor = Color.LightGray
        lblEdtMsgDescription.Location = New Point(12, 25)
        lblEdtMsgDescription.Name = "lblEdtMsgDescription"
        lblEdtMsgDescription.Size = New Size(444, 45)
        lblEdtMsgDescription.TabIndex = 13
        lblEdtMsgDescription.Text = "You can edit the template for the text message sent below. Use |otp-code| as the placeholder for the OTP code in the message. Changes are immediate, you do NOT need to reload the application."
        lblEdtMsgDescription.TextAlign = ContentAlignment.MiddleCenter
        ' 
        ' GroupBox2
        ' 
        GroupBox2.BackColor = Color.FromArgb(CByte(45), CByte(45), CByte(48))
        GroupBox2.Controls.Add(btnSave)
        GroupBox2.Controls.Add(btnCancel)
        GroupBox2.Controls.Add(txtMessage)
        GroupBox2.Dock = DockStyle.Bottom
        GroupBox2.ForeColor = Color.White
        GroupBox2.Location = New Point(0, 74)
        GroupBox2.Name = "GroupBox2"
        GroupBox2.Size = New Size(468, 203)
        GroupBox2.TabIndex = 15
        GroupBox2.TabStop = False
        GroupBox2.Text = "Texting"
        ' 
        ' btnSave
        ' 
        btnSave.BackColor = Color.FromArgb(CByte(60), CByte(60), CByte(60))
        btnSave.ForeColor = Color.White
        btnSave.Location = New Point(306, 168)
        btnSave.Name = "btnSave"
        btnSave.Size = New Size(98, 29)
        btnSave.TabIndex = 18
        btnSave.TabStop = False
        btnSave.Text = "Save"
        btnSave.UseVisualStyleBackColor = False
        ' 
        ' btnCancel
        ' 
        btnCancel.BackColor = Color.FromArgb(CByte(60), CByte(60), CByte(60))
        btnCancel.ForeColor = Color.White
        btnCancel.Location = New Point(61, 168)
        btnCancel.Name = "btnCancel"
        btnCancel.Size = New Size(98, 29)
        btnCancel.TabIndex = 17
        btnCancel.TabStop = False
        btnCancel.Text = "Cancel"
        btnCancel.UseVisualStyleBackColor = False
        ' 
        ' txtMessage
        ' 
        txtMessage.BackColor = SystemColors.Window
        txtMessage.Location = New Point(36, 22)
        txtMessage.Multiline = True
        txtMessage.Name = "txtMessage"
        txtMessage.Size = New Size(393, 140)
        txtMessage.TabIndex = 16
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Font = New Font("Segoe UI", 12F, FontStyle.Bold)
        Label1.ForeColor = Color.LightGray
        Label1.Location = New Point(169, 2)
        Label1.Name = "Label1"
        Label1.Size = New Size(126, 21)
        Label1.TabIndex = 16
        Label1.Text = "Message Editor"
        ' 
        ' EditMessage
        ' 
        AccessibleName = ""
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        BackColor = Color.FromArgb(CByte(30), CByte(30), CByte(30))
        ClientSize = New Size(468, 277)
        Controls.Add(Label1)
        Controls.Add(GroupBox2)
        Controls.Add(lblEdtMsgDescription)
        Icon = CType(resources.GetObject("$this.Icon"), Icon)
        Name = "EditMessage"
        Text = "Message Editor"
        GroupBox2.ResumeLayout(False)
        GroupBox2.PerformLayout()
        ResumeLayout(False)
        PerformLayout()
    End Sub
    Friend WithEvents lblEdtMsgDescription As Label
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents txtMessage As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents btnSave As Button
    Friend WithEvents btnCancel As Button
End Class
