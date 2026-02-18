Imports RingCentral
Imports System.Text
Imports System.Net.Http
Imports System.Net
Imports Newtonsoft.Json.Linq
Imports System.IO
Imports System.Reflection

Public Class Form1

    ' Timer to regenerate OTP every 30 seconds
    Private WithEvents otpTimer As New Timer() With {.Interval = 1000}
    Private countdown As Integer = 30 ' Countdown starts from 30 seconds

    ' Timer to check token expiration
    Private WithEvents tokenRefreshTimer As New Timer() With {.Interval = 60000} ' 60 seconds

    ' Timer to update token countdown every second
    Private WithEvents tokenCountdownTimer As New Timer() With {.Interval = 1000} ' 1 second

    ' Variable to hold the access token in memory
    Private accessToken As String = String.Empty
    Private tokenExpiration As DateTime

    ' Check if the Reauthorization prompt is already visible
    Private isReauthorizationPromptVisible As Boolean = False

    ' Create the message file for the OTP text
    Private Const TEMP_FOLDER As String = "OTPTextingApp"
    Private Const MESSAGE_TEMPLATE_FILE As String = "message_template.txt"

    ' Track if the user has declined reauthorization
    Private userDeclinedReauthorization As Boolean = False

    ' Track if the full-expiration prompt has been shown
    Private isFullExpirationPromptVisible As Boolean = False

    ' Reference to the reauthorization dialog
    Private reauthorizationDialog As CustomPrompt = Nothing

    ' Easter egg click counter
    Private easterEggClickCount As Integer = 0
    Private Const EASTER_EGG_CLICKS_REQUIRED As Integer = 5

    Private listener As HttpListener

    ' Start the HTTP listener for the redirect URI
    Private Sub StartHttpListener()
        Try
            ' Dispose the existing listener if it's not null
            listener?.Close()
            listener = Nothing

            ' Create a new HttpListener instance
            listener = New HttpListener()
            listener.Prefixes.Add("http://localhost:5000/oauth/") ' Match your redirect_uri
            listener.Start()

            Log("HTTP Listener started on http://localhost:5000/oauth/")
            ListenForAuthorizationCode()
        Catch ex As Exception
            MessageBox.Show($"Failed to start HTTP server: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Log($"HTTP Listener Error: {ex.Message}")
        End Try
    End Sub

    ' Stop the HTTP listener
    Private Sub StopHttpListener()
        If listener?.IsListening Then
            listener.Stop()
            listener.Close()
            Log("HTTP Listener stopped.")
        End If
    End Sub

    ' Async method to listen for the authorization code
    Private Async Sub ListenForAuthorizationCode()
        Try
            While listener.IsListening
                Dim context = Await listener.GetContextAsync()
                Dim request = context.Request

                If request.QueryString("code") IsNot Nothing Then
                    Dim authCode As String = request.QueryString("code")
                    Log($"Authorization code received: {authCode}")

                    Using writer = New StreamWriter(context.Response.OutputStream)
                        writer.Write("Authorization successful! You may close this browser window.")
                    End Using

                    StopHttpListener()

                    ' Pass the authorization code to process
                    Await ProcessAuthorizationCode(authCode)

                    ' Close the AuthForm if still open
                    Application.OpenForms.OfType(Of AuthForm).FirstOrDefault()?.Close()
                Else
                    Using writer = New StreamWriter(context.Response.OutputStream)
                        writer.Write("No authorization code found.")
                    End Using
                End If
            End While
        Catch ex As Exception
            Log($"Error in HTTP listener: {ex.Message}")
        End Try
    End Sub

    Private Sub EnsureMessageTemplateExists()
        Dim tempPath As String = Path.Combine(Path.GetTempPath(), TEMP_FOLDER)
        Dim templatePath As String = Path.Combine(tempPath, MESSAGE_TEMPLATE_FILE)

        ' Ensure the folder exists
        Directory.CreateDirectory(tempPath)

        ' Check if the file exists, and create a default one if not
        If Not File.Exists(templatePath) Then
            File.WriteAllText(templatePath, "Your OTP verification code is: |otp-code|")
        End If
    End Sub

    Private Sub UpdateSendOtpButtonState()
        btnSendOtp.Enabled = Not String.IsNullOrEmpty(accessToken) AndAlso tokenExpiration > DateTime.Now
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        StartHttpListener() ' Start the HTTP listener for the redirect URI

        ' Set initial authorization status to Not Authorized with red background
        lblAuthorized.Text = "Not Authorized"
        lblAuthorized.BackColor = Color.Red

        ' Extract WebView2Loader.dll from resources
        ExtractWebView2Loader()

        ' Start OTP timer and generate first OTP immediately
        otpTimer.Start()
        lblCountdown.Text = "Next OTP in:" & Environment.NewLine & countdown & " seconds"
        GenerateAndDisplayOtp()

        ' Initialize progress bar for countdown
        pbCountdown.Minimum = 0
        pbCountdown.Maximum = 30
        pbCountdown.Value = countdown

        ' Start token refresh and countdown timers
        tokenRefreshTimer.Start()
        tokenCountdownTimer.Start()

        ' Verify the message template file exists
        EnsureMessageTemplateExists()

        ' Set up tooltip for lblTokenCounter
        Dim tokenTooltip As New ToolTip()
        tokenTooltip.SetToolTip(lblTokenCounter, "This shows how long you can use the app before having to sign in again." & Environment.NewLine & "The app will prompt you to re-authorize before it expires.")
    End Sub

    Private Sub ExtractWebView2Loader()
        Try
            ' Get the executing assembly
            Dim assembly As Assembly = Assembly.GetExecutingAssembly()

            ' Extract WebView2Loader.dll from the resources
            Using resourceStream = assembly.GetManifestResourceStream("OTP_Texting_App.WebView2Loader.dll")
                If resourceStream IsNot Nothing Then
                    Dim tempPath As String = Path.Combine(Path.GetTempPath(), "OTPTextingApp")
                    Dim dllPath As String = Path.Combine(tempPath, "WebView2Loader.dll")
                    Directory.CreateDirectory(tempPath)
                    Using fileStream = New FileStream(dllPath, FileMode.Create, FileAccess.Write)
                        resourceStream.CopyTo(fileStream)
                    End Using

                    ' Add the temporary path to the DLL search path
                    Environment.SetEnvironmentVariable("PATH", Environment.GetEnvironmentVariable("PATH") & ";" & tempPath)
                Else
                    MessageBox.Show("Failed to load WebView2Loader. Please ensure all dependencies are available.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            End Using
        Catch ex As Exception
            MessageBox.Show("Failed to load WebView2Loader. Please ensure all dependencies are available." & vbCrLf & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub otpTimer_Tick(sender As Object, e As EventArgs) Handles otpTimer.Tick
        If countdown > 0 Then
            countdown -= 1
            lblCountdown.Text = "Next OTP in:" & Environment.NewLine & countdown & " seconds"
            pbCountdown.Value = countdown ' Update progress bar
        Else
            GenerateAndDisplayOtp()
            countdown = 30 ' Reset countdown
            lblCountdown.Text = "Next OTP in:" & Environment.NewLine & countdown & " seconds"
            pbCountdown.Value = countdown ' Reset progress bar
        End If
    End Sub

    ' Timer tick to check token expiration and prompt for re-authorization
    Private Sub tokenRefreshTimer_Tick(sender As Object, e As EventArgs) Handles tokenRefreshTimer.Tick
        If tokenExpiration <> DateTime.MinValue Then
            Dim timeRemaining = tokenExpiration - DateTime.Now
            Log($"tokenRefreshTimer_Tick triggered. Current Time: {DateTime.Now}, tokenExpiration: {tokenExpiration}")
            Log($"Time Remaining: {timeRemaining.TotalSeconds} seconds")

            ' Stop the timer if the token has expired
            If timeRemaining.TotalSeconds <= 0 Then
                Log("Token has expired. Reauthorization is required.")
                lblTokenCounter.Text = "Token has expired. Please re-authorize."
                lblAuthorized.Text = "Authorization Expired"
                lblAuthorized.BackColor = Color.Red
                tokenRefreshTimer.Stop()
            End If
        End If
    End Sub

    Private Sub tokenCountdownTimer_Tick(sender As Object, e As EventArgs) Handles tokenCountdownTimer.Tick
        If tokenExpiration <> DateTime.MinValue Then
            Dim timeRemaining = tokenExpiration - DateTime.Now

            If timeRemaining.TotalSeconds > 0 Then
                ' Update the token expiration label
                lblTokenCounter.Text = $"Token expires in: {timeRemaining.Minutes}m {timeRemaining.Seconds}s."

                ' Change lblAuthorized color based on time remaining
                If timeRemaining.TotalSeconds <= 300 Then ' Less than or equal to 5 minutes
                    lblAuthorized.Text = "Authorization Expiring Soon"
                    lblAuthorized.BackColor = Color.Yellow

                    ' Show the reauthorization prompt if not already visible and user hasn't declined
                    If Not userDeclinedReauthorization AndAlso (reauthorizationDialog Is Nothing OrElse Not reauthorizationDialog.Visible) Then
                        Log("Token is nearing expiration. Prompting for reauthorization.")
                        reauthorizationDialog = New CustomPrompt(
                            "Your session will expire in 5 minutes. To continue after that, you must re-authorize. Re-authorize now?",
                            "Re-Authorization Required", ' Set the window title
                            My.Resources.WarningIcon ' Warning icon
                        )
                        AddHandler reauthorizationDialog.FormClosed, AddressOf ReauthorizationDialog_Closed
                        reauthorizationDialog.Show(Me)
                    End If

                Else
                    lblAuthorized.Text = "Authorized"
                    lblAuthorized.BackColor = Color.Lime
                End If
            Else
                ' Token has expired
                lblTokenCounter.Text = "Token has expired. Please re-authorize."
                lblAuthorized.Text = "Authorization Expired"
                lblAuthorized.BackColor = Color.Red

                ' Ensure the reauthorization dialog is closed
                If reauthorizationDialog IsNot Nothing AndAlso reauthorizationDialog.Visible Then
                    Log("Closing reauthorization prompt before showing the expiration prompt.")
                    reauthorizationDialog.UserResponse = DialogResult.No ' Programmatically select No
                    reauthorizationDialog.Close()
                    reauthorizationDialog = Nothing
                    userDeclinedReauthorization = True ' Treat this as a declined reauthorization
                End If

                ' Show the full-expiration prompt
                If Not isFullExpirationPromptVisible Then
                    isFullExpirationPromptVisible = True
                    Log("Token has fully expired. Prompting for reauthorization.")

                    Dim expirationDialog As New CustomPrompt(
                    "Your authorization has expired. To continue using the app, you must re-authorize. Would you like to re-authorize now?",
                    "Authorization Expired", ' Set the window title
                    My.Resources.ErrorIcon ' Error icon
                    )
                    expirationDialog.ShowDialog(Me)

                    If expirationDialog.UserResponse = DialogResult.Yes Then
                        btnAuthorize.PerformClick() ' Trigger reauthorization
                    Else
                        Log("User declined full reauthorization. Application will remain unauthorized.")
                    End If

                    ' Reset the flag after the prompt is closed
                    isFullExpirationPromptVisible = False
                    Log("Full-expiration prompt closed.")
                End If

                ' Stop the countdown timer
                tokenCountdownTimer.Stop()
            End If
        End If
    End Sub

    Private Sub ReauthorizationDialog_Closed(sender As Object, e As EventArgs)
        If reauthorizationDialog IsNot Nothing Then
            ' If the dialog is closed programmatically, treat it as a "No" response
            If reauthorizationDialog.UserResponse <> DialogResult.Yes Then
                userDeclinedReauthorization = True
                Log("Reauthorization dialog closed without reauthorization. Assuming user declined.")
            End If

            ' Clear the dialog reference
            reauthorizationDialog = Nothing
        End If
    End Sub

    ' Check for token expiration and prompt for re-authentication if expired.
    Private Sub HandleTokenExpiration()
        Log("Token has expired. Redirecting user for reauthorization.")
        MessageBox.Show("Your session has expired. Please log in again.", "Session Expired", MessageBoxButtons.OK, MessageBoxIcon.Error)
        btnAuthorize.PerformClick() ' Trigger reauthorization
    End Sub

    ' Method to generate and display OTP
    Private Sub GenerateAndDisplayOtp()
        Dim otpCode As String = GenerateOtp()
        txtCurrentOTP.Text = otpCode
    End Sub

    Private Async Sub btnSendOtp_Click(sender As Object, e As EventArgs) Handles btnSendOtp.Click
        ' Check if the user is authorized
        If String.IsNullOrEmpty(accessToken) Then
            MessageBox.Show("Please authorize with RingCentral first.", "Authorization Required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Check if the token is expired
        If tokenExpiration <= DateTime.Now Then
            MessageBox.Show("Your session has expired. Please reauthorize to send OTPs.", "Authorization Expired", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Get current OTP from the display box
        Dim otpCode = txtCurrentOTP.Text
        Dim phoneNumber = txtToPhoneNumber.Text
        Dim fromPhoneNumber = lblFromPhoneNumber.Text

        ' Ensure the phone number contains exactly 10 digits and is a valid U.S. phone number
        Dim cleanedPhoneNumber = New String(phoneNumber.Where(Function(c) Char.IsDigit(c)).ToArray)

        ' Validate phone number length and prevent invalid numbers like '555'
        If cleanedPhoneNumber.Length <> 10 OrElse cleanedPhoneNumber.StartsWith("555") Then
            MessageBox.Show("Please enter a valid 10-digit U.S. phone number.", "Invalid Number", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        If String.IsNullOrWhiteSpace(fromPhoneNumber) Then
            MessageBox.Show("From Phone Number is missing.", "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Try
            ' Attempt to send the OTP code
            Await SendOtpAsync(cleanedPhoneNumber, fromPhoneNumber, otpCode)

            ' Update the last sent OTP, phone number, and last sent time only if OTP was sent successfully
            lblOtpSent.Text = "Last OTP Sent: " & otpCode
            lblSentTo.Text = "Sent To: " & phoneNumber
            lblLastSent.Text = Date.Now.ToString("g") ' "g" for general date/time format

            ' Show success message and start hide timer
            lblSuccessMessage.Text = "OTP Code successfully sent!"
            lblSuccessMessage.Visible = True
            hideSuccessMessageTimer.Start()
        Catch ex As Exception
            ' Display a simple error message
            MessageBox.Show("Check the phone number and try again.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub btnAuthorize_Click(sender As Object, e As EventArgs) Handles btnAuthorize.Click
        ' Restart the HTTP listener if it is not running
        If listener Is Nothing OrElse Not listener.IsListening Then
            StartHttpListener()
        End If

        ' Open the AuthForm for authorization
        Dim authForm As New AuthForm(Me)
        authForm.ShowDialog() ' Show the AuthForm as a modal dialog
    End Sub


    ' Method to send OTP using RingCentral with specified from phone number
    Private Async Function SendOtpAsync(toPhoneNumber As String, fromPhoneNumber As String, otpCode As String) As Task
        Dim rc As New RestClient(AppConstants.CLIENT_ID, AppConstants.BASE_URL)

        If String.IsNullOrEmpty(accessToken) Then
            MessageBox.Show("Your session has expired. Please log in again to send the OTP.", "Session Expired", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        rc.token = New TokenInfo With {.access_token = accessToken}

        Try
            Dim tempPath As String = Path.Combine(Path.GetTempPath(), TEMP_FOLDER)
            Dim templatePath As String = Path.Combine(tempPath, MESSAGE_TEMPLATE_FILE)
            Dim messageTemplate As String = File.ReadAllText(templatePath)

            Dim messageText As String = messageTemplate.Replace("|otp-code|", otpCode)

            Dim parameters As New CreateSMSMessage() With {
            .from = New MessageStoreCallerInfoRequest() With {.phoneNumber = $"+1{fromPhoneNumber}"},
            .to = {New MessageStoreCallerInfoRequest() With {.phoneNumber = $"+1{toPhoneNumber}"}},
            .text = messageText
        }

            Await rc.Restapi().Account().Extension().Sms().Post(parameters)
        Catch ex As HttpRequestException
            Log($"Network error: {ex.Message}")
            MessageBox.Show("A network error occurred while sending the OTP. Please check your connection.", "Network Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Catch ex As UnauthorizedAccessException
            Log($"Invalid token: {ex.Message}")
            MessageBox.Show("Invalid or expired token. Please log in again.", "Authorization Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Catch ex As Exception
            Log($"Unexpected error: {ex.Message}")
            MessageBox.Show("An unexpected error occurred while sending the OTP.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function



    ' Method to generate a 6-digit OTP
    Private Function GenerateOtp() As String
        Dim random As New Random()
        Return random.Next(100000, 999999).ToString()
    End Function

    ' Logging method
    Private Sub Log(message As String)
        Try
            ' Redirect log file to the %temp% directory
            Dim tempPath As String = Path.Combine(Path.GetTempPath(), "OTPTextingApp")
            Directory.CreateDirectory(tempPath) ' Ensure the directory exists

            Dim logFilePath As String = Path.Combine(tempPath, "application_log.txt")
            File.AppendAllText(logFilePath, $"{DateTime.Now}: {message}{Environment.NewLine}")
        Catch ex As Exception
            ' Handle logging failure silently
        End Try
    End Sub

    ' Edit the message template file
    Private Sub btnEditMessage_Click(sender As Object, e As EventArgs) Handles btnEditMessage.Click
        Dim tempPath As String = Path.Combine(Path.GetTempPath(), TEMP_FOLDER)
        Dim templatePath As String = Path.Combine(tempPath, MESSAGE_TEMPLATE_FILE)

        ' Open the editor form as a modal dialog
        Dim editor As New EditMessage(templatePath)
        If editor.ShowDialog() = DialogResult.OK Then
        End If
    End Sub



    ' Method to set access token (called from AuthForm)
    Public Async Sub SetAccessToken(token As String, expiration As DateTime)
        accessToken = token
        tokenExpiration = expiration

        ' Reset reauthorization and expiration flags
        userDeclinedReauthorization = False
        isReauthorizationPromptVisible = False
        isFullExpirationPromptVisible = False

        ' Update authorization status in the UI
        lblAuthorized.Text = "Authorized"
        lblAuthorized.BackColor = Color.Lime

        Log($"Access token set. Expiration: {tokenExpiration}, Current Time: {DateTime.Now}")

        ' Start token expiration countdown if not already running
        If Not tokenCountdownTimer.Enabled Then
            tokenCountdownTimer.Start()
        End If

        ' Retrieve and update the user's phone number
        Await GetUserPhoneNumberAsync()
    End Sub
    ' Process the received authorization code
    Private Async Function ProcessAuthorizationCode(authCode As String) As Task
        Try
            ' Exchange the authorization code for tokens
            Dim authForm As New AuthForm(Me)
            Dim tokenResponse = Await authForm.ExchangeCodeForTokens(authCode)

            If tokenResponse IsNot Nothing Then
                ' Set the access token and expiration in Form1
                SetAccessToken(tokenResponse.AccessToken, tokenResponse.Expiration)

                ' Log the success
                Log("Authorization successful. Access token obtained.")

                ' Close the AuthForm if it's still open
                Application.OpenForms.OfType(Of AuthForm).FirstOrDefault()?.Close()
            Else
                Log("Failed to exchange authorization code for tokens.")
            End If
        Catch ex As Exception
            ' Log and handle the error
            Log($"Error processing authorization code: {ex.Message}")
        End Try
    End Function


    ' Method to get the user's phone number after authorization
    Private Async Function GetUserPhoneNumberAsync() As Task
        Using client As New HttpClient()
            client.DefaultRequestHeaders.Authorization = New Headers.AuthenticationHeaderValue("Bearer", accessToken)
            Dim response = Await client.GetAsync("https://platform.ringcentral.com/restapi/v1.0/account/~/extension/~/phone-number")

            If response.IsSuccessStatusCode Then
                Dim json = Await response.Content.ReadAsStringAsync()
                Dim jsonObj = JObject.Parse(json)
                Dim phoneNumber As String = jsonObj("records")(0)("phoneNumber").ToString()

                ' Remove the +1 country code if present
                If phoneNumber.StartsWith("+1") Then
                    phoneNumber = phoneNumber.Substring(2)
                End If

                ' Format the phone number as (###) ###-####
                If phoneNumber.Length = 10 Then
                    phoneNumber = String.Format("({0}) {1}-{2}",
                                                phoneNumber.Substring(0, 3),
                                                phoneNumber.Substring(3, 3),
                                                phoneNumber.Substring(6))
                End If

                ' Set the formatted phone number in the From Phone Number field
                lblFromPhoneNumber.Text = phoneNumber
            Else
                MessageBox.Show("Failed to retrieve phone number. Please check your account settings.")
            End If
        End Using
    End Function

    Private Sub lblDescription_Click(sender As Object, e As EventArgs) Handles lblDescription.Click

    End Sub

    Private Sub lblCreator_Click(sender As Object, e As EventArgs) Handles lblCreator.Click
        easterEggClickCount += 1

        If easterEggClickCount >= EASTER_EGG_CLICKS_REQUIRED Then
            easterEggClickCount = 0 ' Reset counter
            Dim easterEgg As New EasterEggForm()
            easterEgg.ShowDialog(Me)
        End If
    End Sub
End Class
