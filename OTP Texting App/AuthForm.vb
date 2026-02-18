Imports RingCentral
Imports System.Net.Http
Imports Newtonsoft.Json.Linq
Imports System.Threading.Tasks
Imports System.IO
Imports Microsoft.Web.WebView2.Core
Imports System.Net
Imports System.Security.Cryptography
Imports System.Text

Public Class AuthForm
    Private ReadOnly parentFormReference As Form1 ' Use ReadOnly for immutability

    Public Sub New(parent As Form1)
        InitializeComponent()
        parentFormReference = parent
    End Sub

    ' Logging method
    Private Sub Log(message As String)
        Try
            ' Redirect the log file to a temp directory
            Dim tempPath As String = Path.Combine(Path.GetTempPath(), "OTPTextingApp")
            Directory.CreateDirectory(tempPath) ' Ensure the directory exists

            Dim logFilePath As String = Path.Combine(tempPath, "application_log.txt")
            File.AppendAllText(logFilePath, $"{DateTime.Now}: {message}{Environment.NewLine}")
        Catch ex As Exception
            ' Handle logging failure
        End Try
    End Sub

    Private Async Sub AuthForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Log("AuthForm Load Started")

            ' Generate code verifier and code challenge for PKCE
            GenerateCodeVerifierAndChallenge()

            ' Define WebView2 environment for authentication
            Dim tempPath As String = Path.Combine(Path.GetTempPath(), "OTPTextingApp", "WebView2_UserData")
            Directory.CreateDirectory(tempPath)
            Dim env = Await CoreWebView2Environment.CreateAsync(Nothing, tempPath)
            Await wvAuth.EnsureCoreWebView2Async(env)

            ' Construct the authorization URL with PKCE parameters
            Dim authorizationUrl = $"{AppConstants.BASE_URL}/restapi/oauth/authorize" &
                                   $"?response_type=code" &
                                   $"&client_id={AppConstants.CLIENT_ID}" &
                                   $"&redirect_uri={Uri.EscapeDataString(AppConstants.REDIRECT_URI)}" &
                                   $"&state=1234" &
                                   $"&code_challenge={codeChallenge}" &
                                   $"&code_challenge_method=S256"
            Log($"Authorization URL: {authorizationUrl}")

            ' Load the authorization URL in WebView2
            wvAuth.Source = New Uri(authorizationUrl)
        Catch ex As Exception
            MessageBox.Show("Error during form load: " & ex.Message)
            Log("Error during form load: " & ex.Message)
        End Try
    End Sub

    Private Sub wvAuth_NavigationStarting(sender As Object, e As CoreWebView2NavigationStartingEventArgs) Handles wvAuth.NavigationStarting
        Try
            Log($"Navigation Starting: {e.Uri}")
            Dim currentUrl As String = e.Uri

            ' Check if the URL contains the access token in the fragment
            If currentUrl.Contains("#access_token=") Then
                e.Cancel = True ' Stop navigation to avoid redirecting further
                Log("Access token found in URL fragment.")

                ' Extract the access token and expiration from the URL fragment
                Dim fragment = New Uri(currentUrl).Fragment
                Dim queryParams = System.Web.HttpUtility.ParseQueryString(fragment.TrimStart("#"c))
                Dim accessToken = queryParams("access_token")
                Dim expiresIn As Integer = If(Integer.TryParse(queryParams("expires_in"), expiresIn), expiresIn, 3600)

                If Not String.IsNullOrEmpty(accessToken) Then
                    Log("Access token successfully retrieved.")

                    ' Pass the token to the parent form
                    parentFormReference.SetAccessToken(accessToken, DateTime.Now.AddSeconds(expiresIn))

                    ' Close the AuthForm automatically
                    Me.Invoke(Sub() Me.Close())
                Else
                    MessageBox.Show("Failed to retrieve access token.")
                    Log("Access token missing in URL fragment.")
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("An error occurred during navigation: " & ex.Message)
            Log("An error occurred during navigation: " & ex.Message)
        End Try
    End Sub

    Private Shared codeVerifier As String
    Private Shared codeChallenge As String

    Private Sub GenerateCodeVerifierAndChallenge()
        Dim randomBytes(31) As Byte
        Using rng As New RNGCryptoServiceProvider()
            rng.GetBytes(randomBytes)
        End Using

        ' Generate code verifier (base64 URL-safe string)
        codeVerifier = Convert.ToBase64String(randomBytes).Replace("+", "-").Replace("/", "_").Replace("=", "")

        ' Generate code challenge (SHA-256 hash of the verifier, then base64 URL-safe string)
        Using sha256 As SHA256 = SHA256.Create()
            Dim hash = sha256.ComputeHash(Encoding.UTF8.GetBytes(codeVerifier))
            codeChallenge = Convert.ToBase64String(hash).Replace("+", "-").Replace("/", "_").Replace("=", "")
        End Using
    End Sub

    ' Method to exchange authorization code for access token
    Friend Async Function ExchangeCodeForTokens(authCode As String) As Task(Of TokenResponse)
        Try
            Using client As New HttpClient()
                Dim tokenEndpoint = $"{AppConstants.BASE_URL}/restapi/oauth/token"
                Dim parameters = New List(Of KeyValuePair(Of String, String)) From {
                    New KeyValuePair(Of String, String)("grant_type", "authorization_code"),
                    New KeyValuePair(Of String, String)("code", authCode),
                    New KeyValuePair(Of String, String)("redirect_uri", AppConstants.REDIRECT_URI),
                    New KeyValuePair(Of String, String)("client_id", AppConstants.CLIENT_ID),
                    New KeyValuePair(Of String, String)("code_verifier", codeVerifier) ' PKCE Code Verifier
                }

                Dim content = New FormUrlEncodedContent(parameters)
                Dim response = Await client.PostAsync(tokenEndpoint, content)
                Dim json = Await response.Content.ReadAsStringAsync()

                If response.IsSuccessStatusCode Then
                    Dim jsonObj = JObject.Parse(json)
                    If jsonObj("access_token") IsNot Nothing Then
                        Dim expiresIn As Integer = If(Integer.TryParse(jsonObj("expires_in")?.ToString(), expiresIn), expiresIn, 3600)
                        Return New TokenResponse With {
                            .AccessToken = jsonObj("access_token").ToString(),
                            .Expiration = DateTime.Now.AddSeconds(expiresIn)
                        }
                    Else
                        Log($"Error exchanging authorization code: {json}")
                        MessageBox.Show("An error occurred while retrieving the access token. Please try again.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                Else
                    Log($"Error response from server: {response.StatusCode} - {json}")
                    Select Case response.StatusCode
                        Case HttpStatusCode.BadRequest
                            MessageBox.Show("Invalid request. Please check your input or try again.", "Bad Request", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Case HttpStatusCode.Unauthorized
                            MessageBox.Show("Authorization failed. Please check your credentials and try again.", "Unauthorized", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Case Else
                            MessageBox.Show("An unexpected error occurred. Please try again later.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Select
                End If
            End Using
        Catch ex As HttpRequestException
            Log($"Network error: {ex.Message}")
            MessageBox.Show("A network error occurred. Please check your internet connection and try again.", "Network Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Catch ex As Exception
            Log($"Unexpected error: {ex.Message}")
            MessageBox.Show("An unexpected error occurred. Please try again later.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return Nothing
    End Function

    Private Sub wvAuth_Click(sender As Object, e As EventArgs) Handles wvAuth.Click
        ' No implementation needed
    End Sub
End Class
