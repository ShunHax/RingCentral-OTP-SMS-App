Imports Microsoft.VisualBasic.ApplicationServices
Imports System.Threading

Namespace My
    ' The following events are available for MyApplication:
    ' Startup: Raised when the application starts, before the startup form is created.
    ' Shutdown: Raised after all application forms are closed.  This event is not raised if the application terminates abnormally.
    ' UnhandledException: Raised if the application encounters an unhandled exception.
    ' StartupNextInstance: Raised when launching a single-instance application and the application is already active. 
    ' NetworkAvailabilityChanged: Raised when the network connection is connected or disconnected.

    ' **NEW** ApplyApplicationDefaults: Raised when the application queries default values to be set for the application.

    ' Example:
    ' Private Sub MyApplication_ApplyApplicationDefaults(sender As Object, e As ApplyApplicationDefaultsEventArgs) Handles Me.ApplyApplicationDefaults
    '
    '   ' Setting the application-wide default Font:
    '   e.Font = New Font(FontFamily.GenericSansSerif, 12, FontStyle.Regular)
    '
    '   ' Setting the HighDpiMode for the Application:
    '   e.HighDpiMode = HighDpiMode.PerMonitorV2
    '
    '   ' If a splash dialog is used, this sets the minimum display time:
    '   e.MinimumSplashScreenDisplayTime = 4000
    ' End Sub

    Partial Friend Class MyApplication
        Private Shared appMutex As Mutex = Nothing

        Private Sub MyApplication_Startup(sender As Object, e As StartupEventArgs) Handles Me.Startup
            ' Create a unique mutex name for this application
            Dim mutexName As String = "OTP_Texting_App_SingleInstance_Mutex"
            Dim createdNew As Boolean = False

            Try
                ' Try to create a new mutex
                appMutex = New Mutex(True, mutexName, createdNew)

                ' If the mutex already exists, another instance is running
                If Not createdNew Then
                    MessageBox.Show("The OTP Texting App is already running. Please close the existing instance before opening a new one.", 
                                    "Application Already Running", 
                                    MessageBoxButtons.OK, 
                                    MessageBoxIcon.Warning)
                    e.Cancel = True ' Cancel the startup
                    Return
                End If
            Catch ex As Exception
                ' If there's an error creating the mutex, show error and exit
                MessageBox.Show($"An error occurred while checking for running instances: {ex.Message}", 
                                "Startup Error", 
                                MessageBoxButtons.OK, 
                                MessageBoxIcon.Error)
                e.Cancel = True
                Return
            End Try
        End Sub

        Private Sub MyApplication_Shutdown(sender As Object, e As EventArgs) Handles Me.Shutdown
            ' Release the mutex when the application shuts down
            If appMutex IsNot Nothing Then
                appMutex.ReleaseMutex()
                appMutex.Dispose()
                appMutex = Nothing
            End If
        End Sub
    End Class
End Namespace
