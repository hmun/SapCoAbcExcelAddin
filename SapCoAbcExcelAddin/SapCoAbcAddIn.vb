Public Class SapCoAbcAddIn

    Private Sub SapCoAbcAddIn_Startup() Handles Me.Startup
        log4net.Config.XmlConfigurator.Configure()
    End Sub

    Private Sub SapCoAbcAddIn_Shutdown() Handles Me.Shutdown

    End Sub

End Class
