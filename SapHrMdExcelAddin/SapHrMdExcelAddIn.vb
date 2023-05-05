Public Class SapHrMdExcelAddIn

    Private Sub SapHrMdExcelAddIn_Startup() Handles Me.Startup
        log4net.Config.XmlConfigurator.Configure()
    End Sub

    Private Sub SapHrMdExcelAddIn_Shutdown() Handles Me.Shutdown

    End Sub

End Class
