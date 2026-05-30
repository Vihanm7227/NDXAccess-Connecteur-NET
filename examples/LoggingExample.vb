' ============================================================================
' NDXAccess - Câblage du logging (Microsoft.Extensions.Logging)
' ============================================================================
' Le connecteur journalise le cycle de vie des connexions et les actions via
' ILogger. Il suffit de fournir un ILoggerFactory à la factory (ou à la connexion).
'
' NOTE : exemple de documentation (projet compilé).
' Auteur : Nicolas DEOUX <NDXDev@gmail.com>
' ============================================================================

Imports Microsoft.Extensions.Logging
Imports NDXAccess

Namespace NDXAccess.Examples

    Public Module LoggingExample

        ''' <summary>
        ''' Crée une factory de connexions journalisant en console (niveau Debug pour
        ''' voir les actions internes : Open/Close, transactions, AutoClose, etc.).
        ''' </summary>
        Public Function CreateFactoryWithConsoleLogging(databasePath As String) As IAccessConnectionFactory
            Dim lf = LoggerFactory.Create(
                Sub(builder)
                    builder.AddConsole()
                    builder.SetMinimumLevel(LogLevel.Debug)
                End Sub)

            Dim options As New AccessConnectionOptions With {.DatabasePath = databasePath}
            Return New AccessConnectionFactory(options, lf)
        End Function

        ''' <summary>
        ''' Variante : créer une connexion unique avec un logger fourni.
        ''' </summary>
        Public Function CreateConnectionWithLogging(databasePath As String, logger As ILogger(Of AccessConnection)) As IAccessConnection
            Dim options As New AccessConnectionOptions With {.DatabasePath = databasePath}
            Return New AccessConnection(options, logger)
        End Function

    End Module

End Namespace
