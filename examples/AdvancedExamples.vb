' ============================================================================
' NDXAccess - Exemples avancés
' ============================================================================
' Health checks, injection de dépendances, et limites de concurrence d'Access.
'
' NOTE : exemples de documentation, non compilés par les tests.
' Auteur : Nicolas DEOUX <NDXDev@gmail.com>
' ============================================================================

Imports System.Data
Imports Microsoft.Extensions.DependencyInjection
Imports Microsoft.Extensions.Logging
Imports NDXAccess
Imports NDXAccess.Extensions

Namespace NDXAccess.Examples

    Public Module AdvancedExamples

        ' --------------------------------------------------------------------
        ' Health checks
        ' --------------------------------------------------------------------

        Public Async Function HealthCheckAsync(factory As IAccessConnectionFactory) As Task
            Dim healthCheck = New AccessHealthCheck(factory)
            Dim result = Await healthCheck.CheckHealthAsync()

            System.Console.WriteLine($"Santé : {If(result.IsHealthy, "OK", "KO")} - {result.Message}")
            System.Console.WriteLine($"Temps de réponse : {result.ResponseTime.TotalMilliseconds:F1} ms")

            If result.DatabaseInfo IsNot Nothing Then
                Dim info = result.DatabaseInfo
                System.Console.WriteLine($"Fichier : {info.FilePath}")
                System.Console.WriteLine($"Taille : {info.FileSizeMegabytes:F1} Mo ({info.UsagePercent:F1} % de la limite de 2 Go)")
                System.Console.WriteLine($"Tables : {info.UserTableCount} | Moteur : {info.EngineVersion}")
                If info.IsApproachingSizeLimit Then
                    System.Console.WriteLine("ATTENTION : la base approche la limite des 2 Go — compactez ou archivez.")
                End If
            End If
        End Function

        ' --------------------------------------------------------------------
        ' Injection de dépendances
        ' --------------------------------------------------------------------

        Public Function ConfigureServices(databasePath As String) As IServiceProvider
            Dim services As New ServiceCollection()

            services.AddLogging(Sub(b) b.AddConsole())

            ' Enregistre options + factory + IAccessConnection + AccessHealthCheck.
            services.AddNDXAccess(Sub(options)
                                      options.DatabasePath = databasePath
                                      options.AutoCloseTimeoutMs = 30_000
                                  End Sub)

            services.AddScoped(Of IClientRepository, ClientRepository)()

            Return services.BuildServiceProvider()
        End Function

        Public Interface IClientRepository
            Function GetAllAsync() As Task(Of DataTable)
        End Interface

        Public Class ClientRepository
            Implements IClientRepository

            Private ReadOnly _factory As IAccessConnectionFactory
            Private ReadOnly _logger As ILogger(Of ClientRepository)

            Public Sub New(factory As IAccessConnectionFactory, logger As ILogger(Of ClientRepository))
                _factory = factory
                _logger = logger
            End Sub

            Public Async Function GetAllAsync() As Task(Of DataTable) Implements IClientRepository.GetAllAsync
                Using connection = _factory.CreateConnection()
                    _logger.LogDebug("Récupération de tous les clients")
                    Return Await connection.ExecuteQueryAsync("SELECT * FROM clients")
                End Using
            End Function
        End Class

        ' --------------------------------------------------------------------
        ' Concurrence (limites d'Access)
        ' --------------------------------------------------------------------

        ''' <summary>
        ''' Access tient mal au-delà de ~10-15 utilisateurs simultanés sur le même fichier.
        ''' L'async étant "de façade", n'attendez pas de gain de parallélisme : chaque
        ''' connexion sérialise réellement ses accès au fichier.
        ''' Pour de la forte concurrence, migrez vers un SGBD client-serveur.
        ''' </summary>
        Public Async Function SequentialQueriesAsync(factory As IAccessConnectionFactory) As Task
            For i = 0 To 4
                Using connection = factory.CreateConnection()
                    Dim count = Await connection.ExecuteScalarAsync(Of Integer)("SELECT COUNT(*) FROM clients")
                    System.Console.WriteLine($"Itération {i} : {count} clients")
                End Using
            Next
        End Function

    End Module

End Namespace
