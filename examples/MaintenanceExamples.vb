' ============================================================================
' NDXAccess - Exemples de maintenance
' ============================================================================
' Compactage/réparation de la base et détection de l'architecture du provider.
'
' NOTE : exemples de documentation, non compilés par les tests.
' Auteur : Nicolas DEOUX <NDXDev@gmail.com>
' ============================================================================

Imports NDXAccess

Namespace NDXAccess.Examples

    Public Module MaintenanceExamples

        ' --------------------------------------------------------------------
        ' Détection x86 / x64 du provider ACE
        ' --------------------------------------------------------------------

        ''' <summary>
        ''' Vérifie en amont que le provider ACE correspond à l'architecture de l'appli.
        ''' Affiche un message clair plutôt que de laisser planter une connexion.
        ''' </summary>
        Public Sub CheckProvider()
            Console.WriteLine($"Architecture du processus : {AccessProviderHelper.CurrentProcessArchitecture}")
            Console.WriteLine("Providers OLE DB disponibles :")
            For Each p In AccessProviderHelper.GetAvailableProviders()
                Console.WriteLine($"  - {p}")
            Next

            If AccessProviderHelper.IsProviderAvailable("Microsoft.ACE.OLEDB.16.0") Then
                Console.WriteLine("ACE 16.0 disponible.")
            Else
                Try
                    AccessProviderHelper.EnsureProviderAvailable("Microsoft.ACE.OLEDB.16.0")
                Catch ex As AccessProviderNotFoundException
                    Console.WriteLine(ex.Message)
                End Try
            End If
        End Sub

        ' --------------------------------------------------------------------
        ' Compactage / réparation
        ' --------------------------------------------------------------------

        ''' <summary>
        ''' Compacte la base en place (asynchrone réel : opération bloquante déportée).
        ''' La connexion est fermée automatiquement avant le compactage.
        ''' </summary>
        Public Async Function CompactInPlaceAsync(connection As IAccessConnection) As Task
            Console.WriteLine("Compactage en cours...")
            Await connection.CompactDatabaseAsync()
            Console.WriteLine("Compactage terminé.")
        End Function

        ''' <summary>Compactage vers un nouveau fichier (synchrone).</summary>
        Public Sub CompactToNewFile(connection As IAccessConnection, targetPath As String)
            connection.CompactDatabase(targetPath)
            Console.WriteLine($"Base compactée vers : {targetPath}")
        End Sub

        ''' <summary>
        ''' Compactage planifié : Access n'ayant pas d'Event Scheduler, créez une tâche
        ''' dans le Planificateur de tâches Windows qui lance votre exécutable avec un
        ''' argument déclenchant ce code (par ex. chaque nuit).
        ''' </summary>
        Public Async Function ScheduledMaintenanceEntryPointAsync(databasePath As String) As Task
            Dim options = New AccessConnectionOptions With {.DatabasePath = databasePath}
            Using connection As IAccessConnection = New AccessConnection(options)
                Await connection.CompactDatabaseAsync()
            End Using
        End Function

    End Module

End Namespace
