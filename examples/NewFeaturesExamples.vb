' ============================================================================
' NDXAccess - Exemples des fonctionnalités v1.1
' ============================================================================
' Résilience, traduction d'erreurs, helpers de schéma, création de base,
' insertion en masse, mapping objet et paramètres nommés.
'
' NOTE : exemples de documentation, non compilés par les tests.
' Auteur : Nicolas DEOUX <NDXDev@gmail.com>
' ============================================================================

Imports System.Collections.Generic
Imports NDXAccess

Namespace NDXAccess.Examples

    Public Module NewFeaturesExamples

        ''' <summary>Classe cible du mapping objet (propriétés = colonnes).</summary>
        Public Class Client
            Public Property Id As Integer
            Public Property Nom As String
            Public Property Email As String
            Public Property Actif As Boolean
        End Class

        ' --------------------------------------------------------------------
        ' Résilience + erreurs claires
        ' --------------------------------------------------------------------

        Public Async Function ResilientUpdateAsync(connection As IAccessConnection) As Task
            ' Retry automatique sur verrou transitoire (config via les options).
            Try
                Await connection.ExecuteNonQueryAsync(
                    "UPDATE clients SET actif = ? WHERE id = ?", {True, 1})
            Catch ex As AccessQueryException
                Console.WriteLine($"Erreur Access claire : {ex.Message}")
                Console.WriteLine($"Code natif : {ex.NativeError} | Transitoire : {ex.IsTransient}")
                ' ex.InnerException contient l'OleDbException d'origine.
            End Try
        End Function

        ' --------------------------------------------------------------------
        ' Création de base + helpers de schéma
        ' --------------------------------------------------------------------

        Public Async Function CreateAndInspectAsync(path As String) As Task
            Await AccessConnection.CreateDatabaseAsync(path)

            Using connection As IAccessConnection = New AccessConnection(New AccessConnectionOptions With {.DatabasePath = path})
                Await connection.ExecuteNonQueryAsync(
                    "CREATE TABLE clients (id AUTOINCREMENT PRIMARY KEY, nom TEXT(100), email TEXT(200), actif YESNO)")

                Console.WriteLine($"Table 'clients' existe : {connection.TableExists("clients")}")
                For Each t In connection.GetTableNames()
                    Console.WriteLine($"  Table : {t}")
                Next

                Dim columns = connection.GetColumns("clients")
                Console.WriteLine($"Colonnes : {columns.Rows.Count}")
            End Using
        End Function

        ' --------------------------------------------------------------------
        ' Insertion en masse
        ' --------------------------------------------------------------------

        Public Async Function BulkLoadAsync(connection As IAccessConnection) As Task
            Dim rows As New List(Of Object())()
            For i = 1 To 5000
                rows.Add(New Object() {$"Client {i}", $"client{i}@example.com", True})
            Next

            Dim inserted = Await connection.BulkInsertAsync(
                "clients", {"nom", "email", "actif"}, rows)
            Console.WriteLine($"{inserted} lignes insérées (transaction unique)")
        End Function

        ' --------------------------------------------------------------------
        ' Mapping objet
        ' --------------------------------------------------------------------

        Public Async Function GetActiveClientsTypedAsync(connection As IAccessConnection) As Task(Of List(Of Client))
            Return Await connection.ExecuteQueryAsync(Of Client)(
                "SELECT id, nom, email, actif FROM clients WHERE actif = ? ORDER BY nom", {True})
        End Function

        ' --------------------------------------------------------------------
        ' Paramètres nommés (@nom -> ?)
        ' --------------------------------------------------------------------

        Public Async Function NamedParametersAsync(connection As IAccessConnection) As Task
            Dim params As New Dictionary(Of String, Object) From {
                {"actif", True},
                {"prefixe", "A%"}
            }

            Dim dt = Await connection.ExecuteQueryNamedAsync(
                "SELECT nom FROM clients WHERE actif = @actif AND nom LIKE @prefixe ORDER BY nom",
                params)
            Console.WriteLine($"Résultats : {dt.Rows.Count}")
        End Function

    End Module

End Namespace
