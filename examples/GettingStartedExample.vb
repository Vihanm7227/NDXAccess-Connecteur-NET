' ============================================================================
' NDXAccess - Démarrage « tout-en-un »
' ============================================================================
' Flux complet de bout en bout, à copier dans un Sub Main :
'   création de la base -> schéma -> insertion -> lecture -> affichage.
'
' NOTE : exemple de documentation (projet compilé et testé).
' Auteur : Nicolas DEOUX <NDXDev@gmail.com>
' ============================================================================

Imports System.IO
Imports NDXAccess

Namespace NDXAccess.Examples

    Public Module GettingStartedExample

        ''' <summary>
        ''' Exécute le scénario complet sur <paramref name="databasePath"/> (créé s'il
        ''' n'existe pas) et retourne le nombre de clients en base à la fin.
        ''' </summary>
        Public Async Function RunAsync(databasePath As String) As Task(Of Integer)
            ' 1. Créer le fichier .accdb s'il n'existe pas encore.
            If Not File.Exists(databasePath) Then
                Await AccessConnection.CreateDatabaseAsync(databasePath)
            End If

            Dim options As New AccessConnectionOptions With {.DatabasePath = databasePath}
            Using connection As IAccessConnection = New AccessConnection(options)

                ' 2. Créer le schéma si nécessaire.
                If Not connection.TableExists("clients") Then
                    Await connection.ExecuteNonQueryAsync(
                        "CREATE TABLE clients (id AUTOINCREMENT PRIMARY KEY, nom TEXT(100), email TEXT(200), actif YESNO)")
                End If

                ' 3. Insérer des données (paramètres positionnels '?').
                Await connection.ExecuteNonQueryAsync(
                    "INSERT INTO clients (nom, email, actif) VALUES (?, ?, ?)",
                    {"Jean Dupont", "jean@example.com", True})
                Dim newId = Await connection.ExecuteScalarAsync(Of Integer)("SELECT @@IDENTITY")
                Console.WriteLine($"Client créé avec l'id {newId}")

                ' 4. Lire et afficher.
                Dim clients = Await connection.ExecuteQueryAsync("SELECT id, nom, email FROM clients ORDER BY nom")
                For Each row As Data.DataRow In clients.Rows
                    Console.WriteLine($"  [{row("id")}] {row("nom")} <{row("email")}>")
                Next

                Return Await connection.ExecuteScalarAsync(Of Integer)("SELECT COUNT(*) FROM clients")
            End Using
        End Function

    End Module

End Namespace
