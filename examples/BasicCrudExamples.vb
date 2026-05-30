' ============================================================================
' NDXAccess - Exemples CRUD de base
' ============================================================================
' Opérations Create / Read / Update / Delete avec NDXAccess (VB.NET).
'
' NOTE : exemples de documentation, non compilés par les tests.
' Rappel Access : paramètres POSITIONNELS ('?'), dans l'ordre.
'
' Auteur : Nicolas DEOUX <NDXDev@gmail.com>
' ============================================================================

Imports System.Data
Imports NDXAccess

Namespace NDXAccess.Examples

    Public Module BasicCrudExamples

        ' --------------------------------------------------------------------
        ' Configuration
        ' --------------------------------------------------------------------

        Public Function GetOptions() As AccessConnectionOptions
            Return New AccessConnectionOptions With {
                .DatabasePath = "C:\data\ma_base.accdb",
                .Password = "",          ' mot de passe éventuel
                .AutoCloseTimeoutMs = 60_000
            }
        End Function

        ' --------------------------------------------------------------------
        ' CREATE
        ' --------------------------------------------------------------------

        ''' <summary>Insertion simple (synchrone).</summary>
        Public Sub InsertSimple(connection As IAccessConnection)
            Dim rows = connection.ExecuteNonQuery(
                "INSERT INTO clients (nom, email) VALUES (?, ?)",
                "Jean Dupont", "jean.dupont@example.com")
            Console.WriteLine($"Lignes insérées : {rows}")
        End Sub

        ''' <summary>Insertion (asynchrone) puis récupération de l'identifiant auto-incrémenté.</summary>
        Public Async Function InsertAndGetIdAsync(connection As IAccessConnection) As Task(Of Integer)
            Await connection.ExecuteNonQueryAsync(
                "INSERT INTO clients (nom, email) VALUES (?, ?)",
                {"Marie Martin", "marie.martin@example.com"})

            ' En Access, @@IDENTITY renvoie le dernier compteur (AUTOINCREMENT) inséré
            ' sur la connexion courante.
            Dim newId = Await connection.ExecuteScalarAsync(Of Integer)("SELECT @@IDENTITY")
            Console.WriteLine($"Nouveau client créé avec l'ID : {newId}")
            Return newId
        End Function

        ''' <summary>Insertion avec plusieurs types (date, montant).</summary>
        Public Async Function InsertWithTypesAsync(connection As IAccessConnection) As Task
            ' Decimal -> CURRENCY et Date -> DATETIME sont gérés automatiquement.
            Await connection.ExecuteNonQueryAsync(
                "INSERT INTO produits (nom, prix, quantite, actif, date_creation) VALUES (?, ?, ?, ?, ?)",
                {"Laptop Pro", 1299.99D, 50, True, DateTime.Now})
        End Function

        ' --------------------------------------------------------------------
        ' READ
        ' --------------------------------------------------------------------

        ''' <summary>Lecture scalaire typée.</summary>
        Public Async Function GetCountAsync(connection As IAccessConnection) As Task(Of Integer)
            Dim count = Await connection.ExecuteScalarAsync(Of Integer)("SELECT COUNT(*) FROM clients")
            Console.WriteLine($"Nombre de clients : {count}")
            Return count
        End Function

        ''' <summary>Lecture filtrée vers un DataTable.</summary>
        Public Async Function GetActiveClientsAsync(connection As IAccessConnection) As Task(Of DataTable)
            Dim result = Await connection.ExecuteQueryAsync(
                "SELECT id, nom, email FROM clients WHERE actif = ? ORDER BY nom",
                {True})
            Console.WriteLine($"Clients actifs : {result.Rows.Count}")
            Return result
        End Function

        ''' <summary>Lecture ligne par ligne avec un DataReader.</summary>
        Public Sub ProcessWithReader(connection As IAccessConnection)
            connection.Open()
            Using reader = connection.ExecuteReader("SELECT id, nom FROM clients WHERE actif = ?", True)
                While reader.Read()
                    Console.WriteLine($"Client {reader.GetInt32(0)} : {reader.GetString(1)}")
                End While
            End Using
        End Sub

        ' --------------------------------------------------------------------
        ' UPDATE
        ' --------------------------------------------------------------------

        Public Async Function UpdateEmailAsync(connection As IAccessConnection, clientId As Integer, email As String) As Task(Of Integer)
            Dim rows = Await connection.ExecuteNonQueryAsync(
                "UPDATE clients SET email = ? WHERE id = ?",
                {email, clientId})
            Console.WriteLine($"Client {clientId} mis à jour : {rows} ligne(s)")
            Return rows
        End Function

        ' --------------------------------------------------------------------
        ' DELETE
        ' --------------------------------------------------------------------

        Public Async Function DeleteClientAsync(connection As IAccessConnection, clientId As Integer) As Task(Of Integer)
            Dim rows = Await connection.ExecuteNonQueryAsync(
                "DELETE FROM clients WHERE id = ?", {clientId})
            Console.WriteLine($"Client {clientId} supprimé : {rows} ligne(s)")
            Return rows
        End Function

        ' --------------------------------------------------------------------
        ' Exemple complet
        ' --------------------------------------------------------------------

        Public Async Function FullExampleAsync() As Task
            Using connection As IAccessConnection = New AccessConnection(GetOptions())
                Await connection.OpenAsync()

                Dim newId = Await InsertAndGetIdAsync(connection)
                Await GetActiveClientsAsync(connection)
                Await UpdateEmailAsync(connection, newId, "nouveau@example.com")
                Await DeleteClientAsync(connection, newId)

                Await connection.CloseAsync()
            End Using
        End Function

    End Module

End Namespace
