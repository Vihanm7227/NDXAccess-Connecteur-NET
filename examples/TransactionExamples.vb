' ============================================================================
' NDXAccess - Exemples de transactions
' ============================================================================
' Access/ACE supporte les transactions (OleDbTransaction) via Begin/Commit/Rollback.
' Le niveau d'isolation pris en charge est essentiellement ReadCommitted.
'
' NOTE : exemples de documentation, non compilés par les tests.
' Auteur : Nicolas DEOUX <NDXDev@gmail.com>
' ============================================================================

Imports System.Data
Imports NDXAccess

Namespace NDXAccess.Examples

    Public Module TransactionExamples

        ''' <summary>Transaction simple avec commit (asynchrone de façade).</summary>
        Public Async Function SimpleTransactionAsync(connection As IAccessConnection) As Task
            Await connection.BeginTransactionAsync()
            ' VB interdit Await dans un Catch : on capture l'exception puis on annule hors du Catch.
            Dim failure As Exception = Nothing
            Try
                Await connection.ExecuteNonQueryAsync(
                    "INSERT INTO clients (nom, email) VALUES (?, ?)",
                    {"Client Test", "test@example.com"})
                Await connection.ExecuteNonQueryAsync(
                    "INSERT INTO journal (action, message) VALUES (?, ?)",
                    {"CREATE_CLIENT", "Nouveau client créé"})

                Await connection.CommitAsync()
                Console.WriteLine("Transaction validée")
            Catch ex As Exception
                failure = ex
            End Try

            If failure IsNot Nothing Then
                Await connection.RollbackAsync()
                Console.WriteLine($"Transaction annulée : {failure.Message}")
                Throw failure
            End If
        End Function

        ''' <summary>Transaction synchrone avec rollback conditionnel.</summary>
        Public Sub TransferMoney(connection As IAccessConnection, fromId As Integer, toId As Integer, amount As Decimal)
            If amount <= 0D Then Throw New ArgumentException("Le montant doit être positif", NameOf(amount))

            connection.BeginTransaction()
            Try
                Dim solde = connection.ExecuteScalar(Of Decimal)(
                    "SELECT solde FROM comptes WHERE id = ?", fromId)

                If solde < amount Then
                    connection.Rollback()
                    Console.WriteLine("Transfert refusé : solde insuffisant")
                    Return
                End If

                connection.ExecuteNonQuery("UPDATE comptes SET solde = solde - ? WHERE id = ?", amount, fromId)
                connection.ExecuteNonQuery("UPDATE comptes SET solde = solde + ? WHERE id = ?", amount, toId)

                connection.Commit()
                Console.WriteLine($"Transfert de {amount:C} effectué")
            Catch
                connection.Rollback()
                Throw
            End Try
        End Sub

        ''' <summary>Insertion en masse dans une transaction (bonnes performances en Access).</summary>
        Public Async Function BulkInsertAsync(connection As IAccessConnection, clients As IEnumerable(Of (Nom As String, Email As String))) As Task(Of Integer)
            Await connection.BeginTransactionAsync()
            Dim inserted = 0
            Dim failure As Exception = Nothing
            Try
                For Each c In clients
                    inserted += Await connection.ExecuteNonQueryAsync(
                        "INSERT INTO clients (nom, email) VALUES (?, ?)",
                        {c.Nom, c.Email})
                Next
                Await connection.CommitAsync()
            Catch ex As Exception
                failure = ex
            End Try

            If failure IsNot Nothing Then
                Await connection.RollbackAsync()
                Throw failure
            End If

            Console.WriteLine($"{inserted} clients insérés")
            Return inserted
        End Function

    End Module

End Namespace
