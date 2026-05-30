' ============================================================================
' NDXAccess - Exemples de requêtes enregistrées (stored queries)
' ============================================================================
' Access n'a PAS de vraies procédures stockées. Il dispose de "requêtes
' enregistrées" (QueryDefs) qui acceptent uniquement des paramètres d'ENTRÉE.
' Il n'existe NI paramètre OUT, NI paramètre INOUT.
'
' Pour récupérer une valeur calculée, exécutez une requête SELECT classique
' (ex. SELECT @@IDENTITY après une insertion).
'
' NOTE : exemples de documentation, non compilés par les tests.
' Auteur : Nicolas DEOUX <NDXDev@gmail.com>
' ============================================================================

Imports System.Data
Imports NDXAccess

Namespace NDXAccess.Examples

    Public Module StoredQueryExamples

        ''' <summary>
        ''' Création d'une requête enregistrée paramétrée (paramètre IN uniquement).
        ''' À faire une fois (par exemple à l'initialisation de la base).
        ''' </summary>
        Public Async Function CreateStoredQueryAsync(connection As IAccessConnection) As Task
            ' CREATE PROCEDURE crée une requête enregistrée dans le fichier .accdb.
            Await connection.ExecuteNonQueryAsync(
                "CREATE PROCEDURE qClientsParStatut (prmActif BIT) AS " &
                "SELECT id, nom, email FROM clients WHERE actif = prmActif ORDER BY nom")
        End Function

        ''' <summary>
        ''' Appel d'une requête enregistrée renvoyant des lignes, avec paramètre IN.
        ''' Les paramètres sont positionnels (dans l'ordre de déclaration de la requête).
        ''' </summary>
        Public Async Function CallSelectStoredQueryAsync(connection As IAccessConnection) As Task(Of DataTable)
            Dim result = Await connection.ExecuteStoredQueryAsync("qClientsParStatut", {True})
            Console.WriteLine($"Clients actifs : {result.Rows.Count}")
            Return result
        End Function

        ''' <summary>Appel synchrone équivalent.</summary>
        Public Function CallSelectStoredQuery(connection As IAccessConnection) As DataTable
            Return connection.ExecuteStoredQuery("qClientsParStatut", True)
        End Function

        ''' <summary>
        ''' Requête enregistrée d'action (UPDATE/DELETE/INSERT) : renvoie le nombre
        ''' de lignes affectées.
        ''' </summary>
        Public Async Function CallActionStoredQueryAsync(connection As IAccessConnection) As Task
            Await connection.ExecuteNonQueryAsync(
                "CREATE PROCEDURE qDesactiverClient (prmId LONG) AS " &
                "UPDATE clients SET actif = No WHERE id = prmId")

            Dim affected = Await connection.ExecuteStoredQueryNonQueryAsync("qDesactiverClient", {42})
            Console.WriteLine($"Lignes désactivées : {affected}")
        End Function

        ''' <summary>
        ''' "Émulation" d'un paramètre de sortie : Access ne sait pas le faire dans une
        ''' requête enregistrée. On exécute donc une requête de calcul classique.
        ''' </summary>
        Public Async Function GetComputedValueAsync(connection As IAccessConnection) As Task(Of Integer)
            ' Pas de paramètre OUT : on lit directement le résultat d'un SELECT.
            Return Await connection.ExecuteScalarAsync(Of Integer)(
                "SELECT COUNT(*) FROM clients WHERE actif = ?", {True})
        End Function

    End Module

End Namespace
