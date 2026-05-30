' ============================================================================
' NDXAccess - Pagination « façon Access »
' ============================================================================
' Access (ACE/Jet) n'a NI LIMIT NI OFFSET. On pagine avec SELECT TOP n, et pour
' les pages suivantes on exclut les n*(page-1) premières lignes via NOT IN.
'
' ATTENTION : SELECT TOP inclut les ex-æquo. Pour une pagination déterministe,
' triez sur une colonne (ou combinaison) UNIQUE — typiquement la clé primaire.
'
' NOTE : exemple de documentation (projet compilé et testé).
' Auteur : Nicolas DEOUX <NDXDev@gmail.com>
' ============================================================================

Imports System.Data
Imports NDXAccess

Namespace NDXAccess.Examples

    Public Module PaginationExample

        ''' <summary>
        ''' Retourne la page <paramref name="page"/> (1-based) de taille
        ''' <paramref name="pageSize"/>, triée par <paramref name="keyColumn"/> (doit être
        ''' une colonne unique, p.ex. la clé primaire).
        ''' </summary>
        Public Function GetPage(connection As IAccessConnection, tableName As String, keyColumn As String, pageSize As Integer, page As Integer) As DataTable
            Dim skip = pageSize * (page - 1)

            Dim sql As String
            If skip <= 0 Then
                sql = $"SELECT TOP {pageSize} * FROM [{tableName}] ORDER BY [{keyColumn}]"
            Else
                sql = $"SELECT TOP {pageSize} * FROM [{tableName}] " &
                      $"WHERE [{keyColumn}] NOT IN (SELECT TOP {skip} [{keyColumn}] FROM [{tableName}] ORDER BY [{keyColumn}]) " &
                      $"ORDER BY [{keyColumn}]"
            End If

            Return connection.ExecuteQuery(sql)
        End Function

    End Module

End Namespace
