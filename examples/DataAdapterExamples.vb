' ============================================================================
' NDXAccess - Exemples DataAdapter (édition déconnectée par lot)
' ============================================================================
' Pattern ADO.NET classique : remplir un DataTable (Fill), le modifier hors ligne,
' puis pousser toutes les modifications d'un coup (Update) via OleDbCommandBuilder
' qui génère automatiquement les commandes INSERT / UPDATE / DELETE.
'
' NOTE : exemple de documentation (projet compilé).
' Auteur : Nicolas DEOUX <NDXDev@gmail.com>
' ============================================================================

Imports System.Data
Imports System.Data.OleDb
Imports NDXAccess

Namespace NDXAccess.Examples

    Public Module DataAdapterExamples

        ''' <summary>
        ''' Charge la table dans un DataTable, applique des modifications déconnectées,
        ''' puis les répercute en base. Retourne le nombre de lignes affectées par Update.
        ''' La table doit avoir une clé primaire (pour que le CommandBuilder génère
        ''' UPDATE/DELETE) et les colonnes nom / email / actif.
        ''' </summary>
        Public Function FillModifyUpdate(connection As IAccessConnection, tableName As String) As Integer
            connection.Open()

            Dim table As New DataTable()
            Using adapter As New OleDbDataAdapter($"SELECT * FROM [{tableName}]", connection.Connection)
                ' Le CommandBuilder doit rester vivant tant qu'on utilise l'adapter.
                Using builder As New OleDbCommandBuilder(adapter) With {.QuotePrefix = "[", .QuoteSuffix = "]"}
                    adapter.Fill(table)

                    ' --- Modifications 100% déconnectées (aucun aller-retour réseau) ---
                    If table.Rows.Count > 0 Then
                        table.Rows(0)("nom") = "Nom modifié"
                    End If

                    Dim nouvelle = table.NewRow()
                    nouvelle("nom") = "Nouveau via adapter"
                    nouvelle("email") = "adapter@example.com"
                    nouvelle("actif") = True
                    table.Rows.Add(nouvelle)

                    ' --- Pousse INSERT + UPDATE générés automatiquement, en un seul appel ---
                    Return adapter.Update(table)
                End Using
            End Using
        End Function

        ''' <summary>Charge plusieurs tables dans un DataSet via des adapters successifs.</summary>
        Public Function FillDataSet(connection As IAccessConnection, ParamArray tableNames As String()) As DataSet
            connection.Open()
            Dim ds As New DataSet()
            For Each name In tableNames
                Using adapter As New OleDbDataAdapter($"SELECT * FROM [{name}]", connection.Connection)
                    adapter.Fill(ds, name)
                End Using
            Next
            Return ds
        End Function

    End Module

End Namespace
