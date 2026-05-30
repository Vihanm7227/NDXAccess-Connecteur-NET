' ============================================================================
' NDXAccess - Liaison WinForms (DataGridView éditable)
' ============================================================================
' Scénario classique VB.NET + Access : afficher une table dans un DataGridView,
' laisser l'utilisateur éditer, puis enregistrer toutes les modifications via
' OleDbDataAdapter.Update (+ OleDbCommandBuilder).
'
' On utilise une connexion PRINCIPALE (pas de fermeture auto) maintenue ouverte
' tant que le binder vit, car l'adapter en a besoin pour Fill et Update.
'
' NOTE : exemple de documentation (projet compilé ; non exécuté car interface).
' Auteur : Nicolas DEOUX <NDXDev@gmail.com>
' ============================================================================

Imports System.Data
Imports System.Data.OleDb
Imports System.Windows.Forms
Imports NDXAccess

Namespace NDXAccess.Examples

    ''' <summary>
    ''' Lie la table "clients" d'une base Access à un DataGridView et sauvegarde
    ''' les modifications par lot. À disposer quand le formulaire se ferme.
    ''' </summary>
    Public NotInheritable Class ClientsGridBinder
        Implements IDisposable

        Private ReadOnly _connection As IAccessConnection
        Private ReadOnly _adapter As OleDbDataAdapter
        Private ReadOnly _builder As OleDbCommandBuilder
        Private ReadOnly _table As New DataTable()

        Public Sub New(databasePath As String)
            ' Connexion principale : reste ouverte pour l'adapter (pas d'auto-close).
            _connection = New AccessConnection(New AccessConnectionOptions With {
                .DatabasePath = databasePath,
                .IsPrimaryConnection = True
            })
            _connection.Open()

            _adapter = New OleDbDataAdapter("SELECT * FROM clients", _connection.Connection)
            _builder = New OleDbCommandBuilder(_adapter) With {.QuotePrefix = "[", .QuoteSuffix = "]"}
            _adapter.Fill(_table)
        End Sub

        ''' <summary>Lie le DataTable au DataGridView (édition directe dans la grille).</summary>
        Public Sub BindTo(grid As DataGridView)
            grid.AutoGenerateColumns = True
            grid.DataSource = _table
        End Sub

        ''' <summary>Enregistre en base toutes les modifications de la grille. Retourne le nb de lignes.</summary>
        Public Function Save() As Integer
            ' Valide l'édition de cellule en cours avant l'enregistrement.
            Return _adapter.Update(_table)
        End Function

        Public Sub Dispose() Implements IDisposable.Dispose
            _builder?.Dispose()
            _adapter?.Dispose()
            _connection?.Dispose()
        End Sub

    End Class

End Namespace
