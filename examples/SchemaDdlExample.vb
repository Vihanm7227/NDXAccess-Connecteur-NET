' ============================================================================
' NDXAccess - DDL : tables, index et relations (clés étrangères)
' ============================================================================
' Access (ACE/Jet) supporte un sous-ensemble du DDL SQL : CREATE TABLE avec
' contraintes (PRIMARY KEY, NOT NULL, UNIQUE, FOREIGN KEY) et CREATE INDEX.
'
' NOTE : exemple de documentation (projet compilé et testé).
' Auteur : Nicolas DEOUX <NDXDev@gmail.com>
' ============================================================================

Imports NDXAccess

Namespace NDXAccess.Examples

    Public Module SchemaDdlExample

        ''' <summary>
        ''' Crée un schéma relationnel : clients (1) — (N) commandes, avec un index
        ''' unique sur l'email et une clé étrangère avec intégrité référentielle.
        ''' </summary>
        Public Sub CreateRelationalSchema(connection As IAccessConnection)
            connection.Open()

            connection.ExecuteNonQuery(
                "CREATE TABLE clients (" &
                "  id AUTOINCREMENT PRIMARY KEY, " &
                "  nom TEXT(100) NOT NULL, " &
                "  email TEXT(200))")

            ' Index unique : empêche les emails en double.
            connection.ExecuteNonQuery("CREATE UNIQUE INDEX idx_clients_email ON clients (email)")

            ' Table enfant avec clé étrangère vers clients(id).
            connection.ExecuteNonQuery(
                "CREATE TABLE commandes (" &
                "  id AUTOINCREMENT PRIMARY KEY, " &
                "  fkClient LONG NOT NULL, " &
                "  montant CURRENCY, " &
                "  CONSTRAINT fk_commande_client FOREIGN KEY (fkClient) REFERENCES clients (id))")
        End Sub

    End Module

End Namespace
