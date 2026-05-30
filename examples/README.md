# Exemples NDXAccess

Ce dossier contient des exemples d'utilisation de la bibliothèque **NDXAccess** (VB.NET).

> **Note** : ces exemples forment un projet compilé (`NDXAccess.Examples.vbproj`, inclus dans  
> la solution) — ils sont donc **garantis de compiler** à chaque build/CI. Ils ne sont pas  
> exécutés (pas de base réelle) et ne sont pas packagés.

## Structure des fichiers

| Fichier | Description |
|---------|-------------|
| `BasicCrudExamples.vb` | Opérations CRUD de base (Create, Read, Update, Delete) |
| `StoredQueryExamples.vb` | Requêtes enregistrées Access (paramètres IN uniquement) |
| `TransactionExamples.vb` | Transactions et opérations en masse |
| `AdvancedExamples.vb` | Health checks, injection de dépendances, parallélisme |
| `MaintenanceExamples.vb` | Compactage/réparation et détection x86/x64 |
| `NewFeaturesExamples.vb` | v1.1 : résilience, schéma, CreateDatabase, BulkInsert, mapping, paramètres nommés |
| `EntityCrudExample.vb` | Entité métier « Active Record » : `OleDbCommand`, paramètres, `OleDbDataAdapter.Fill`, `@@IDENTITY`, `IDisposable` |
| `DataAdapterExamples.vb` | Édition déconnectée par lot : `OleDbDataAdapter` `Fill`/`Update` + `OleDbCommandBuilder`, et `DataSet` multi-tables |
| `GettingStartedExample.vb` | Démarrage tout-en-un : création base → schéma → CRUD → lecture (à copier dans un `Sub Main`) |
| `PaginationExample.vb` | Pagination « façon Access » : `SELECT TOP n` + astuce `NOT IN` (ni `LIMIT` ni `OFFSET`) |
| `SchemaDdlExample.vb` | DDL : tables, `CREATE UNIQUE INDEX`, relations (clé étrangère) |
| `ExcelCsvExample.vb` | Lire des fichiers CSV / Excel via le provider ACE (`Extended Properties`) |
| `LoggingExample.vb` | Câbler `Microsoft.Extensions.Logging` pour voir les logs du connecteur |
| `WinFormsBindingExample.vb` | Liaison `DataGridView` éditable + sauvegarde par lot (`adapter.Update`) |

## Points clés spécifiques à Access

- **Paramètres positionnels** : Access/OLE DB n'utilise **pas** de paramètres nommés.
  Utilisez `?` dans le SQL et passez les valeurs **dans l'ordre**.

  ```vb
  Await connection.ExecuteNonQueryAsync(
      "INSERT INTO clients (nom, email) VALUES (?, ?)",
      {"Jean Dupont", "jean@example.com"})
  ```

- **Async de façade** : les méthodes `...Async` existent pour la cohérence d'API mais
  s'exécutent de manière synchrone (OLE DB n'a pas de vraie asynchronie). Seule
  `CompactDatabaseAsync` est réellement déportée sur un thread.

- **Pas d'Event Scheduler** : Access n'a pas de planificateur. Utilisez le
  **Planificateur de tâches Windows** pour exécuter un programme/script à intervalle régulier.

- **Pas de vraies procédures stockées** : seulement des **requêtes enregistrées**
  paramétrées, avec des paramètres d'**entrée** uniquement (pas de OUT/INOUT).

- **Windows uniquement** : le provider `Microsoft.ACE.OLEDB.16.0` n'existe pas ailleurs,
  et son architecture (x86/x64) doit correspondre à celle de l'application.

## Configuration de la connexion

```vb
Dim options As New AccessConnectionOptions With {
    .DatabasePath = "C:\data\ma_base.accdb",
    .Password = "mot_de_passe_optionnel"
}

' VB.NET n'a pas de "Await Using" : on utilise un bloc Using (Dispose synchrone, OK ici)…
Using connection As IAccessConnection = New AccessConnection(options)
    Await connection.OpenAsync()
End Using

' …ou, pour une libération asynchrone explicite :
Dim conn As IAccessConnection = New AccessConnection(options)  
Try  
    Await conn.OpenAsync()
Finally
    Await conn.DisposeAsync()
End Try
```

## Pattern « entité métier » (`EntityCrudExample.vb`)

Reprise du style Active Record avec `OleDbCommand` / `OleDbParameter` /
`OleDbDataAdapter.Fill`, adaptée aux **spécificités Access** :

- **Paramètres positionnels** : le SQL utilise `?` et l'**ordre** des paramètres ajoutés
  doit suivre celui des `?`. En `UPDATE … WHERE id = ?`, le paramètre `id` est ajouté
  **en dernier**.
- **Pas de multi-instructions** : `INSERT …; SELECT @@IDENTITY;` est interdit. On exécute
  **deux commandes** (l'`INSERT` puis `SELECT @@IDENTITY`) sur la **même** connexion.
- **`SELECT TOP 1 …`** au lieu de `… LIMIT 1`.
- **Dates** typées `OleDbType.Date`.
- **Connexion partagée vs possédée** : le constructeur accepte une `IAccessConnection`
  partagée (non disposée) ou un chemin de base (connexion créée et disposée par l'entité).

```vb
' INSERT puis récupération de l'identité (deux commandes)
Using cmd = connection.CreateCommand("INSERT INTO tb_Etapes (Libelle) VALUES (?)")
    cmd.Parameters.Add("Libelle", OleDbType.VarWChar).Value = "Préparation"
    cmd.ExecuteNonQuery()
End Using  
Using idCmd = connection.CreateCommand("SELECT @@IDENTITY")  
    Dim newId = Convert.ToInt32(idCmd.ExecuteScalar())
End Using

' Lecture avec OleDbDataAdapter.Fill
Dim dt As New DataTable()  
Using cmd = connection.CreateCommand("SELECT TOP 1 * FROM tb_Etapes WHERE id = ?")  
    cmd.Parameters.Add("id", OleDbType.Integer).Value = 1
    Using da As New OleDbDataAdapter(cmd) : da.Fill(dt) : End Using
End Using
```

> 💡 Le fichier montre aussi la variante **haut niveau** (`_access.ExecuteQuery(...)`),  
> qui ajoute gratuitement le **retry** sur verrou et la **traduction d'erreurs**.

## Auteur

**Nicolas DEOUX** — NDXDev@gmail.com
