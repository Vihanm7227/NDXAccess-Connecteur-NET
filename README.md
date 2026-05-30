# NDXAccess

**Connecteur Microsoft Access (.accdb / .mdb) moderne pour VB.NET**

[![.NET](https://img.shields.io/badge/.NET-8.0--windows-512BD4?style=flat-square&logo=dotnet)](https://dotnet.microsoft.com/)
[![Access](https://img.shields.io/badge/Provider-ACE.OLEDB.16.0-A4373A?style=flat-square&logo=microsoftaccess)](https://www.microsoft.com/download/details.aspx?id=54920)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg?style=flat-square)](LICENSE)
[![Platform](https://img.shields.io/badge/Platform-Windows%20(x86%20%7C%20x64)-blue?style=flat-square)]()
[![Tests](https://img.shields.io/badge/Tests-98%20passed-brightgreen?style=flat-square)]()

---

## Pourquoi NDXAccess ?

L'accès aux bases Access en .NET repose encore souvent sur du code ADO.NET répétitif et
bas niveau. NDXAccess apporte une **API fluide et moderne** pour VB.NET — injection de
dépendances, logging, health checks et compactage — tout en encapsulant les **pièges
spécifiques d'Access** (architecture x86/x64, verrouillage, limite des 2 Go, etc.).

- **API fluide** : CRUD paramétré, transactions, requêtes enregistrées
- **Sync ET async** : double API (l'async est « de façade », voir plus bas — on est honnête)
- **Résilience** : retry automatique avec back-off sur les verrous transitoires d'Access
- **Erreurs claires** : les `OleDbException` cryptiques deviennent des `AccessQueryException` lisibles
- **Détection x86/x64** : message clair si le provider ACE ne correspond pas à l'appli
- **Versions exposées** : version de la bibliothèque, du **moteur ACE** et du format de fichier
- **Helpers de schéma** : `TableExists`, `GetTableNames`, `GetQueryNames`, `GetColumns`
- **Création de base** : `CreateDatabase` (génère un `.accdb` vide via ADOX)
- **Insertion en masse** : `BulkInsert` (transaction unique, bien plus rapide)
- **Mapping objet** : `ExecuteQuery(Of T)` (micro-ORM par réflexion)
- **Paramètres nommés** (option) : `@nom` traduits en `?` positionnels
- **Compactage encapsulé** : `CompactDatabase` / `CompactDatabaseAsync` (via DAO)
- **Health checks** : connectivité + surveillance de la limite des **2 Go**
- **Mot de passe** : encapsulé proprement dans les options
- **Logging** : compatible `Microsoft.Extensions.Logging`
- **Injection de dépendances** : `AddNDXAccess(...)`

> ⚠️ **Honnêteté avant tout** : Access n'a pas les mêmes capacités qu'un SGBD client-serveur.
> Lisez la section [« Ce qu'Access sait et ne sait pas faire »](#ce-quaccess-sait-et-ne-sait-pas-faire).

---

## Prérequis

- **Windows** (le provider ACE n'existe pas ailleurs)
- **.NET 8.0** (cible `net8.0-windows`)
- **Microsoft Access Database Engine 2016 Redistributable** (gratuit), dans la **même
  architecture (x86/x64)** que votre application
  → [Guide d'installation détaillé](docs/installation/README.md)

---

## Installation rapide

```powershell
git clone https://github.com/NDXDeveloper/NDXAccess-Connecteur-NET.git
dotnet add reference chemin/vers/src/NDXAccess/NDXAccess.vbproj
```

Ou copiez le dossier `src/NDXAccess` dans votre solution.

---

## Démarrage en 30 secondes

```vb
Imports NDXAccess

Dim options As New AccessConnectionOptions With {
    .DatabasePath = "C:\data\ma_base.accdb"
}

Using connection As IAccessConnection = New AccessConnection(options)
    Dim count = Await connection.ExecuteScalarAsync(Of Integer)("SELECT COUNT(*) FROM clients")
    Console.WriteLine($"Nombre de clients : {count}")
End Using
```

> 💡 **Paramètres positionnels** : Access utilise `?` (pas `@nom`). Les valeurs sont
> passées **dans l'ordre**.

---

## Fonctionnalités principales

### CRUD paramétré (sync + async)

```vb
' INSERT (les '?' sont remplis dans l'ordre)
Await connection.ExecuteNonQueryAsync(
    "INSERT INTO clients (nom, email) VALUES (?, ?)",
    {"Jean", "jean@example.com"})

' SELECT -> DataTable
Dim clients = Await connection.ExecuteQueryAsync(
    "SELECT * FROM clients WHERE actif = ?", {True})

' SELECT scalaire typé
Dim total = Await connection.ExecuteScalarAsync(Of Integer)("SELECT COUNT(*) FROM clients")

' UPDATE / DELETE
Await connection.ExecuteNonQueryAsync("UPDATE clients SET email = ? WHERE id = ?", {"x@y.z", 1})
Await connection.ExecuteNonQueryAsync("DELETE FROM clients WHERE id = ?", {1})

' Variantes synchrones (ParamArray)
Dim n = connection.ExecuteNonQuery("DELETE FROM clients WHERE id = ?", 1)
```

### Transactions

```vb
Await connection.BeginTransactionAsync()
Try
    Await connection.ExecuteNonQueryAsync("UPDATE comptes SET solde = solde - ? WHERE id = ?", {100D, 1})
    Await connection.ExecuteNonQueryAsync("UPDATE comptes SET solde = solde + ? WHERE id = ?", {100D, 2})
    Await connection.CommitAsync()
Catch
    Await connection.RollbackAsync()
    Throw
End Try
```

### Requêtes enregistrées (paramètres IN uniquement)

```vb
' Création d'une requête enregistrée paramétrée
connection.ExecuteNonQuery(
    "CREATE PROCEDURE qParStatut (prmActif BIT) AS " &
    "SELECT * FROM clients WHERE actif = prmActif")

' Appel avec paramètre d'entrée
Dim dt = Await connection.ExecuteStoredQueryAsync("qParStatut", {True})
```

### Détection x86 / x64 du provider

```vb
Console.WriteLine(AccessProviderHelper.CurrentProcessArchitecture)  ' "x86" ou "x64"
AccessProviderHelper.EnsureProviderAvailable("Microsoft.ACE.OLEDB.16.0")
' -> AccessProviderNotFoundException avec message clair si mismatch d'architecture
```

### Health check (avec surveillance des 2 Go)

```vb
Dim hc = New AccessHealthCheck(factory)
Dim result = Await hc.CheckHealthAsync()

If result.IsHealthy Then
    Dim info = result.DatabaseInfo
    Console.WriteLine($"{info.FileSizeMegabytes:F1} Mo / 2 Go ({info.UsagePercent:F1} %)")
End If
```

### Compactage / réparation

```vb
' Vrai async : opération bloquante déportée sur un thread
Await connection.CompactDatabaseAsync()

' Vers un nouveau fichier (synchrone)
connection.CompactDatabase("C:\data\ma_base_compactee.accdb")
```

### Injection de dépendances

```vb
services.AddNDXAccess(Sub(options)
                          options.DatabasePath = "C:\data\ma_base.accdb"
                      End Sub)

' Puis injectez IAccessConnectionFactory, IAccessConnection ou AccessHealthCheck
```

### Mot de passe

```vb
Dim options As New AccessConnectionOptions With {
    .DatabasePath = "C:\data\securisee.accdb",
    .Password = "mon_mot_de_passe"   ' -> Jet OLEDB:Database Password
}
```

---

## Fonctionnalités avancées

### Résilience (retry) + erreurs claires

```vb
' Le retry sur verrous transitoires est actif par défaut (jamais dans une transaction).
Dim options As New AccessConnectionOptions With {
    .DatabasePath = "C:\data\partage.accdb",
    .EnableRetryOnTransientErrors = True,   ' défaut
    .MaxRetries = 3,                        ' défaut
    .RetryBaseDelayMs = 100                 ' back-off exponentiel
}

Try
    Await connection.ExecuteNonQueryAsync("UPDATE t SET v = ? WHERE id = ?", {1, 2})
Catch ex As AccessQueryException
    ' Message clair + code natif ACE + IsTransient + OleDbException d'origine préservée
    Console.WriteLine($"{ex.Message} (transitoire={ex.IsTransient})")
End Try
```

### Mapping objet — `ExecuteQuery(Of T)`

```vb
Public Class Client
    Public Property Id As Integer
    Public Property Nom As String
    Public Property Email As String
End Class

Dim clients = Await connection.ExecuteQueryAsync(Of Client)(
    "SELECT id, nom, email FROM clients WHERE actif = ?", {True})
```

### Insertion en masse — `BulkInsert`

```vb
Dim rows As New List(Of Object())()
For i = 1 To 10000
    rows.Add(New Object() {$"Nom{i}", i * 1.5D})
Next
Dim inserted = Await connection.BulkInsertAsync("clients", {"nom", "montant"}, rows)
```

### Helpers de schéma

```vb
If Not connection.TableExists("clients") Then ...
For Each t In connection.GetTableNames() : Console.WriteLine(t) : Next
Dim queries = connection.GetQueryNames()      ' requêtes enregistrées
Dim cols = connection.GetColumns("clients")   ' DataTable des colonnes
```

### Créer une base par programme

```vb
AccessConnection.CreateDatabase("C:\data\nouvelle.accdb")           ' synchrone
Await AccessConnection.CreateDatabaseAsync("C:\data\autre.accdb", password:="secret")
```

### Paramètres nommés (optionnel)

```vb
Dim p As New Dictionary(Of String, Object) From {{"actif", True}}
Dim dt = Await connection.ExecuteQueryNamedAsync(
    "SELECT nom FROM clients WHERE actif = @actif", p)   ' @actif -> ? automatiquement
```

### Version de la bibliothèque

```vb
Console.WriteLine(AccessConnection.Version)              ' "1.2.0" (raccourci)
Console.WriteLine(NDXAccessInfo.InformationalVersion)    ' "1.2.0"
Console.WriteLine(NDXAccessInfo.Version)                 ' "1.2.0.0" (version d'assembly)
Console.WriteLine(NDXAccessInfo.ProductName)             ' "NDXAccess"
```

### Version du moteur ACE

```vb
' Sur la connexion (aucune ouverture requise — lecture du registre / nom du provider)
Console.WriteLine(connection.ProviderName)    ' "Microsoft.ACE.OLEDB.16.0"
Console.WriteLine(connection.EngineVersion)   ' "16.0.5011.1000" (version réelle du DLL ACE)

' Ou via AccessProviderHelper, sans connexion
Console.WriteLine(AccessProviderHelper.GetEngineVersion())   ' "16.0.5011.1000"

' Via le health check : moteur ACE vs format de fichier
Dim info = New AccessHealthCheck(factory).GetDatabaseInfo()
Console.WriteLine(info.EngineVersion)      ' "16.0.5011.1000" (moteur ACE)
Console.WriteLine(info.FileFormatVersion)  ' "04.00.0000"     (format du fichier)
```

> **Trois versions distinctes** : celle de la **bibliothèque** NDXAccess (`AccessConnection.Version`),
> celle du **moteur ACE** (`connection.EngineVersion`), et celle du **format de fichier**
> (`DatabaseInfo.FileFormatVersion`).

---

## Ce qu'Access sait et ne sait pas faire

| Capacité | Access (NDXAccess) |
|----------|--------------------|
| CRUD paramétré, transactions | ✅ Oui (paramètres positionnels `?`) |
| Requêtes enregistrées | ✅ Oui, **paramètres IN uniquement** |
| Procédures stockées IN/OUT/INOUT | ❌ Non (pas de OUT/INOUT dans Access) |
| Event Scheduler / tâches planifiées | ❌ Non → **Planificateur de tâches Windows** |
| Vrai parallélisme async | ❌ Non — async **« de façade »** (sauf `CompactDatabaseAsync`) |
| Compactage/réparation | ✅ Oui (via DAO) |
| Multiplateforme | ❌ **Windows uniquement** (provider ACE) |
| Forte concurrence | ⚠️ ~10-15 utilisateurs simultanés maximum |
| Taille de base | ⚠️ **2 Go maximum** par fichier |

### L'async « de façade », expliqué

`OleDbCommand.ExecuteNonQueryAsync` & co. **ne sont pas réellement asynchrones** : ce
sont des wrappers qui s'exécutent de manière synchrone sur le thread courant. NDXAccess
expose quand même une API `...Async` pour la cohérence et l'usage de `Await`, mais
**n'attendez pas de gain de parallélisme**. La seule exception est `CompactDatabaseAsync`,
réellement déportée sur le pool de threads car l'opération est longue et bloquante.

---

## Structure du projet

```
NDXAccess-Connecteur-NET/
├── src/NDXAccess/              # Bibliothèque (VB.NET)
│   ├── AccessConnection.vb
│   ├── AccessConnectionOptions.vb
│   ├── AccessConnectionFactory.vb
│   ├── AccessHealthCheck.vb
│   ├── AccessProviderHelper.vb
│   ├── AccessMaintenance.vb
│   ├── AccessExceptions.vb
│   └── Extensions/
├── tests/NDXAccess.Tests/      # 44 tests (unitaires + intégration)
├── examples/                   # Exemples d'utilisation
└── docs/                       # Documentation (source, tests, installation)
```

---

## Construire en x86 ou x64

Le binaire est **AnyCPU** ; c'est l'application hôte qui impose le bitness (et donc l'ACE
requis). Le projet expose `AnyCPU`, `x86` et `x64` :

```powershell
dotnet build NDXAccess.sln                 # AnyCPU
dotnet build NDXAccess.sln -p:Platform=x86 # 32 bits (ACE x86)
dotnet build NDXAccess.sln -p:Platform=x64 # 64 bits (ACE x64)
```

---

## Tests

```powershell
dotnet test                                  # tout
dotnet test --filter "Category=Integration"  # tests d'intégration uniquement
dotnet test --filter "Category=Unit"         # tests unitaires uniquement
```

**98 tests** : 37 unitaires (sans base) + 61 d'intégration (sur une base `.accdb` réelle).
La plupart des exemples (entité Active Record, DataAdapter, démarrage tout-en-un, pagination,
DDL/relations, lecture CSV) ainsi que la résilience, le mot de passe et l'injection de
dépendances sont **testés de bout en bout**.
Les tests d'intégration sont **automatiquement ignorés** si le provider ACE n'est pas
disponible pour l'architecture courante. Voir [docs/tests](docs/tests/README.md).

L'intégration continue ([.github/workflows/ci.yml](.github/workflows/ci.yml)) compile
AnyCPU/x86/x64 sur Windows et publie le package NuGet (`.nupkg` + symboles `.snupkg`).

---

## Documentation

Sommaire complet : [SOMMAIRE.md](SOMMAIRE.md).

- [Installation & x86/x64](docs/installation/README.md)
- [Dépannage & FAQ](docs/troubleshooting/README.md)
- [Documentation source](docs/source/README.md)
- [Documentation des tests](docs/tests/README.md)
- [Exemples](examples/README.md)
- [CHANGELOG](CHANGELOG.md) · [CONTRIBUTING](CONTRIBUTING.md) · [SECURITY](SECURITY.md)

---

## Prérequis (résumé)

- Windows + .NET 8.0
- Microsoft Access Database Engine 2016 (x86 ou x64 selon l'appli)
- Visual Studio Community 2022 (ou `dotnet` CLI)

---

## Auteur

**Nicolas DEOUX**
- Email : [NDXDev@gmail.com](mailto:NDXDev@gmail.com)
- LinkedIn : [nicolas-deoux-ab295980](https://www.linkedin.com/in/nicolas-deoux-ab295980/)

---

## Licence

Projet sous licence MIT. Voir [LICENSE](LICENSE).

---

<p align="center"><b>Fait avec passion en France</b></p>
