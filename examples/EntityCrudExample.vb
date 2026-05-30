' ============================================================================
' NDXAccess - Exemple : entité métier "Active Record" sur Access
' ============================================================================
' Démontre, sur une base Access, le pattern d'entité CRUD « Active Record » en
' utilisant directement OleDbCommand / OleDbParameter / OleDbDataAdapter.Fill /
' ExecuteScalar / ExecuteNonQuery, tout en s'appuyant sur NDXAccess pour le cycle de
' vie de la connexion, les transactions, le retry et la traduction des erreurs.
'
' PIÈGES ACCESS mis en évidence (spécificités à connaître) :
'   - Paramètres POSITIONNELS : le SQL utilise '?', et l'ORDRE des paramètres
'     ajoutés doit suivre l'ordre des '?' (les noms sont ignorés par OLE DB).
'   - Pas de multi-instructions : "INSERT ...; SELECT @@IDENTITY;" est INTERDIT.
'     On exécute donc DEUX commandes : l'INSERT puis "SELECT @@IDENTITY".
'   - "SELECT TOP 1 ..." au lieu de "... LIMIT 1".
'   - Les DateTime sont typés OleDbType.Date (sinon "Data type mismatch").
'
' NOTE : exemple de documentation (projet compilé, non exécuté).
' Auteur : Nicolas DEOUX <NDXDev@gmail.com>
' ============================================================================

Imports System.Data
Imports System.Data.OleDb
Imports NDXAccess

Namespace NDXAccess.Examples

    ''' <summary>
    ''' Entité liée à une étape (table tb_Etapes). Pattern Active Record :
    ''' Load / Save / Remove + collection enfant chargée à la demande.
    ''' </summary>
    Public Class Etape
        Implements IDisposable

        Public Enum EnumFonction
            Continuer = 0
            Auto = 1
            Bloquante = 2
        End Enum

        Private Shared ReadOnly DefaultDate As DateTime = New DateTime(1900, 1, 1)

#Region "Variables et propriétés"

        ' Connexion : transmise (partagée, non disposée) OU possédée (créée ici, disposée).
        Private ReadOnly _ownsConnection As Boolean
        Private _access As IAccessConnection

        ''' <summary>Connexion NDXAccess sous-jacente.</summary>
        Public ReadOnly Property Access As IAccessConnection
            Get
                Return _access
            End Get
        End Property

        Private _id As Integer
        ''' <summary>Identifiant (AutoNumber). -1 si non persisté.</summary>
        Public Property Id As Integer
            Get
                Return _id
            End Get
            Set(value As Integer)
                If value >= -1 Then _id = value
            End Set
        End Property

        Private _libelle As String
        Public Property Libelle As String
            Get
                Return _libelle
            End Get
            Set(value As String)
                _libelle = value
            End Set
        End Property

        Private _fkChapitre As Integer
        Public Property FkChapitre As Integer
            Get
                Return _fkChapitre
            End Get
            Set(value As Integer)
                If value >= 0 Then _fkChapitre = value
            End Set
        End Property

        Private _fonction As EnumFonction
        Public Property Fonction As EnumFonction
            Get
                Return _fonction
            End Get
            Set(value As EnumFonction)
                _fonction = value
            End Set
        End Property

        Private _numero As Integer
        Public Property Numero As Integer
            Get
                Return _numero
            End Get
            Set(value As Integer)
                If value >= 0 Then _numero = value
            End Set
        End Property

        Private _parQui As String
        Public ReadOnly Property ParQui As String
            Get
                Return _parQui
            End Get
        End Property

        Private _dateCrea As DateTime
        Public ReadOnly Property DateCrea As DateTime
            Get
                Return _dateCrea
            End Get
        End Property

        Private _dateModif As DateTime
        Public ReadOnly Property DateModif As DateTime
            Get
                Return _dateModif
            End Get
        End Property

        Private _listeModules As List(Of Modl)
        ''' <summary>Modules de l'étape, chargés à la demande (lazy).</summary>
        Public ReadOnly Property ListeModules As List(Of Modl)
            Get
                If _listeModules Is Nothing Then
                    _listeModules = New List(Of Modl)()
                    ListerModules()
                End If
                Return _listeModules
            End Get
        End Property

#End Region

#Region "Construction / Destruction"

        ''' <summary>Crée l'entité avec une connexion PARTAGÉE (non disposée par l'entité).</summary>
        Public Sub New(sharedConnection As IAccessConnection)
            ArgumentNullException.ThrowIfNull(sharedConnection)
            _access = sharedConnection
            _ownsConnection = False
            Raz()
        End Sub

        ''' <summary>Crée l'entité avec sa PROPRE connexion (disposée par l'entité).</summary>
        Public Sub New(databasePath As String, Optional password As String = Nothing)
            _access = New AccessConnection(New AccessConnectionOptions With {
                .DatabasePath = databasePath,
                .Password = If(password, String.Empty)
            })
            _ownsConnection = True
            Raz()
        End Sub

        Private disposedValue As Boolean

        Protected Overridable Sub Dispose(disposing As Boolean)
            If disposedValue Then Return
            If disposing Then
                ' On ne dispose que la connexion que l'on possède.
                If _ownsConnection AndAlso _access IsNot Nothing Then
                    _access.Dispose()
                End If
                _access = Nothing
            End If
            disposedValue = True
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub

#End Region

#Region "Initialisation"

        Private Sub SetDefaultValues()
            _libelle = ""
            _dateCrea = DefaultDate
            _dateModif = DefaultDate
            _parQui = ""
            _numero = -1
            _fonction = EnumFonction.Continuer
            _fkChapitre = -1
        End Sub

        Private Sub Raz()
            _id = -1
            SetDefaultValues()
        End Sub

        ''' <summary>Recharge l'entité depuis la base à partir de son Id.</summary>
        Public Function Refresh() As Boolean
            SetDefaultValues()
            Return LoadFromID(_id)
        End Function

#End Region

#Region "CRUD - lecture (OleDbCommand + OleDbDataAdapter.Fill)"

        ''' <summary>Charge une étape par son Id (SELECT TOP 1 + da.Fill).</summary>
        Public Function LoadFromID(inId As Integer) As Boolean
            Dim retour = False
            ' '?' positionnel : un seul paramètre ici (l'Id).
            Const sql As String = "SELECT TOP 1 * FROM tb_Etapes WHERE id = ?"
            Dim dt As New DataTable()

            Try
                _access.Open()
                Using cmd = _access.CreateCommand(sql)
                    cmd.Parameters.Add("id", OleDbType.Integer).Value = inId
                    Using da As New OleDbDataAdapter(cmd)
                        da.Fill(dt)
                    End Using
                End Using

                If dt.Rows.Count > 0 Then
                    retour = True
                    MapRow(dt.Rows(0))
                End If
            Catch ex As OleDbException
                ' En passant par OleDbCommand directement, l'exception n'est pas traduite
                ' par NDXAccess : on la gère ici (ou on routerait via _access.ExecuteQuery).
                Console.WriteLine($"LoadFromID a échoué : {ex.Message}")
            End Try

            Return retour
        End Function

        ''' <summary>Charge une étape par (Numéro, Chapitre), puis délègue à LoadFromID.</summary>
        Public Function LoadFrom(inNumEtape As Integer, inChapitre As Integer) As Boolean
            Dim retour = False
            ' Deux '?' : ORDRE = (Numero, fkChapitre), comme dans le SQL.
            Const sql As String = "SELECT TOP 1 id FROM tb_Etapes WHERE Numero = ? AND fkChapitre = ?"
            Dim dt As New DataTable()

            Try
                _access.Open()
                Using cmd = _access.CreateCommand(sql)
                    cmd.Parameters.Add("Numero", OleDbType.Integer).Value = inNumEtape
                    cmd.Parameters.Add("fkChapitre", OleDbType.Integer).Value = inChapitre
                    Using da As New OleDbDataAdapter(cmd)
                        da.Fill(dt)
                    End Using
                End Using

                If dt.Rows.Count > 0 AndAlso Not Convert.IsDBNull(dt.Rows(0)("id")) Then
                    retour = LoadFromID(Convert.ToInt32(dt.Rows(0)("id")))
                End If
            Catch ex As OleDbException
                Console.WriteLine($"LoadFrom a échoué : {ex.Message}")
            End Try

            Return retour
        End Function

        Private Sub MapRow(row As DataRow)
            If Not Convert.IsDBNull(row("id")) Then _id = Convert.ToInt32(row("id"))
            If Not Convert.IsDBNull(row("Libelle")) Then _libelle = Convert.ToString(row("Libelle"))
            If Not Convert.IsDBNull(row("fkChapitre")) Then _fkChapitre = Convert.ToInt32(row("fkChapitre"))
            If Not Convert.IsDBNull(row("Fonction")) Then _fonction = CType(Convert.ToInt32(row("Fonction")), EnumFonction)
            If Not Convert.IsDBNull(row("Numero")) Then _numero = Convert.ToInt32(row("Numero"))
            If Not Convert.IsDBNull(row("DateCrea")) Then _dateCrea = Convert.ToDateTime(row("DateCrea"))
            If Not Convert.IsDBNull(row("DateModif")) Then _dateModif = Convert.ToDateTime(row("DateModif"))
            If Not Convert.IsDBNull(row("ParQui")) Then _parQui = Convert.ToString(row("ParQui"))
        End Sub

#End Region

#Region "CRUD - écriture (paramètres typés + @@IDENTITY)"

        ''' <summary>
        ''' INSERT (si Id = -1) ou UPDATE. En INSERT, récupère le nouvel Id via une
        ''' SECONDE commande "SELECT @@IDENTITY" (Access n'accepte pas le multi-statement).
        ''' </summary>
        Public Function Save() As Boolean
            Dim retour = False

            Const sqlInsert As String =
                "INSERT INTO tb_Etapes (Libelle, Numero, Fonction, fkChapitre, DateCrea, ParQui) " &
                "VALUES (?, ?, ?, ?, ?, ?)"

            ' ATTENTION : en UPDATE, le paramètre du WHERE (id) vient EN DERNIER,
            ' car '?' est positionnel.
            Const sqlUpdate As String =
                "UPDATE tb_Etapes SET Libelle = ?, Numero = ?, Fonction = ?, fkChapitre = ?, " &
                "DateModif = ?, ParQui = ? WHERE id = ?"

            Try
                _access.Open()

                If _id = -1 Then
                    ' --- INSERT ---
                    Using cmd = _access.CreateCommand(sqlInsert)
                        cmd.Parameters.Add("Libelle", OleDbType.VarWChar).Value = _libelle
                        cmd.Parameters.Add("Numero", OleDbType.Integer).Value = _numero
                        cmd.Parameters.Add("Fonction", OleDbType.Integer).Value = CInt(_fonction)
                        cmd.Parameters.Add("fkChapitre", OleDbType.Integer).Value = _fkChapitre
                        cmd.Parameters.Add("DateCrea", OleDbType.Date).Value = DateTime.Now
                        cmd.Parameters.Add("ParQui", OleDbType.VarWChar).Value = Environment.UserName
                        cmd.ExecuteNonQuery()
                    End Using

                    ' Récupération de l'identité sur la MÊME connexion (commande séparée).
                    Using idCmd = _access.CreateCommand("SELECT @@IDENTITY")
                        Dim newId = idCmd.ExecuteScalar()
                        If newId IsNot Nothing AndAlso Not Convert.IsDBNull(newId) Then
                            _id = Convert.ToInt32(newId)
                            retour = _id > 0
                        End If
                    End Using
                Else
                    ' --- UPDATE ---
                    Using cmd = _access.CreateCommand(sqlUpdate)
                        cmd.Parameters.Add("Libelle", OleDbType.VarWChar).Value = _libelle
                        cmd.Parameters.Add("Numero", OleDbType.Integer).Value = _numero
                        cmd.Parameters.Add("Fonction", OleDbType.Integer).Value = CInt(_fonction)
                        cmd.Parameters.Add("fkChapitre", OleDbType.Integer).Value = _fkChapitre
                        cmd.Parameters.Add("DateModif", OleDbType.Date).Value = DateTime.Now
                        cmd.Parameters.Add("ParQui", OleDbType.VarWChar).Value = Environment.UserName
                        cmd.Parameters.Add("id", OleDbType.Integer).Value = _id   ' WHERE id = ? -> en dernier
                        retour = cmd.ExecuteNonQuery() > 0
                    End Using
                End If
            Catch ex As OleDbException
                retour = False
                Console.WriteLine($"Save a échoué : {ex.Message}")
            End Try

            Return retour
        End Function

        ''' <summary>Supprime l'étape courante.</summary>
        Public Function Remove() As Boolean
            If _id = -1 Then Return False
            Dim retour = False
            Const sql As String = "DELETE FROM tb_Etapes WHERE id = ?"

            Try
                _access.Open()
                Using cmd = _access.CreateCommand(sql)
                    cmd.Parameters.Add("id", OleDbType.Integer).Value = _id
                    retour = cmd.ExecuteNonQuery() > 0
                End Using
            Catch ex As OleDbException
                Console.WriteLine($"Remove a échoué : {ex.Message}")
            End Try

            Return retour
        End Function

        ''' <summary>Charge les modules liés (table de jointure tb_Etape_Modules).</summary>
        Private Sub ListerModules()
            Const sql As String = "SELECT fkModl FROM tb_Etape_Modules WHERE fkEtape = ?"
            Dim dt As New DataTable()

            Try
                _access.Open()
                Using cmd = _access.CreateCommand(sql)
                    cmd.Parameters.Add("fkEtape", OleDbType.Integer).Value = _id
                    Using da As New OleDbDataAdapter(cmd)
                        da.Fill(dt)
                    End Using
                End Using

                For Each row As DataRow In dt.Rows
                    If Not Convert.IsDBNull(row("fkModl")) Then
                        Dim md As New Modl(_access)
                        If md.LoadFromID(Convert.ToInt32(row("fkModl"))) Then
                            _listeModules.Add(md)
                        End If
                    End If
                Next
            Catch ex As OleDbException
                Console.WriteLine($"ListerModules a échoué : {ex.Message}")
            End Try
        End Sub

#End Region

#Region "Variante haut niveau (équivalent via l'API NDXAccess)"

        ''' <summary>
        ''' Équivalent de LoadFromID en passant par l'API haut niveau : retry automatique
        ''' sur verrou transitoire + traduction des erreurs (AccessQueryException) en bonus.
        ''' </summary>
        Public Function LoadFromIDViaConnector(inId As Integer) As Boolean
            Try
                ' Mapping objet direct, ou utilisez ExecuteQuery (DataTable) selon le besoin.
                Dim dt = _access.ExecuteQuery("SELECT TOP 1 * FROM tb_Etapes WHERE id = ?", inId)
                If dt.Rows.Count = 0 Then Return False
                MapRow(dt.Rows(0))
                Return True
            Catch ex As AccessQueryException
                Console.WriteLine($"Erreur claire : {ex.Message} (transitoire={ex.IsTransient})")
                Return False
            End Try
        End Function

#End Region

    End Class

    ''' <summary>Entité minimale liée à un module (pour l'exemple de collection enfant).</summary>
    Public Class Modl

        Private ReadOnly _access As IAccessConnection

        Public Property Id As Integer = -1
        Public Property Libelle As String = ""

        Public Sub New(sharedConnection As IAccessConnection)
            ArgumentNullException.ThrowIfNull(sharedConnection)
            _access = sharedConnection
        End Sub

        ''' <summary>Charge le module par Id via l'API haut niveau (retry + erreurs claires).</summary>
        Public Function LoadFromID(inId As Integer) As Boolean
            Try
                Dim dt = _access.ExecuteQuery("SELECT TOP 1 * FROM tb_Modules WHERE id = ?", inId)
                If dt.Rows.Count = 0 Then Return False
                Dim row = dt.Rows(0)
                If Not Convert.IsDBNull(row("id")) Then Id = Convert.ToInt32(row("id"))
                If Not Convert.IsDBNull(row("Libelle")) Then Libelle = Convert.ToString(row("Libelle"))
                Return True
            Catch ex As AccessQueryException
                Console.WriteLine($"Modl.LoadFromID a échoué : {ex.Message}")
                Return False
            End Try
        End Function

    End Class

End Namespace
