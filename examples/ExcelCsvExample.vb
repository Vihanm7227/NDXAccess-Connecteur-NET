' ============================================================================
' NDXAccess - Lire des fichiers Excel / CSV via le provider ACE
' ============================================================================
' Capacité méconnue du provider ACE : il sait requêter des classeurs Excel (.xlsx)
' et des fichiers texte (.csv) comme s'il s'agissait de tables, via une chaîne de
' connexion personnalisée (option ConnectionString) avec "Extended Properties".
'
' - CSV  : Data Source = le DOSSIER ; table = [fichier.csv]
' - Excel: Data Source = le FICHIER .xlsx ; table = [Feuille1$]
'
' NOTE : exemple de documentation (projet compilé ; le CSV est testé).
' Auteur : Nicolas DEOUX <NDXDev@gmail.com>
' ============================================================================

Imports System.Data
Imports NDXAccess

Namespace NDXAccess.Examples

    Public Module ExcelCsvExample

        ''' <summary>
        ''' Lit un fichier CSV (1re ligne = en-têtes) via ACE. <paramref name="folderPath"/>
        ''' est le DOSSIER contenant le fichier ; <paramref name="fileName"/> le nom du .csv.
        ''' </summary>
        Public Function ReadCsv(folderPath As String, fileName As String) As DataTable
            Dim cs = $"Provider=Microsoft.ACE.OLEDB.16.0;Data Source={folderPath};" &
                     "Extended Properties=""text;HDR=Yes;FMT=Delimited"""

            Using connection As IAccessConnection = New AccessConnection(New AccessConnectionOptions With {.ConnectionString = cs})
                Return connection.ExecuteQuery($"SELECT * FROM [{fileName}]")
            End Using
        End Function

        ''' <summary>
        ''' Lit une feuille d'un classeur Excel (.xlsx) via ACE.
        ''' <paramref name="sheet"/> est le nom de la feuille suivi de '$' (ex. "Feuil1$").
        ''' </summary>
        Public Function ReadExcelSheet(workbookPath As String, sheet As String) As DataTable
            Dim cs = $"Provider=Microsoft.ACE.OLEDB.16.0;Data Source={workbookPath};" &
                     "Extended Properties=""Excel 12.0 Xml;HDR=YES"""

            Using connection As IAccessConnection = New AccessConnection(New AccessConnectionOptions With {.ConnectionString = cs})
                Return connection.ExecuteQuery($"SELECT * FROM [{sheet}]")
            End Using
        End Function

    End Module

End Namespace
