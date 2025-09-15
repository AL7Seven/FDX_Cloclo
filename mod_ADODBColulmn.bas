Attribute VB_Name = "mod_ADODBColulmn"
' ============================================================================
' COUCHE ADODB D�DI�E - R�SOLUTION WORKBOOK FERM�
' ============================================================================
' Module sp�cialis� pour acc�s Excel via ADODB sans ouvrir le workbook
' R�solution plages nomm�es et tableaux structur�s
' Int�gration transparente avec le syst�me de parsing existant
' ============================================================================

Option Explicit

' ============================================================================
' TYPES D�DI�S ADODB
' ============================================================================
' Note: ExcelResolutionContext et ComparisonContext sont d�finis dans mod_Global

Public Type ADODBConnectionInfo
    Provider As String
    DataSource As String
    ExtendedProperties As String
    ConnectionString As String
    IsValid As Boolean
    LastError As String
End Type

Public Type ADODBQueryResult
    Success As Boolean
    RecordCount As Long
    FieldCount As Long
    ErrorMessage As String
    QueryTime As Single
End Type

' ============================================================================
' CONSTANTES ADODB
' ============================================================================

Private Const EXCEL_PROVIDER_ACE As String = "Microsoft.ACE.OLEDB.12.0"
Private Const EXCEL_PROVIDER_JET As String = "Microsoft.Jet.OLEDB.4.0"
Private Const EXCEL_EXTENDED_PROPS As String = "Excel 12.0 Xml;HDR=Yes;IMEX=1"
Private Const EXCEL_EXTENDED_PROPS_LEGACY As String = "Excel 8.0;HDR=Yes;IMEX=1"

Private Const ADODB_TIMEOUT_SECONDS As Long = 30
Private Const ADODB_MAX_RETRIES As Long = 3

' ============================================================================
' ResolveFromClosedWorkbookADODB - R�SOLUTION COMPL�TE VIA ADODB
' ============================================================================
Public Function ResolveFromClosedWorkbookADODB(registry As Object, context As ExcelResolutionContext) As Boolean
    On Error GoTo ADODBError
    
    ResolveFromClosedWorkbookADODB = False
    
    ' Validation entr�e
    If Not ValidateADODBContext(actualContext) Then
        SetParsingError ERR_PARSING_EXCEL_ACCESS_FAILED, "Contexte ADODB invalide", actualContext.workbookPath
        Exit Function
    End If
    
    ' �tablir connexion ADODB
    Dim connInfo As ADODBConnectionInfo
    If Not EstablishADODBConnection(actualContext.workbookPath, connInfo) Then
        SetParsingError ERR_PARSING_EXCEL_ACCESS_FAILED, "Impossible d'�tablir connexion ADODB: " & connInfo.LastError, actualContext.workbookPath
        Exit Function
    End If
    
    ' Analyser structure workbook
    Dim workbookStructure As Object
    Set workbookStructure = AnalyzeWorkbookStructure(connInfo)
    
    If workbookStructure Is Nothing Then
        SetParsingError ERR_PARSING_EXCEL_ACCESS_FAILED, "Impossible d'analyser structure workbook", context.workbookPath
        GoTo CleanupConnection
    End If
    
    ' R�soudre r�f�rences nomm�es
    If Not ResolveNamedReferencesADODB(registry, connInfo, workbookStructure) Then
        GoTo CleanupConnection ' Erreur d�j� d�finie
    End If
    
    ' Mettre � jour registry apr�s r�solution
    UpdateRegistryAfterADODBResolution registry
    
    ResolveFromClosedWorkbookADODB = True
    
CleanupConnection:
    CloseADODBConnection connInfo
    Exit Function
    
ADODBError:
    SetParsingError ERR_PARSING_EXCEL_ACCESS_FAILED, "Erreur ADODB: " & err.Description, context.workbookPath
    CloseADODBConnection connInfo
    ResolveFromClosedWorkbookADODB = False
End Function

' ============================================================================
' EstablishADODBConnection - �TABLISSEMENT CONNEXION S�CURIS�E
' ============================================================================
Private Function EstablishADODBConnection(workbookPath As String, ByRef connInfo As ADODBConnectionInfo) As Boolean
    On Error GoTo ConnectionError
    
    EstablishADODBConnection = False
    
    ' R�initialiser structure
    With connInfo
        .Provider = ""
        .DataSource = workbookPath
        .ExtendedProperties = ""
        .ConnectionString = ""
        .IsValid = False
        .LastError = ""
    End With
    
    ' Validation fichier
    If Not FileExists(workbookPath) Then
        connInfo.LastError = "Fichier inexistant: " & workbookPath
        Exit Function
    End If
    
    ' D�terminer provider optimal
    If Not DetermineOptimalProvider(workbookPath, connInfo) Then
        connInfo.LastError = "Impossible de d�terminer provider ADODB adapt�"
        Exit Function
    End If
    
    ' Construire cha�ne connexion
    BuildConnectionString connInfo
    
    ' Tester connexion avec retry
    If TestADODBConnection(connInfo) Then
        EstablishADODBConnection = True
    End If
    
    Exit Function
    
ConnectionError:
    connInfo.LastError = "Erreur �tablissement connexion: " & err.Description
    EstablishADODBConnection = False
End Function

' ============================================================================
' DetermineOptimalProvider - D�TECTION PROVIDER OPTIMAL
' ============================================================================
Private Function DetermineOptimalProvider(filePath As String, ByRef connInfo As ADODBConnectionInfo) As Boolean
    On Error GoTo ProviderError
    
    DetermineOptimalProvider = False
    
    ' Analyser extension fichier
    Dim fileExt As String: fileExt = UCase(Right(filePath, 4))
    
    Select Case fileExt
        Case ".XLS"
            ' Excel 97-2003 - essayer JET puis ACE
            If IsProviderAvailable(EXCEL_PROVIDER_JET) Then
                connInfo.Provider = EXCEL_PROVIDER_JET
                connInfo.ExtendedProperties = EXCEL_EXTENDED_PROPS_LEGACY
                DetermineOptimalProvider = True
            ElseIf IsProviderAvailable(EXCEL_PROVIDER_ACE) Then
                connInfo.Provider = EXCEL_PROVIDER_ACE
                connInfo.ExtendedProperties = EXCEL_EXTENDED_PROPS_LEGACY
                DetermineOptimalProvider = True
            End If
            
        Case ".XLX", ".LSM" ' .XLSX, .XLSM
            ' Excel 2007+ - ACE uniquement
            If IsProviderAvailable(EXCEL_PROVIDER_ACE) Then
                connInfo.Provider = EXCEL_PROVIDER_ACE
                connInfo.ExtendedProperties = EXCEL_EXTENDED_PROPS
                DetermineOptimalProvider = True
            End If
            
        Case Else
            ' Extension inconnue - essayer ACE par d�faut
            If IsProviderAvailable(EXCEL_PROVIDER_ACE) Then
                connInfo.Provider = EXCEL_PROVIDER_ACE
                connInfo.ExtendedProperties = EXCEL_EXTENDED_PROPS
                DetermineOptimalProvider = True
            End If
    End Select
    
    Exit Function
    
ProviderError:
    DetermineOptimalProvider = False
End Function

' ============================================================================
' IsProviderAvailable - TEST DISPONIBILIT� PROVIDER
' ============================================================================
Private Function IsProviderAvailable(providerName As String) As Boolean
    On Error Resume Next
    
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    
    If Not conn Is Nothing Then
        conn.Provider = providerName
        IsProviderAvailable = (err.Number = 0)
        conn.Close
    Else
        IsProviderAvailable = False
    End If
    
    Set conn = Nothing
    On Error GoTo 0
End Function

' ============================================================================
' BuildConnectionString - CONSTRUCTION CHA�NE CONNEXION
' ============================================================================
Private Sub BuildConnectionString(ByRef connInfo As ADODBConnectionInfo)
    With connInfo
        .ConnectionString = "Provider=" & .Provider & ";" & _
                           "Data Source=" & .DataSource & ";" & _
                           "Extended Properties=""" & .ExtendedProperties & """;"
    End With
End Sub

' ============================================================================
' TestADODBConnection - TEST CONNEXION AVEC RETRY
' ============================================================================
Private Function TestADODBConnection(ByRef connInfo As ADODBConnectionInfo) As Boolean
    Dim retryCount As Long: retryCount = 0
    Dim conn As Object
    
    Do While retryCount < ADODB_MAX_RETRIES
        On Error Resume Next
        
        Set conn = CreateObject("ADODB.Connection")
        
        If Not conn Is Nothing Then
            conn.CommandTimeout = ADODB_TIMEOUT_SECONDS
            conn.ConnectionTimeout = ADODB_TIMEOUT_SECONDS
            conn.Open connInfo.ConnectionString
            
            If err.Number = 0 And conn.State = 1 Then ' adStateOpen
                ' Test query simple
                Dim rs As Object
                Set rs = conn.Execute("SELECT * FROM [Sheet1$] WHERE 1=0")
                
                If err.Number = 0 Then
                    connInfo.IsValid = True
                    TestADODBConnection = True
                    rs.Close
                    conn.Close
                    Exit Function
                End If
                
                If Not rs Is Nothing Then rs.Close
            End If
            
            If conn.State = 1 Then conn.Close
        End If
        
        connInfo.LastError = "Tentative " & (retryCount + 1) & ": " & err.Description
        retryCount = retryCount + 1
        
        ' Attendre avant retry
        If retryCount < ADODB_MAX_RETRIES Then
            Application.Wait DateAdd("s", 1, Now)
        End If
        
        Set conn = Nothing
        On Error GoTo 0
    Loop
    
    TestADODBConnection = False
End Function

' ============================================================================
' AnalyzeWorkbookStructure - ANALYSE STRUCTURE VIA ADODB
' ============================================================================
Private Function AnalyzeWorkbookStructure(connInfo As ADODBConnectionInfo) As Object
    On Error GoTo StructureError
    
    Set AnalyzeWorkbookStructure = CreateObject("Scripting.Dictionary")
    
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    conn.Open connInfo.ConnectionString
    
    ' Obtenir liste feuilles/tables
    Dim sheets As Object
    Set sheets = GetWorksheetList(conn)
    AnalyzeWorkbookStructure.Add "Worksheets", sheets
    
    ' Analyser plages nomm�es (via sch�ma OLEDB)
    Dim namedRanges As Object
    Set namedRanges = GetNamedRangesList(conn)
    AnalyzeWorkbookStructure.Add "NamedRanges", namedRanges
    
    ' Analyser en-t�tes colonnes premi�re ligne chaque feuille
    Dim headers As Object
    Set headers = GetColumnHeadersList(conn, sheets)
    AnalyzeWorkbookStructure.Add "ColumnHeaders", headers
    
    conn.Close
    Set conn = Nothing
    Exit Function
    
StructureError:
    Set AnalyzeWorkbookStructure = Nothing
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If
End Function

' ============================================================================
' GetWorksheetList - LISTE FEUILLES DISPONIBLES
' ============================================================================
Private Function GetWorksheetList(conn As Object) As Object
    On Error GoTo WorksheetError
    
    Set GetWorksheetList = CreateObject("Scripting.Dictionary")
    
    Dim rs As Object
    Set rs = conn.OpenSchema(20) ' adSchemaTables
    
    Do While Not rs.EOF
        Dim tableName As String: tableName = rs.Fields("TABLE_NAME").value
        
        ' Filtrer tables Excel (feuilles se terminent par $)
        If Right(tableName, 1) = "$" And InStr(tableName, "'") = 0 Then
            Dim sheetName As String: sheetName = Left(tableName, Len(tableName) - 1)
            GetWorksheetList.Add sheetName, tableName
        End If
        
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    Exit Function
    
WorksheetError:
    Set GetWorksheetList = CreateObject("Scripting.Dictionary") ' Vide si erreur
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
End Function

' ============================================================================
' GetNamedRangesList - LISTE PLAGES NOMM�ES (TENTATIVE VIA SCH�MA)
' ============================================================================
Private Function GetNamedRangesList(conn As Object) As Object
    On Error Resume Next ' Non critique si �choue
    
    Set GetNamedRangesList = CreateObject("Scripting.Dictionary")
    
    ' Note: ADODB ne peut pas toujours acc�der aux plages nomm�es Excel
    ' Cette fonction tente diff�rentes approches mais peut retourner vide
    
    ' Approche 1: Essayer sch�ma proc�dures (parfois contient plages nomm�es)
    Dim rs As Object
    Set rs = conn.OpenSchema(16) ' adSchemaProcedures
    
    Do While Not rs.EOF And err.Number = 0
        Dim procName As String: procName = rs.Fields("PROCEDURE_NAME").value
        
        ' Les plages nomm�es apparaissent parfois comme proc�dures
        If InStr(procName, "$") = 0 And Len(procName) > 0 Then
            GetNamedRangesList.Add procName, "NAMED_RANGE"
        End If
        
        rs.MoveNext
    Loop
    
    If Not rs Is Nothing And rs.State = 1 Then rs.Close
    Set rs = Nothing
    
    On Error GoTo 0
End Function

' ============================================================================
' GetColumnHeadersList - EN-T�TES COLONNES PREMI�RE LIGNE
' ============================================================================
Private Function GetColumnHeadersList(conn As Object, sheets As Object) As Object
    On Error Resume Next ' Non critique si �choue pour certaines feuilles
    
    Set GetColumnHeadersList = CreateObject("Scripting.Dictionary")
    
    Dim sheetKey As Variant
    For Each sheetKey In sheets.Keys
        Dim tableName As String: tableName = sheets(sheetKey)
        Dim headers As Object: Set headers = CreateObject("Scripting.Dictionary")
        
        ' Requ�te premi�re ligne seulement
        Dim sql As String: sql = "SELECT TOP 1 * FROM [" & tableName & "]"
        Dim rs As Object: Set rs = conn.Execute(sql)
        
        If Not rs Is Nothing And err.Number = 0 Then
            Dim fieldIndex As Long
            For fieldIndex = 0 To rs.Fields.Count - 1
                Dim fieldName As String: fieldName = rs.Fields(fieldIndex).name
                Dim fieldValue As String
                
                ' Obtenir valeur premi�re ligne si disponible
                If Not rs.EOF Then
                    fieldValue = CStr(rs.Fields(fieldIndex).value)
                Else
                    fieldValue = fieldName ' Utiliser nom champ par d�faut
                End If
                
                headers.Add fieldName, fieldValue
            Next fieldIndex
            
            rs.Close
        End If
        
        GetColumnHeadersList.Add CStr(sheetKey), headers
        Set rs = Nothing
        err.Clear
    Next sheetKey
    
    On Error GoTo 0
End Function

' ============================================================================
' ResolveNamedReferencesADODB - R�SOLUTION R�F�RENCES NOMM�ES VIA ADODB
' ============================================================================
Private Function ResolveNamedReferencesADODB(registry As Object, connInfo As ADODBConnectionInfo, workbookStructure As Object) As Boolean
    On Error GoTo ResolveError
    
    ResolveNamedReferencesADODB = False
    
    Dim readFields As Object: Set readFields = registry("READ_FIELDS")
    Dim readOrder As Object: Set readOrder = registry("READ_ORDER")
    Dim resolvedFields As Object: Set resolvedFields = CreateObject("Scripting.Dictionary")
    
    Dim key As Variant
    For Each key In readFields.Keys
        Dim keyStr As String: keyStr = CStr(key)
        
        If Right(keyStr, 6) = "_NAMED" Then
            Dim cleanRef As String: cleanRef = Left(keyStr, Len(keyStr) - 6)
            cleanRef = Mid(cleanRef, 2) ' Enlever @
            
            Dim originalOrder As Long
            If readOrder.Exists(keyStr) Then originalOrder = readOrder(keyStr)
            
            ' Essayer r�solution par diff�rentes m�thodes
            If InStr(cleanRef, ":") > 0 Then
                ' Range nomm�e : Date:Facture
                If ResolveNamedRangeADODB(cleanRef, workbookStructure, resolvedFields, originalOrder) Then
                    ' Succ�s - marquer pour suppression
                    readFields.Remove keyStr
                    If readOrder.Exists(keyStr) Then readOrder.Remove keyStr
                End If
            Else
                ' R�f�rence unique : MonTableau, Clients
                If ResolveSingleNamedItemADODB(cleanRef, workbookStructure, resolvedFields, originalOrder) Then
                    ' Succ�s - marquer pour suppression
                    readFields.Remove keyStr
                    If readOrder.Exists(keyStr) Then readOrder.Remove keyStr
                End If
            End If
        End If
    Next key
    
    ' Ajouter champs r�solus au registry
    For Each key In resolvedFields.Keys
        readFields.Add CStr(key), True
        readOrder.Add CStr(key), resolvedFields(key)
    Next key
    
    ResolveNamedReferencesADODB = True
    Exit Function
    
ResolveError:
    SetParsingError ERR_PARSING_RESOLUTION_FAILED, "Erreur r�solution ADODB: " & err.Description, ""
    ResolveNamedReferencesADODB = False
End Function

' ============================================================================
' ResolveNamedRangeADODB - R�SOLUTION RANGE NOMM�E VIA HEADERS
' ============================================================================
Private Function ResolveNamedRangeADODB(rangeSpec As String, workbookStructure As Object, resolvedFields As Object, originalOrder As Long) As Boolean
    On Error GoTo ResolveRangeError
    
    ResolveNamedRangeADODB = False
    
    Dim rangeParts As Variant: rangeParts = Split(rangeSpec, ":")
    If UBound(rangeParts) < 1 Then Exit Function
    
    Dim startName As String: startName = Trim(CStr(rangeParts(0)))
    Dim endName As String: endName = Trim(CStr(rangeParts(1)))
    
    ' Chercher dans headers colonnes
    Dim headers As Object: Set headers = workbookStructure("ColumnHeaders")
    
    Dim startCol As String: startCol = ""
    Dim endCol As String: endCol = ""
    
    ' Rechercher noms dans toutes les feuilles
    Dim sheetKey As Variant
    For Each sheetKey In headers.Keys
        Dim sheetHeaders As Object: Set sheetHeaders = headers(sheetKey)
        
        Dim headerKey As Variant
        For Each headerKey In sheetHeaders.Keys
            Dim headerValue As String: headerValue = UCase(CStr(sheetHeaders(headerKey)))
            
            If UCase(startName) = headerValue And startCol = "" Then
                startCol = ExtractColumnFromFieldName(CStr(headerKey))
            End If
            
            If UCase(endName) = headerValue And endCol = "" Then
                endCol = ExtractColumnFromFieldName(CStr(headerKey))
            End If
        Next headerKey
        
        ' Si trouv� les deux, sortir
        If startCol <> "" And endCol <> "" Then Exit For
    Next sheetKey
    
    ' Si trouv�, cr�er range
    If startCol <> "" And endCol <> "" Then
        Dim startIdx As Long: startIdx = ColumnToIndexSafe(startCol)
        Dim endIdx As Long: endIdx = ColumnToIndexSafe(endCol)
        
        If startIdx > 0 And endIdx > 0 Then
            If startIdx > endIdx Then SwapLongsSafe startIdx, endIdx
            
            Dim orderOffset As Long: orderOffset = 0
            Dim col As Long
            For col = startIdx To endIdx
                Dim fieldRef As String: fieldRef = "@" & IndexToColumnSafe(col)
                resolvedFields.Add fieldRef, originalOrder + orderOffset
                orderOffset = orderOffset + 1
            Next col
            
            ResolveNamedRangeADODB = True
        End If
    End If
    
    Exit Function
    
ResolveRangeError:
    ResolveNamedRangeADODB = False
End Function

' ============================================================================
' ResolveSingleNamedItemADODB - R�SOLUTION ITEM UNIQUE VIA HEADERS
' ============================================================================
Private Function ResolveSingleNamedItemADODB(itemName As String, workbookStructure As Object, resolvedFields As Object, originalOrder As Long) As Boolean
    On Error GoTo ResolveSingleError
    
    ResolveSingleNamedItemADODB = False
    
    ' Chercher dans headers colonnes
    Dim headers As Object: Set headers = workbookStructure("ColumnHeaders")
    
    Dim sheetKey As Variant
    For Each sheetKey In headers.Keys
        Dim sheetHeaders As Object: Set sheetHeaders = headers(sheetKey)
        
        Dim headerKey As Variant
        For Each headerKey In sheetHeaders.Keys
            Dim headerValue As String: headerValue = UCase(CStr(sheetHeaders(headerKey)))
            
            If UCase(itemName) = headerValue Then
                Dim colName As String: colName = ExtractColumnFromFieldName(CStr(headerKey))
                If colName <> "" Then
                    Dim fieldRef As String: fieldRef = "@" & colName
                    resolvedFields.Add fieldRef, originalOrder
                    ResolveSingleNamedItemADODB = True
                    Exit Function
                End If
            End If
        Next headerKey
    Next sheetKey
    
    Exit Function
    
ResolveSingleError:
    ResolveSingleNamedItemADODB = False
End Function

' ============================================================================
' UTILITAIRES SUPPORT ADODB
' ============================================================================

Private Function ValidateADODBContext(context As ExcelResolutionContext) As Boolean
    ValidateADODBContext = (Len(context.workbookPath) > 0 And FileExists(context.workbookPath))
End Function

Private Function FileExists(filePath As String) As Boolean
    On Error Resume Next
    FileExists = (Dir(filePath) <> "")
    On Error GoTo 0
End Function

Private Sub CloseADODBConnection(ByRef connInfo As ADODBConnectionInfo)
    ' Nettoyage s�curis� - rien � faire avec les infos statiques
    connInfo.IsValid = False
End Sub

Private Sub UpdateRegistryAfterADODBResolution(registry As Object)
    ' Reconstruire mappings apr�s r�solution ADODB
    BuildUnionAndMappingsSafe registry
End Sub

Private Function ExtractColumnFromFieldName(fieldName As String) As String
    ' Extraire lettre colonne depuis nom champ ADODB (ex: "F1" -> "F")
    ExtractColumnFromFieldName = ExtractColumnLettersSafe(fieldName)
End Function

' ============================================================================
' API PUBLIQUE ADODB
' ============================================================================

Public Function IsADODBAvailable() As Boolean
    On Error Resume Next
    
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    IsADODBAvailable = (Not conn Is Nothing And err.Number = 0)
    
    Set conn = Nothing
    On Error GoTo 0
End Function

Public Function GetADODBProviderInfo() As String
    Dim info As String: info = "Providers ADODB disponibles:" & vbCrLf
    
    If IsProviderAvailable(EXCEL_PROVIDER_ACE) Then
        info = info & "- " & EXCEL_PROVIDER_ACE & " (recommand� Excel 2007+)" & vbCrLf
    End If
    
    If IsProviderAvailable(EXCEL_PROVIDER_JET) Then
        info = info & "- " & EXCEL_PROVIDER_JET & " (Excel 97-2003)" & vbCrLf
    End If
    
    GetADODBProviderInfo = info
End Function

' ============================================================================
' DOCUMENTATION COUCHE ADODB
' ============================================================================
'
' UTILISATION:
'
' 1. V�rifier disponibilit� ADODB:
'    If Not IsADODBAvailable() Then
'        MsgBox "ADODB non disponible"
'        Exit Sub
'    End If
'
' 2. Utilisation automatique via ResolveExcelDynamicReferences:
'    Dim registry As Object
'    Set registry = BuildColumnRegistry("@Test=1", "Clients:Montants")
'
'    Dim context As ExcelResolutionContext
'    context.WorkbookPath = "C:\MonFichier.xlsx"
'
'    If ResolveExcelDynamicReferences(registry, context) Then
'        ' R�solution r�ussie (utilise ADODB si workbook ferm�)
'    End If
'
' LIMITATIONS:
' - Acc�s plages nomm�es Excel limit� via ADODB
' - R�solution principalement bas�e sur en-t�tes colonnes
' - N�cessite providers ADODB install�s sur syst�me
' - Performance moindre qu'Excel ouvert
'
' AVANTAGES:
' - Pas besoin d'ouvrir Excel
' - Acc�s multi-feuilles
' - Int�gration transparente
' - Fallback automatique si �chec
'
' ============================================================================

