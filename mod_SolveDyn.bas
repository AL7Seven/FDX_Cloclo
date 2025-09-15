Attribute VB_Name = "mod_SolveDyn"
' ============================================================================
' DOCUMENTATION MODULE ÉTENDU
' ============================================================================
' MODULE COMPLET AVEC:
' 1. RÉSOLUTION DYNAMIQUE EXCEL:
'    - Plages nommées workbook/worksheet
'    - Tableaux structurés Excel
'    - En-têtes colonnes automatiques
'    - Gestion workbook ouvert/fermé
'
' 2. DÉTECTION COMPARAISONS SOPHISTIQUÉE:
'    - Opérateurs basiques (=, >, <, >=, <=, <>)
'    - Clauses IN avec valeurs multiples
'    - Clauses LIKE avec patterns
'    - Clauses BETWEEN avec ranges
'    - Subqueries (IN SELECT, = SELECT...)
'    - Comparaisons JOIN (@Field1 = @Field2)
'
' 3. MAPPINGS POSITION COMPLETS:
'    - Source positions depuis données réelles
'    - Extract positions avec ordre optimisé
'    - Index inverses pour performance O(1)
'    - Priorisation champs comparaison
'    - Validation cohérence mappings
'    - Statistiques et diagnostics
'    - Recommandations optimisation
' ============================================================================
' ============================================================================
' V2: MODULE RÉSOLUTION DYNAMIQUE EXCEL ET DÉTECTION AVANCÉE
' ============================================================================
' Résolution plages nommées, tableaux structurés Excel
' Détection sophistiquée comparaisons dynamiques dans expressions
' Mappings position tableau avec gestion ordre et types
' ============================================================================

Option Explicit

' ============================================================================
' TYPES POUR RÉSOLUTION DYNAMIQUE
' ============================================================================
'Public Type ExcelResolutionContext
'    WorkbookPath As String
'    WorksheetName As String
'    IsWorkbookOpen As Boolean
'    HasNamedRanges As Boolean
'    HasStructuredTables As Boolean
'End Type

Public Type ComparisonContext
    fieldName As String
    Operator As String
    comparedValue As String
    contextType As String  ' "FILTER", "JOIN", "SUBQUERY", etc.
    Position As Long
End Type

' ============================================================================
' ResolveExcelDynamicReferences - RÉSOLUTION DYNAMIQUE COMPLÈTE
' ============================================================================
' Résout les références nommées Excel en adresses réelles
Public Function ResolveExcelDynamicReferences(registry As Object, context As ExcelResolutionContext) As Boolean
    On Error GoTo ErrorHandler
    
    ResolveExcelDynamicReferences = False
    
    ' Vérifier si résolution nécessaire
    If Not HasNamedReferences(registry) Then
        ResolveExcelDynamicReferences = True
        Exit Function
    End If
    
    ' Initialiser contexte si non fourni
    If context.workbookPath = "" Then
        context = DetectExcelContext()
    End If
    
    ' Stratégies de résolution selon contexte
    If context.IsWorkbookOpen Then
        ResolveExcelDynamicReferences = ResolveFromOpenWorkbook(registry, context)
    Else
        ResolveExcelDynamicReferences = ResolveFromClosedWorkbook(registry, context)
    End If
    
    ' Nettoyer marqueurs temporaires
    CleanupNamedMarkers registry
    
    ' Reconstruire mappings après résolution
    BuildCompleteUnionAndMappings registry
    
    Exit Function
    
ErrorHandler:
    If GetConfigValue("DebugMode") Then
        Debug.Print "ERREUR ResolveExcelDynamicReferences: " & err.Description
    End If
    ResolveExcelDynamicReferences = False
End Function

' ============================================================================
' ResolveFromOpenWorkbook - RÉSOLUTION WORKBOOK OUVERT
' ============================================================================
Function ResolveFromOpenWorkbook(registry As Object, context As ExcelResolutionContext) As Boolean
    On Error GoTo ErrorHandler
    
    Dim wb As Workbook
    Dim ws As Worksheet
    
    ' Obtenir références workbook/worksheet
    If context.workbookPath <> "" Then
        Set wb = GetWorkbookByPath(context.workbookPath)
    Else
        Set wb = ActiveWorkbook
    End If
    
    If context.WorksheetName <> "" Then
        Set ws = wb.Worksheets(context.WorksheetName)
    Else
        Set ws = wb.ActiveSheet
    End If
    
    ' Résoudre chaque référence nommée
    Dim readFields As Object: Set readFields = registry("READ_FIELDS")
    Dim readOrder As Object: Set readOrder = registry("READ_ORDER")
    
    Dim key As Variant
    Dim keysToRemove As New Collection
    Dim keysToAdd As Object: Set keysToAdd = CreateObject("Scripting.Dictionary")
    
    For Each key In readFields.Keys
        If Right(CStr(key), 6) = "_NAMED" Then
            keysToRemove.Add key
            ResolveNamedReference CStr(key), wb, ws, keysToAdd, readOrder
        End If
    Next key
    
    ' Mettre à jour registry
    UpdateRegistryAfterResolution registry, keysToRemove, keysToAdd
    
    ResolveFromOpenWorkbook = True
    Exit Function
    
ErrorHandler:
    ResolveFromOpenWorkbook = False
    If GetConfigValue("DebugMode") Then
        Debug.Print "ERREUR ResolveFromOpenWorkbook: " & err.Description
    End If
End Function

' ============================================================================
' ResolveNamedReference - RÉSOLUTION RÉFÉRENCE NOMMÉE SPÉCIFIQUE
' ============================================================================
Private Sub ResolveNamedReference(namedRef As String, wb As Workbook, ws As Worksheet, keysToAdd As Object, readOrder As Object)
    On Error GoTo ErrorHandler
    
    ' Parser référence : @Date:Facture_NAMED ou @MonTableau_NAMED
    Dim cleanRef As String: cleanRef = Replace(namedRef, "_NAMED", "")
    cleanRef = Mid(cleanRef, 2) ' Enlever @
    
    Dim originalOrder As Long
    If readOrder.Exists(namedRef) Then originalOrder = readOrder(namedRef)
    
    If InStr(cleanRef, ":") > 0 Then
        ' Range nommée : Date:Facture
        ResolveNamedRange cleanRef, wb, ws, keysToAdd, originalOrder
    Else
        ' Élément unique : Tableau, PlageNommée
        ResolveSingleNamedItem cleanRef, wb, ws, keysToAdd, originalOrder
    End If
    
    Exit Sub
    
ErrorHandler:
    If GetConfigValue("DebugMode") Then
        Debug.Print "ERREUR ResolveNamedReference pour " & namedRef & ": " & err.Description
    End If
End Sub

' ============================================================================
' ResolveNamedRange - RÉSOLUTION RANGE NOMMÉE (Date:Facture)
' ============================================================================
Private Sub ResolveNamedRange(rangeSpec As String, wb As Workbook, ws As Worksheet, keysToAdd As Object, originalOrder As Long)
    Dim rangeParts As Variant: rangeParts = Split(rangeSpec, ":")
    If UBound(rangeParts) < 1 Then Exit Sub
    
    Dim startName As String: startName = Trim(CStr(rangeParts(0)))
    Dim endName As String: endName = Trim(CStr(rangeParts(1)))
    
    Dim startAddr As String, endAddr As String
    
    ' Résoudre noms en adresses
    startAddr = ResolveItemToAddress(startName, wb, ws)
    endAddr = ResolveItemToAddress(endName, wb, ws)
    
    If startAddr <> "" And endAddr <> "" Then
        ' Convertir en range et ajouter colonnes
        Dim startCol As String: startCol = ExtractColumnLetters(startAddr)
        Dim endCol As String: endCol = ExtractColumnLetters(endAddr)
        
        If startCol <> "" And endCol <> "" Then
            AddResolvedRange startCol, endCol, keysToAdd, originalOrder
        End If
    End If
End Sub

' ============================================================================
' ResolveSingleNamedItem - RÉSOLUTION ÉLÉMENT UNIQUE
' ============================================================================
Private Sub ResolveSingleNamedItem(itemName As String, wb As Workbook, ws As Worksheet, keysToAdd As Object, originalOrder As Long)
    Dim resolvedAddr As String
    
    ' Essayer différents types de résolution
    resolvedAddr = ResolveItemToAddress(itemName, wb, ws)
    
    If resolvedAddr <> "" Then
        ' Traiter selon type de résolution
        If InStr(resolvedAddr, ":") > 0 Then
            ' Range résolue : A1:C10
            Dim rangeParts As Variant: rangeParts = Split(resolvedAddr, ":")
            Dim startCol As String: startCol = ExtractColumnLetters(CStr(rangeParts(0)))
            Dim endCol As String: endCol = ExtractColumnLetters(CStr(rangeParts(1)))
            AddResolvedRange startCol, endCol, keysToAdd, originalOrder
        Else
            ' Cellule unique : A1
            Dim colLetter As String: colLetter = ExtractColumnLetters(resolvedAddr)
            If colLetter <> "" Then
                keysToAdd("@" & colLetter) = originalOrder
            End If
        End If
    End If
End Sub

' ============================================================================
' ResolveItemToAddress - RÉSOLUTION ITEM ? ADRESSE
' ============================================================================
Function ResolveItemToAddress(itemName As String, wb As Workbook, ws As Worksheet) As String
    On Error Resume Next
    
    Dim resolvedAddr As String: resolvedAddr = ""
    
    ' 1. Essayer plage nommée workbook
    Dim namedRange As name
    Set namedRange = wb.Names(itemName)
    If Not namedRange Is Nothing Then
        resolvedAddr = namedRange.RefersTo
        resolvedAddr = CleanExcelAddress(resolvedAddr)
        If resolvedAddr <> "" Then
            ResolveItemToAddress = resolvedAddr
            Exit Function
        End If
    End If
    
    ' 2. Essayer plage nommée worksheet
    Set namedRange = ws.Names(itemName)
    If Not namedRange Is Nothing Then
        resolvedAddr = namedRange.RefersTo
        resolvedAddr = CleanExcelAddress(resolvedAddr)
        If resolvedAddr <> "" Then
            ResolveItemToAddress = resolvedAddr
            Exit Function
        End If
    End If
    
    ' 3. Essayer tableau structuré
    resolvedAddr = ResolveStructuredTableReference(itemName, wb, ws)
    If resolvedAddr <> "" Then
        ResolveItemToAddress = resolvedAddr
        Exit Function
    End If
    
    ' 4. Essayer en-tête de colonne
    resolvedAddr = ResolveColumnHeaderReference(itemName, ws)
    If resolvedAddr <> "" Then
        ResolveItemToAddress = resolvedAddr
        Exit Function
    End If
    
    On Error GoTo 0
    ResolveItemToAddress = ""
End Function

' ============================================================================
' ResolveStructuredTableReference - RÉSOLUTION TABLEAU STRUCTURÉ
' ============================================================================
Function ResolveStructuredTableReference(itemName As String, wb As Workbook, ws As Worksheet) As String
    On Error Resume Next
    
    ' Rechercher dans tous les tableaux structurés
    Dim tbl As ListObject
    For Each tbl In ws.ListObjects
        If UCase(tbl.name) = UCase(itemName) Then
            ' Tableau entier
            ResolveStructuredTableReference = tbl.range.Address
            Exit Function
        End If
        
        ' Rechercher colonne dans tableau
        Dim col As ListColumn
        For Each col In tbl.ListColumns
            If UCase(col.name) = UCase(itemName) Then
                ResolveStructuredTableReference = col.range.Address
                Exit Function
            End If
        Next col
    Next tbl
    
    On Error GoTo 0
    ResolveStructuredTableReference = ""
End Function

' ============================================================================
' ResolveColumnHeaderReference - RÉSOLUTION EN-TÊTE COLONNE
' ============================================================================
Function ResolveColumnHeaderReference(itemName As String, ws As Worksheet) As String
    On Error Resume Next
    
    ' Recherche en-tête dans première ligne (comportement par défaut)
    Dim searchRange As range
    Set searchRange = ws.Rows(1)
    
    Dim foundCell As range
    Set foundCell = searchRange.Find(itemName, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not foundCell Is Nothing Then
        ResolveColumnHeaderReference = foundCell.Address
    Else
        ResolveColumnHeaderReference = ""
    End If
    
    On Error GoTo 0
End Function

' ============================================================================
' DÉTECTION COMPARAISONS DYNAMIQUES AVANCÉE
' ============================================================================

' AnalyzeComparisonContexts - ANALYSE SOPHISTIQUÉE COMPARAISONS
Public Function AnalyzeComparisonContexts(expression As String) As Collection
    On Error GoTo ErrorHandler
    
    Set AnalyzeComparisonContexts = New Collection
    
    If Len(Trim(expression)) = 0 Then Exit Function
    
    Dim expr As String: expr = UCase(Trim(expression))
    Dim contexts As Collection: Set contexts = New Collection
    
    ' Détecter différents types de comparaisons
    DetectBasicComparisons expr, contexts
    DetectInClauses expr, contexts
    DetectLikeClauses expr, contexts
    DetectBetweenClauses expr, contexts
    DetectSubqueryComparisons expr, contexts
    DetectJoinComparisons expr, contexts
    
    Set AnalyzeComparisonContexts = contexts
    Exit Function
    
ErrorHandler:
    Set AnalyzeComparisonContexts = New Collection
End Function

' ============================================================================
' DetectBasicComparisons - DÉTECTION COMPARAISONS BASIQUES
' ============================================================================
Private Sub DetectBasicComparisons(expr As String, contexts As Collection)
    Dim operators As Variant
    operators = Array(">=", "<=", "<>", "=", ">", "<")
    
    Dim i As Long, j As Long
    
    For i = 0 To UBound(operators)
        Dim op As String: op = CStr(operators(i))
        Dim pos As Long: pos = 1
        
        Do
            pos = InStr(pos, expr, op)
            If pos > 0 Then
                Dim context As ComparisonContext
                context = ExtractComparisonContext(expr, pos, op, "FILTER")
                
                If context.fieldName <> "" Then
                    contexts.Add context
                End If
                
                pos = pos + Len(op)
            End If
        Loop While pos > 0 And pos < Len(expr)
    Next i
End Sub

' ============================================================================
' DetectInClauses - DÉTECTION CLAUSES IN
' ============================================================================
Private Sub DetectInClauses(expr As String, contexts As Collection)
    Dim pos As Long: pos = 1
    
    Do
        pos = InStr(pos, expr, " IN ")
        If pos > 0 Then
            Dim context As ComparisonContext
            context = ExtractInClauseContext(expr, pos)
            
            If context.fieldName <> "" Then
                contexts.Add context
            End If
            
            pos = pos + 4 ' Longueur " IN "
        End If
    Loop While pos > 0 And pos < Len(expr)
End Sub

' ============================================================================
' DetectLikeClauses - DÉTECTION CLAUSES LIKE
' ============================================================================
Private Sub DetectLikeClauses(expr As String, contexts As Collection)
    Dim pos As Long: pos = 1
    
    Do
        pos = InStr(pos, expr, " LIKE ")
        If pos > 0 Then
            Dim context As ComparisonContext
            context = ExtractLikeClauseContext(expr, pos)
            
            If context.fieldName <> "" Then
                contexts.Add context
            End If
            
            pos = pos + 6 ' Longueur " LIKE "
        End If
    Loop While pos > 0 And pos < Len(expr)
End Sub

' ============================================================================
' DetectBetweenClauses - DÉTECTION CLAUSES BETWEEN
' ============================================================================
Private Sub DetectBetweenClauses(expr As String, contexts As Collection)
    Dim pos As Long: pos = 1
    
    Do
        pos = InStr(pos, expr, " BETWEEN ")
        If pos > 0 Then
            Dim context As ComparisonContext
            context = ExtractBetweenClauseContext(expr, pos)
            
            If context.fieldName <> "" Then
                contexts.Add context
            End If
            
            pos = pos + 9 ' Longueur " BETWEEN "
        End If
    Loop While pos > 0 And pos < Len(expr)
End Sub

' ============================================================================
' DetectSubqueryComparisons - DÉTECTION COMPARAISONS SUBQUERY
' ============================================================================
Private Sub DetectSubqueryComparisons(expr As String, contexts As Collection)
    ' Détecter patterns comme @Field IN (SELECT...) ou @Field = (SELECT...)
    Dim patterns As Variant
    patterns = Array(" IN (SELECT", " = (SELECT", " > (SELECT", " < (SELECT")
    
    Dim i As Long
    For i = 0 To UBound(patterns)
        Dim pattern As String: pattern = CStr(patterns(i))
        Dim pos As Long: pos = 1
        
        Do
            pos = InStr(pos, expr, pattern)
            If pos > 0 Then
                Dim context As ComparisonContext
                context = ExtractSubqueryContext(expr, pos, pattern)
                
                If context.fieldName <> "" Then
                    contexts.Add context
                End If
                
                pos = pos + Len(pattern)
            End If
        Loop While pos > 0 And pos < Len(expr)
    Next i
End Sub

' ============================================================================
' DetectJoinComparisons - DÉTECTION COMPARAISONS JOIN
' ============================================================================
Private Sub DetectJoinComparisons(expr As String, contexts As Collection)
    ' Détecter patterns de jointure @Field1 = @Field2
    Dim pos As Long: pos = 1
    
    Do
        pos = InStr(pos, expr, "@")
        If pos > 0 Then
            ' Chercher = avec autre @Field après
            Dim nextAt As Long: nextAt = InStr(pos + 1, expr, "@")
            If nextAt > 0 Then
                Dim betweenText As String: betweenText = Mid(expr, pos, nextAt - pos)
                If InStr(betweenText, "=") > 0 Then
                    Dim context As ComparisonContext
                    context = ExtractJoinContext(expr, pos, nextAt)
                    
                    If context.fieldName <> "" Then
                        contexts.Add context
                    End If
                End If
            End If
            
            pos = pos + 1
        End If
    Loop While pos > 0 And pos < Len(expr)
End Sub

' ============================================================================
' FONCTIONS EXTRACTION CONTEXTE COMPARAISON
' ============================================================================

Function ExtractComparisonContext(expr As String, pos As Long, op As String, contextType As String) As ComparisonContext
    Dim context As ComparisonContext
    
    ' Extraire champ avant opérateur
    Dim fieldName As String: fieldName = ExtractFieldBeforePosition(expr, pos)
    
    ' Extraire valeur après opérateur
    Dim comparedValue As String: comparedValue = ExtractValueAfterPosition(expr, pos + Len(op))
    
    context.fieldName = fieldName
    context.Operator = op
    context.comparedValue = comparedValue
    context.contextType = contextType
    context.Position = pos
    
    ExtractComparisonContext = context
End Function

Function ExtractInClauseContext(expr As String, pos As Long) As ComparisonContext
    Dim context As ComparisonContext
    
    context.fieldName = ExtractFieldBeforePosition(expr, pos)
    context.Operator = "IN"
    context.comparedValue = ExtractInValues(expr, pos + 4)
    context.contextType = "IN_CLAUSE"
    context.Position = pos
    
    ExtractInClauseContext = context
End Function

Function ExtractLikeClauseContext(expr As String, pos As Long) As ComparisonContext
    Dim context As ComparisonContext
    
    context.fieldName = ExtractFieldBeforePosition(expr, pos)
    context.Operator = "LIKE"
    context.comparedValue = ExtractValueAfterPosition(expr, pos + 6)
    context.contextType = "LIKE_CLAUSE"
    context.Position = pos
    
    ExtractLikeClauseContext = context
End Function

Function ExtractBetweenClauseContext(expr As String, pos As Long) As ComparisonContext
    Dim context As ComparisonContext
    
    context.fieldName = ExtractFieldBeforePosition(expr, pos)
    context.Operator = "BETWEEN"
    context.comparedValue = ExtractBetweenValues(expr, pos + 9)
    context.contextType = "BETWEEN_CLAUSE"
    context.Position = pos
    
    ExtractBetweenClauseContext = context
End Function

Function ExtractSubqueryContext(expr As String, pos As Long, pattern As String) As ComparisonContext
    Dim context As ComparisonContext
    
    context.fieldName = ExtractFieldBeforePosition(expr, pos)
    context.Operator = Trim(Replace(pattern, "(SELECT", ""))
    context.comparedValue = ExtractSubqueryText(expr, pos + Len(pattern))
    context.contextType = "SUBQUERY"
    context.Position = pos
    
    ExtractSubqueryContext = context
End Function

Function ExtractJoinContext(expr As String, pos1 As Long, pos2 As Long) As ComparisonContext
    Dim context As ComparisonContext
    
    context.fieldName = ExtractFieldAtPosition(expr, pos1)
    context.Operator = "="
    context.comparedValue = ExtractFieldAtPosition(expr, pos2)
    context.contextType = "JOIN"
    context.Position = pos1
    
    ExtractJoinContext = context
End Function

' ============================================================================
' UTILITAIRES EXTRACTION AVANCÉS
' ============================================================================

Function ExtractFieldBeforePosition(expr As String, pos As Long) As String
    ' Recherche @Field avant position donnée
    Dim searchStart As Long: searchStart = IIf(pos - 50 > 1, pos - 50, 1)
    Dim segment As String: segment = Mid(expr, searchStart, pos - searchStart)
    
    Dim lastAt As Long: lastAt = 0
    Dim i As Long
    For i = Len(segment) To 1 Step -1
        If Mid(segment, i, 1) = "@" Then
            lastAt = i
            Exit For
        End If
    Next i
    
    If lastAt > 0 Then
        ExtractFieldBeforePosition = ExtractFieldAtPosition(expr, searchStart + lastAt - 1)
    Else
        ExtractFieldBeforePosition = ""
    End If
End Function

Function ExtractFieldAtPosition(expr As String, pos As Long) As String
    If pos > Len(expr) Or Mid(expr, pos, 1) <> "@" Then
        ExtractFieldAtPosition = ""
        Exit Function
    End If
    
    Dim fieldName As String: fieldName = "@"
    Dim i As Long: i = pos + 1
    
    Do While i <= Len(expr)
        Dim char As String: char = Mid(expr, i, 1)
        If char >= "A" And char <= "Z" Then
            fieldName = fieldName & char
            i = i + 1
        Else
            Exit Do
        End If
    Loop
    
    ExtractFieldAtPosition = fieldName
End Function

Function ExtractValueAfterPosition(expr As String, pos As Long) As String
    ' Extraire valeur après position (gérer quotes, espaces, parenthèses)
    If pos > Len(expr) Then
        ExtractValueAfterPosition = ""
        Exit Function
    End If
    
    Dim i As Long: i = pos
    
    ' Ignorer espaces
    Do While i <= Len(expr) And Mid(expr, i, 1) = " "
        i = i + 1
    Loop
    
    If i > Len(expr) Then
        ExtractValueAfterPosition = ""
        Exit Function
    End If
    
    Dim value As String: value = ""
    Dim inQuotes As Boolean: inQuotes = False
    Dim quoteChar As String
    
    Do While i <= Len(expr)
        Dim char As String: char = Mid(expr, i, 1)
        
        If Not inQuotes Then
            If char = "'" Or char = """" Then
                inQuotes = True
                quoteChar = char
                value = value & char
            ElseIf char = " " And Len(value) > 0 Then
                ' Fin de valeur
                Exit Do
            ElseIf char <> " " Then
                value = value & char
            End If
        Else
            value = value & char
            If char = quoteChar Then
                inQuotes = False
                Exit Do
            End If
        End If
        
        i = i + 1
    Loop
    
    ExtractValueAfterPosition = value
End Function

' ============================================================================
' MAPPINGS POSITION TABLEAU COMPLETS
' ============================================================================

' BuildAdvancedPositionMappings - CONSTRUCTION MAPPINGS AVANCÉS
Public Sub BuildAdvancedPositionMappings(registry As Object, sourceDataArray As Variant)
    On Error GoTo ErrorHandler
    
    ' Construire mappings source depuis données réelles
    BuildSourcePositionMappings registry, sourceDataArray
    
    ' Optimiser ordre extraction selon utilisation
    OptimizeExtractionOrder registry
    
    ' Créer index inverse performants
    BuildPerformanceIndexes registry
    
    Exit Sub
    
ErrorHandler:
    If GetConfigValue("DebugMode") Then
        Debug.Print "ERREUR BuildAdvancedPositionMappings: " & err.Description
    End If
End Sub

Private Sub BuildSourcePositionMappings(registry As Object, sourceArray As Variant)
    If Not IsArray(sourceArray) Then Exit Sub
    If UBound(sourceArray, 1) < 1 Then Exit Sub ' Pas de données
    
    Dim sourcePositions As Object: Set sourcePositions = registry("SOURCE_POSITIONS")
    Dim positionToField As Object: Set positionToField = registry("POSITION_TO_FIELD")
    
    ' Analyser première ligne (en-têtes)
    Dim col As Long
    For col = LBound(sourceArray, 2) To UBound(sourceArray, 2)
        Dim headerValue As String: headerValue = CStr(sourceArray(1, col))
        Dim fieldRef As String: fieldRef = "@" & UCase(headerValue)
        
        ' Si ce champ est requis, mapper sa position
        If registry("ALL_REQUIRED").Exists(fieldRef) Then
            sourcePositions(fieldRef) = col
            positionToField(col) = fieldRef
        End If
    Next col
End Sub

Private Sub OptimizeExtractionOrder(registry As Object)
    ' Réorganiser ordre extraction pour performance
    Dim extractPos As Object: Set extractPos = registry("EXTRACT_POSITIONS")
    Dim readOrder As Object: Set readOrder = registry("READ_ORDER")
    Dim compFields As Object: Set compFields = registry("COMPARISON_FIELDS")
    
    ' Prioriser champs de comparaison (accès fréquent)
    Dim newOrder As Long: newOrder = 1
    Dim key As Variant
    
    ' 1. Champs de comparaison d'abord
    For Each key In compFields.Keys
        If extractPos.Exists(key) Then
            extractPos(key) = newOrder
            registry("EXTRACT_TO_FIELD")(newOrder) = key
            newOrder = newOrder + 1
        End If
    Next key
    
    ' 2. Autres champs ensuite (ordre préservé)
    For Each key In readOrder.Keys
        If Not compFields.Exists(key) Then
            extractPos(key) = newOrder
            registry("EXTRACT_TO_FIELD")(newOrder) = key
            newOrder = newOrder + 1
        End If
    Next key
End Sub

Private Sub BuildPerformanceIndexes(registry As Object)
    ' Créer index inverses pour accès rapide O(1)
    registry.Add "FIELD_TO_SOURCE_INDEX", CreateObject("Scripting.Dictionary")
    registry.Add "FIELD_TO_EXTRACT_INDEX", CreateObject("Scripting.Dictionary")
    registry.Add "SOURCE_TO_EXTRACT_MAP", CreateObject("Scripting.Dictionary")
    
    Dim fieldToSource As Object: Set fieldToSource = registry("FIELD_TO_SOURCE_INDEX")
    Dim fieldToExtract As Object: Set fieldToExtract = registry("FIELD_TO_EXTRACT_INDEX")
    Dim sourceToExtract As Object: Set sourceToExtract = registry("SOURCE_TO_EXTRACT_MAP")
    
    Dim key As Variant
    For Each key In registry("ALL_REQUIRED").Keys
        If registry("SOURCE_POSITIONS").Exists(key) And registry("EXTRACT_POSITIONS").Exists(key) Then
            Dim sourcePos As Long: sourcePos = registry("SOURCE_POSITIONS")(key)
            Dim extractPos As Long: extractPos = registry("EXTRACT_POSITIONS")(key)
            
            fieldToSource(key) = sourcePos
            fieldToExtract(key) = extractPos
            sourceToExtract(sourcePos) = extractPos
        End If
    Next key
End Sub

' ============================================================================
' UTILITAIRES SUPPORT
' ============================================================================

Function HasNamedReferences(registry As Object) As Boolean
    Dim key As Variant
    For Each key In registry("READ_FIELDS").Keys
        If Right(CStr(key), 6) = "_NAMED" Then
            HasNamedReferences = True
            Exit Function
        End If
    Next key
    HasNamedReferences = False
End Function

Function DetectExcelContext() As ExcelResolutionContext
    Dim context As ExcelResolutionContext
    
    On Error Resume Next
    context.IsWorkbookOpen = (Not ActiveWorkbook Is Nothing)
    context.workbookPath = IIf(context.IsWorkbookOpen, ActiveWorkbook.FullName, "")
    context.WorksheetName = IIf(context.IsWorkbookOpen, ActiveSheet.name, "")
    On Error GoTo 0
    
    DetectExcelContext = context
End Function

Function GetWorkbookByPath(path As String) As Workbook
    On Error Resume Next
    Set GetWorkbookByPath = Workbooks(Dir(path))
    On Error GoTo 0
End Function

Function CleanExcelAddress(rawAddress As String) As String
    ' Nettoyer adresse Excel (enlever $, =, etc.)
    Dim cleaned As String: cleaned = rawAddress
    cleaned = Replace(cleaned, "=", "")
    cleaned = Replace(cleaned, "$", "")
    cleaned = Replace(cleaned, " ", "")
    
    CleanExcelAddress = cleaned
End Function

Private Sub AddResolvedRange(startCol As String, endCol As String, keysToAdd As Object, originalOrder As Long)
    Dim startIdx As Long: startIdx = ColumnToIndex(startCol)
    Dim endIdx As Long: endIdx = ColumnToIndex(endCol)
    
    If startIdx > endIdx Then SwapLongs startIdx, endIdx
    
    Dim col As Long, orderOffset As Long: orderOffset = 0
    For col = startIdx To endIdx
        keysToAdd("@" & IndexToColumn(col)) = originalOrder + orderOffset
        orderOffset = orderOffset + 1
    Next col
End Sub

Private Sub UpdateRegistryAfterResolution(registry As Object, keysToRemove As Collection, keysToAdd As Object)
    Dim readFields As Object: Set readFields = registry("READ_FIELDS")
    Dim readOrder As Object: Set readOrder = registry("READ_ORDER")
    
    ' Supprimer références nommées non résolues
    Dim i As Long
    For i = 1 To keysToRemove.Count
        Dim keyToRemove As String: keyToRemove = CStr(keysToRemove(i))
        If readFields.Exists(keyToRemove) Then readFields.Remove keyToRemove
        If readOrder.Exists(keyToRemove) Then readOrder.Remove keyToRemove
    Next i
    
    ' Ajouter références résolues
    Dim key As Variant
    For Each key In keysToAdd.Keys
        readFields(key) = True
        readOrder(key) = keysToAdd(key)
    Next key
End Sub

Private Sub CleanupNamedMarkers(registry As Object)
    ' Nettoyer tous les marqueurs _NAMED restants non résolus
    Dim collections() As String
    collections = Split("READ_FIELDS,ALL_REQUIRED,READ_ORDER", ",")
    
    Dim i As Long
    For i = 0 To UBound(collections)
        Dim dict As Object: Set dict = registry(collections(i))
        Dim keysToRemove As New Collection
        
        Dim key As Variant
        For Each key In dict.Keys
            If Right(CStr(key), 6) = "_NAMED" Then
                keysToRemove.Add key
            End If
        Next key
        
        ' Supprimer marqueurs non résolus
        Dim j As Long
        For j = 1 To keysToRemove.Count
            If dict.Exists(keysToRemove(j)) Then
                dict.Remove keysToRemove(j)
            End If
        Next j
    Next i
End Sub

Function ResolveFromClosedWorkbook(registry As Object, context As ExcelResolutionContext) As Boolean
    ' Pour workbook fermé, tentative avec ADODB ou accès limité
    On Error GoTo ErrorHandler
    
    ' Stratégie 1: Essayer ouverture temporaire en mode lecture seule
    If TryTemporaryOpen(context.workbookPath) Then
        ResolveFromClosedWorkbook = ResolveFromOpenWorkbook(registry, context)
        ' Refermer si ouvert temporairement
        If context.workbookPath <> "" Then
            Application.Workbooks(Dir(context.workbookPath)).Close False
        End If
    Else
        ' Stratégie 2: Marquer comme non résolu et continuer
        MarkUnresolvedReferences registry
        ResolveFromClosedWorkbook = True ' Continuer malgré non résolution
    End If
    
    Exit Function
    
ErrorHandler:
    ResolveFromClosedWorkbook = False
End Function

Function TryTemporaryOpen(filePath As String) As Boolean
    On Error Resume Next
    Application.Workbooks.Open filePath, ReadOnly:=True, UpdateLinks:=False
    TryTemporaryOpen = (err.Number = 0)
    On Error GoTo 0
End Function

Private Sub MarkUnresolvedReferences(registry As Object)
    ' Marquer références non résolues pour traitement ultérieur
    If Not registry.Exists("UNRESOLVED_REFERENCES") Then
        registry.Add "UNRESOLVED_REFERENCES", CreateObject("Scripting.Dictionary")
    End If
    
    Dim unresolvedRefs As Object: Set unresolvedRefs = registry("UNRESOLVED_REFERENCES")
    Dim key As Variant
    
    For Each key In registry("READ_FIELDS").Keys
        If Right(CStr(key), 6) = "_NAMED" Then
            unresolvedRefs(key) = "PENDING_RESOLUTION"
        End If
    Next key
End Sub

' ============================================================================
' FONCTIONS UTILITAIRES EXTRACTION VALEURS COMPLÈTES
' ============================================================================

Function ExtractInValues(expr As String, startPos As Long) As String
    ' Extraire valeurs dans clause IN (...)
    Dim pos As Long: pos = startPos
    
    ' Chercher parenthèse ouvrante
    Do While pos <= Len(expr) And Mid(expr, pos, 1) <> "("
        pos = pos + 1
    Loop
    
    If pos > Len(expr) Then
        ExtractInValues = ""
        Exit Function
    End If
    
    pos = pos + 1 ' Après (
    Dim values As String: values = ""
    Dim parenCount As Long: parenCount = 1
    
    Do While pos <= Len(expr) And parenCount > 0
        Dim char As String: char = Mid(expr, pos, 1)
        If char = "(" Then
            parenCount = parenCount + 1
        ElseIf char = ")" Then
            parenCount = parenCount - 1
        End If
        
        If parenCount > 0 Then values = values & char
        pos = pos + 1
    Loop
    
    ExtractInValues = Trim(values)
End Function

Function ExtractBetweenValues(expr As String, startPos As Long) As String
    ' Extraire valeurs dans clause BETWEEN val1 AND val2
    Dim pos As Long: pos = startPos
    Dim values As String: values = ""
    
    ' Ignorer espaces
    Do While pos <= Len(expr) And Mid(expr, pos, 1) = " "
        pos = pos + 1
    Loop
    
    ' Extraire jusqu'à AND
    Do While pos <= Len(expr)
        If pos + 4 <= Len(expr) And Mid(expr, pos, 4) = " AND" Then
            Exit Do
        End If
        values = values & Mid(expr, pos, 1)
        pos = pos + 1
    Loop
    
    ' Ajouter " AND "
    values = values & " AND "
    pos = pos + 4
    
    ' Extraire valeur après AND
    Do While pos <= Len(expr) And Mid(expr, pos, 1) = " "
        pos = pos + 1
    Loop
    
    ' Extraire jusqu'à prochain délimiteur
    Do While pos <= Len(expr)
        Dim char As String: char = Mid(expr, pos, 1)
        If char = " " Or char = ")" Or char = ";" Then
            Exit Do
        End If
        values = values & char
        pos = pos + 1
    Loop
    
    ExtractBetweenValues = Trim(values)
End Function

Function ExtractSubqueryText(expr As String, startPos As Long) As String
    ' Extraire texte subquery SELECT...)
    Dim pos As Long: pos = startPos
    Dim subquery As String: subquery = ""
    Dim parenCount As Long: parenCount = 1
    
    ' Déjà après (SELECT
    Do While pos <= Len(expr) And parenCount > 0
        Dim char As String: char = Mid(expr, pos, 1)
        If char = "(" Then
            parenCount = parenCount + 1
        ElseIf char = ")" Then
            parenCount = parenCount - 1
        End If
        
        If parenCount > 0 Then subquery = subquery & char
        pos = pos + 1
    Loop
    
    ExtractSubqueryText = "SELECT" & Trim(subquery)
End Function

' ============================================================================
' API PUBLIQUE ÉTENDUE POUR RÉSOLUTION ET COMPARAISONS
' ============================================================================

' Vérifier si références résolues
Public Function AreAllReferencesResolved(registry As Object) As Boolean
    AreAllReferencesResolved = True
    
    If registry.Exists("UNRESOLVED_REFERENCES") Then
        AreAllReferencesResolved = (registry("UNRESOLVED_REFERENCES").Count = 0)
    End If
    
    ' Vérifier aussi marqueurs _NAMED restants
    Dim key As Variant
    For Each key In registry("ALL_REQUIRED").Keys
        If Right(CStr(key), 6) = "_NAMED" Then
            AreAllReferencesResolved = False
            Exit Function
        End If
    Next key
End Function

' Obtenir contextes comparaison pour un champ
Public Function GetComparisonContextsForField(registry As Object, fieldRef As String) As Collection
    Set GetComparisonContextsForField = New Collection
    
    If Not registry.Exists("COMPARISON_CONTEXTS") Then
        Exit Function
    End If
    
    Dim allContexts As Collection: Set allContexts = registry("COMPARISON_CONTEXTS")
    Dim i As Long
    
    For i = 1 To allContexts.Count
        Dim context As ComparisonContext: context = allContexts(i)
        If context.fieldName = fieldRef Then
            GetComparisonContextsForField.Add context
        End If
    Next i
End Function

' Obtenir mapping complet source ? extract pour performance
Public Function GetSourceToExtractMapping(registry As Object) As Object
    If registry.Exists("SOURCE_TO_EXTRACT_MAP") Then
        Set GetSourceToExtractMapping = registry("SOURCE_TO_EXTRACT_MAP")
    Else
        Set GetSourceToExtractMapping = CreateObject("Scripting.Dictionary")
    End If
End Function

' Obtenir position optimale pour accès données
Public Function GetOptimalAccessOrder(registry As Object) As Collection
    Set GetOptimalAccessOrder = New Collection
    
    If Not registry.Exists("EXTRACT_TO_FIELD") Then
        Exit Function
    End If
    
    Dim extractToField As Object: Set extractToField = registry("EXTRACT_TO_FIELD")
    Dim i As Long
    
    ' Retourner dans l'ordre optimal (comparaisons d'abord)
    For i = 1 To extractToField.Count
        If extractToField.Exists(i) Then
            GetOptimalAccessOrder.Add extractToField(i)
        End If
    Next i
End Function

' Valider cohérence mappings
Public Function ValidateMappingConsistency(registry As Object) As Boolean
    On Error GoTo ErrorHandler
    
    ValidateMappingConsistency = True
    
    ' Vérifier cohérence source ? extract
    Dim key As Variant
    For Each key In registry("ALL_REQUIRED").Keys
        Dim hasSource As Boolean: hasSource = registry("SOURCE_POSITIONS").Exists(key)
        Dim hasExtract As Boolean: hasExtract = registry("EXTRACT_POSITIONS").Exists(key)
        
        ' Les deux doivent exister ou aucun
        If hasSource <> hasExtract Then
            ValidateMappingConsistency = False
            If GetConfigValue("DebugMode") Then
                Debug.Print "INCOHÉRENCE mapping pour " & key & ": Source=" & hasSource & ", Extract=" & hasExtract
            End If
        End If
    Next key
    
    Exit Function
    
ErrorHandler:
    ValidateMappingConsistency = False
End Function

' ============================================================================
' FONCTIONS ANALYSE PERFORMANCE ET STATISTIQUES
' ============================================================================

Public Function GetRegistryStatistics(registry As Object) As Object
    Set GetRegistryStatistics = CreateObject("Scripting.Dictionary")
    
    Dim stats As Object: Set stats = GetRegistryStatistics
    
    stats("WHAT_FIELDS_COUNT") = registry("WHAT_FIELDS").Count
    stats("READ_FIELDS_COUNT") = registry("READ_FIELDS").Count
    stats("COMPARISON_FIELDS_COUNT") = registry("COMPARISON_FIELDS").Count
    stats("ALL_REQUIRED_COUNT") = registry("ALL_REQUIRED").Count
    stats("READ_EQUALS_WHAT") = registry("READ_EQUALS_WHAT")
    
    ' Calcul taux résolution
    Dim resolvedCount As Long: resolvedCount = 0
    Dim totalCount As Long: totalCount = registry("ALL_REQUIRED").Count
    
    Dim key As Variant
    For Each key In registry("ALL_REQUIRED").Keys
        If Right(CStr(key), 6) <> "_NAMED" Then
            resolvedCount = resolvedCount + 1
        End If
    Next key
    
    stats("RESOLUTION_RATE") = IIf(totalCount > 0, resolvedCount / totalCount, 1)
    stats("RESOLVED_COUNT") = resolvedCount
    stats("UNRESOLVED_COUNT") = totalCount - resolvedCount
    
    ' Statistiques ordre
    stats("HAS_CUSTOM_ORDER") = (registry("READ_ORDER").Count > 0)
    stats("MEMORY_EFFICIENCY") = IIf(registry("READ_EQUALS_WHAT"), "HIGH", "NORMAL")
    
End Function

' Recommandations optimisation
Public Function GetOptimizationRecommendations(registry As Object) As Collection
    Set GetOptimizationRecommendations = New Collection
    
    Dim stats As Object: Set stats = GetRegistryStatistics(registry)
    
    ' Recommandation mémoire
    If Not CBool(stats("READ_EQUALS_WHAT")) And stats("WHAT_FIELDS_COUNT") = stats("READ_FIELDS_COUNT") Then
        GetOptimizationRecommendations.Add "CONSIDER_READ_EQUALS_WHAT: Même colonnes WHAT/READ, optimisation mémoire possible"
    End If
    
    ' Recommandation résolution
    If CSng(stats("RESOLUTION_RATE")) < 1 Then
        GetOptimizationRecommendations.Add "UNRESOLVED_REFERENCES: " & stats("UNRESOLVED_COUNT") & " références non résolues"
    End If
    
    ' Recommandation performance accès
    If stats("COMPARISON_FIELDS_COUNT") > stats("ALL_REQUIRED_COUNT") * 0.5 Then
        GetOptimizationRecommendations.Add "HIGH_COMPARISON_RATIO: Considérer index ou cache pour champs comparaison"
    End If
    
    ' Recommandation ordre
    If Not CBool(stats("HAS_CUSTOM_ORDER")) And stats("READ_FIELDS_COUNT") > 10 Then
        GetOptimizationRecommendations.Add "NO_CUSTOM_ORDER: Ordre READ non spécifié avec nombreuses colonnes"
    End If
End Function

' ============================================================================
' FONCTIONS DEBUG ET DIAGNOSTICS
' ============================================================================

Public Sub DiagnoseRegistry(registry As Object)
    If Not GetConfigValue("DebugMode") Then Exit Sub
    
    Debug.Print "========================================="
    Debug.Print "DIAGNOSTIC REGISTRY COMPLET"
    Debug.Print "========================================="
    
    ' Statistiques générales
    Dim stats As Object: Set stats = GetRegistryStatistics(registry)
    Dim key As Variant
    
    For Each key In stats.Keys
        Debug.Print key & ": " & stats(key)
    Next key
    
    Debug.Print "========================================="
    
    ' Détail mappings
    Debug.Print "MAPPINGS DÉTAILLÉS:"
    For Each key In registry("ALL_REQUIRED").Keys
        Dim sourcePos As String: sourcePos = "N/A"
        Dim extractPos As String: extractPos = "N/A"
        
        If registry("SOURCE_POSITIONS").Exists(key) Then
            sourcePos = CStr(registry("SOURCE_POSITIONS")(key))
        End If
        
        If registry("EXTRACT_POSITIONS").Exists(key) Then
            extractPos = CStr(registry("EXTRACT_POSITIONS")(key))
        End If
        
        Debug.Print "  " & key & " ? Source:" & sourcePos & ", Extract:" & extractPos
    Next key
    
    Debug.Print "========================================="
    
    ' Recommandations
    Dim recommendations As Collection: Set recommendations = GetOptimizationRecommendations(registry)
    If recommendations.Count > 0 Then
        Debug.Print "RECOMMANDATIONS:"
        Dim i As Long
        For i = 1 To recommendations.Count
            Debug.Print "  • " & recommendations(i)
        Next i
    End If
    
    Debug.Print "========================================="
End Sub

' ============================================================================
' UTILITAIRES FINAUX
' ============================================================================

Function ColumnToIndex(colLetter As String) As Long
    ' Réutilisation fonction existante
    Dim result As Long: result = 0
    Dim i As Long
    
    colLetter = UCase(Trim(colLetter))
    If Len(colLetter) = 0 Then
        ColumnToIndex = 1
        Exit Function
    End If
    
    For i = Len(colLetter) To 1 Step -1
        Dim char As String: char = Mid(colLetter, i, 1)
        result = result + (Asc(char) - Asc("A") + 1) * (26 ^ (Len(colLetter) - i))
    Next i
    
    ColumnToIndex = result
End Function

Function IndexToColumn(colIndex As Long) As String
    ' Réutilisation fonction existante
    If colIndex < 1 Then
        IndexToColumn = "A"
        Exit Function
    End If
    
    Dim result As String: result = ""
    Dim tempNum As Long: tempNum = colIndex
    
    Do While tempNum > 0
        tempNum = tempNum - 1
        result = Chr(Asc("A") + (tempNum Mod 26)) + result
        tempNum = tempNum \ 26
    Loop
    
    IndexToColumn = result
End Function

Private Sub SwapLongs(ByRef a As Long, ByRef b As Long)
    Dim temp As Long: temp = a
    a = b: b = temp
End Sub

Function ExtractColumnLetters(cellRef As String) As String
    Dim result As String: result = ""
    Dim i As Long
    
    cellRef = UCase(Trim(cellRef))
    For i = 1 To Len(cellRef)
        Dim char As String: char = Mid(cellRef, i, 1)
        If char >= "A" And char <= "Z" Then
            result = result & char
        Else
            Exit For
        End If
    Next i
    
    ExtractColumnLetters = result
End Function

Function GetConfigValue(configKey As String) As Variant
    On Error GoTo DefaultValue
    
    ' Référence configuration externe supposée
    GetConfigValue = True ' Valeur par défaut pour éviter erreurs
    Exit Function
    
DefaultValue:
    Select Case configKey
        Case "DebugMode": GetConfigValue = False
        Case Else: GetConfigValue = False
    End Select
End Function

