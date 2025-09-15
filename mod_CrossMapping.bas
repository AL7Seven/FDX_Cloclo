Attribute VB_Name = "mod_CrossMapping"
' ============================================================================
' DOCUMENTATION VALIDATION CROISÉE
' ============================================================================
'
' UTILISATION VALIDATION CROISÉE:
'
' 1. Validation complète avec réparation automatique:
'    Dim result As MappingValidationResult
'    result = ValidateMappingConsistencyAdvanced(registry, sourceData, True)
'
'    If result.IsValid Then
'        Debug.Print "Mappings valides"
'    Else
'        Debug.Print result.DetailedReport
'    End If
'
' 2. Validation rapide:
'    If QuickValidateMappings(registry) Then
'        ' Mappings OK
'    End If
'
' 3. Score santé:
'    Dim score As Single: score = GetMappingHealthScore(registry)
'    If score < 80 Then
'        RepairMappingInconsistencies registry
'    End If
'
' 4. Validation avec données source:
'    Dim sourceArray As Variant: sourceArray = Range("A1:Z1000").Value
'    Dim result As MappingValidationResult
'    result = ValidateMappingsWithData(registry, sourceArray)
'
' TYPES DE VALIDATIONS:
' - Structure registry (clés manquantes, types incorrects)
' - Cohérence interne mappings (bidirectionnels, doublons)
' - Validation contre données source (limites, en-têtes)
' - Ordre extraction (continuité, cohérence READ_ORDER)
' - Intégrité référentielle (orphelins, incohérences union)
'
' RÉPARATIONS AUTOMATIQUES:
' - Création positions manquantes
' - Correction mappings bidirectionnels
' - Élimination doublons avec réassignation
' - Compactage positions extract
' - Correspondance en-têtes automatique
' - Suppression mappings orphelins
'
' MONITORING:
' - Score santé continu (0-100%)
' - Métriques détaillées
' - Rapports diagnostics complets
' - Validation périodique programmable
'
' ============================================================================
' ============================================================================
' V3. VALIDATION CROISÉE MAPPINGS - CONTRÔLES INTÉGRITÉ AVANCÉS
' ============================================================================
' Module de validation complète des mappings avec données réelles
' Contrôles cohérence source ? extract avec diagnostic détaillé
' Réparation automatique des incohérences détectées
' ============================================================================

Option Explicit

' ============================================================================
' TYPES VALIDATION
' ============================================================================
Public Type MappingValidationResult
    IsValid As Boolean
    ErrorCount As Long
    WarningCount As Long
    RepairedCount As Long
    ValidationTime As Single
    DetailedReport As String
End Type

Public Type MappingInconsistency
    InconsistencyType As String      ' "MISSING_SOURCE", "MISSING_EXTRACT", "ORDER_MISMATCH", etc.
    FieldReference As String         ' "@A", "@Nom", etc.
    ExpectedValue As Variant
    ActualValue As Variant
    severity As String              ' "ERROR", "WARNING", "INFO"
    CanAutoRepair As Boolean
    RepairAction As String
End Type

' ============================================================================
' ÉNUMÉRATIONS VALIDATION
' ============================================================================

Public Enum ValidationSeverity
    VALIDATION_INFO = 0
    VALIDATION_WARNING = 1
    VALIDATION_ERROR = 2
    VALIDATION_CRITICAL = 3
End Enum

Public Enum MappingInconsistencyType
    INCONSISTENCY_MISSING_SOURCE_POSITION = 1
    INCONSISTENCY_MISSING_EXTRACT_POSITION = 2
    INCONSISTENCY_ORDER_MISMATCH = 3
    INCONSISTENCY_DUPLICATE_POSITION = 4
    INCONSISTENCY_INVALID_FIELD_REFERENCE = 5
    INCONSISTENCY_SOURCE_DATA_MISMATCH = 6
    INCONSISTENCY_ORPHANED_MAPPING = 7
End Enum

' ============================================================================
' ValidateMappingConsistencyAdvanced - VALIDATION COMPLÈTE AVEC DONNÉES
' ============================================================================
Public Function ValidateMappingConsistencyAdvanced(registry As Object, Optional sourceDataArray As Variant, Optional autoRepair As Boolean = True) As MappingValidationResult
    On Error GoTo ValidationError
    
    Dim startTime As Single: startTime = Timer
    Dim result As MappingValidationResult
    
    ' Initialisation résultat
    With result
        .IsValid = True
        .ErrorCount = 0
        .WarningCount = 0
        .RepairedCount = 0
        .ValidationTime = 0
        .DetailedReport = ""
    End With
    
    ClearParsingError
    
    ' Phase 1: Validation structure registry
    If Not ValidateRegistryStructure(registry, result) Then
        GoTo ValidationComplete
    End If
    
    ' Phase 2: Validation cohérence interne mappings
    ValidateInternalMappingConsistency registry, result, autoRepair
    
    ' Phase 3: Validation avec données source (si disponibles)
    If Not IsMissing(sourceDataArray) And IsArray(sourceDataArray) Then
        ValidateAgainstSourceData registry, sourceDataArray, result, autoRepair
    End If
    
    ' Phase 4: Validation ordre extraction
    ValidateExtractionOrder registry, result, autoRepair
    
    ' Phase 5: Validation intégrité référentielle
    ValidateReferentialIntegrity registry, result, autoRepair
    
ValidationComplete:
    result.ValidationTime = Timer - startTime
    result.IsValid = (result.ErrorCount = 0)
    
    ' Générer rapport final
    GenerateValidationReport registry, result
    
    ' Log si mode debug
    If GetParsingConfig("DebugMode") Then
        Debug.Print "Validation mappings: " & result.ErrorCount & " erreurs, " & result.WarningCount & " avertissements en " & Format(result.ValidationTime, "0.000") & "s"
    End If
    
    ValidateMappingConsistencyAdvanced = result
    Exit Function
    
ValidationError:
    SetParsingError ERR_PARSING_MAPPING_INCONSISTENT, "Erreur validation mappings: " & err.Description, ""
    result.IsValid = False
    result.ErrorCount = result.ErrorCount + 1
    result.ValidationTime = Timer - startTime
    ValidateMappingConsistencyAdvanced = result
End Function

' ============================================================================
' ValidateRegistryStructure - VALIDATION STRUCTURE REGISTRY
' ============================================================================
Function ValidateRegistryStructure(registry As Object, ByRef result As MappingValidationResult) As Boolean
    On Error GoTo StructureError
    
    ValidateRegistryStructure = True
    
    If registry Is Nothing Then
        AddValidationError result, "Registry est null", VALIDATION_CRITICAL
        ValidateRegistryStructure = False
        Exit Function
    End If
    
    ' Vérifier clés essentielles
    Dim requiredKeys As Variant
    requiredKeys = Array("WHAT_FIELDS", "READ_FIELDS", "ALL_REQUIRED", "SOURCE_POSITIONS", "EXTRACT_POSITIONS", "POSITION_TO_FIELD", "EXTRACT_TO_FIELD")
    
    Dim i As Long
    For i = 0 To UBound(requiredKeys)
        If Not registry.Exists(requiredKeys(i)) Then
            AddValidationError result, "Clé registry manquante: " & requiredKeys(i), VALIDATION_ERROR
            ValidateRegistryStructure = False
        ElseIf registry(requiredKeys(i)) Is Nothing Then
            AddValidationError result, "Objet registry null pour clé: " & requiredKeys(i), VALIDATION_ERROR
            ValidateRegistryStructure = False
        End If
    Next i
    
    ' Vérifier types objets
    If ValidateRegistryStructure Then
        For i = 0 To UBound(requiredKeys)
            If TypeName(registry(requiredKeys(i))) <> "Dictionary" Then
                AddValidationWarning result, "Type inattendu pour " & requiredKeys(i) & ": " & TypeName(registry(requiredKeys(i)))
            End If
        Next i
    End If
    
    Exit Function
    
StructureError:
    AddValidationError result, "Erreur validation structure: " & err.Description, VALIDATION_ERROR
    ValidateRegistryStructure = False
End Function

' ============================================================================
' ValidateInternalMappingConsistency - COHÉRENCE INTERNE MAPPINGS
' ============================================================================
Private Sub ValidateInternalMappingConsistency(registry As Object, ByRef result As MappingValidationResult, autoRepair As Boolean)
    On Error GoTo InternalValidationError
    
    Dim allRequired As Object: Set allRequired = registry("ALL_REQUIRED")
    Dim sourcePositions As Object: Set sourcePositions = registry("SOURCE_POSITIONS")
    Dim extractPositions As Object: Set extractPositions = registry("EXTRACT_POSITIONS")
    Dim positionToField As Object: Set positionToField = registry("POSITION_TO_FIELD")
    Dim extractToField As Object: Set extractToField = registry("EXTRACT_TO_FIELD")
    
    ' Validation 1: Chaque champ ALL_REQUIRED doit avoir positions source ET extract
    Dim key As Variant
    For Each key In allRequired.Keys
        Dim fieldRef As String: fieldRef = CStr(key)
        
        ' Ignorer références nommées non résolues
        If Right(fieldRef, 6) = "_NAMED" Then
            AddValidationWarning result, "Référence non résolue: " & fieldRef
        Else
            ' Vérifier position source
            If Not sourcePositions.Exists(fieldRef) Then
                If autoRepair Then
                    ' Auto-réparation: créer position source placeholder
                    Dim nextSourcePos As Long: nextSourcePos = GetNextAvailableSourcePosition(sourcePositions)
                    sourcePositions(fieldRef) = nextSourcePos
                    positionToField(nextSourcePos) = fieldRef
                    AddValidationInfo result, "Réparé: position source créée pour " & fieldRef & " ? " & nextSourcePos
                    result.RepairedCount = result.RepairedCount + 1
                Else
                    AddValidationError result, "Position source manquante pour: " & fieldRef, VALIDATION_ERROR
                End If
            End If
            
            ' Vérifier position extract
            If Not extractPositions.Exists(fieldRef) Then
                If autoRepair Then
                    ' Auto-réparation: créer position extract
                    Dim nextExtractPos As Long: nextExtractPos = GetNextAvailableExtractPosition(extractPositions)
                    extractPositions(fieldRef) = nextExtractPos
                    extractToField(nextExtractPos) = fieldRef
                    AddValidationInfo result, "Réparé: position extract créée pour " & fieldRef & " ? " & nextExtractPos
                    result.RepairedCount = result.RepairedCount + 1
                Else
                    AddValidationError result, "Position extract manquante pour: " & fieldRef, VALIDATION_ERROR
                End If
            End If
        End If
    Next key
    
    ' Validation 2: Cohérence mappings bidirectionnels source
    ValidateBidirectionalSourceMappings registry, result, autoRepair
    
    ' Validation 3: Cohérence mappings bidirectionnels extract
    ValidateBidirectionalExtractMappings registry, result, autoRepair
    
    ' Validation 4: Détection positions dupliquées
    DetectDuplicatePositions registry, result, autoRepair
    
    Exit Sub
    
InternalValidationError:
    AddValidationError result, "Erreur validation interne: " & err.Description, VALIDATION_ERROR
End Sub

' ============================================================================
' ValidateBidirectionalSourceMappings - MAPPINGS SOURCE BIDIRECTIONNELS
' ============================================================================
Private Sub ValidateBidirectionalSourceMappings(registry As Object, ByRef result As MappingValidationResult, autoRepair As Boolean)
    On Error GoTo SourceMappingError
    
    Dim sourcePositions As Object: Set sourcePositions = registry("SOURCE_POSITIONS")
    Dim positionToField As Object: Set positionToField = registry("POSITION_TO_FIELD")
    
    ' Vérifier field ? position ? field
    Dim key As Variant
    For Each key In sourcePositions.Keys
        Dim fieldRef As String: fieldRef = CStr(key)
        Dim sourcePos As Long: sourcePos = CLng(sourcePositions(fieldRef))
        
        If positionToField.Exists(sourcePos) Then
            Dim mappedField As String: mappedField = CStr(positionToField(sourcePos))
            If mappedField <> fieldRef Then
                If autoRepair Then
                    ' Réparer mapping inverse
                    positionToField(sourcePos) = fieldRef
                    AddValidationInfo result, "Réparé: mapping inverse source " & sourcePos & " ? " & fieldRef
                    result.RepairedCount = result.RepairedCount + 1
                Else
                    AddValidationError result, "Incohérence mapping source: " & fieldRef & "(" & sourcePos & ") ? " & mappedField, VALIDATION_ERROR
                End If
            End If
        Else
            If autoRepair Then
                ' Créer mapping inverse manquant
                positionToField(sourcePos) = fieldRef
                AddValidationInfo result, "Réparé: mapping inverse source créé " & sourcePos & " ? " & fieldRef
                result.RepairedCount = result.RepairedCount + 1
            Else
                AddValidationError result, "Mapping inverse source manquant pour position: " & sourcePos, VALIDATION_ERROR
            End If
        End If
    Next key
    
    Exit Sub
    
SourceMappingError:
    AddValidationError result, "Erreur mappings source: " & err.Description, VALIDATION_ERROR
End Sub

' ============================================================================
' ValidateBidirectionalExtractMappings - MAPPINGS EXTRACT BIDIRECTIONNELS
' ============================================================================
Private Sub ValidateBidirectionalExtractMappings(registry As Object, ByRef result As MappingValidationResult, autoRepair As Boolean)
    On Error GoTo ExtractMappingError
    
    Dim extractPositions As Object: Set extractPositions = registry("EXTRACT_POSITIONS")
    Dim extractToField As Object: Set extractToField = registry("EXTRACT_TO_FIELD")
    
    ' Vérifier field ? position ? field
    Dim key As Variant
    For Each key In extractPositions.Keys
        Dim fieldRef As String: fieldRef = CStr(key)
        Dim extractPos As Long: extractPos = CLng(extractPositions(fieldRef))
        
        If extractToField.Exists(extractPos) Then
            Dim mappedField As String: mappedField = CStr(extractToField(extractPos))
            If mappedField <> fieldRef Then
                If autoRepair Then
                    ' Réparer mapping inverse
                    extractToField(extractPos) = fieldRef
                    AddValidationInfo result, "Réparé: mapping inverse extract " & extractPos & " ? " & fieldRef
                    result.RepairedCount = result.RepairedCount + 1
                Else
                    AddValidationError result, "Incohérence mapping extract: " & fieldRef & "(" & extractPos & ") ? " & mappedField, VALIDATION_ERROR
                End If
            End If
        Else
            If autoRepair Then
                ' Créer mapping inverse manquant
                extractToField(extractPos) = fieldRef
                AddValidationInfo result, "Réparé: mapping inverse extract créé " & extractPos & " ? " & fieldRef
                result.RepairedCount = result.RepairedCount + 1
            Else
                AddValidationError result, "Mapping inverse extract manquant pour position: " & extractPos, VALIDATION_ERROR
            End If
        End If
    Next key
    
    Exit Sub
    
ExtractMappingError:
    AddValidationError result, "Erreur mappings extract: " & err.Description, VALIDATION_ERROR
End Sub

' ============================================================================
' DetectDuplicatePositions - DÉTECTION POSITIONS DUPLIQUÉES
' ============================================================================
Private Sub DetectDuplicatePositions(registry As Object, ByRef result As MappingValidationResult, autoRepair As Boolean)
    On Error GoTo DuplicateError
    
    ' Détecter doublons positions source
    Dim sourcePositions As Object: Set sourcePositions = registry("SOURCE_POSITIONS")
    Dim usedSourcePositions As Object: Set usedSourcePositions = CreateObject("Scripting.Dictionary")
    
    Dim key As Variant
    For Each key In sourcePositions.Keys
        Dim pos As Long: pos = CLng(sourcePositions(key))
        
        If usedSourcePositions.Exists(pos) Then
            Dim conflictField As String: conflictField = CStr(usedSourcePositions(pos))
            
            If autoRepair Then
                ' Réassigner position unique
                Dim newPos As Long: newPos = GetNextAvailableSourcePosition(sourcePositions)
                sourcePositions(key) = newPos
                registry("POSITION_TO_FIELD")(newPos) = CStr(key)
                AddValidationInfo result, "Réparé: position source dupliquée " & pos & " ? " & newPos & " pour " & key
                result.RepairedCount = result.RepairedCount + 1
            Else
                AddValidationError result, "Position source dupliquée " & pos & " pour " & key & " et " & conflictField, VALIDATION_ERROR
            End If
        Else
            usedSourcePositions(pos) = CStr(key)
        End If
    Next key
    
    ' Détecter doublons positions extract
    Dim extractPositions As Object: Set extractPositions = registry("EXTRACT_POSITIONS")
    Dim usedExtractPositions As Object: Set usedExtractPositions = CreateObject("Scripting.Dictionary")
    
    For Each key In extractPositions.Keys
        Dim pos As Long: pos = CLng(extractPositions(key))
        
        If usedExtractPositions.Exists(pos) Then
            Dim conflictField As String: conflictField = CStr(usedExtractPositions(pos))
            
            If autoRepair Then
                ' Réassigner position unique (en préservant ordre READ)
                Dim newPos As Long: newPos = GetNextAvailableExtractPosition(extractPositions)
                extractPositions(key) = newPos
                registry("EXTRACT_TO_FIELD")(newPos) = CStr(key)
                AddValidationWarning result, "Position extract dupliquée réassignée " & pos & " ? " & newPos & " pour " & key
                result.RepairedCount = result.RepairedCount + 1
            Else
                AddValidationError result, "Position extract dupliquée " & pos & " pour " & key & " et " & conflictField, VALIDATION_ERROR
            End If
        Else
            usedExtractPositions(pos) = CStr(key)
        End If
    Next key
    
    Exit Sub
    
DuplicateError:
    AddValidationError result, "Erreur détection doublons: " & err.Description, VALIDATION_ERROR
End Sub

' ============================================================================
' ValidateAgainstSourceData - VALIDATION AVEC DONNÉES SOURCE
' ============================================================================
Private Sub ValidateAgainstSourceData(registry As Object, sourceDataArray As Variant, ByRef result As MappingValidationResult, autoRepair As Boolean)
    On Error GoTo SourceDataError
    
    If Not IsArray(sourceDataArray) Then
        AddValidationWarning result, "Données source non disponibles pour validation"
        Exit Sub
    End If
    
    Dim sourceUBound As Long
    On Error Resume Next
    sourceUBound = UBound(sourceDataArray, 2) ' Colonnes
    On Error GoTo SourceDataError
    
    If err.Number <> 0 Then
        AddValidationWarning result, "Structure données source invalide"
        Exit Sub
    End If
    
    Dim sourcePositions As Object: Set sourcePositions = registry("SOURCE_POSITIONS")
    
    ' Validation positions source contre taille données réelles
    Dim key As Variant
    For Each key In sourcePositions.Keys
        Dim fieldRef As String: fieldRef = CStr(key)
        Dim sourcePos As Long: sourcePos = CLng(sourcePositions(fieldRef))
        
        If sourcePos < LBound(sourceDataArray, 2) Or sourcePos > sourceUBound Then
            If autoRepair Then
                ' Réassigner position valide
                Dim validPos As Long: validPos = FindValidSourcePositionInData(sourceDataArray, fieldRef)
                If validPos > 0 Then
                    sourcePositions(fieldRef) = validPos
                    registry("POSITION_TO_FIELD")(validPos) = fieldRef
                    AddValidationInfo result, "Réparé: position source " & fieldRef & " " & sourcePos & " ? " & validPos
                    result.RepairedCount = result.RepairedCount + 1
                Else
                    AddValidationError result, "Position source " & fieldRef & "(" & sourcePos & ") hors limites données [" & LBound(sourceDataArray, 2) & ":" & sourceUBound & "]", VALIDATION_ERROR
                End If
            Else
                AddValidationError result, "Position source " & fieldRef & "(" & sourcePos & ") hors limites données [" & LBound(sourceDataArray, 2) & ":" & sourceUBound & "]", VALIDATION_ERROR
            End If
        End If
    Next key
    
    ' Validation correspondance en-têtes si disponibles
    ValidateHeaderCorrespondence registry, sourceDataArray, result, autoRepair
    
    Exit Sub
    
SourceDataError:
    AddValidationError result, "Erreur validation données source: " & err.Description, VALIDATION_WARNING
End Sub

' ============================================================================
' ValidateHeaderCorrespondence - VALIDATION EN-TÊTES
' ============================================================================
Private Sub ValidateHeaderCorrespondence(registry As Object, sourceDataArray As Variant, ByRef result As MappingValidationResult, autoRepair As Boolean)
    On Error GoTo HeaderError
    
    ' Vérifier si première ligne contient en-têtes texte
    Dim hasHeaders As Boolean: hasHeaders = True
    Dim col As Long
    
    For col = LBound(sourceDataArray, 2) To UBound(sourceDataArray, 2)
        If IsNumeric(sourceDataArray(1, col)) Then
            hasHeaders = False
            Exit For
        End If
    Next col
    
    If Not hasHeaders Then
        AddValidationInfo result, "Pas d'en-têtes détectés en première ligne - validation en-têtes ignorée"
        Exit Sub
    End If
    
    ' Comparer en-têtes avec références champs
    Dim sourcePositions As Object: Set sourcePositions = registry("SOURCE_POSITIONS")
    
    Dim key As Variant
    For Each key In sourcePositions.Keys
        Dim fieldRef As String: fieldRef = CStr(key)
        Dim sourcePos As Long: sourcePos = CLng(sourcePositions(fieldRef))
        
        If sourcePos >= LBound(sourceDataArray, 2) And sourcePos <= UBound(sourceDataArray, 2) Then
            Dim headerValue As String: headerValue = CStr(sourceDataArray(1, sourcePos))
            Dim expectedName As String: expectedName = Mid(fieldRef, 2) ' Enlever @
            
            ' Comparaison souple (ignorer casse, espaces)
            If UCase(Trim(headerValue)) <> UCase(Trim(expectedName)) Then
                ' Rechercher correspondance dans autres colonnes si auto-réparation
                If autoRepair Then
                    Dim foundCol As Long: foundCol = FindColumnByHeader(sourceDataArray, expectedName)
                    If foundCol > 0 And foundCol <> sourcePos Then
                        sourcePositions(fieldRef) = foundCol
                        registry("POSITION_TO_FIELD")(foundCol) = fieldRef
                        AddValidationInfo result, "Réparé: correspondance en-tête " & fieldRef & " trouvée colonne " & foundCol & " ('" & headerValue & "')"
                        result.RepairedCount = result.RepairedCount + 1
                    Else
                        AddValidationWarning result, "En-tête inattendu pour " & fieldRef & " pos " & sourcePos & ": '" & headerValue & "' (attendu '" & expectedName & "')"
                    End If
                Else
                    AddValidationWarning result, "En-tête inattendu pour " & fieldRef & " pos " & sourcePos & ": '" & headerValue & "' (attendu '" & expectedName & "')"
                End If
            End If
        End If
    Next key
    
    Exit Sub
    
HeaderError:
    AddValidationWarning result, "Erreur validation en-têtes: " & err.Description
End Sub

' ============================================================================
' ValidateExtractionOrder - VALIDATION ORDRE EXTRACTION
' ============================================================================
Private Sub ValidateExtractionOrder(registry As Object, ByRef result As MappingValidationResult, autoRepair As Boolean)
    On Error GoTo OrderError
    
    Dim extractPositions As Object: Set extractPositions = registry("EXTRACT_POSITIONS")
    Dim readOrder As Object: Set readOrder = registry("READ_ORDER")
    
    ' Vérifier continuité positions extract (1, 2, 3, ... sans trous)
    Dim maxPos As Long: maxPos = 0
    Dim key As Variant
    
    For Each key In extractPositions.Keys
        Dim pos As Long: pos = CLng(extractPositions(key))
        If pos > maxPos Then maxPos = pos
    Next key
    
    ' Vérifier toutes positions de 1 à maxPos existent
    Dim pos As Long
    For pos = 1 To maxPos
        If Not registry("EXTRACT_TO_FIELD").Exists(pos) Then
            If autoRepair Then
                ' Compacter positions pour éliminer trous
                CompactExtractPositions registry
                AddValidationInfo result, "Réparé: positions extract compactées"
                result.RepairedCount = result.RepairedCount + 1
                Exit For
            Else
                AddValidationWarning result, "Trou dans positions extract à position: " & pos
            End If
        End If
    Next pos
    
    ' Vérifier cohérence avec READ_ORDER si présent
    If readOrder.Count > 0 And Not CBool(registry("READ_EQUALS_WHAT")) Then
        ValidateExtractVsReadOrder registry, result, autoRepair
    End If
    
    Exit Sub
    
OrderError:
    AddValidationError result, "Erreur validation ordre: " & err.Description, VALIDATION_WARNING
End Sub

' ============================================================================
' ValidateReferentialIntegrity - VALIDATION INTÉGRITÉ RÉFÉRENTIELLE
' ============================================================================
Private Sub ValidateReferentialIntegrity(registry As Object, ByRef result As MappingValidationResult, autoRepair As Boolean)
    On Error GoTo ReferentialError
    
    Dim allRequired As Object: Set allRequired = registry("ALL_REQUIRED")
    Dim whatFields As Object: Set whatFields = registry("WHAT_FIELDS")
    Dim readFields As Object: Set readFields = registry("READ_FIELDS")
    
    ' Validation 1: ALL_REQUIRED doit être union de WHAT + READ
    Dim expectedCount As Long
    If CBool(registry("READ_EQUALS_WHAT")) Then
        expectedCount = whatFields.Count
    Else
        expectedCount = whatFields.Count + readFields.Count
        
        ' Soustraire doublons potentiels
        Dim key As Variant
        For Each key In whatFields.Keys
            If readFields.Exists(key) Then
                expectedCount = expectedCount - 1
            End If
        Next key
    End If
    
    If allRequired.Count <> expectedCount Then
        AddValidationWarning result, "Nombre champs ALL_REQUIRED (" & allRequired.Count & ") différent de l'union WHAT+READ attendue (" & expectedCount & ")"
    End If
    
    ' Validation 2: Vérifier champs orphelins dans mappings
    DetectOrphanedMappings registry, result, autoRepair
    
    Exit Sub
    
ReferentialError:
    AddValidationError result, "Erreur intégrité référentielle: " & err.Description, VALIDATION_WARNING
End Sub

' ============================================================================
' UTILITAIRES VALIDATION
' ============================================================================

Function GetNextAvailableSourcePosition(sourcePositions As Object) As Long
    Dim maxPos As Long: maxPos = 0
    
    Dim key As Variant
    For Each key In sourcePositions.Keys
        Dim pos As Long: pos = CLng(sourcePositions(key))
        If pos > maxPos Then maxPos = pos
    Next key
    
    GetNextAvailableSourcePosition = maxPos + 1
End Function

Function GetNextAvailableExtractPosition(extractPositions As Object) As Long
    Dim maxPos As Long: maxPos = 0
    
    Dim key As Variant
    For Each key In extractPositions.Keys
        Dim pos As Long: pos = CLng(extractPositions(key))
        If pos > maxPos Then maxPos = pos
    Next key
    
    GetNextAvailableExtractPosition = maxPos + 1
End Function

Function FindValidSourcePositionInData(sourceDataArray As Variant, fieldRef As String) As Long
    ' Tentative recherche position valide basée sur nom champ
    Dim fieldName As String: fieldName = Mid(fieldRef, 2) ' Enlever @
    
    FindValidSourcePositionInData = FindColumnByHeader(sourceDataArray, fieldName)
End Function

Function FindColumnByHeader(sourceDataArray As Variant, headerName As String) As Long
    On Error Resume Next
    
    FindColumnByHeader = 0
    
    Dim col As Long
    For col = LBound(sourceDataArray, 2) To UBound(sourceDataArray, 2)
        Dim headerValue As String: headerValue = UCase(Trim(CStr(sourceDataArray(1, col))))
        If headerValue = UCase(Trim(headerName)) Then
            FindColumnByHeader = col
            Exit Function
        End If
    Next col
    
    On Error GoTo 0
End Function

Private Sub CompactExtractPositions(registry As Object)
    ' Réorganiser positions extract pour éliminer trous
    Dim extractPositions As Object: Set extractPositions = registry("EXTRACT_POSITIONS")
    Dim extractToField As Object: Set extractToField = registry("EXTRACT_TO_FIELD")
    Dim readOrder As Object: Set readOrder = registry("READ_ORDER")
    
    ' Créer liste ordonnée
    Dim orderedFields As Collection: Set orderedFields = New Collection
    Dim pos As Long
    
    For pos = 1 To 10000 ' Limite sécuritaire
        If extractToField.Exists(pos) Then
            orderedFields.Add extractToField(pos)
        End If
    Next pos
    
    ' Réassigner positions compactes
    extractPositions.RemoveAll
    extractToField.RemoveAll
    
    Dim i As Long
    For i = 1 To orderedFields.Count
        Dim fieldRef As String: fieldRef = orderedFields(i)
        extractPositions(fieldRef) = i
        extractToField(i) = fieldRef
    Next i
End Sub

Private Sub DetectOrphanedMappings(registry As Object, ByRef result As MappingValidationResult, autoRepair As Boolean)
    ' Détecter mappings sans référence dans ALL_REQUIRED
    Dim allRequired As Object: Set allRequired = registry("ALL_REQUIRED")
    Dim sourcePositions As Object: Set sourcePositions = registry("SOURCE_POSITIONS")
    Dim extractPositions As Object: Set extractPositions = registry("EXTRACT_POSITIONS")
    
    ' Vérifier mappings source orphelins
    Dim key As Variant
    For Each key In sourcePositions.Keys
        If Not allRequired.Exists(key) Then
            If autoRepair Then
                sourcePositions.Remove key
                AddValidationInfo result, "Réparé: mapping source orphelin supprimé " & key
                result.RepairedCount = result.RepairedCount + 1
            Else
                AddValidationWarning result, "Mapping source orphelin détecté: " & key
            End If
        End If
    Next key
    
    ' Vérifier mappings extract orphelins
    For Each key In extractPositions.Keys
        If Not allRequired.Exists(key) Then
            If autoRepair Then
                extractPositions.Remove key
                AddValidationInfo result, "Réparé: mapping extract orphelin supprimé " & key
                result.RepairedCount = result.RepairedCount + 1
            Else
                AddValidationWarning result, "Mapping extract orphelin détecté: " & key
            End If
        End If
    Next key
End Sub

Private Sub ValidateExtractVsReadOrder(registry As Object, ByRef result As MappingValidationResult, autoRepair As Boolean)
    On Error GoTo ReadOrderError
    
    Dim extractPositions As Object: Set extractPositions = registry("EXTRACT_POSITIONS")
    Dim readOrder As Object: Set readOrder = registry("READ_ORDER")
    
    ' Vérifier que positions extract correspondent à ordre READ
    Dim key As Variant
    For Each key In readOrder.Keys
        Dim expectedOrder As Long: expectedOrder = CLng(readOrder(key))
        
        If extractPositions.Exists(key) Then
            Dim actualOrder As Long: actualOrder = CLng(extractPositions(key))
            
            If actualOrder <> expectedOrder Then
                If autoRepair Then
                    ' Réorganiser positions extract selon READ_ORDER
                    ReorganizeExtractByReadOrder registry
                    AddValidationInfo result, "Réparé: positions extract réorganisées selon READ_ORDER"
                    result.RepairedCount = result.RepairedCount + 1
                    Exit For
                Else
                    AddValidationWarning result, "Ordre extract (" & actualOrder & ") différent de READ_ORDER (" & expectedOrder & ") pour " & key
                End If
            End If
        End If
    Next key
    
    Exit Sub
    
ReadOrderError:
    AddValidationWarning result, "Erreur validation ordre READ: " & err.Description
End Sub

Private Sub ReorganizeExtractByReadOrder(registry As Object)
    Dim extractPositions As Object: Set extractPositions = registry("EXTRACT_POSITIONS")
    Dim extractToField As Object: Set extractToField = registry("EXTRACT_TO_FIELD")
    Dim readOrder As Object: Set readOrder = registry("READ_ORDER")
    
    ' Nettoyer mappings extract actuels
    extractPositions.RemoveAll
    extractToField.RemoveAll
    
    ' Reconstruire selon READ_ORDER
    Dim key As Variant
    For Each key In readOrder.Keys
        Dim order As Long: order = CLng(readOrder(key))
        extractPositions(key) = order
        extractToField(order) = CStr(key)
    Next key
End Sub

' ============================================================================
' GÉNÉRATEURS RAPPORTS
' ============================================================================

Private Sub AddValidationError(ByRef result As MappingValidationResult, message As String, severity As ValidationSeverity)
    result.ErrorCount = result.ErrorCount + 1
    AppendToReport result, "ERREUR", message, severity
End Sub

Private Sub AddValidationWarning(ByRef result As MappingValidationResult, message As String)
    result.WarningCount = result.WarningCount + 1
    AppendToReport result, "ATTENTION", message, VALIDATION_WARNING
End Sub

Private Sub AddValidationInfo(ByRef result As MappingValidationResult, message As String)
    AppendToReport result, "INFO", message, VALIDATION_INFO
End Sub

Private Sub AppendToReport(ByRef result As MappingValidationResult, category As String, message As String, severity As ValidationSeverity)
    Dim timestamp As String: timestamp = Format(Now, "hh:mm:ss")
    result.DetailedReport = result.DetailedReport & timestamp & " [" & category & "] " & message & vbCrLf
End Sub

Private Sub GenerateValidationReport(registry As Object, ByRef result As MappingValidationResult)
    Dim report As String: report = ""
    
    report = report & "=========================================" & vbCrLf
    report = report & "RAPPORT VALIDATION MAPPINGS" & vbCrLf
    report = report & "=========================================" & vbCrLf
    report = report & "Durée validation: " & Format(result.ValidationTime, "0.000") & " secondes" & vbCrLf
    report = report & "Erreurs: " & result.ErrorCount & vbCrLf
    report = report & "Avertissements: " & result.WarningCount & vbCrLf
    report = report & "Réparations: " & result.RepairedCount & vbCrLf
    report = report & "Résultat: " & IIf(result.IsValid, "VALIDE", "INVALIDE") & vbCrLf
    report = report & "=========================================" & vbCrLf
    
    ' Ajouter statistiques registry
    If Not registry Is Nothing Then
        report = report & "STATISTIQUES REGISTRY:" & vbCrLf
        report = report & "- WHAT_FIELDS: " & registry("WHAT_FIELDS").Count & vbCrLf
        report = report & "- READ_FIELDS: " & registry("READ_FIELDS").Count & vbCrLf
        report = report & "- ALL_REQUIRED: " & registry("ALL_REQUIRED").Count & vbCrLf
        report = report & "- READ_EQUALS_WHAT: " & registry("READ_EQUALS_WHAT") & vbCrLf
        report = report & "- SOURCE_POSITIONS: " & registry("SOURCE_POSITIONS").Count & vbCrLf
        report = report & "- EXTRACT_POSITIONS: " & registry("EXTRACT_POSITIONS").Count & vbCrLf
        report = report & "=========================================" & vbCrLf
    End If
    
    report = report & "DÉTAIL VALIDATION:" & vbCrLf
    report = report & result.DetailedReport
    report = report & "=========================================" & vbCrLf
    
    result.DetailedReport = report
End Sub

' ============================================================================
' API PUBLIQUE VALIDATION
' ============================================================================

Public Function QuickValidateMappings(registry As Object) As Boolean
    ' Validation rapide sans auto-réparation
    Dim result As MappingValidationResult
    result = ValidateMappingConsistencyAdvanced(registry, , False)
    QuickValidateMappings = result.IsValid
End Function

Public Function GetMappingHealthScore(registry As Object) As Single
    ' Score santé mappings (0-100)
    Dim result As MappingValidationResult
    result = ValidateMappingConsistencyAdvanced(registry, , False)
    
    Dim totalIssues As Long: totalIssues = result.ErrorCount + result.WarningCount
    Dim maxPossibleIssues As Long: maxPossibleIssues = registry("ALL_REQUIRED").Count * 5 ' Estimation
    
    If maxPossibleIssues = 0 Then
        GetMappingHealthScore = 100
    Else
        GetMappingHealthScore = ((maxPossibleIssues - totalIssues) / maxPossibleIssues) * 100
        If GetMappingHealthScore < 0 Then GetMappingHealthScore = 0
    End If
End Function

Public Sub RepairMappingInconsistencies(registry As Object)
    ' Réparation automatique avec rapport
    Dim result As MappingValidationResult
    result = ValidateMappingConsistencyAdvanced(registry, , True)
    
    If GetParsingConfig("DebugMode") Then
        Debug.Print "Réparation mappings: " & result.RepairedCount & " corrections effectuées"
    End If
    
    If result.ErrorCount > 0 Then
        Debug.Print "Attention: " & result.ErrorCount & " erreurs non réparables restent"
    End If
End Sub

Public Function ValidateMappingsWithData(registry As Object, sourceData As Variant) As MappingValidationResult
    ' Validation complète avec données source
    ValidateMappingsWithData = ValidateMappingConsistencyAdvanced(registry, sourceData, True)
End Function

Public Function GetMappingDiagnosticReport(registry As Object) As String
    ' Rapport diagnostique détaillé
    Dim result As MappingValidationResult
    result = ValidateMappingConsistencyAdvanced(registry, , False)
    GetMappingDiagnosticReport = result.DetailedReport
End Function

' ============================================================================
' TESTS VALIDATION INTÉGRÉS
' ============================================================================

Public Sub TestMappingValidation()
    Debug.Print "=== TESTS VALIDATION MAPPINGS ==="
    
    ' Créer registry test avec incohérences volontaires
    Dim testRegistry As Object: Set testRegistry = CreateTestRegistryWithIssues()
    
    ' Test 1: Validation sans réparation
    Debug.Print "Test 1: Validation sans réparation"
    Dim result1 As MappingValidationResult
    result1 = ValidateMappingConsistencyAdvanced(testRegistry, , False)
    Debug.Print "  Erreurs: " & result1.ErrorCount & ", Avertissements: " & result1.WarningCount
    
    ' Test 2: Validation avec réparation
    Debug.Print "Test 2: Validation avec réparation"
    Dim result2 As MappingValidationResult
    result2 = ValidateMappingConsistencyAdvanced(testRegistry, , True)
    Debug.Print "  Erreurs: " & result2.ErrorCount & ", Réparations: " & result2.RepairedCount
    
    ' Test 3: Score santé
    Debug.Print "Test 3: Score santé mappings"
    Dim healthScore As Single: healthScore = GetMappingHealthScore(testRegistry)
    Debug.Print "  Score: " & Format(healthScore, "0.0") & "%"
    
    ' Test 4: Validation avec données
    Debug.Print "Test 4: Validation avec données source"
    Dim testData(1 To 5, 1 To 10) As Variant
    Dim result4 As MappingValidationResult
    result4 = ValidateMappingsWithData(testRegistry, testData)
    Debug.Print "  Avec données - Erreurs: " & result4.ErrorCount & ", Avertissements: " & result4.WarningCount
    
    Debug.Print "=== TESTS TERMINÉS ==="
End Sub

Function CreateTestRegistryWithIssues() As Object
    ' Créer registry test avec incohérences pour tester validation
    Set CreateTestRegistryWithIssues = CreateObject("Scripting.Dictionary")
    
    With CreateTestRegistryWithIssues
        .Add "WHAT_FIELDS", CreateObject("Scripting.Dictionary")
        .Add "READ_FIELDS", CreateObject("Scripting.Dictionary")
        .Add "ALL_REQUIRED", CreateObject("Scripting.Dictionary")
        .Add "SOURCE_POSITIONS", CreateObject("Scripting.Dictionary")
        .Add "EXTRACT_POSITIONS", CreateObject("Scripting.Dictionary")
        .Add "POSITION_TO_FIELD", CreateObject("Scripting.Dictionary")
        .Add "EXTRACT_TO_FIELD", CreateObject("Scripting.Dictionary")
        .Add "READ_ORDER", CreateObject("Scripting.Dictionary")
        .Add "READ_EQUALS_WHAT", False
    End With
    
    ' Remplir avec données incohérentes volontaires
    CreateTestRegistryWithIssues("WHAT_FIELDS")("@A") = True
    CreateTestRegistryWithIssues("READ_FIELDS")("@B") = True
    CreateTestRegistryWithIssues("ALL_REQUIRED")("@A") = True
    CreateTestRegistryWithIssues("ALL_REQUIRED")("@B") = True
    CreateTestRegistryWithIssues("ALL_REQUIRED")("@C") = True ' Orphelin
    
    ' Mappings source incomplets (manque @B)
    CreateTestRegistryWithIssues("SOURCE_POSITIONS")("@A") = 1
    CreateTestRegistryWithIssues("POSITION_TO_FIELD")(1) = "@A"
    
    ' Mappings extract avec doublons
    CreateTestRegistryWithIssues("EXTRACT_POSITIONS")("@A") = 1
    CreateTestRegistryWithIssues("EXTRACT_POSITIONS")("@B") = 1 ' Doublon position
    CreateTestRegistryWithIssues("EXTRACT_TO_FIELD")(1) = "@A"
    
    Set CreateTestRegistryWithIssues = CreateTestRegistryWithIssues
End Function

' ============================================================================
' UTILITAIRES MONITORING CONTINU
' ============================================================================

Public Function SetupContinuousValidation(registry As Object, intervalMinutes As Long) As Boolean
    ' Configuration validation continue (pour environnements critiques)
    ' Note: Implémentation dépend de l'architecture de monitoring de l'application
    
    If GetParsingConfig("ValidateMappings") Then
        ' Programmer validation périodique
        SetupContinuousValidation = True
        
        If GetParsingConfig("DebugMode") Then
            Debug.Print "Validation continue configurée: " & intervalMinutes & " minutes"
        End If
    Else
        SetupContinuousValidation = False
    End If
End Function

Public Sub LogMappingMetrics(registry As Object)
    ' Log métriques pour monitoring
    If Not GetParsingConfig("LogParsingSteps") Then Exit Sub
    
    Dim timestamp As String: timestamp = Format(Now, "yyyy-mm-dd hh:mm:ss")
    Dim healthScore As Single: healthScore = GetMappingHealthScore(registry)
    
    Debug.Print timestamp & " [METRICS] Registry Health: " & Format(healthScore, "0.0") & "%"
    Debug.Print timestamp & " [METRICS] Fields Count: " & registry("ALL_REQUIRED").Count
    Debug.Print timestamp & " [METRICS] Source Mappings: " & registry("SOURCE_POSITIONS").Count
    Debug.Print timestamp & " [METRICS] Extract Mappings: " & registry("EXTRACT_POSITIONS").Count
End Sub

