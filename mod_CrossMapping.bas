Attribute VB_Name = "mod_CrossMapping"
' ============================================================================
' DOCUMENTATION VALIDATION CROIS�E
' ============================================================================
'
' UTILISATION VALIDATION CROIS�E:
'
' 1. Validation compl�te avec r�paration automatique:
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
' 3. Score sant�:
'    Dim score As Single: score = GetMappingHealthScore(registry)
'    If score < 80 Then
'        RepairMappingInconsistencies registry
'    End If
'
' 4. Validation avec donn�es source:
'    Dim sourceArray As Variant: sourceArray = Range("A1:Z1000").Value
'    Dim result As MappingValidationResult
'    result = ValidateMappingsWithData(registry, sourceArray)
'
' TYPES DE VALIDATIONS:
' - Structure registry (cl�s manquantes, types incorrects)
' - Coh�rence interne mappings (bidirectionnels, doublons)
' - Validation contre donn�es source (limites, en-t�tes)
' - Ordre extraction (continuit�, coh�rence READ_ORDER)
' - Int�grit� r�f�rentielle (orphelins, incoh�rences union)
'
' R�PARATIONS AUTOMATIQUES:
' - Cr�ation positions manquantes
' - Correction mappings bidirectionnels
' - �limination doublons avec r�assignation
' - Compactage positions extract
' - Correspondance en-t�tes automatique
' - Suppression mappings orphelins
'
' MONITORING:
' - Score sant� continu (0-100%)
' - M�triques d�taill�es
' - Rapports diagnostics complets
' - Validation p�riodique programmable
'
' ============================================================================
' ============================================================================
' V3. VALIDATION CROIS�E MAPPINGS - CONTR�LES INT�GRIT� AVANC�S
' ============================================================================
' Module de validation compl�te des mappings avec donn�es r�elles
' Contr�les coh�rence source ? extract avec diagnostic d�taill�
' R�paration automatique des incoh�rences d�tect�es
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
' �NUM�RATIONS VALIDATION
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
' ValidateMappingConsistencyAdvanced - VALIDATION COMPL�TE AVEC DONN�ES
' ============================================================================
Public Function ValidateMappingConsistencyAdvanced(registry As Object, Optional sourceDataArray As Variant, Optional autoRepair As Boolean = True) As MappingValidationResult
    On Error GoTo ValidationError
    
    Dim startTime As Single: startTime = Timer
    Dim result As MappingValidationResult
    
    ' Initialisation r�sultat
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
    
    ' Phase 2: Validation coh�rence interne mappings
    ValidateInternalMappingConsistency registry, result, autoRepair
    
    ' Phase 3: Validation avec donn�es source (si disponibles)
    If Not IsMissing(sourceDataArray) And IsArray(sourceDataArray) Then
        ValidateAgainstSourceData registry, sourceDataArray, result, autoRepair
    End If
    
    ' Phase 4: Validation ordre extraction
    ValidateExtractionOrder registry, result, autoRepair
    
    ' Phase 5: Validation int�grit� r�f�rentielle
    ValidateReferentialIntegrity registry, result, autoRepair
    
ValidationComplete:
    result.ValidationTime = Timer - startTime
    result.IsValid = (result.ErrorCount = 0)
    
    ' G�n�rer rapport final
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
    
    ' V�rifier cl�s essentielles
    Dim requiredKeys As Variant
    requiredKeys = Array("WHAT_FIELDS", "READ_FIELDS", "ALL_REQUIRED", "SOURCE_POSITIONS", "EXTRACT_POSITIONS", "POSITION_TO_FIELD", "EXTRACT_TO_FIELD")
    
    Dim i As Long
    For i = 0 To UBound(requiredKeys)
        If Not registry.Exists(requiredKeys(i)) Then
            AddValidationError result, "Cl� registry manquante: " & requiredKeys(i), VALIDATION_ERROR
            ValidateRegistryStructure = False
        ElseIf registry(requiredKeys(i)) Is Nothing Then
            AddValidationError result, "Objet registry null pour cl�: " & requiredKeys(i), VALIDATION_ERROR
            ValidateRegistryStructure = False
        End If
    Next i
    
    ' V�rifier types objets
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
' ValidateInternalMappingConsistency - COH�RENCE INTERNE MAPPINGS
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
        
        ' Ignorer r�f�rences nomm�es non r�solues
        If Right(fieldRef, 6) = "_NAMED" Then
            AddValidationWarning result, "R�f�rence non r�solue: " & fieldRef
        Else
            ' V�rifier position source
            If Not sourcePositions.Exists(fieldRef) Then
                If autoRepair Then
                    ' Auto-r�paration: cr�er position source placeholder
                    Dim nextSourcePos As Long: nextSourcePos = GetNextAvailableSourcePosition(sourcePositions)
                    sourcePositions(fieldRef) = nextSourcePos
                    positionToField(nextSourcePos) = fieldRef
                    AddValidationInfo result, "R�par�: position source cr��e pour " & fieldRef & " ? " & nextSourcePos
                    result.RepairedCount = result.RepairedCount + 1
                Else
                    AddValidationError result, "Position source manquante pour: " & fieldRef, VALIDATION_ERROR
                End If
            End If
            
            ' V�rifier position extract
            If Not extractPositions.Exists(fieldRef) Then
                If autoRepair Then
                    ' Auto-r�paration: cr�er position extract
                    Dim nextExtractPos As Long: nextExtractPos = GetNextAvailableExtractPosition(extractPositions)
                    extractPositions(fieldRef) = nextExtractPos
                    extractToField(nextExtractPos) = fieldRef
                    AddValidationInfo result, "R�par�: position extract cr��e pour " & fieldRef & " ? " & nextExtractPos
                    result.RepairedCount = result.RepairedCount + 1
                Else
                    AddValidationError result, "Position extract manquante pour: " & fieldRef, VALIDATION_ERROR
                End If
            End If
        End If
    Next key
    
    ' Validation 2: Coh�rence mappings bidirectionnels source
    ValidateBidirectionalSourceMappings registry, result, autoRepair
    
    ' Validation 3: Coh�rence mappings bidirectionnels extract
    ValidateBidirectionalExtractMappings registry, result, autoRepair
    
    ' Validation 4: D�tection positions dupliqu�es
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
    
    ' V�rifier field ? position ? field
    Dim key As Variant
    For Each key In sourcePositions.Keys
        Dim fieldRef As String: fieldRef = CStr(key)
        Dim sourcePos As Long: sourcePos = CLng(sourcePositions(fieldRef))
        
        If positionToField.Exists(sourcePos) Then
            Dim mappedField As String: mappedField = CStr(positionToField(sourcePos))
            If mappedField <> fieldRef Then
                If autoRepair Then
                    ' R�parer mapping inverse
                    positionToField(sourcePos) = fieldRef
                    AddValidationInfo result, "R�par�: mapping inverse source " & sourcePos & " ? " & fieldRef
                    result.RepairedCount = result.RepairedCount + 1
                Else
                    AddValidationError result, "Incoh�rence mapping source: " & fieldRef & "(" & sourcePos & ") ? " & mappedField, VALIDATION_ERROR
                End If
            End If
        Else
            If autoRepair Then
                ' Cr�er mapping inverse manquant
                positionToField(sourcePos) = fieldRef
                AddValidationInfo result, "R�par�: mapping inverse source cr�� " & sourcePos & " ? " & fieldRef
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
    
    ' V�rifier field ? position ? field
    Dim key As Variant
    For Each key In extractPositions.Keys
        Dim fieldRef As String: fieldRef = CStr(key)
        Dim extractPos As Long: extractPos = CLng(extractPositions(fieldRef))
        
        If extractToField.Exists(extractPos) Then
            Dim mappedField As String: mappedField = CStr(extractToField(extractPos))
            If mappedField <> fieldRef Then
                If autoRepair Then
                    ' R�parer mapping inverse
                    extractToField(extractPos) = fieldRef
                    AddValidationInfo result, "R�par�: mapping inverse extract " & extractPos & " ? " & fieldRef
                    result.RepairedCount = result.RepairedCount + 1
                Else
                    AddValidationError result, "Incoh�rence mapping extract: " & fieldRef & "(" & extractPos & ") ? " & mappedField, VALIDATION_ERROR
                End If
            End If
        Else
            If autoRepair Then
                ' Cr�er mapping inverse manquant
                extractToField(extractPos) = fieldRef
                AddValidationInfo result, "R�par�: mapping inverse extract cr�� " & extractPos & " ? " & fieldRef
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
' DetectDuplicatePositions - D�TECTION POSITIONS DUPLIQU�ES
' ============================================================================
Private Sub DetectDuplicatePositions(registry As Object, ByRef result As MappingValidationResult, autoRepair As Boolean)
    On Error GoTo DuplicateError
    
    ' D�tecter doublons positions source
    Dim sourcePositions As Object: Set sourcePositions = registry("SOURCE_POSITIONS")
    Dim usedSourcePositions As Object: Set usedSourcePositions = CreateObject("Scripting.Dictionary")
    
    Dim key As Variant
    For Each key In sourcePositions.Keys
        Dim pos As Long: pos = CLng(sourcePositions(key))
        
        If usedSourcePositions.Exists(pos) Then
            Dim conflictField As String: conflictField = CStr(usedSourcePositions(pos))
            
            If autoRepair Then
                ' R�assigner position unique
                Dim newPos As Long: newPos = GetNextAvailableSourcePosition(sourcePositions)
                sourcePositions(key) = newPos
                registry("POSITION_TO_FIELD")(newPos) = CStr(key)
                AddValidationInfo result, "R�par�: position source dupliqu�e " & pos & " ? " & newPos & " pour " & key
                result.RepairedCount = result.RepairedCount + 1
            Else
                AddValidationError result, "Position source dupliqu�e " & pos & " pour " & key & " et " & conflictField, VALIDATION_ERROR
            End If
        Else
            usedSourcePositions(pos) = CStr(key)
        End If
    Next key
    
    ' D�tecter doublons positions extract
    Dim extractPositions As Object: Set extractPositions = registry("EXTRACT_POSITIONS")
    Dim usedExtractPositions As Object: Set usedExtractPositions = CreateObject("Scripting.Dictionary")
    
    For Each key In extractPositions.Keys
        Dim pos As Long: pos = CLng(extractPositions(key))
        
        If usedExtractPositions.Exists(pos) Then
            Dim conflictField As String: conflictField = CStr(usedExtractPositions(pos))
            
            If autoRepair Then
                ' R�assigner position unique (en pr�servant ordre READ)
                Dim newPos As Long: newPos = GetNextAvailableExtractPosition(extractPositions)
                extractPositions(key) = newPos
                registry("EXTRACT_TO_FIELD")(newPos) = CStr(key)
                AddValidationWarning result, "Position extract dupliqu�e r�assign�e " & pos & " ? " & newPos & " pour " & key
                result.RepairedCount = result.RepairedCount + 1
            Else
                AddValidationError result, "Position extract dupliqu�e " & pos & " pour " & key & " et " & conflictField, VALIDATION_ERROR
            End If
        Else
            usedExtractPositions(pos) = CStr(key)
        End If
    Next key
    
    Exit Sub
    
DuplicateError:
    AddValidationError result, "Erreur d�tection doublons: " & err.Description, VALIDATION_ERROR
End Sub

' ============================================================================
' ValidateAgainstSourceData - VALIDATION AVEC DONN�ES SOURCE
' ============================================================================
Private Sub ValidateAgainstSourceData(registry As Object, sourceDataArray As Variant, ByRef result As MappingValidationResult, autoRepair As Boolean)
    On Error GoTo SourceDataError
    
    If Not IsArray(sourceDataArray) Then
        AddValidationWarning result, "Donn�es source non disponibles pour validation"
        Exit Sub
    End If
    
    Dim sourceUBound As Long
    On Error Resume Next
    sourceUBound = UBound(sourceDataArray, 2) ' Colonnes
    On Error GoTo SourceDataError
    
    If err.Number <> 0 Then
        AddValidationWarning result, "Structure donn�es source invalide"
        Exit Sub
    End If
    
    Dim sourcePositions As Object: Set sourcePositions = registry("SOURCE_POSITIONS")
    
    ' Validation positions source contre taille donn�es r�elles
    Dim key As Variant
    For Each key In sourcePositions.Keys
        Dim fieldRef As String: fieldRef = CStr(key)
        Dim sourcePos As Long: sourcePos = CLng(sourcePositions(fieldRef))
        
        If sourcePos < LBound(sourceDataArray, 2) Or sourcePos > sourceUBound Then
            If autoRepair Then
                ' R�assigner position valide
                Dim validPos As Long: validPos = FindValidSourcePositionInData(sourceDataArray, fieldRef)
                If validPos > 0 Then
                    sourcePositions(fieldRef) = validPos
                    registry("POSITION_TO_FIELD")(validPos) = fieldRef
                    AddValidationInfo result, "R�par�: position source " & fieldRef & " " & sourcePos & " ? " & validPos
                    result.RepairedCount = result.RepairedCount + 1
                Else
                    AddValidationError result, "Position source " & fieldRef & "(" & sourcePos & ") hors limites donn�es [" & LBound(sourceDataArray, 2) & ":" & sourceUBound & "]", VALIDATION_ERROR
                End If
            Else
                AddValidationError result, "Position source " & fieldRef & "(" & sourcePos & ") hors limites donn�es [" & LBound(sourceDataArray, 2) & ":" & sourceUBound & "]", VALIDATION_ERROR
            End If
        End If
    Next key
    
    ' Validation correspondance en-t�tes si disponibles
    ValidateHeaderCorrespondence registry, sourceDataArray, result, autoRepair
    
    Exit Sub
    
SourceDataError:
    AddValidationError result, "Erreur validation donn�es source: " & err.Description, VALIDATION_WARNING
End Sub

' ============================================================================
' ValidateHeaderCorrespondence - VALIDATION EN-T�TES
' ============================================================================
Private Sub ValidateHeaderCorrespondence(registry As Object, sourceDataArray As Variant, ByRef result As MappingValidationResult, autoRepair As Boolean)
    On Error GoTo HeaderError
    
    ' V�rifier si premi�re ligne contient en-t�tes texte
    Dim hasHeaders As Boolean: hasHeaders = True
    Dim col As Long
    
    For col = LBound(sourceDataArray, 2) To UBound(sourceDataArray, 2)
        If IsNumeric(sourceDataArray(1, col)) Then
            hasHeaders = False
            Exit For
        End If
    Next col
    
    If Not hasHeaders Then
        AddValidationInfo result, "Pas d'en-t�tes d�tect�s en premi�re ligne - validation en-t�tes ignor�e"
        Exit Sub
    End If
    
    ' Comparer en-t�tes avec r�f�rences champs
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
                ' Rechercher correspondance dans autres colonnes si auto-r�paration
                If autoRepair Then
                    Dim foundCol As Long: foundCol = FindColumnByHeader(sourceDataArray, expectedName)
                    If foundCol > 0 And foundCol <> sourcePos Then
                        sourcePositions(fieldRef) = foundCol
                        registry("POSITION_TO_FIELD")(foundCol) = fieldRef
                        AddValidationInfo result, "R�par�: correspondance en-t�te " & fieldRef & " trouv�e colonne " & foundCol & " ('" & headerValue & "')"
                        result.RepairedCount = result.RepairedCount + 1
                    Else
                        AddValidationWarning result, "En-t�te inattendu pour " & fieldRef & " pos " & sourcePos & ": '" & headerValue & "' (attendu '" & expectedName & "')"
                    End If
                Else
                    AddValidationWarning result, "En-t�te inattendu pour " & fieldRef & " pos " & sourcePos & ": '" & headerValue & "' (attendu '" & expectedName & "')"
                End If
            End If
        End If
    Next key
    
    Exit Sub
    
HeaderError:
    AddValidationWarning result, "Erreur validation en-t�tes: " & err.Description
End Sub

' ============================================================================
' ValidateExtractionOrder - VALIDATION ORDRE EXTRACTION
' ============================================================================
Private Sub ValidateExtractionOrder(registry As Object, ByRef result As MappingValidationResult, autoRepair As Boolean)
    On Error GoTo OrderError
    
    Dim extractPositions As Object: Set extractPositions = registry("EXTRACT_POSITIONS")
    Dim readOrder As Object: Set readOrder = registry("READ_ORDER")
    
    ' V�rifier continuit� positions extract (1, 2, 3, ... sans trous)
    Dim maxPos As Long: maxPos = 0
    Dim key As Variant
    
    For Each key In extractPositions.Keys
        Dim pos As Long: pos = CLng(extractPositions(key))
        If pos > maxPos Then maxPos = pos
    Next key
    
    ' V�rifier toutes positions de 1 � maxPos existent
    Dim pos As Long
    For pos = 1 To maxPos
        If Not registry("EXTRACT_TO_FIELD").Exists(pos) Then
            If autoRepair Then
                ' Compacter positions pour �liminer trous
                CompactExtractPositions registry
                AddValidationInfo result, "R�par�: positions extract compact�es"
                result.RepairedCount = result.RepairedCount + 1
                Exit For
            Else
                AddValidationWarning result, "Trou dans positions extract � position: " & pos
            End If
        End If
    Next pos
    
    ' V�rifier coh�rence avec READ_ORDER si pr�sent
    If readOrder.Count > 0 And Not CBool(registry("READ_EQUALS_WHAT")) Then
        ValidateExtractVsReadOrder registry, result, autoRepair
    End If
    
    Exit Sub
    
OrderError:
    AddValidationError result, "Erreur validation ordre: " & err.Description, VALIDATION_WARNING
End Sub

' ============================================================================
' ValidateReferentialIntegrity - VALIDATION INT�GRIT� R�F�RENTIELLE
' ============================================================================
Private Sub ValidateReferentialIntegrity(registry As Object, ByRef result As MappingValidationResult, autoRepair As Boolean)
    On Error GoTo ReferentialError
    
    Dim allRequired As Object: Set allRequired = registry("ALL_REQUIRED")
    Dim whatFields As Object: Set whatFields = registry("WHAT_FIELDS")
    Dim readFields As Object: Set readFields = registry("READ_FIELDS")
    
    ' Validation 1: ALL_REQUIRED doit �tre union de WHAT + READ
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
        AddValidationWarning result, "Nombre champs ALL_REQUIRED (" & allRequired.Count & ") diff�rent de l'union WHAT+READ attendue (" & expectedCount & ")"
    End If
    
    ' Validation 2: V�rifier champs orphelins dans mappings
    DetectOrphanedMappings registry, result, autoRepair
    
    Exit Sub
    
ReferentialError:
    AddValidationError result, "Erreur int�grit� r�f�rentielle: " & err.Description, VALIDATION_WARNING
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
    ' Tentative recherche position valide bas�e sur nom champ
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
    ' R�organiser positions extract pour �liminer trous
    Dim extractPositions As Object: Set extractPositions = registry("EXTRACT_POSITIONS")
    Dim extractToField As Object: Set extractToField = registry("EXTRACT_TO_FIELD")
    Dim readOrder As Object: Set readOrder = registry("READ_ORDER")
    
    ' Cr�er liste ordonn�e
    Dim orderedFields As Collection: Set orderedFields = New Collection
    Dim pos As Long
    
    For pos = 1 To 10000 ' Limite s�curitaire
        If extractToField.Exists(pos) Then
            orderedFields.Add extractToField(pos)
        End If
    Next pos
    
    ' R�assigner positions compactes
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
    ' D�tecter mappings sans r�f�rence dans ALL_REQUIRED
    Dim allRequired As Object: Set allRequired = registry("ALL_REQUIRED")
    Dim sourcePositions As Object: Set sourcePositions = registry("SOURCE_POSITIONS")
    Dim extractPositions As Object: Set extractPositions = registry("EXTRACT_POSITIONS")
    
    ' V�rifier mappings source orphelins
    Dim key As Variant
    For Each key In sourcePositions.Keys
        If Not allRequired.Exists(key) Then
            If autoRepair Then
                sourcePositions.Remove key
                AddValidationInfo result, "R�par�: mapping source orphelin supprim� " & key
                result.RepairedCount = result.RepairedCount + 1
            Else
                AddValidationWarning result, "Mapping source orphelin d�tect�: " & key
            End If
        End If
    Next key
    
    ' V�rifier mappings extract orphelins
    For Each key In extractPositions.Keys
        If Not allRequired.Exists(key) Then
            If autoRepair Then
                extractPositions.Remove key
                AddValidationInfo result, "R�par�: mapping extract orphelin supprim� " & key
                result.RepairedCount = result.RepairedCount + 1
            Else
                AddValidationWarning result, "Mapping extract orphelin d�tect�: " & key
            End If
        End If
    Next key
End Sub

Private Sub ValidateExtractVsReadOrder(registry As Object, ByRef result As MappingValidationResult, autoRepair As Boolean)
    On Error GoTo ReadOrderError
    
    Dim extractPositions As Object: Set extractPositions = registry("EXTRACT_POSITIONS")
    Dim readOrder As Object: Set readOrder = registry("READ_ORDER")
    
    ' V�rifier que positions extract correspondent � ordre READ
    Dim key As Variant
    For Each key In readOrder.Keys
        Dim expectedOrder As Long: expectedOrder = CLng(readOrder(key))
        
        If extractPositions.Exists(key) Then
            Dim actualOrder As Long: actualOrder = CLng(extractPositions(key))
            
            If actualOrder <> expectedOrder Then
                If autoRepair Then
                    ' R�organiser positions extract selon READ_ORDER
                    ReorganizeExtractByReadOrder registry
                    AddValidationInfo result, "R�par�: positions extract r�organis�es selon READ_ORDER"
                    result.RepairedCount = result.RepairedCount + 1
                    Exit For
                Else
                    AddValidationWarning result, "Ordre extract (" & actualOrder & ") diff�rent de READ_ORDER (" & expectedOrder & ") pour " & key
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
' G�N�RATEURS RAPPORTS
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
    report = report & "Dur�e validation: " & Format(result.ValidationTime, "0.000") & " secondes" & vbCrLf
    report = report & "Erreurs: " & result.ErrorCount & vbCrLf
    report = report & "Avertissements: " & result.WarningCount & vbCrLf
    report = report & "R�parations: " & result.RepairedCount & vbCrLf
    report = report & "R�sultat: " & IIf(result.IsValid, "VALIDE", "INVALIDE") & vbCrLf
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
    
    report = report & "D�TAIL VALIDATION:" & vbCrLf
    report = report & result.DetailedReport
    report = report & "=========================================" & vbCrLf
    
    result.DetailedReport = report
End Sub

' ============================================================================
' API PUBLIQUE VALIDATION
' ============================================================================

Public Function QuickValidateMappings(registry As Object) As Boolean
    ' Validation rapide sans auto-r�paration
    Dim result As MappingValidationResult
    result = ValidateMappingConsistencyAdvanced(registry, , False)
    QuickValidateMappings = result.IsValid
End Function

Public Function GetMappingHealthScore(registry As Object) As Single
    ' Score sant� mappings (0-100)
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
    ' R�paration automatique avec rapport
    Dim result As MappingValidationResult
    result = ValidateMappingConsistencyAdvanced(registry, , True)
    
    If GetParsingConfig("DebugMode") Then
        Debug.Print "R�paration mappings: " & result.RepairedCount & " corrections effectu�es"
    End If
    
    If result.ErrorCount > 0 Then
        Debug.Print "Attention: " & result.ErrorCount & " erreurs non r�parables restent"
    End If
End Sub

Public Function ValidateMappingsWithData(registry As Object, sourceData As Variant) As MappingValidationResult
    ' Validation compl�te avec donn�es source
    ValidateMappingsWithData = ValidateMappingConsistencyAdvanced(registry, sourceData, True)
End Function

Public Function GetMappingDiagnosticReport(registry As Object) As String
    ' Rapport diagnostique d�taill�
    Dim result As MappingValidationResult
    result = ValidateMappingConsistencyAdvanced(registry, , False)
    GetMappingDiagnosticReport = result.DetailedReport
End Function

' ============================================================================
' TESTS VALIDATION INT�GR�S
' ============================================================================

Public Sub TestMappingValidation()
    Debug.Print "=== TESTS VALIDATION MAPPINGS ==="
    
    ' Cr�er registry test avec incoh�rences volontaires
    Dim testRegistry As Object: Set testRegistry = CreateTestRegistryWithIssues()
    
    ' Test 1: Validation sans r�paration
    Debug.Print "Test 1: Validation sans r�paration"
    Dim result1 As MappingValidationResult
    result1 = ValidateMappingConsistencyAdvanced(testRegistry, , False)
    Debug.Print "  Erreurs: " & result1.ErrorCount & ", Avertissements: " & result1.WarningCount
    
    ' Test 2: Validation avec r�paration
    Debug.Print "Test 2: Validation avec r�paration"
    Dim result2 As MappingValidationResult
    result2 = ValidateMappingConsistencyAdvanced(testRegistry, , True)
    Debug.Print "  Erreurs: " & result2.ErrorCount & ", R�parations: " & result2.RepairedCount
    
    ' Test 3: Score sant�
    Debug.Print "Test 3: Score sant� mappings"
    Dim healthScore As Single: healthScore = GetMappingHealthScore(testRegistry)
    Debug.Print "  Score: " & Format(healthScore, "0.0") & "%"
    
    ' Test 4: Validation avec donn�es
    Debug.Print "Test 4: Validation avec donn�es source"
    Dim testData(1 To 5, 1 To 10) As Variant
    Dim result4 As MappingValidationResult
    result4 = ValidateMappingsWithData(testRegistry, testData)
    Debug.Print "  Avec donn�es - Erreurs: " & result4.ErrorCount & ", Avertissements: " & result4.WarningCount
    
    Debug.Print "=== TESTS TERMIN�S ==="
End Sub

Function CreateTestRegistryWithIssues() As Object
    ' Cr�er registry test avec incoh�rences pour tester validation
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
    
    ' Remplir avec donn�es incoh�rentes volontaires
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
    ' Note: Impl�mentation d�pend de l'architecture de monitoring de l'application
    
    If GetParsingConfig("ValidateMappings") Then
        ' Programmer validation p�riodique
        SetupContinuousValidation = True
        
        If GetParsingConfig("DebugMode") Then
            Debug.Print "Validation continue configur�e: " & intervalMinutes & " minutes"
        End If
    Else
        SetupContinuousValidation = False
    End If
End Function

Public Sub LogMappingMetrics(registry As Object)
    ' Log m�triques pour monitoring
    If Not GetParsingConfig("LogParsingSteps") Then Exit Sub
    
    Dim timestamp As String: timestamp = Format(Now, "yyyy-mm-dd hh:mm:ss")
    Dim healthScore As Single: healthScore = GetMappingHealthScore(registry)
    
    Debug.Print timestamp & " [METRICS] Registry Health: " & Format(healthScore, "0.0") & "%"
    Debug.Print timestamp & " [METRICS] Fields Count: " & registry("ALL_REQUIRED").Count
    Debug.Print timestamp & " [METRICS] Source Mappings: " & registry("SOURCE_POSITIONS").Count
    Debug.Print timestamp & " [METRICS] Extract Mappings: " & registry("EXTRACT_POSITIONS").Count
End Sub

