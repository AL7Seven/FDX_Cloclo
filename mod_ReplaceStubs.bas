Attribute VB_Name = "mod_ReplaceStubs"
' ============================================================================
' DOCUMENTATION COMPLÉTION
' ============================================================================
'
' FONCTIONNALITÉS COMPLÉTÉES:
'
' 1. BuildUnionAndMappingsSafe - Construction robuste union + mappings
' 2. ParseRangeWithOrderSafe - Support multi-ranges avec ordre préservé
' 3. ParseSingleColumnWithOrderSafe - Colonnes simples sécurisées
' 4. Support complet formats: Excel (A:B, A1:B2), numérique (1:3), nommé (Date:Facture)
' 5. Validation rigoureuse avec gestion erreurs spécifiques
' 6. Limites sécuritaires sur toutes les opérations
' 7. Logging complet pour debugging
'
' TESTS À EFFECTUER:
' - Multi-ranges: "A1:B2,EF10:EG10,C5"
' - Ranges numériques: "1:3,5,8:10"
' - Références nommées: "Clients:Montants"
' - Mixte: "[A,C:E,10:12]"
' - Cas limites: ranges vides, colonnes > XFD, etc.
'
' ============================================================================
' ============================================================================
' COMPLÉTION STUBS PRIORITÉ 1 - IMPLÉMENTATION COMPLÈTE
' ============================================================================
' Implémentation sécurisée des fonctions manquantes
' Support multi-ranges réels avec ordre préservé
' ============================================================================

Option Explicit

' ============================================================================
' BuildUnionAndMappingsSafe - CONSTRUCTION UNION ET MAPPINGS SÉCURISÉE
' ============================================================================
Function BuildUnionAndMappingsSafe(registry As Object) As Boolean
    On Error GoTo MappingError
    
    BuildUnionAndMappingsSafe = False
    
    Dim whatFields As Object: Set whatFields = registry("WHAT_FIELDS")
    Dim readFields As Object: Set readFields = registry("READ_FIELDS")
    Dim allRequired As Object: Set allRequired = registry("ALL_REQUIRED")
    Dim colIndex As Object: Set colIndex = registry("COLUMN_INDEX")
    Dim indexCol As Object: Set indexCol = registry("INDEX_COLUMN")
    Dim extractPos As Object: Set extractPos = registry("EXTRACT_POSITIONS")
    Dim extractToField As Object: Set extractToField = registry("EXTRACT_TO_FIELD")
    
    ' Validation objets nécessaires
    If whatFields Is Nothing Or readFields Is Nothing Or allRequired Is Nothing Then
        SetParsingError ERR_PARSING_MAPPING_INCONSISTENT, "Objets registry manquants", ""
        Exit Function
    End If
    
    ' Union des champs WHAT
    Dim key As Variant
    For Each key In whatFields.Keys
        If Not SafeAddToDictionary(allRequired, CStr(key), True) Then
            SetParsingError ERR_PARSING_MEMORY_EXCEEDED, "Echec ajout champ WHAT: " & key, ""
            Exit Function
        End If
    Next key
    
    ' Union des champs READ (seulement si READ ? WHAT)
    If Not CBool(registry("READ_EQUALS_WHAT")) Then
        For Each key In readFields.Keys
            If Not SafeAddToDictionary(allRequired, CStr(key), True) Then
                SetParsingError ERR_PARSING_MEMORY_EXCEEDED, "Echec ajout champ READ: " & key, ""
                Exit Function
            End If
        Next key
    End If
    
    ' S'assurer qu'on a au minimum un champ
    If allRequired.Count = 0 Then
        If Not SafeAddToDictionary(allRequired, "@A", True) Then
            SetParsingError ERR_PARSING_MEMORY_EXCEEDED, "Impossible d'ajouter fallback @A", ""
            Exit Function
        End If
    End If
    
    ' Construction mappings bidirectionnels lettre ? index
    For Each key In allRequired.Keys
        Dim keyStr As String: keyStr = CStr(key)
        
        ' Vérifier format @Lettre (ignorer marqueurs _NAMED)
        If Left(keyStr, 1) = "@" And Right(keyStr, 6) <> "_NAMED" Then
            Dim colLetter As String: colLetter = Mid(keyStr, 2)
            
            If Len(colLetter) > 0 And Len(colLetter) <= 3 Then
                Dim colIdx As Long: colIdx = ColumnToIndexSafe(colLetter)
                
                If colIdx > 0 Then
                    If Not SafeAddToDictionary(colIndex, colLetter, colIdx) Then
                        SetParsingError ERR_PARSING_MEMORY_EXCEEDED, "Echec mapping colonne->index: " & colLetter, ""
                        Exit Function
                    End If
                    
                    If Not SafeAddToDictionary(indexCol, CStr(colIdx), colLetter) Then
                        SetParsingError ERR_PARSING_MEMORY_EXCEEDED, "Echec mapping index->colonne: " & colIdx, ""
                        Exit Function
                    End If
                End If
            End If
        End If
    Next key
    
    ' Construction mappings positions tableau extrait
    If Not BuildExtractPositionMappingsSafe(registry) Then
        Exit Function ' Erreur déjà définie
    End If
    
    BuildUnionAndMappingsSafe = True
    Exit Function
    
MappingError:
    SetParsingError ERR_PARSING_MAPPING_INCONSISTENT, "Erreur construction mappings: " & err.Description, ""
    BuildUnionAndMappingsSafe = False
End Function

' ============================================================================
' BuildExtractPositionMappingsSafe - MAPPINGS POSITIONS TABLEAU EXTRAIT
' ============================================================================
Function BuildExtractPositionMappingsSafe(registry As Object) As Boolean
    On Error GoTo ExtractMappingError
    
    BuildExtractPositionMappingsSafe = False
    
    Dim readOrder As Object: Set readOrder = registry("READ_ORDER")
    Dim extractPos As Object: Set extractPos = registry("EXTRACT_POSITIONS")
    Dim extractToField As Object: Set extractToField = registry("EXTRACT_TO_FIELD")
    Dim whatFields As Object: Set whatFields = registry("WHAT_FIELDS")
    
    ' Si READ_EQUALS_WHAT, utiliser ordre WHAT (ordre découverte dans expression)
    If CBool(registry("READ_EQUALS_WHAT")) Then
        Dim pos As Long: pos = 1
        Dim key As Variant
        
        For Each key In whatFields.Keys
            Dim keyStr As String
            keyStr = CStr(key)
            
            If Not SafeAddToDictionary(extractPos, keyStr, pos) Then
                SetParsingError ERR_PARSING_MEMORY_EXCEEDED, "Echec mapping extract position: " & keyStr, ""
                Exit Function
            End If
            
            If Not SafeAddToDictionary(extractToField, CStr(pos), keyStr) Then
                SetParsingError ERR_PARSING_MEMORY_EXCEEDED, "Echec mapping position->field: " & pos, ""
                Exit Function
            End If
            
            pos = pos + 1
        Next key
    Else
        ' Utiliser ordre READ préservé (IMPORTANT: respecter ordre utilisateur)
        For Each key In readOrder.Keys
'            Dim keyStr As String
            keyStr = CStr(key)
            Dim orderPos As Variant
            orderPos = readOrder(keyStr)
            
            ' Validation position
            If IsNumeric(orderPos) Then
                Dim orderPosLong As Long: orderPosLong = CLng(orderPos)
                
                If orderPosLong > 0 And orderPosLong <= 10000 Then ' Limite sécuritaire
                    If Not SafeAddToDictionary(extractPos, keyStr, orderPosLong) Then
                        SetParsingError ERR_PARSING_MEMORY_EXCEEDED, "Echec mapping extract ordered: " & keyStr, ""
                        Exit Function
                    End If
                    
                    If Not SafeAddToDictionary(extractToField, CStr(orderPosLong), keyStr) Then
                        SetParsingError ERR_PARSING_MEMORY_EXCEEDED, "Echec mapping ordered->field: " & orderPosLong, ""
                        Exit Function
                    End If
                End If
            End If
        Next key
    End If
    
    BuildExtractPositionMappingsSafe = True
    Exit Function
    
ExtractMappingError:
    SetParsingError ERR_PARSING_MAPPING_INCONSISTENT, "Erreur mappings extract: " & err.Description, ""
    BuildExtractPositionMappingsSafe = False
End Function

' ============================================================================
' ParseRangeWithOrderSafe - PARSE RANGE AVEC ORDRE PRÉSERVÉ SÉCURISÉ
' ============================================================================
Function ParseRangeWithOrderSafe(rangeSpec As String, readFields As Object, readOrder As Object, ByRef orderIndex As Long) As Boolean
    On Error GoTo ParseRangeError
    
    ParseRangeWithOrderSafe = False
    
    If Len(rangeSpec) = 0 Then Exit Function
    
    ' Parser range: A1:B2 ou A:B ou 1:3 ou Date:Facture
    Dim rangeParts As Variant
    rangeParts = Split(rangeSpec, ":")
    
    If UBound(rangeParts) < 1 Then Exit Function
    
    Dim startPart As String: startPart = Trim(CStr(rangeParts(0)))
    Dim endPart As String: endPart = Trim(CStr(rangeParts(1)))
    
    If Len(startPart) = 0 Or Len(endPart) = 0 Then Exit Function
    
    ' Détecter type de range et traiter en conséquence
    If IsExcelColumnFormatSafe(startPart) And IsExcelColumnFormatSafe(endPart) Then
        ' Format Excel : A1:B2 ou A:B
        ParseRangeWithOrderSafe = ParseExcelRangeOrderedSafe(startPart, endPart, readFields, readOrder, orderIndex)
    ElseIf IsNumericRange(startPart) And IsNumericRange(endPart) Then
        ' Format numérique : 1:3, 5:8
        ParseRangeWithOrderSafe = ParseNumericRangeOrderedSafe(startPart, endPart, readFields, readOrder, orderIndex)
    Else
        ' Format nommé : Date:Facture (marquer pour résolution dynamique)
        ParseRangeWithOrderSafe = AddNamedRangeForResolutionSafe(startPart, endPart, readFields, readOrder, orderIndex)
    End If
    
    Exit Function
    
ParseRangeError:
    SetParsingError ERR_PARSING_INVALID_READ_SPEC, "Erreur parsing range: " & err.Description, "range=" & rangeSpec
    ParseRangeWithOrderSafe = False
End Function

' ============================================================================
' ParseExcelRangeOrderedSafe - PARSE RANGE EXCEL SÉCURISÉ
' ============================================================================
Function ParseExcelRangeOrderedSafe(startSpec As String, endSpec As String, readFields As Object, readOrder As Object, ByRef orderIndex As Long) As Boolean
    On Error GoTo ParseExcelError
    
    ParseExcelRangeOrderedSafe = False
    
    Dim startCol As String: startCol = ExtractColumnLettersSafe(startSpec)
    Dim endCol As String: endCol = ExtractColumnLettersSafe(endSpec)
    
    If Len(startCol) = 0 Or Len(endCol) = 0 Then Exit Function
    
    ' Gérer ranges cellules vs colonnes
    Dim startRow As Long: startRow = ExtractRowNumberSafe(startSpec)
    Dim endRow As Long: endRow = ExtractRowNumberSafe(endSpec)
    
    If startRow > 1 Or endRow > 1 Then
        ' Range cellules A1:B2 - extraire colonnes uniques
        ParseExcelRangeOrderedSafe = ParseCellRangeOrderedSafe(startCol, endCol, startRow, endRow, readFields, readOrder, orderIndex)
    Else
        ' Range colonnes A:B
        ParseExcelRangeOrderedSafe = ParseColumnRangeOrderedSafe(startCol, endCol, readFields, readOrder, orderIndex)
    End If
    
    Exit Function
    
ParseExcelError:
    ParseExcelRangeOrderedSafe = False
End Function

' ============================================================================
' ParseCellRangeOrderedSafe - PARSE RANGE CELLULES SÉCURISÉ
' ============================================================================
Function ParseCellRangeOrderedSafe(startCol As String, endCol As String, startRow As Long, endRow As Long, readFields As Object, readOrder As Object, ByRef orderIndex As Long) As Boolean
    On Error GoTo ParseCellError
    
    ParseCellRangeOrderedSafe = False
    
    Dim startColIdx As Long: startColIdx = ColumnToIndexSafe(startCol)
    Dim endColIdx As Long: endColIdx = ColumnToIndexSafe(endCol)
    
    If startColIdx = -1 Or endColIdx = -1 Then Exit Function
    
    ' S'assurer ordre correct
    If startColIdx > endColIdx Then SwapLongsSafe startColIdx, endColIdx
    If startRow > endRow Then SwapLongsSafe startRow, endRow
    
    ' Limites sécuritaires
    If (endColIdx - startColIdx + 1) * (endRow - startRow + 1) > PARSING_MAX_RANGE_SIZE Then
        endColIdx = startColIdx + 10  ' Limite colonnes
        endRow = startRow + 10        ' Limite lignes
    End If
    
    ' Extraire colonnes uniques (ignorer lignes pour parsing colonnes)
    Dim col As Long
    For col = startColIdx To endColIdx
        Dim colLetter As String: colLetter = IndexToColumnSafe(col)
        
        If Len(colLetter) > 0 Then
            Dim fieldRef As String: fieldRef = "@" & colLetter
            
            If Not SafeAddToDictionary(readFields, fieldRef, True) Then
                SetParsingError ERR_PARSING_MEMORY_EXCEEDED, "Echec ajout cellule: " & fieldRef, ""
                Exit Function
            End If
            
            If Not readOrder.Exists(fieldRef) Then
                If Not SafeAddToDictionary(readOrder, fieldRef, orderIndex) Then
                    SetParsingError ERR_PARSING_MEMORY_EXCEEDED, "Echec ordre cellule: " & fieldRef, ""
                    Exit Function
                End If
                orderIndex = orderIndex + 1
            End If
        End If
    Next col
    
    ParseCellRangeOrderedSafe = True
    Exit Function
    
ParseCellError:
    ParseCellRangeOrderedSafe = False
End Function

' ============================================================================
' ParseColumnRangeOrderedSafe - PARSE RANGE COLONNES SÉCURISÉ
' ============================================================================
Function ParseColumnRangeOrderedSafe(startCol As String, endCol As String, readFields As Object, readOrder As Object, ByRef orderIndex As Long) As Boolean
    On Error GoTo ParseColError
    
    ParseColumnRangeOrderedSafe = False
    
    Dim startIdx As Long: startIdx = ColumnToIndexSafe(startCol)
    Dim endIdx As Long: endIdx = ColumnToIndexSafe(endCol)
    
    If startIdx = -1 Or endIdx = -1 Then Exit Function
    
    If startIdx > endIdx Then SwapLongsSafe startIdx, endIdx
    
    ' Limite sécuritaire
    If endIdx - startIdx > PARSING_MAX_RANGE_SIZE Then
        endIdx = startIdx + PARSING_MAX_RANGE_SIZE
    End If
    
    Dim col As Long
    For col = startIdx To endIdx
        Dim colLetter As String: colLetter = IndexToColumnSafe(col)
        
        If Len(colLetter) > 0 Then
            Dim fieldRef As String: fieldRef = "@" & colLetter
            
            If Not SafeAddToDictionary(readFields, fieldRef, True) Then
                SetParsingError ERR_PARSING_MEMORY_EXCEEDED, "Echec ajout colonne: " & fieldRef, ""
                Exit Function
            End If
            
            If Not SafeAddToDictionary(readOrder, fieldRef, orderIndex) Then
                SetParsingError ERR_PARSING_MEMORY_EXCEEDED, "Echec ordre colonne: " & fieldRef, ""
                Exit Function
            End If
            
            orderIndex = orderIndex + 1
        End If
    Next col
    
    ParseColumnRangeOrderedSafe = True
    Exit Function
    
ParseColError:
    ParseColumnRangeOrderedSafe = False
End Function

' ============================================================================
' ParseNumericRangeOrderedSafe - PARSE RANGE NUMÉRIQUE SÉCURISÉ
' ============================================================================
Function ParseNumericRangeOrderedSafe(startSpec As String, endSpec As String, readFields As Object, readOrder As Object, ByRef orderIndex As Long) As Boolean
    On Error GoTo ParseNumError
    
    ParseNumericRangeOrderedSafe = False
    
    Dim startNum As Long: startNum = ExtractNumberSafe(startSpec)
    Dim endNum As Long: endNum = ExtractNumberSafe(endSpec)
    
    If startNum = -1 Or endNum = -1 Then Exit Function
    If startNum < 1 Or endNum < 1 Then Exit Function
    
    If startNum > endNum Then SwapLongsSafe startNum, endNum
    
    ' Limite sécuritaire
    If endNum - startNum > PARSING_MAX_RANGE_SIZE Then
        endNum = startNum + PARSING_MAX_RANGE_SIZE
    End If
    
    Dim col As Long
    For col = startNum To endNum
        Dim colLetter As String: colLetter = IndexToColumnSafe(col)
        
        If Len(colLetter) > 0 Then
            Dim fieldRef As String: fieldRef = "@" & colLetter
            
            If Not SafeAddToDictionary(readFields, fieldRef, True) Then
                SetParsingError ERR_PARSING_MEMORY_EXCEEDED, "Echec ajout numérique: " & fieldRef, ""
                Exit Function
            End If
            
            If Not SafeAddToDictionary(readOrder, fieldRef, orderIndex) Then
                SetParsingError ERR_PARSING_MEMORY_EXCEEDED, "Echec ordre numérique: " & fieldRef, ""
                Exit Function
            End If
            
            orderIndex = orderIndex + 1
        End If
    Next col
    
    ParseNumericRangeOrderedSafe = True
    Exit Function
    
ParseNumError:
    ParseNumericRangeOrderedSafe = False
End Function

' ============================================================================
' ParseSingleColumnWithOrderSafe - PARSE COLONNE SIMPLE SÉCURISÉ
' ============================================================================
Function ParseSingleColumnWithOrderSafe(colSpec As String, readFields As Object, readOrder As Object, ByRef orderIndex As Long) As Boolean
    On Error GoTo ParseSingleError
    
    ParseSingleColumnWithOrderSafe = False
    
    If Len(colSpec) = 0 Or Len(colSpec) > 255 Then Exit Function
    
    Dim fieldRef As String
    
    If IsExcelColumnFormatSafe(colSpec) Then
        ' Format Excel : A1 ou A
        Dim colLetter As String: colLetter = ExtractColumnLettersSafe(colSpec)
        If Len(colLetter) > 0 And IsValidColumnRefSafe(colLetter) Then
            fieldRef = "@" & colLetter
        End If
    ElseIf IsNumericRange(colSpec) Then
        ' Format numérique : 3
        Dim colNum As Long: colNum = ExtractNumberSafe(colSpec)
        If colNum > 0 Then
            Dim colLetter2 As String: colLetter2 = IndexToColumnSafe(colNum)
            If Len(colLetter2) > 0 Then
                fieldRef = "@" & colLetter2
            End If
        End If
    Else
        ' Format nommé : Date (marquer pour résolution dynamique)
        fieldRef = "@" & UCase(colSpec) & "_NAMED"
    End If
    
    If Len(fieldRef) > 0 Then
        If Not SafeAddToDictionary(readFields, fieldRef, True) Then
            SetParsingError ERR_PARSING_MEMORY_EXCEEDED, "Echec ajout simple: " & fieldRef, ""
            Exit Function
        End If
        
        If Not SafeAddToDictionary(readOrder, fieldRef, orderIndex) Then
            SetParsingError ERR_PARSING_MEMORY_EXCEEDED, "Echec ordre simple: " & fieldRef, ""
            Exit Function
        End If
        
        orderIndex = orderIndex + 1
        ParseSingleColumnWithOrderSafe = True
    End If
    
    Exit Function
    
ParseSingleError:
    ParseSingleColumnWithOrderSafe = False
End Function

' ============================================================================
' AddNamedRangeForResolutionSafe - MARQUAGE RANGE NOMMÉE SÉCURISÉ
' ============================================================================
Function AddNamedRangeForResolutionSafe(startName As String, endName As String, readFields As Object, readOrder As Object, ByRef orderIndex As Long) As Boolean
    On Error GoTo AddNamedError
    
    AddNamedRangeForResolutionSafe = False
    
    If Len(startName) = 0 Or Len(endName) = 0 Then Exit Function
    If Len(startName) > 255 Or Len(endName) > 255 Then Exit Function
    
    ' Validation caractères nom (Excel naming rules basiques)
    If Not IsValidExcelName(startName) Or Not IsValidExcelName(endName) Then Exit Function
    
    ' Marquer pour résolution dynamique ultérieure
    Dim namedRange As String: namedRange = "@" & UCase(startName) & ":" & UCase(endName) & "_NAMED"
    
    If Not SafeAddToDictionary(readFields, namedRange, True) Then
        SetParsingError ERR_PARSING_MEMORY_EXCEEDED, "Echec ajout nommé: " & namedRange, ""
        Exit Function
    End If
    
    If Not SafeAddToDictionary(readOrder, namedRange, orderIndex) Then
        SetParsingError ERR_PARSING_MEMORY_EXCEEDED, "Echec ordre nommé: " & namedRange, ""
        Exit Function
    End If
    
    orderIndex = orderIndex + 1
    AddNamedRangeForResolutionSafe = True
    Exit Function
    
AddNamedError:
    AddNamedRangeForResolutionSafe = False
End Function

' ============================================================================
' UTILITAIRES SÉCURISÉS SUPPLÉMENTAIRES
' ============================================================================

Function IsExcelColumnFormatSafe(spec As String) As Boolean
    On Error GoTo ExcelFormatError
    
    IsExcelColumnFormatSafe = False
    
    If Len(spec) = 0 Or Len(spec) > 10 Then Exit Function
    
    Dim letters As String: letters = ExtractColumnLettersSafe(spec)
    IsExcelColumnFormatSafe = (Len(letters) > 0 And IsValidColumnRefSafe(letters))
    Exit Function
    
ExcelFormatError:
    IsExcelColumnFormatSafe = False
End Function

Function IsNumericRange(spec As String) As Boolean
    On Error GoTo NumericError
    
    IsNumericRange = False
    
    If Len(spec) = 0 Then Exit Function
    
    Dim numPart As String: numPart = ExtractNumbersSafe(spec)
    IsNumericRange = (Len(numPart) > 0 And IsNumeric(numPart) And CLng(numPart) > 0)
    Exit Function
    
NumericError:
    IsNumericRange = False
End Function

Function ExtractColumnLettersSafe(cellRef As String) As String
    On Error GoTo ExtractLettersError
    
    ExtractColumnLettersSafe = ""
    
    If Len(cellRef) = 0 Or Len(cellRef) > 10 Then Exit Function
    
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
    
    ExtractColumnLettersSafe = result
    Exit Function
    
ExtractLettersError:
    ExtractColumnLettersSafe = ""
End Function

Function ExtractRowNumberSafe(cellAddress As String) As Long
    On Error GoTo ExtractRowError
    
    ExtractRowNumberSafe = 1 ' Défaut ligne 1
    
    If Len(cellAddress) = 0 Then Exit Function
    
    Dim numPart As String: numPart = ExtractNumbersSafe(cellAddress)
    If Len(numPart) > 0 And IsNumeric(numPart) Then
        Dim rowNum As Long: rowNum = CLng(numPart)
        If rowNum > 0 And rowNum <= 1048576 Then ' Limite Excel
            ExtractRowNumberSafe = rowNum
        End If
    End If
    
    Exit Function
    
ExtractRowError:
    ExtractRowNumberSafe = 1
End Function

Function ExtractNumbersSafe(text As String) As String
    On Error GoTo ExtractNumError
    
    ExtractNumbersSafe = ""
    
    If Len(text) = 0 Or Len(text) > 50 Then Exit Function
    
    Dim result As String: result = ""
    Dim i As Long
    
    For i = 1 To Len(text)
        Dim char As String: char = Mid(text, i, 1)
        If char >= "0" And char <= "9" Then
            result = result & char
        End If
    Next i
    
    ExtractNumbersSafe = result
    Exit Function
    
ExtractNumError:
    ExtractNumbersSafe = ""
End Function

Function ExtractNumberSafe(text As String) As Long
    On Error GoTo ExtractSingleNumError
    
    ExtractNumberSafe = -1 ' Valeur erreur
    
    Dim numPart As String: numPart = ExtractNumbersSafe(text)
    If Len(numPart) > 0 And IsNumeric(numPart) Then
        Dim num As Long: num = CLng(numPart)
        If num > 0 And num <= PARSING_MAX_COLUMNS Then
            ExtractNumberSafe = num
        End If
    End If
    
    Exit Function
    
ExtractSingleNumError:
    ExtractNumberSafe = -1
End Function

Function IndexToColumnSafe(colIndex As Long) As String
    On Error GoTo IndexToColError
    
    IndexToColumnSafe = ""
    
    If colIndex < 1 Or colIndex > PARSING_MAX_COLUMNS Then Exit Function
    
    Dim result As String: result = ""
    Dim tempNum As Long: tempNum = colIndex
    
    Do While tempNum > 0
        tempNum = tempNum - 1
        result = Chr(Asc("A") + (tempNum Mod 26)) + result
        tempNum = tempNum \ 26
    Loop
    
    IndexToColumnSafe = result
    Exit Function
    
IndexToColError:
    IndexToColumnSafe = ""
End Function

Function IsValidExcelName(name As String) As Boolean
    On Error GoTo ValidateNameError
    
    IsValidExcelName = False
    
    If Len(name) = 0 Or Len(name) > 255 Then Exit Function
    
    ' Premier caractère doit être lettre ou underscore
    Dim firstChar As String: firstChar = Left(name, 1)
    If Not ((firstChar >= "A" And firstChar <= "Z") Or (firstChar >= "a" And firstChar <= "z") Or firstChar = "_") Then
        Exit Function
    End If
    
    ' Autres caractères: lettres, chiffres, underscore
    Dim i As Long
    For i = 2 To Len(name)
        Dim char As String: char = Mid(name, i, 1)
        If Not ((char >= "A" And char <= "Z") Or (char >= "a" And char <= "z") Or _
                (char >= "0" And char <= "9") Or char = "_") Then
            Exit Function
        End If
    Next i
    
    IsValidExcelName = True
    Exit Function
    
ValidateNameError:
    IsValidExcelName = False
End Function

Private Sub SwapLongsSafe(ByRef a As Long, ByRef b As Long)
    Dim temp As Long: temp = a
    a = b: b = temp
End Sub

Private Sub LogRegistryContentsSecure(registry As Object)
    On Error Resume Next ' Non critique
    
    If GetParsingConfig("DebugMode") Then
        Debug.Print "=== REGISTRY SÉCURISÉ ==="
        Debug.Print "WHAT_FIELDS: " & registry("WHAT_FIELDS").Count
        Debug.Print "READ_FIELDS: " & registry("READ_FIELDS").Count
        Debug.Print "ALL_REQUIRED: " & registry("ALL_REQUIRED").Count
        Debug.Print "READ_EQUALS_WHAT: " & registry("READ_EQUALS_WHAT")
        
        If GetParsingConfig("VerboseLogging") Then
            Dim key As Variant
            For Each key In registry("ALL_REQUIRED").Keys
                Debug.Print "  ? " & key
            Next key
        End If
    End If
    
    On Error GoTo 0
End Sub


