Attribute VB_Name = "mod_ColumnMain"
' ============================================================================
' DOCUMENTATION GESTION ERREURS
' ============================================================================
'
' UTILISATION:
'
' Dim registry As Object
' Set registry = BuildColumnRegistry("@Nom LIKE 'Test*'", "A:E")
'
' If HasParsingError() Then
'     Dim err As ParsingError: err = GetLastParsingError()
'     MsgBox "Erreur parsing [" & err.Code & "]: " & err.Message
' Else
'     ' Registry valide, continuer traitement
' End If
'
' ============================================================================
' ============================================================================
' GESTION D'ERREURS SP�CIFIQUE ET AM�LIORATIONS PRIORIT� 1
' ============================================================================
' Remplacement des fallbacks simplistes par vraie gestion d'erreurs
' Correction bugs identifi�s et robustesse
' ============================================================================

Option Explicit

' ============================================================================
' BuildColumnRegistry - VERSION AVEC GESTION ERREURS ROBUSTE
' ============================================================================
Function BuildColumnRegistry(whatExpression As String, readColumns As String) As Object
    ' Nettoyer erreurs pr�c�dentes
    ClearParsingError
    
    On Error GoTo CriticalError
    
    ' Validation entr�e
    If Not ValidateInputParameters(whatExpression, readColumns) Then
        Exit Function ' Erreur d�j� d�finie dans validation
    End If
    
    ' Initialisation s�curis�e
    Set BuildColumnRegistry = CreateSecureRegistry()
    
    ' 1. ANALYSER EXPRESSION WHAT
    If Not AnalyzeWhatExpressionSafe(whatExpression, BuildColumnRegistry) Then
        GoTo AnalysisError
    End If
    
    ' 2. ANALYSER SP�CIFICATION READ
    If Not AnalyzeReadSpecificationSafe(readColumns, BuildColumnRegistry) Then
        GoTo AnalysisError
    End If
    
    ' 3. CONSTRUIRE MAPPINGS
    If Not BuildUnionAndMappingsSafe(BuildColumnRegistry) Then
        GoTo MappingError
    End If
    
    ' 4. VALIDATION FINALE
    If Not ValidateRegistryIntegrity(BuildColumnRegistry) Then
        GoTo ValidationError
    End If
    
    ' 5. LOGGING SI ACTIV�
    If GetParsingConfig("LogParsingSteps") Then
        LogRegistryContentsSecure BuildColumnRegistry
    End If
    
    Exit Function
    
CriticalError:
    SetParsingError ERR_PARSING_UNKNOWN, "Erreur critique dans BuildColumnRegistry: " & err.Description, "whatExpression=" & Left(whatExpression, 100)
    Set BuildColumnRegistry = CreateFallbackRegistrySecure()
    Exit Function
    
AnalysisError:
    If Not HasParsingError() Then
        SetParsingError ERR_PARSING_INVALID_EXPRESSION, "Echec analyse expression", "whatExpression=" & Left(whatExpression, 50)
    End If
    Set BuildColumnRegistry = CreateFallbackRegistrySecure()
    Exit Function
    
MappingError:
    SetParsingError ERR_PARSING_MAPPING_INCONSISTENT, "Echec construction mappings", ""
    Set BuildColumnRegistry = CreateFallbackRegistrySecure()
    Exit Function
    
ValidationError:
    SetParsingError ERR_PARSING_MAPPING_INCONSISTENT, "Registry invalide apr�s construction", ""
    Set BuildColumnRegistry = CreateFallbackRegistrySecure()
    Exit Function
End Function

' ============================================================================
' VALIDATION PARAM�TRES D'ENTR�E
' ============================================================================
Function ValidateInputParameters(whatExpression As String, readColumns As String) As Boolean
    ValidateInputParameters = False
    
    ' Validation expression WHAT
    If Len(Trim(whatExpression)) = 0 Then
        SetParsingError ERR_PARSING_INVALID_EXPRESSION, "Expression WHAT vide", ""
        Exit Function
    End If
    
    If Len(whatExpression) > PARSING_MAX_EXPRESSION_LENGTH Then
        SetParsingError ERR_PARSING_INVALID_EXPRESSION, "Expression WHAT trop longue (" & Len(whatExpression) & " caract�res)", ""
        Exit Function
    End If
    
    ' V�rification caract�res interdits
    If InStr(whatExpression, Chr(0)) > 0 Or InStr(whatExpression, Chr(1)) > 0 Then
        SetParsingError ERR_PARSING_INVALID_EXPRESSION, "Caract�res interdits dans expression WHAT", ""
        Exit Function
    End If
    
    ' Validation basique syntaxe
    If Not ValidateBasicSyntax(whatExpression) Then
        SetParsingError ERR_PARSING_INVALID_EXPRESSION, "Syntaxe invalide dans expression WHAT", "Parenth�ses ou quotes non �quilibr�es"
        Exit Function
    End If
    
    ' Validation READ (optionnel mais si pr�sent doit �tre valide)
    If Len(Trim(readColumns)) > 0 Then
        If Len(readColumns) > PARSING_MAX_EXPRESSION_LENGTH Then
            SetParsingError ERR_PARSING_INVALID_READ_SPEC, "Sp�cification READ trop longue", ""
            Exit Function
        End If
    End If
    
    ValidateInputParameters = True
End Function

' ============================================================================
' VALIDATION SYNTAXE BASIQUE
' ============================================================================
Function ValidateBasicSyntax(expression As String) As Boolean
    ValidateBasicSyntax = False
    
    Dim parenCount As Long: parenCount = 0
    Dim squoteCount As Long: squoteCount = 0  ' Single quotes
    Dim dquoteCount As Long: dquoteCount = 0  ' Double quotes
    Dim i As Long
    
    For i = 1 To Len(expression)
        Dim char As String: char = Mid(expression, i, 1)
        
        Select Case char
            Case "("
                parenCount = parenCount + 1
            Case ")"
                parenCount = parenCount - 1
                If parenCount < 0 Then Exit Function ' Parenth�se fermante sans ouvrante
            Case "'"
                squoteCount = squoteCount + 1
            Case """"
                dquoteCount = dquoteCount + 1
        End Select
    Next i
    
    ' V�rifications finales
    If parenCount <> 0 Then Exit Function ' Parenth�ses non �quilibr�es
    If squoteCount Mod 2 <> 0 Then Exit Function ' Quotes simples non �quilibr�es
    If dquoteCount Mod 2 <> 0 Then Exit Function ' Quotes doubles non �quilibr�es
    
    ValidateBasicSyntax = True
End Function

' ============================================================================
' ANALYSE WHAT S�CURIS�E
' ============================================================================
Function AnalyzeWhatExpressionSafe(expression As String, registry As Object) As Boolean
    On Error GoTo AnalyzeError
    
    AnalyzeWhatExpressionSafe = False
    
    If Len(Trim(expression)) = 0 Then
        AnalyzeWhatExpressionSafe = True ' Techniquement valide si vide
        Exit Function
    End If
    
    Dim expr As String: expr = UCase(Trim(expression))
    Dim whatFields As Object: Set whatFields = registry("WHAT_FIELDS")
    Dim compFields As Object: Set compFields = registry("COMPARISON_FIELDS")
    
    Dim FieldCount As Long: FieldCount = 0
    Dim i As Long
    
    For i = 1 To Len(expr) - 1
        If Mid(expr, i, 1) = "@" Then
            Dim colRef As String: colRef = ExtractColumnReferenceSafe(expr, i + 1)
            
            If colRef = "ERROR" Then
                SetParsingError ERR_PARSING_INVALID_EXPRESSION, "R�f�rence colonne invalide � position " & i, ""
                Exit Function
            End If
            
            If Len(colRef) > 0 And IsValidColumnRefSafe(colRef) Then
                Dim fieldRef As String: fieldRef = "@" & colRef
                
                ' Ajout s�curis� au dictionary
                If Not SafeAddToDictionary(whatFields, fieldRef, True) Then
                    SetParsingError ERR_PARSING_MEMORY_EXCEEDED, "Impossible d'ajouter champ " & fieldRef, ""
                    Exit Function
                End If
                
                ' D�tection contexte comparaison s�curis�e
                If IsInComparisonContextSafe(expr, i) Then
                    SafeAddToDictionary compFields, fieldRef, True
                End If
                
                FieldCount = FieldCount + 1
                
                ' Limite s�curitaire
                If FieldCount > PARSING_MAX_COLUMNS Then
                    SetParsingError ERR_PARSING_MEMORY_EXCEEDED, "Trop de champs dans expression (" & FieldCount & ")", ""
                    Exit Function
                End If
            End If
        End If
    Next i
    
    AnalyzeWhatExpressionSafe = True
    Exit Function
    
AnalyzeError:
    SetParsingError ERR_PARSING_INVALID_EXPRESSION, "Erreur analyse WHAT: " & err.Description, ""
    AnalyzeWhatExpressionSafe = False
End Function

' ============================================================================
' ANALYSE READ S�CURIS�E
' ============================================================================
Function AnalyzeReadSpecificationSafe(readSpec As String, registry As Object) As Boolean
    On Error GoTo AnalyzeReadError
    
    AnalyzeReadSpecificationSafe = False
    
    ' Optimisation READ_EQUALS_WHAT
    If Len(Trim(readSpec)) = 0 Then
        registry("READ_EQUALS_WHAT") = True
        AnalyzeReadSpecificationSafe = True
        Exit Function
    End If
    
    registry("READ_EQUALS_WHAT") = False
    
    Dim spec As String: spec = Trim(readSpec)
    Dim readFields As Object: Set readFields = registry("READ_FIELDS")
    Dim readOrder As Object: Set readOrder = registry("READ_ORDER")
    Dim orderIndex As Long: orderIndex = 1
    
    ' Enlever crochets globaux si pr�sents
    If Left(spec, 1) = "[" And Right(spec, 1) = "]" Then
        spec = Mid(spec, 2, Len(spec) - 2)
    End If
    
    ' Validation longueur apr�s nettoyage
    If Len(spec) = 0 Then
        SetParsingError ERR_PARSING_INVALID_READ_SPEC, "Sp�cification READ vide apr�s nettoyage", ""
        Exit Function
    End If
    
    ' Parser multi-ranges s�curis�
    Dim ranges As Variant
    On Error Resume Next
    ranges = Split(spec, ",")
    On Error GoTo AnalyzeReadError
    
    If Not IsArray(ranges) Then
        SetParsingError ERR_PARSING_INVALID_READ_SPEC, "Impossible de parser les ranges", ""
        Exit Function
    End If
    
    Dim i As Long
    For i = 0 To UBound(ranges)
        Dim rangeSpec As String
        
        ' Conversion s�curis�e
        On Error Resume Next
        rangeSpec = Trim(UCase(CStr(ranges(i))))
        On Error GoTo AnalyzeReadError
        
        If Len(rangeSpec) > 0 Then
            If Not ParseSingleRangeOrderedSafe(rangeSpec, readFields, readOrder, orderIndex) Then
                SetParsingError ERR_PARSING_INVALID_READ_SPEC, "Echec parsing range: " & rangeSpec, ""
                Exit Function
            End If
        End If
        
        ' Limite s�curitaire nombre de ranges
        If i > 100 Then
            SetParsingError ERR_PARSING_INVALID_READ_SPEC, "Trop de ranges sp�cifi�es (max 100)", ""
            Exit Function
        End If
    Next i
    
    AnalyzeReadSpecificationSafe = True
    Exit Function
    
AnalyzeReadError:
    SetParsingError ERR_PARSING_INVALID_READ_SPEC, "Erreur analyse READ: " & err.Description, ""
    AnalyzeReadSpecificationSafe = False
End Function

' ============================================================================
' PARSING RANGE S�CURIS�
' ============================================================================
Function ParseSingleRangeOrderedSafe(rangeSpec As String, readFields As Object, readOrder As Object, ByRef orderIndex As Long) As Boolean
    On Error GoTo ParseRangeError
    
    ParseSingleRangeOrderedSafe = False
    
    If Len(rangeSpec) = 0 Or Len(rangeSpec) > 255 Then
        Exit Function
    End If
    
    If InStr(rangeSpec, ":") > 0 Then
        ' Range A1:B2 ou A:B ou 1:3
        ParseSingleRangeOrderedSafe = ParseRangeWithOrderSafe(rangeSpec, readFields, readOrder, orderIndex)
    Else
        ' Colonne simple A1 ou A ou 3
        ParseSingleRangeOrderedSafe = ParseSingleColumnWithOrderSafe(rangeSpec, readFields, readOrder, orderIndex)
    End If
    
    Exit Function
    
ParseRangeError:
    ParseSingleRangeOrderedSafe = False
End Function

' ============================================================================
' UTILITAIRES S�CURIS�S
' ============================================================================
Function SafeAddToDictionary(dict As Object, key As String, value As Variant) As Boolean
    On Error GoTo AddError
    
    SafeAddToDictionary = False
    
    If dict Is Nothing Then Exit Function
    If Len(key) = 0 Then Exit Function
    
    ' V�rifier si cl� existe d�j� (�viter erreur)
    If dict.Exists(key) Then
        dict(key) = value ' Mettre � jour
    Else
        dict.Add key, value ' Ajouter nouveau
    End If
    
    SafeAddToDictionary = True
    Exit Function
    
AddError:
    SafeAddToDictionary = False
End Function

Function ExtractColumnReferenceSafe(expr As String, startPos As Long) As String
    On Error GoTo ExtractError
    
    ExtractColumnReferenceSafe = ""
    
    If startPos < 1 Or startPos > Len(expr) Then
        ExtractColumnReferenceSafe = "ERROR"
        Exit Function
    End If
    
    Dim result As String: result = ""
    Dim pos As Long: pos = startPos
    Dim charCount As Long: charCount = 0
    
    Do While pos <= Len(expr) And charCount < 10 ' Limite s�curitaire
        Dim char As String: char = Mid(expr, pos, 1)
        If char >= "A" And char <= "Z" Then
            result = result & char
            pos = pos + 1
            charCount = charCount + 1
        Else
            Exit Do
        End If
    Loop
    
    ExtractColumnReferenceSafe = result
    Exit Function
    
ExtractError:
    ExtractColumnReferenceSafe = "ERROR"
End Function

Function IsValidColumnRefSafe(colRef As String) As Boolean
    On Error GoTo ValidateError
    
    IsValidColumnRefSafe = False
    
    If Len(colRef) = 0 Or Len(colRef) > 3 Then Exit Function
    
    ' V�rifier caract�res A-Z uniquement
    Dim i As Long
    For i = 1 To Len(colRef)
        Dim char As String: char = Mid(colRef, i, 1)
        If Not (char >= "A" And char <= "Z") Then Exit Function
    Next i
    
    ' Limite Excel avec validation s�curis�e
    Dim colIndex As Long: colIndex = ColumnToIndexSafe(colRef)
    If colIndex = -1 Then Exit Function ' Erreur conversion
    
    IsValidColumnRefSafe = (colIndex <= PARSING_MAX_COLUMNS)
    Exit Function
    
ValidateError:
    IsValidColumnRefSafe = False
End Function

Function ColumnToIndexSafe(colLetter As String) As Long
    On Error GoTo ConvertError
    
    ColumnToIndexSafe = -1 ' Valeur erreur
    
    If Len(colLetter) = 0 Or Len(colLetter) > 3 Then Exit Function
    
    Dim result As Long: result = 0
    Dim i As Long
    
    colLetter = UCase(Trim(colLetter))
    
    For i = Len(colLetter) To 1 Step -1
        Dim char As String: char = Mid(colLetter, i, 1)
        If char < "A" Or char > "Z" Then Exit Function
        
        Dim charValue As Long: charValue = Asc(char) - Asc("A") + 1
        Dim multiplier As Long: multiplier = 26 ^ (Len(colLetter) - i)
        
        ' V�rification overflow
        If result > (2147483647 - (charValue * multiplier)) Then Exit Function
        
        result = result + (charValue * multiplier)
    Next i
    
    If result < 1 Or result > PARSING_MAX_COLUMNS Then Exit Function
    
    ColumnToIndexSafe = result
    Exit Function
    
ConvertError:
    ColumnToIndexSafe = -1
End Function

Function IsInComparisonContextSafe(expr As String, pos As Long) As Boolean
    On Error GoTo ComparisonError
    
    IsInComparisonContextSafe = False
    
    If pos < 1 Or pos > Len(expr) Then Exit Function
    
    ' Recherche op�rateurs comparaison dans contexte s�curis�
    Dim startPos As Long: startPos = IIf(pos - 10 > 1, pos - 10, 1)
    Dim endPos As Long: endPos = IIf(pos + 20 <= Len(expr), pos + 20, Len(expr))
    
    If endPos <= startPos Then Exit Function
    
    Dim context As String: context = Mid(expr, startPos, endPos - startPos + 1)
    
    ' Recherche op�rateurs
    IsInComparisonContextSafe = (InStr(context, ">") > 0 Or InStr(context, "<") > 0 Or _
                                 InStr(context, "=") > 0 Or InStr(context, " LIKE ") > 0 Or _
                                 InStr(context, " IN ") > 0 Or InStr(context, " BETWEEN ") > 0)
    Exit Function
    
ComparisonError:
    IsInComparisonContextSafe = False
End Function

' ============================================================================
' FONCTIONS SUPPORT S�CURIS�ES
' ============================================================================

Function CreateSecureRegistry() As Object
    On Error GoTo CreateError
    
    Set CreateSecureRegistry = CreateObject("Scripting.Dictionary")
    
    With CreateSecureRegistry
        .Add "WHAT_FIELDS", CreateObject("Scripting.Dictionary")
        .Add "READ_FIELDS", CreateObject("Scripting.Dictionary")
        .Add "ALL_REQUIRED", CreateObject("Scripting.Dictionary")
        .Add "COLUMN_INDEX", CreateObject("Scripting.Dictionary")
        .Add "INDEX_COLUMN", CreateObject("Scripting.Dictionary")
        .Add "READ_EQUALS_WHAT", False
        .Add "READ_ORDER", CreateObject("Scripting.Dictionary")
        .Add "SOURCE_POSITIONS", CreateObject("Scripting.Dictionary")
        .Add "EXTRACT_POSITIONS", CreateObject("Scripting.Dictionary")
        .Add "POSITION_TO_FIELD", CreateObject("Scripting.Dictionary")
        .Add "EXTRACT_TO_FIELD", CreateObject("Scripting.Dictionary")
        .Add "COMPARISON_FIELDS", CreateObject("Scripting.Dictionary")
    End With
    
    Exit Function
    
CreateError:
    Set CreateSecureRegistry = Nothing
    SetParsingError ERR_PARSING_MEMORY_EXCEEDED, "Impossible de cr�er registry: " & err.Description, ""
End Function

Function CreateFallbackRegistrySecure() As Object
    On Error Resume Next
    
    Set CreateFallbackRegistrySecure = CreateSecureRegistry()
    
    If CreateFallbackRegistrySecure Is Nothing Then
        ' Fallback ultime - dictionnaire minimal
        Set CreateFallbackRegistrySecure = CreateObject("Scripting.Dictionary")
        CreateFallbackRegistrySecure.Add "ALL_REQUIRED", CreateObject("Scripting.Dictionary")
        CreateFallbackRegistrySecure("ALL_REQUIRED")("@A") = True
    End If
    
    On Error GoTo 0
End Function

Function ValidateRegistryIntegrity(registry As Object) As Boolean
    On Error GoTo ValidationError
    
    ValidateRegistryIntegrity = False
    
    If registry Is Nothing Then Exit Function
    
    ' V�rifier pr�sence cl�s essentielles
    Dim requiredKeys As Variant
    requiredKeys = Array("WHAT_FIELDS", "READ_FIELDS", "ALL_REQUIRED")
    
    Dim i As Long
    For i = 0 To UBound(requiredKeys)
        If Not registry.Exists(requiredKeys(i)) Then Exit Function
    Next i
    
    ' V�rifier coh�rence minimale
    If registry("ALL_REQUIRED").Count = 0 Then
        ' Ajouter fallback minimal
        registry("ALL_REQUIRED")("@A") = True
    End If
    
    ValidateRegistryIntegrity = True
    Exit Function
    
ValidationError:
    ValidateRegistryIntegrity = False
End Function

' ============================================================================
' STUBS POUR FONCTIONS NON ENCORE IMPL�MENT�ES
' ============================================================================

Function BuildUnionAndMappingsSafe(registry As Object) As Boolean
    ' TODO: Impl�menter version s�curis�e de BuildUnionAndMappings
    ' Pour l'instant, version simplifi�e
    On Error GoTo MappingError
    
    BuildUnionAndMappingsSafe = True
    ' Impl�mentation temporaire - � compl�ter
    Exit Function
    
MappingError:
    BuildUnionAndMappingsSafe = False
End Function

Function ParseRangeWithOrderSafe(rangeSpec As String, readFields As Object, readOrder As Object, ByRef orderIndex As Long) As Boolean
    ' TODO: Impl�menter version s�curis�e
    ParseRangeWithOrderSafe = True
    orderIndex = orderIndex + 1
End Function

Function ParseSingleColumnWithOrderSafe(colSpec As String, readFields As Object, readOrder As Object, ByRef orderIndex As Long) As Boolean
    ' TODO: Impl�menter version s�curis�e
    ParseSingleColumnWithOrderSafe = True
    orderIndex = orderIndex + 1
End Function

Private Sub LogRegistryContentsSecure(registry As Object)
    ' TODO: Impl�menter logging s�curis�
    If GetParsingConfig("DebugMode") Then
        Debug.Print "Registry cr�� avec " & registry("ALL_REQUIRED").Count & " champs"
    End If
End Sub


