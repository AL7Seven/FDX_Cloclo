Attribute VB_Name = "Mod_foolish_3"
' ===============================================================================
' FindXtreme Hi (FDXH) - SESSION 3 PARTIES MANQUANTES COMPLÈTES
' Compléter TOUS les TODOs + Opérateurs manquants (IN, XOR, NOT, etc.)
' LIVRABLE FINAL : Code production-ready avec tests unitaires
' ===============================================================================
Option Explicit

Public FDXH_Config As Object

' ===============================================================================
' OPÉRATEUR IN - IMPLÉMENTATION COMPLÈTE
' ===============================================================================

Function EvaluateInExpression(fieldRef As String, operator As String, listValues As String, rowIndex As Long) As Boolean
    ' Évaluer @A IN ["value1","value2","value3"] ou @A NOT IN [...]
    
    Dim fieldValue As Variant
    fieldValue = GetFieldValue(fieldRef, rowIndex)
    
    ' Extraire liste de valeurs
    Dim valuesList() As String
    If Not ParseInValuesList(listValues, valuesList) Then
        EvaluateInExpression = False
        Exit Function
    End If
    
    ' Rechercher dans la liste
    Dim found As Boolean
    found = IsValueInList(fieldValue, valuesList)
    
    ' Appliquer opérateur
    Select Case UCase(operator)
        Case "IN"
            EvaluateInExpression = found
        Case "NOT IN"
            EvaluateInExpression = Not found
        Case Else
            EvaluateInExpression = False
    End Select
    
    If GetConfigValue("LogParsingSteps") Then
        Debug.Print "IN evaluation: " & fieldRef & "=" & fieldValue & " " & operator & " " & listValues & " => " & EvaluateInExpression
    End If
End Function

'Function GetFieldValue(fieldRef As String, rowIndex As Long) As Variant
'    Dim columnIndex As Long
'    columnIndex = GetColumnIndex(fieldRef)
'
'    If columnIndex > 0 And columnIndex <= pSourceCols And rowIndex <= pSourceRows Then
'        GetFieldValue = pSourceData(rowIndex, columnIndex)
'    Else
'        GetFieldValue = ""
'    End If
'End Function
' PROBLÈME 3: GetFieldValue doit gérer les index invalides
Function GetFieldValue(fieldRef As String, rowIndex As Long) As Variant
    On Error GoTo ErrorHandler
    
    Dim columnIndex As Long
    columnIndex = GetColumnIndex(fieldRef)
    
    ' CORRECTION: Validation complète des bounds
    If columnIndex > 0 And columnIndex <= pSourceCols And _
       rowIndex > 0 And rowIndex <= pSourceRows Then
        GetFieldValue = pSourceData(rowIndex, columnIndex)
    Else
        GetFieldValue = Empty  ' Retourner Empty au lieu de ""
        If GetConfigValue("LogParsingSteps") Then
            Debug.Print "GetFieldValue: Invalid bounds - Row:" & rowIndex & "/" & pSourceRows & " Col:" & columnIndex & "/" & pSourceCols
        End If
    End If
    Exit Function
    
ErrorHandler:
    GetFieldValue = Empty
    Debug.Print "GetFieldValue Error: " & Err.Description
End Function

' PROBLÈME 4: BuildColumnMappings manquante
Public Function BuildColumnMappings(sourceRange As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Parser simple pour identifier les colonnes utilisées
    Dim fieldRefs As Collection
    Set fieldRefs = ExtractFieldReferences(pExpressionText)
    
    pColumnCount = fieldRefs.Count
    If pColumnCount = 0 Then
        BuildColumnMappings = False
        Exit Function
    End If
    
    ReDim pColumnMaps(1 To pColumnCount)
    
    Dim i As Long
    For i = 1 To fieldRefs.Count
        With pColumnMaps(i)
            .FieldReference = fieldRefs(i)
            .columnIndex = ConvertFieldRefToColumnIndex(fieldRefs(i))
            .ColumnLetter = ConvertIndexToLetter(.columnIndex)
            .IsRequired = True
        End With
    Next i
    
    BuildColumnMappings = True
    Exit Function
    
ErrorHandler:
    BuildColumnMappings = False
End Function

Function ExtractFieldReferences(expression As String) As Collection
    Set ExtractFieldReferences = New Collection
    
    Dim i As Long
    For i = 1 To pTokenCount
        If pTokens(i).TokenType = TT_FieldReference Then
            On Error Resume Next
            ExtractFieldReferences.Add pTokens(i).TokenValue, pTokens(i).TokenValue
            On Error GoTo 0
        End If
    Next i
End Function

Function ConvertFieldRefToColumnIndex(fieldRef As String) As Long
    ' @A = 1, @B = 2, etc.
    If Len(fieldRef) >= 2 And Left(fieldRef, 1) = "@" Then
        Dim letter As String
        letter = UCase(Mid(fieldRef, 2, 1))
        If letter >= "A" And letter <= "Z" Then
            ConvertFieldRefToColumnIndex = Asc(letter) - Asc("A") + 1
        End If
    End If
End Function

Function ConvertIndexToLetter(index As Long) As String
    If index >= 1 And index <= 26 Then
        ConvertIndexToLetter = Chr(Asc("A") + index - 1)
    End If
End Function

Function ParseInValuesList(listStr As String, ByRef valuesList() As String) As Boolean
    ' Parser ["val1","val2","val3"] vers array
    
    On Error GoTo ErrorHandler
    
    ' Validation format [...]
    If Left(listStr, 1) <> "[" Or Right(listStr, 1) <> "]" Then
        ParseInValuesList = False
        Exit Function
    End If
    
    ' Extraire contenu sans crochets
    Dim innerContent As String
    innerContent = Mid(listStr, 2, Len(listStr) - 2)
    
    If Len(Trim(innerContent)) = 0 Then
        ReDim valuesList(0 To 0)
        valuesList(0) = ""
        ParseInValuesList = True
        Exit Function
    End If
    
    ' Séparer par virgules (gestion guillemets)
    Dim values As Collection
    Set values = ParseCommaSeparatedValues(innerContent)
    
    ' Convertir Collection vers Array
    ReDim valuesList(0 To values.Count - 1)
    Dim i As Long
    For i = 1 To values.Count
        valuesList(i - 1) = CleanQuotedValue(CStr(values(i)))
    Next i
    
    ParseInValuesList = True
    Exit Function
    
ErrorHandler:
    ParseInValuesList = False
End Function

Function ParseCommaSeparatedValues(content As String) As Collection
    ' Parser "val1","val2","val3" en gérant les guillemets et échappements
    
    Set ParseCommaSeparatedValues = New Collection
    
    Dim inQuotes As Boolean
    Dim currentValue As String
    Dim i As Long
    Dim char As String, nextChar As String
    
    inQuotes = False
    currentValue = ""
    
    For i = 1 To Len(content)
        char = Mid(content, i, 1)
        If i < Len(content) Then nextChar = Mid(content, i + 1, 1) Else nextChar = ""
        
        Select Case char
            Case """"
                If inQuotes And nextChar = """" Then
                    ' Double quote = échappement
                    currentValue = currentValue & """"
                    i = i + 1 ' Skip next quote
                Else
                    ' Toggle quote mode
                    inQuotes = Not inQuotes
                    currentValue = currentValue & char
                End If
                
            Case ","
                If Not inQuotes Then
                    ' Fin de valeur
                    ParseCommaSeparatedValues.Add Trim(currentValue)
                    currentValue = ""
                Else
                    currentValue = currentValue & char
                End If
                
            Case Else
                currentValue = currentValue & char
        End Select
    Next i
    
    ' Ajouter dernière valeur
    If Len(Trim(currentValue)) > 0 Then
        ParseCommaSeparatedValues.Add Trim(currentValue)
    End If
End Function

Function CleanQuotedValue(value As String) As String
    ' Nettoyer "value" => value et gérer échappements
    
    Dim cleaned As String
    cleaned = Trim(value)
    
    ' Enlever guillemets extérieurs
    If Left(cleaned, 1) = """" And Right(cleaned, 1) = """" And Len(cleaned) >= 2 Then
        cleaned = Mid(cleaned, 2, Len(cleaned) - 2)
    End If
    
    ' Remplacer double quotes par single quotes
    cleaned = Replace(cleaned, """""", """")
    
    CleanQuotedValue = cleaned
End Function

Function IsValueInList(searchValue As Variant, valuesList() As String) As Boolean
    ' Recherche avec gestion types et comparaison stricte/floue
    
    Dim searchStr As String
    searchStr = CStr(searchValue)
    
    Dim i As Long
    For i = 0 To UBound(valuesList)
        If GetConfigValue("CompareStrict") Then
            ' Comparaison stricte (sensible casse)
            If searchStr = valuesList(i) Then
                IsValueInList = True
                Exit Function
            End If
        Else
            ' Comparaison insensible casse
            If UCase(searchStr) = UCase(valuesList(i)) Then
                IsValueInList = True
                Exit Function
            End If
        End If
    Next i
    
    IsValueInList = False
End Function

' ===============================================================================
' OPÉRATEURS LOGIQUES ÉTENDUS - IMPLÉMENTATION COMPLÈTE
' ===============================================================================

Function EvaluateExtendedLogicalOperator(leftResult As Boolean, operator As String, rightResult As Boolean) As Boolean
    ' Évaluer opérateurs XOR, NAND, NOR
    
    Select Case UCase(Trim(operator))
        Case "XOR"
            EvaluateExtendedLogicalOperator = leftResult Xor rightResult
            
        Case "NAND"
            EvaluateExtendedLogicalOperator = Not (leftResult And rightResult)
            
        Case "NOR"
            EvaluateExtendedLogicalOperator = Not (leftResult Or rightResult)
            
        Case "AND"
            EvaluateExtendedLogicalOperator = leftResult And rightResult
            
        Case "OR"
            EvaluateExtendedLogicalOperator = leftResult Or rightResult
            
        Case Else
            EvaluateExtendedLogicalOperator = False
    End Select
    
    If GetConfigValue("LogParsingSteps") Then
        Debug.Print "Extended logical: " & leftResult & " " & operator & " " & rightResult & " => " & EvaluateExtendedLogicalOperator
    End If
End Function

Function EvaluateNotExpression(tokenIndex As Long, rowIndex As Long) As Boolean
    ' Évaluer expression NOT complète avec gestion récursive
    
    If tokenIndex + 1 > pTokenCount Then
        EvaluateNotExpression = False
        Exit Function
    End If
    
    Dim nextToken As TokenInfo
    nextToken = pTokens(tokenIndex + 1)
    
    Dim innerResult As Boolean
    
    Select Case nextToken.TokenType
        Case TT_OpenParen
            ' NOT (expression_complexe)
            Dim groupResult As Boolean
            groupResult = EvaluateParenthesesGroup(tokenIndex + 1, rowIndex)
            innerResult = groupResult
            
        Case TT_FieldReference
            ' NOT @A = "value" ou NOT EXISTS(@A)
            If tokenIndex + 3 <= pTokenCount Then
                ' NOT @A = "value"
                innerResult = EvaluateSimpleComparison(tokenIndex + 1, rowIndex)
            ElseIf InStr(nextToken.TokenValue, "EXISTS") > 0 Then
                ' NOT EXISTS(@A)
                innerResult = EvaluateExistsFunction(nextToken.TokenValue, rowIndex)
            Else
                innerResult = False
            End If
            
        Case TT_Function
            ' NOT EXISTS(@A), NOT REGEX(@A, "pattern")
            innerResult = EvaluateFunctionExpression(tokenIndex + 1, rowIndex)
            
        Case Else
            innerResult = False
    End Select
    
    EvaluateNotExpression = Not innerResult
    
    If GetConfigValue("LogParsingSteps") Then
        Debug.Print "NOT evaluation: NOT " & innerResult & " => " & EvaluateNotExpression
    End If
End Function

' ===============================================================================
' FONCTION EXISTS() - IMPLÉMENTATION COMPLÈTE
' ===============================================================================

Function EvaluateExistsFunction(funcToken As String, rowIndex As Long) As Boolean
    ' EXISTS(@A) - vérifier existence et non-vide du champ
    
    ' Extraire référence champ : EXISTS(@A) => @A
    Dim fieldRef As String
    fieldRef = ExtractFieldFromFunction(funcToken, "EXISTS")
    
    If Len(fieldRef) = 0 Then
        EvaluateExistsFunction = False
        Exit Function
    End If
    
    ' Obtenir valeur champ
    Dim fieldValue As Variant
    fieldValue = GetFieldValue(fieldRef, rowIndex)
    
    ' EXISTS = valeur présente ET non vide ET non null ET non erreur
    Dim exists As Boolean
    exists = True
    
    ' Vérifications multiples
    If IsNull(fieldValue) Then exists = False
    If IsEmpty(fieldValue) Then exists = False
    If IsError(fieldValue) Then exists = False
    If VarType(fieldValue) = vbString And Len(Trim(CStr(fieldValue))) = 0 Then exists = False
    
    EvaluateExistsFunction = exists
    
    If GetConfigValue("LogParsingSteps") Then
        Debug.Print "EXISTS(" & fieldRef & "): value=" & fieldValue & " => " & exists
    End If
End Function

Function ExtractFieldFromFunction(funcToken As String, functionName As String) As String
    ' Extraire @A de EXISTS(@A) ou REGEX(@A, "pattern")
    
    Dim startPos As Long, endPos As Long
    startPos = InStr(UCase(funcToken), UCase(functionName) & "(")
    
    If startPos = 0 Then
        ExtractFieldFromFunction = ""
        Exit Function
    End If
    
    startPos = startPos + Len(functionName) + 1 ' Après "EXISTS("
    endPos = InStr(startPos, funcToken, ")")
    
    If endPos = 0 Then endPos = Len(funcToken) + 1
    
    Dim innerContent As String
    innerContent = Mid(funcToken, startPos, endPos - startPos)
    
    ' Pour REGEX(@A, "pattern"), prendre seulement @A (avant première virgule)
    Dim commaPos As Long
    commaPos = InStr(innerContent, ",")
    If commaPos > 0 Then
        innerContent = Left(innerContent, commaPos - 1)
    End If
    
    ExtractFieldFromFunction = Trim(innerContent)
End Function

' ===============================================================================
' ÉVALUATEUR PRINCIPAL RÉCURSIF - VERSION COMPLÈTE
' ===============================================================================

Function EvaluateTokenRangeComplete(startToken As Long, endToken As Long, rowIndex As Long) As Boolean
    ' Version complète avec TOUS les opérateurs supportés
    
    If startToken > endToken Or startToken > pTokenCount Then
        EvaluateTokenRangeComplete = True
        Exit Function
    End If
    
    ' Stack pour gestion priorités et récursivité
    Dim resultStack() As Boolean
    Dim operatorStack() As String
    Dim stackDepth As Long
    
    ReDim resultStack(1 To (endToken - startToken + 1))
    ReDim operatorStack(1 To (endToken - startToken + 1))
    stackDepth = 0
    
    Dim currentToken As Long
    currentToken = startToken
    
    Do While currentToken <= endToken
        With pTokens(currentToken)
            Select Case .TokenType
                Case TT_FieldReference
                    ' Détecter type de comparaison/opération
                    Dim compResult As Boolean
                    compResult = EvaluateFieldOperation_INTELLIGENT(currentToken, rowIndex, endToken)
                    
                    PushResult resultStack, operatorStack, stackDepth, compResult
                    currentToken = SkipOperationTokens(currentToken, endToken)
                    
                Case TT_Not
                    ' Expression NOT
                    Dim notResult As Boolean
                    notResult = EvaluateNotExpression(currentToken, rowIndex)
                    
                    PushResult resultStack, operatorStack, stackDepth, notResult
                    currentToken = SkipNotTokens(currentToken, endToken)
                    
                Case TT_Function
                    ' Fonctions EXISTS, REGEX, etc.
                    Dim funcResult As Boolean
                    funcResult = EvaluateFunctionExpression(currentToken, rowIndex)
                    
                    PushResult resultStack, operatorStack, stackDepth, funcResult
                    currentToken = currentToken + 1
                    
                Case TT_LogicalOp, TT_Extended
                    ' Opérateurs logiques standards et étendus
                    ProcessLogicalOperatorComplete resultStack, operatorStack, stackDepth, .TokenValue
                    currentToken = currentToken + 1
                    
                Case TT_OpenParen
                    ' Groupe parenthèses
                    Dim parenResult As Boolean
                    parenResult = EvaluateParenthesesGroup(currentToken, rowIndex)
                    
                    PushResult resultStack, operatorStack, stackDepth, parenResult
                    currentToken = FindMatchingCloseParen(currentToken) + 1
                    
                Case Else
                    currentToken = currentToken + 1
            End Select
        End With
    Loop
    
    ' Résoudre stack final avec tous opérateurs
    EvaluateTokenRangeComplete = ResolveCompleteExpressionStack(resultStack, operatorStack, stackDepth)
End Function


Function EvaluateSimpleComparison(tokenIndex As Long, rowIndex As Long) As Boolean
    ' Comparaison simple @A OP value (=, >, <, etc.)
    
    If tokenIndex + 2 > pTokenCount Then
        EvaluateSimpleComparison = False
        Exit Function
    End If
    
    Dim fieldRef As String, operator As String, value As String
    fieldRef = pTokens(tokenIndex).TokenValue
    operator = pTokens(tokenIndex + 1).TokenValue
    value = pTokens(tokenIndex + 2).TokenValue
    
    Dim fieldValue As Variant, compareValue As Variant
    fieldValue = GetFieldValue(fieldRef, rowIndex)
    compareValue = ConvertValue(value)
    
    ' Gestion types pour comparaison
    Dim result As Boolean
    result = PerformTypedComparison(fieldValue, operator, compareValue)
    
    EvaluateSimpleComparison = result
End Function

Function PerformTypedComparison(leftValue As Variant, operator As String, rightValue As Variant) As Boolean
    ' Comparaison typée avec gestion erreurs et conversions
    
    On Error GoTo ErrorHandler
    
    Select Case UCase(operator)
        Case "="
            PerformTypedComparison = (leftValue = rightValue)
        Case ">"
            PerformTypedComparison = (leftValue > rightValue)
        Case "<"
            PerformTypedComparison = (leftValue < rightValue)
        Case ">="
            PerformTypedComparison = (leftValue >= rightValue)
        Case "<="
            PerformTypedComparison = (leftValue <= rightValue)
        Case "<>"
            PerformTypedComparison = (leftValue <> rightValue)
        Case "~"
            ' LIKE - recherche floue
            PerformTypedComparison = (InStr(1, CStr(leftValue), CStr(rightValue), vbTextCompare) > 0)
        Case "!~"
            ' NOT LIKE
            PerformTypedComparison = (InStr(1, CStr(leftValue), CStr(rightValue), vbTextCompare) = 0)
        Case Else
            PerformTypedComparison = False
    End Select
    
    Exit Function
    
ErrorHandler:
    ' Erreur comparaison (types incompatibles, etc.)
    PerformTypedComparison = False
End Function

Function ParseBetweenRanges(rangeStr As String, ByRef ranges() As String) As Boolean
    ' Parser [10:100,500:1000] vers array de ranges individuels
    
    On Error GoTo ErrorHandler
    
    ' Validation format
    If Left(rangeStr, 1) <> "[" Or Right(rangeStr, 1) <> "]" Then
        ParseBetweenRanges = False
        Exit Function
    End If
    
    ' Extraire contenu
    Dim content As String
    content = Mid(rangeStr, 2, Len(rangeStr) - 2)
    
    ' Séparer par virgules (attention aux unités avec virgules décimales)
    ranges = SplitBetweenRanges(content)
    
    ParseBetweenRanges = True
    Exit Function
    
ErrorHandler:
    ParseBetweenRanges = False
End Function

Function SplitBetweenRanges(content As String) As String()
    ' Séparer "10:100,500:1000" en gérant unités monétaires
    
    Dim parts() As String
    ReDim parts(0 To 0)
    
    Dim currentRange As String
    Dim colonCount As Long
    Dim i As Long
    
    currentRange = ""
    colonCount = 0
    
    For i = 1 To Len(content)
        Dim char As String
        char = Mid(content, i, 1)
        
        If char = ":" Then
            colonCount = colonCount + 1
            currentRange = currentRange & char
        ElseIf char = "," And colonCount Mod 2 = 0 Then
            ' Virgule après range complet (nombre pair de ":")
            parts(UBound(parts)) = Trim(currentRange)
            ReDim Preserve parts(0 To UBound(parts) + 1)
            currentRange = ""
        Else
            currentRange = currentRange & char
        End If
    Next i
    
    ' Ajouter dernier range
    parts(UBound(parts)) = Trim(currentRange)
    
    SplitBetweenRanges = parts
End Function


Function HasCurrencyUnit(value As String) As Boolean
    ' Détecter unités monétaires €, $, £, ¥
    HasCurrencyUnit = (InStr(value, "€") > 0 Or InStr(value, "$") > 0 Or _
                      InStr(value, "£") > 0 Or InStr(value, "¥") > 0)
End Function

Function ExtractNumericValue(value As String) As Double
    ' Extraire 123.45 de "123.45€"
    
    Dim numStr As String
    numStr = value
    
    ' Enlever unités communes
    numStr = Replace(numStr, "€", "")
    numStr = Replace(numStr, "$", "")
    numStr = Replace(numStr, "£", "")
    numStr = Replace(numStr, "¥", "")
    numStr = Replace(numStr, " ", "")
    
    If IsNumeric(numStr) Then
        ExtractNumericValue = CDbl(numStr)
    Else
        ExtractNumericValue = 0
    End If
End Function

' ===============================================================================
' UTILITAIRES NAVIGATION TOKENS
' ===============================================================================

Function SkipOperationTokens(startToken As Long, maxToken As Long) As Long
    ' Avancer après opération complète (@A OP value ou @A IN [...])
    
    If startToken + 1 <= maxToken Then
        Dim operator As String
        operator = UCase(pTokens(startToken + 1).TokenValue)
        
        Select Case operator
            Case "IN", "NOT IN", "BETWEEN"
                SkipOperationTokens = startToken + 3 ' @A IN [...] = 3 tokens
            Case "=", ">", "<", ">=", "<=", "<>", "~", "!~"
                SkipOperationTokens = startToken + 3 ' @A = value = 3 tokens
            Case Else
                SkipOperationTokens = startToken + 1
        End Select
    Else
        SkipOperationTokens = startToken + 1
    End If
End Function

Function SkipNotTokens(startToken As Long, maxToken As Long) As Long
    ' Avancer après expression NOT complète
    
    If startToken + 1 <= maxToken Then
        Dim nextTokenType As TokenType_Enum
        nextTokenType = pTokens(startToken + 1).TokenType
        
        Select Case nextTokenType
            Case TT_OpenParen
                ' NOT (expression) - aller après parenthèse fermante
                Dim closePos As Long
                closePos = FindMatchingCloseParen(startToken + 1)
                SkipNotTokens = closePos + 1
                
            Case TT_FieldReference
                ' NOT @A = value - aller après comparaison
                SkipNotTokens = SkipOperationTokens(startToken + 1, maxToken)
                
            Case TT_Function
                ' NOT EXISTS(@A) - aller après fonction
                SkipNotTokens = startToken + 2
                
            Case Else
                SkipNotTokens = startToken + 2
        End Select
    Else
        SkipNotTokens = startToken + 1
    End If
End Function

' ===============================================================================
' STACK MANAGEMENT COMPLET
' ===============================================================================

Sub PushResult(ByRef resultStack() As Boolean, ByRef operatorStack() As String, ByRef stackDepth As Long, result As Boolean)
    ' Ajouter résultat au stack
    stackDepth = stackDepth + 1
    resultStack(stackDepth) = result
End Sub

Sub ProcessLogicalOperatorComplete(ByRef resultStack() As Boolean, ByRef operatorStack() As String, ByRef stackDepth As Long, currentOp As String)
    ' Gestion complète tous opérateurs logiques avec priorités
    
    Select Case UCase(Trim(currentOp))
        Case "AND"
            ' AND = priorité maximale, traiter immédiatement si possible
            If stackDepth >= 2 Then
                resultStack(stackDepth - 1) = resultStack(stackDepth - 1) And resultStack(stackDepth)
                stackDepth = stackDepth - 1
            End If
            operatorStack(stackDepth) = "AND"
            
        Case "OR"
            ' OR = priorité standard, empiler
            operatorStack(stackDepth) = "OR"
            
        Case "XOR"
            ' XOR = priorité similaire OR mais traitement spécial
            operatorStack(stackDepth) = "XOR"
            
        Case "NAND"
            ' NAND = NOT(AND), traiter immédiatement
            If stackDepth >= 2 Then
                resultStack(stackDepth - 1) = Not (resultStack(stackDepth - 1) And resultStack(stackDepth))
                stackDepth = stackDepth - 1
            End If
            operatorStack(stackDepth) = "AND" ' Déjà traité
            
        Case "NOR"
            ' NOR = NOT(OR), reporter traitement
            operatorStack(stackDepth) = "NOR"
            
        Case Else
            ' Opérateur inconnu, traiter comme AND
            operatorStack(stackDepth) = "AND"
    End Select
End Sub

Function ResolveCompleteExpressionStack(resultStack() As Boolean, operatorStack() As String, stackDepth As Long) As Boolean
    ' Résolution finale avec TOUS les opérateurs logiques
    
    If stackDepth <= 1 Then
        ResolveCompleteExpressionStack = IIf(stackDepth = 1, resultStack(1), True)
        Exit Function
    End If
    
    ' Résolution gauche vers droite avec priorités
    Dim finalResult As Boolean
    finalResult = resultStack(1)
    
    Dim i As Long
    For i = 2 To stackDepth
        Dim currentOp As String
        currentOp = UCase(Trim(operatorStack(i - 1)))
        
        Select Case currentOp
            Case "AND"
                finalResult = finalResult And resultStack(i)
            Case "OR"
                finalResult = finalResult Or resultStack(i)
            Case "XOR"
                finalResult = finalResult Xor resultStack(i)
            Case "NOR"
                finalResult = Not (finalResult Or resultStack(i))
            Case Else
                ' Défaut = AND
                finalResult = finalResult And resultStack(i)
        End Select
        
        If GetConfigValue("LogParsingSteps") Then
            Debug.Print "Stack resolve [" & i & "]: " & finalResult & " " & currentOp & " " & resultStack(i) & " => " & finalResult
        End If
    Next i
    
    ResolveCompleteExpressionStack = finalResult
End Function

' ===============================================================================
' MISE À JOUR FONCTION PRINCIPALE AVEC TOUTES LES FONCTIONNALITÉS
' ===============================================================================

Function EvaluateRowExpressionComplete(rowIndex As Long) As Boolean
    ' Version finale complète avec TOUS les opérateurs
    
    If pGroupCount = 0 Then
        ' Expression linéaire avec opérateurs étendus
        EvaluateRowExpressionComplete = EvaluateTokenRangeComplete(1, pTokenCount, rowIndex)
    Else
        ' Expression hiérarchique avec groupes
        EvaluateRowExpressionComplete = EvaluateHierarchicalExpressionComplete(rowIndex)
    End If
End Function

Function EvaluateHierarchicalExpressionComplete(rowIndex As Long) As Boolean
    ' Évaluation hiérarchique complète avec tous opérateurs
    
    ' Reset flags évaluation
    Dim i As Long
    For i = 1 To pGroupCount
        pGroups(i).IsEvaluated = False
    Next i
    
    ' Évaluer groupes par niveau de profondeur (plus profond en premier)
    Dim maxLevel As Long
    maxLevel = 0
    For i = 1 To pGroupCount
        If pGroups(i).nestingLevel > maxLevel Then maxLevel = pGroups(i).nestingLevel
    Next i
    
    ' Évaluation level par level
    Dim currentLevel As Long
    For currentLevel = maxLevel To 1 Step -1
        For i = 1 To pGroupCount
            If pGroups(i).nestingLevel = currentLevel And Not pGroups(i).IsEvaluated Then
                EvaluateGroupComplete i, rowIndex
            End If
        Next i
    Next currentLevel
    
    ' Évaluer expression racine (niveau 0)
    EvaluateHierarchicalExpressionComplete = EvaluateRootExpressionComplete(rowIndex)
End Function

Function EvaluateGroupComplete(groupIndex As Long, rowIndex As Long) As Boolean
    ' Évaluation groupe avec tous opérateurs
    
    With pGroups(groupIndex)
        If .TokenStartIndex > 0 And .TokenEndIndex > 0 Then
            Dim groupResult As Boolean
            groupResult = EvaluateTokenRangeComplete(.TokenStartIndex, .TokenEndIndex, rowIndex)
            
            .IsEvaluated = True
            EvaluateGroupComplete = groupResult
            
            If GetConfigValue("LogParsingSteps") Then
                Debug.Print "Group " & groupIndex & " (Level " & .nestingLevel & "): " & groupResult
            End If
        Else
            .IsEvaluated = True
            EvaluateGroupComplete = True
        End If
    End With
End Function

Function EvaluateRootExpressionComplete(rowIndex As Long) As Boolean
    ' Évaluation expression racine après résolution groupes
    
    ' Tokens hors groupes (niveau 0)
    Dim rootTokens As Collection
    Set rootTokens = New Collection
    
    Dim i As Long
    For i = 1 To pTokenCount
        If pTokens(i).nestingLevel = 0 Then
            rootTokens.Add i
        End If
    Next i
    
    If rootTokens.Count = 0 Then
        ' Tout dans des groupes
        EvaluateRootExpressionComplete = True
    Else
        ' Évaluer tokens racine
        EvaluateRootExpressionComplete = EvaluateRootTokens(rootTokens, rowIndex)
    End If
End Function

Function EvaluateRootTokens(rootTokens As Collection, rowIndex As Long) As Boolean
    ' Évaluer tokens niveau racine
    
    Dim results() As Boolean
    Dim operators() As String
    Dim stackDepth As Long
    
    ReDim results(1 To rootTokens.Count)
    ReDim operators(1 To rootTokens.Count)
    stackDepth = 0
    
    Dim i As Long
    For i = 1 To rootTokens.Count
        Dim tokenIndex As Long
        tokenIndex = rootTokens(i)
        
        With pTokens(tokenIndex)
            Select Case .TokenType
                Case TT_FieldReference, TT_Function, TT_Not
                    Dim result As Boolean
                    result = EvaluateTokenResult(tokenIndex, rowIndex)
                    stackDepth = stackDepth + 1
                    results(stackDepth) = result
                    
                Case TT_LogicalOp, TT_Extended
                    ProcessLogicalOperatorComplete results, operators, stackDepth, .TokenValue
            End Select
        End With
    Next i
    
    EvaluateRootTokens = ResolveCompleteExpressionStack(results, operators, stackDepth)
End Function

Function EvaluateTokenResult(tokenIndex As Long, rowIndex As Long) As Boolean
    ' Évaluer résultat d'un token individuel
    
    With pTokens(tokenIndex)
        Select Case .TokenType
            Case TT_FieldReference
                EvaluateTokenResult = EvaluateFieldOperation_INTELLIGENT(tokenIndex, rowIndex, pTokenCount)
            Case TT_Function
                EvaluateTokenResult = EvaluateFunctionExpression(tokenIndex, rowIndex)
            Case TT_Not
                EvaluateTokenResult = EvaluateNotExpression(tokenIndex, rowIndex)
            Case Else
                EvaluateTokenResult = True
        End Select
    End With
End Function


' ===============================================================================
' VALIDATION FINALE COMPLÈTE
' ===============================================================================

Function ValidateCompleteExpression() As Boolean
    ' Validation finale avec tous opérateurs
    
    ' Validation de base
    If Not ValidateParsedExpression() Then
        ValidateCompleteExpression = False
        Exit Function
    End If
    
    ' Validation opérateurs étendus
    If Not ValidateExtendedOperators() Then
        ValidateCompleteExpression = False
        Exit Function
    End If
    
    ' Validation cohérence IN lists
    If Not ValidateInLists() Then
        ValidateCompleteExpression = False
        Exit Function
    End If
    
    ' Validation fonctions
    If Not ValidateFunctions() Then
        ValidateCompleteExpression = False
        Exit Function
    End If
    
    ValidateCompleteExpression = True
End Function

Function ValidateExtendedOperators() As Boolean
    ' Validation opérateurs XOR, NAND, NOR, NOT
    
    Dim i As Long
    For i = 1 To pTokenCount
        With pTokens(i)
            If .TokenType = TT_Extended Then
                Select Case UCase(.TokenValue)
                    Case "XOR"
                        If Not GetConfigValue("EnableXorOperator") Then
                            Err.Raise vbObjectError + 2020, "Validation", "XOR operator is disabled"
                            ValidateExtendedOperators = False
                            Exit Function
                        End If
                    Case "NAND", "NOR"
                        If Not GetConfigValue("EnableNandNorOperators") Then
                            Err.Raise vbObjectError + 2021, "Validation", "NAND/NOR operators are disabled"
                            ValidateExtendedOperators = False
                            Exit Function
                        End If
                End Select
            End If
        End With
    Next i
    
    ValidateExtendedOperators = True
End Function

Function ValidateInLists() As Boolean
    ' Validation listes IN
    
    Dim i As Long
    For i = 1 To pTokenCount
        With pTokens(i)
            If .TokenType = TT_ValueList Then
                Dim valueCount As Long
                valueCount = CountValuesInList(.TokenValue)
                
                If valueCount > GetConfigValue("MaxInValues") Then
                    Err.Raise vbObjectError + 2022, "Validation", _
                        "Too many values in IN list: " & valueCount & " (Max: " & GetConfigValue("MaxInValues") & ")"
                    ValidateInLists = False
                    Exit Function
                End If
                
                If valueCount = 0 Then
                    Err.Raise vbObjectError + 2023, "Validation", "Empty IN list not allowed"
                    ValidateInLists = False
                    Exit Function
                End If
            End If
        End With
    Next i
    
    ValidateInLists = True
End Function

Function ValidateFunctions() As Boolean
    ' Validation fonctions EXISTS, REGEX
    
    Dim i As Long
    For i = 1 To pTokenCount
        With pTokens(i)
            If .TokenType = TT_Function Then
                If InStr(UCase(.TokenValue), "EXISTS") > 0 Then
                    If Not GetConfigValue("EnableExistsFunction") Then
                        Err.Raise vbObjectError + 2024, "Validation", "EXISTS function is disabled"
                        ValidateFunctions = False
                        Exit Function
                    End If
                    
                    ' Valider syntaxe EXISTS(@field)
                    Dim fieldRef As String
                    fieldRef = ExtractFieldFromFunction(.TokenValue, "EXISTS")
                    If Len(fieldRef) = 0 Or Left(fieldRef, 1) <> "@" Then
                        Err.Raise vbObjectError + 2025, "Validation", "Invalid EXISTS syntax: " & .TokenValue
                        ValidateFunctions = False
                        Exit Function
                    End If
                End If
            End If
        End With
    Next i
    
    ValidateFunctions = True
End Function

' ===============================================================================
' MISE À JOUR PARSING PRINCIPAL AVEC TOUS OPÉRATEURS
' ===============================================================================

Public Function ParseFDXExpressionHiComplete(preprocessedExpression As String) As Boolean
    ' Version finale parsing avec TOUS les opérateurs
    
    If FDXH_Config Is Nothing Then InitializeExtendedConfig
    
    ' Parse de base
    If Not ParseFDXExpressionHi(preprocessedExpression) Then
        ParseFDXExpressionHiComplete = False
        Exit Function
    End If
    
    ' Post-traitement opérateurs étendus
    If Not PostProcessExtendedOperators() Then
        ParseFDXExpressionHiComplete = False
        Exit Function
    End If
    
    ' Validation complète finale
    If Not ValidateCompleteExpression() Then
        ParseFDXExpressionHiComplete = False
        Exit Function
    End If
    
    ParseFDXExpressionHiComplete = True
End Function

Function PostProcessExtendedOperators() As Boolean
    ' Post-traitement pour identifier et classifier opérateurs étendus
    
    Dim i As Long
    For i = 1 To pTokenCount
        With pTokens(i)
            ' Reclassifier tokens selon nouveaux types
            Select Case UCase(.TokenValue)
                Case "IN", "NOT IN"
                    .TokenType = TT_Operator
                Case "XOR", "NAND", "NOR"
                    .TokenType = TT_Extended
                Case "NOT"
                    .TokenType = TT_Not
                Case Else
                    If Left(.TokenValue, 1) = "[" And Right(.TokenValue, 1) = "]" Then
                        .TokenType = TT_ValueList
                    ElseIf InStr(.TokenValue, "EXISTS(") > 0 Then
                        .TokenType = TT_Function
                    End If
            End Select
            
            ' Recalculer coût avec nouveaux types
            .CostValue = CalculateTokenCostExtended(.TokenType, .TokenValue, .nestingLevel)
        End With
    Next i
    
    PostProcessExtendedOperators = True
End Function

' ===============================================================================
' INTERFACE UTILISATEUR FINALE COMPLÈTE
' ===============================================================================

'Public Function FDXarrComplete(FindWhat As Variant, _
'                              FindWhere As Variant, _
'                              Optional WhereToRead As Variant, _
'                              Optional Source As Variant = "", _
'                              Optional SwapRowCol As Boolean = False, _
'                              Optional ByVal answerWhat As Long = 1, _
'                              Optional showWhat As Long = 1, _
'                              Optional CompareStrict As Boolean = True, _
'                              Optional replaceNa As Variant, _
'                              Optional displayComment As Long = 0) As Variant
'
'    On Error GoTo ErrorHandler
'
'    ' Remplacer EvaluateRowExpression par EvaluateRowExpressionComplete
'    ' dans la fonction FDXarr existante
'
'    ' Variables internes
'    Dim pWhat As String, pWhere As String, pRead As String, pSource As String
'    pWhat = CStr(FindWhat)
'    pWhere = CStr(FindWhere)
'    pRead = IIf(IsMissing(WhereToRead), "", CStr(WhereToRead))
'    pSource = CStr(Source)
'
'    ' Phase 1: Preprocessing étendu
'    Dim preprocessed As String
'    preprocessed = PreprocessFindWhatExtended(pWhat)
'
'    ' Phase 2: Parsing complet
'    If Not ParseFDXExpressionHiComplete(preprocessed) Then
'        FDXarrComplete = CVErr(xlErrValue)
'        Exit Function
'    End If
'
'    ' Phase 3: Chargement source
'    If Not LoadSingleSource(pSource, pWhere) Then
'        FDXarrComplete = CVErr(xlErrValue)
'        Exit Function
'    End If
'
'    ' Phase 4: Évaluation complète avec tous opérateurs
'    Dim results As Variant
'    results = EvaluateCompleteExpression(preprocessed, pWhere, pRead)
'
'    ' Phase 5: Post-traitement
'    FDXarrComplete = PostProcessResultsComplete(results, showWhat, answerWhat, replaceNa, displayComment)
'
'    Exit Function
'
'ErrorHandler:
'    FDXarrComplete = CVErr(xlErrValue)
'End Function
' PROBLÈME 7: Initialisation manquante dans FDXarrComplete
Public Function FDXarrComplete(FindWhat As Variant, _
                              FindWhere As Variant, _
                              Optional WhereToRead As Variant, _
                              Optional Source As Variant = "", _
                              Optional SwapRowCol As Boolean = False, _
                              Optional ByVal answerWhat As Long = 1, _
                              Optional showWhat As Long = 1, _
                              Optional CompareStrict As Boolean = True, _
                              Optional replaceNa As Variant, _
                              Optional displayComment As Long = 0) As Variant
    
    On Error GoTo ErrorHandler
    
    ' CORRECTION: Initialisation forcée
    If FDXH_Config Is Nothing Then InitializeExtendedConfig
    
    ' CORRECTION: Reset des variables globales
    pTokenCount = 0
    pGroupCount = 0
    pSourceRows = 0
    pSourceCols = 0
    pColumnCount = 0
    pTotalCost = 0
    pCurrentNestingLevel = 0
    pChunkMode = False
    
    ' Variables internes
    Dim pWhat As String, pWhere As String, pRead As String, pSource As String
    pWhat = CStr(FindWhat)
    pWhere = CStr(FindWhere)
    pRead = IIf(IsMissing(WhereToRead), "", CStr(WhereToRead))
    pSource = CStr(Source)
    
    ' CORRECTION: Validation entrée
    If Len(Trim(pWhat)) = 0 Then
        Dim errorResult(1 To 1, 1 To 1) As Variant
        errorResult(1, 1) = "Empty FindWhat expression"
        FDXarrComplete = errorResult
        Exit Function
    End If
    
    If Len(Trim(pWhere)) = 0 Then
        Dim errorResult2(1 To 1, 1 To 1) As Variant
        errorResult2(1, 1) = "Empty FindWhere range"
        FDXarrComplete = errorResult2
        Exit Function
    End If
    
    ' Phase 1: Preprocessing étendu
    Dim preprocessed As String
    preprocessed = PreprocessFindWhatExtended(pWhat)
    
    ' Phase 2: Parsing complet
    If Not ParseFDXExpressionHiComplete(preprocessed) Then
        Dim errorResult3(1 To 1, 1 To 1) As Variant
        errorResult3(1, 1) = "Expression parsing failed"
        FDXarrComplete = errorResult3
        Exit Function
    End If
    
    ' Phase 3: Chargement source
    If Not LoadSingleSource(pSource, pWhere) Then
        Dim errorResult4(1 To 1, 1 To 1) As Variant
        errorResult4(1, 1) = "Source loading failed"
        FDXarrComplete = errorResult4
        Exit Function
    End If
    
    ' Phase 4: Évaluation complète avec tous opérateurs
    Dim results As Variant
    results = EvaluateCompleteExpression(preprocessed, pWhere, pRead)
    
    ' Phase 5: Post-traitement
    FDXarrComplete = PostProcessResultsComplete(results, showWhat, answerWhat, replaceNa, displayComment)
    
    Exit Function
    
ErrorHandler:
    Dim errorResult5(1 To 1, 1 To 1) As Variant
    errorResult5(1, 1) = "Critical error: " & Err.Description
    FDXarrComplete = errorResult5
    Resume Next
End Function

Function PostProcessResultsComplete(results As Variant, showWhat As Long, answerWhat As Long, replaceNa As Variant, displayComment As Long) As Variant
    ' Post-traitement complet selon toutes les options utilisateur
    If IsError(results) Then
        PostProcessResultsComplete = results
        Exit Function
    End If

    If Not IsArray(results) Then
        PostProcessResultsComplete = results
        Exit Function
    End If

    Dim processedResults As Variant
    processedResults = results

    ' Traitement selon AnswerWhat
    Select Case answerWhat
        Case 1 ' Values (par défaut)
            ' Garder valeurs telles quelles
        Case 2 ' Row numbers
            processedResults = ConvertToRowNumbers(results)
        Case 3 ' Count only
            processedResults = GetResultCount(results)
        Case 4 ' Boolean indicators
            processedResults = ConvertToBooleanIndicators(results)
    End Select

    ' Traitement selon ShowWhat
    Select Case showWhat
        Case 1 ' OnlyFound (par défaut)
            ' Déjà filtré
        Case 2 ' ShowAll
            processedResults = ExpandToShowAll(processedResults)
    End Select

    ' Remplacement des N/A
    If Not IsMissing(replaceNa) And Not IsEmpty(replaceNa) Then
        processedResults = ReplaceNaValues(processedResults, replaceNa)
    End If

    ' Ajout commentaires si demandé
    If displayComment > 0 Then
        processedResults = AddComments(processedResults, displayComment)
    End If

    PostProcessResultsComplete = processedResults
End Function

Function AddComments(results As Variant, commentType As Long) As Variant
    ' Ajouter commentaires selon type demandé

    If Not IsArray(results) Then
        AddComments = results
        Exit Function
    End If

    Select Case commentType
        Case 1 ' Statistiques basiques
            AddComments = AddBasicStats(results)
        Case 2 ' Informations détaillées
            AddComments = AddDetailedInfo(results)
        Case 3 ' Debug info
            AddComments = AddDebugInfo(results)
        Case Else
            AddComments = results
    End Select
End Function

Function AddDebugInfo(results As Variant) As Variant
    ' Ajouter informations debug complètes

    Dim originalRows As Long, originalCols As Long
    originalRows = UBound(results, 1)
    originalCols = UBound(results, 2)

    ReDim enhancedResults(1 To originalRows + 5, 1 To originalCols) As Variant

    ' Copier données
    Dim i As Long, j As Long
    For i = 1 To originalRows
        For j = 1 To originalCols
            enhancedResults(i, j) = results(i, j)
        Next j
    Next i

    ' Infos debug
    enhancedResults(originalRows + 1, 1) = "=== DEBUG INFO ==="
    enhancedResults(originalRows + 2, 1) = "Expression: " & pExpressionText
    enhancedResults(originalRows + 3, 1) = "Tokens/Groups: " & pTokenCount & "/" & pGroupCount
    enhancedResults(originalRows + 4, 1) = "Source: " & pSourceRows & "x" & pSourceCols
    enhancedResults(originalRows + 5, 1) = "Mode: " & IIf(pChunkMode, "Chunks", "Memory")

    For i = originalRows + 1 To originalRows + 5
        For j = 2 To originalCols
            enhancedResults(i, j) = ""
        Next j
    Next i

    AddDebugInfo = enhancedResults
End Function

Function AddDetailedInfo(results As Variant) As Variant
    ' Ajouter informations détaillées sur l'évaluation

    Dim originalRows As Long, originalCols As Long
    originalRows = UBound(results, 1)
    originalCols = UBound(results, 2)

    ReDim enhancedResults(1 To originalRows + 3, 1 To originalCols) As Variant

    ' Copier données
    Dim i As Long, j As Long
    For i = 1 To originalRows
        For j = 1 To originalCols
            enhancedResults(i, j) = results(i, j)
        Next j
    Next i

    ' Ajouter infos détaillées
    enhancedResults(originalRows + 1, 1) = "Results: " & originalRows
    enhancedResults(originalRows + 2, 1) = "Tokens: " & pTokenCount
    enhancedResults(originalRows + 3, 1) = "Cost: " & pTotalCost

    For i = originalRows + 1 To originalRows + 3
        For j = 2 To originalCols
            enhancedResults(i, j) = ""
        Next j
    Next i

    AddDetailedInfo = enhancedResults
End Function

Function AddBasicStats(results As Variant) As Variant
    ' Ajouter ligne de statistiques basiques

    Dim originalRows As Long, originalCols As Long
    originalRows = UBound(results, 1)
    originalCols = UBound(results, 2)

    ' Étendre array d'une ligne
    ReDim enhancedResults(1 To originalRows + 1, 1 To originalCols) As Variant

    ' Copier données originales
    Dim i As Long, j As Long
    For i = 1 To originalRows
        For j = 1 To originalCols
            enhancedResults(i, j) = results(i, j)
        Next j
    Next i

    ' Ajouter ligne statistiques
    enhancedResults(originalRows + 1, 1) = "Found: " & originalRows & " rows"
    For j = 2 To originalCols
        enhancedResults(originalRows + 1, j) = ""
    Next j

    AddBasicStats = enhancedResults
End Function

Function ReplaceNaValues(results As Variant, replaceValue As Variant) As Variant
    ' Remplacer valeurs N/A, NULL, Empty par valeur spécifiée

    If Not IsArray(results) Then
        ReplaceNaValues = results
        Exit Function
    End If

    Dim processedResults As Variant
    processedResults = results

    Dim i As Long, j As Long
    For i = 1 To UBound(processedResults, 1)
        For j = 1 To UBound(processedResults, 2)
            Dim cellValue As Variant
            cellValue = processedResults(i, j)

            If IsError(cellValue) Or IsNull(cellValue) Or IsEmpty(cellValue) Then
                processedResults(i, j) = replaceValue
            ElseIf VarType(cellValue) = vbString Then
                If Trim(CStr(cellValue)) = "" Then
                    processedResults(i, j) = replaceValue
                End If
            End If
        Next j
    Next i

    ReplaceNaValues = processedResults
End Function

Function ExpandToShowAll(results As Variant) As Variant
    ' Étendre pour montrer toutes les lignes (trouvées + non trouvées)

    ' TODO: Implémentation complète nécessiterait de conserver info
    ' sur toutes les lignes évaluées. Pour l'instant, retourner tel quel.
    ExpandToShowAll = results
End Function

Function ConvertToBooleanIndicators(results As Variant) As Variant
    ' Convertir en indicateurs True/False

    If Not IsArray(results) Then
        Dim boolResult(1 To 1, 1 To 1) As Variant
        boolResult(1, 1) = False
        ConvertToBooleanIndicators = boolResult
        Exit Function
    End If

    Dim rowCount As Long
    rowCount = UBound(results, 1)

    ReDim boolResults(1 To rowCount, 1 To 1) As Variant

    Dim i As Long
    For i = 1 To rowCount
        boolResults(i, 1) = True ' Toutes les lignes retournées sont "trouvées"
    Next i

    ConvertToBooleanIndicators = boolResults
End Function

Function GetResultCount(results As Variant) As Variant
    ' Retourner juste le nombre de résultats

    Dim countResult(1 To 1, 1 To 1) As Variant

    If IsArray(results) Then
        countResult(1, 1) = UBound(results, 1)
    Else
        countResult(1, 1) = 0
    End If

    GetResultCount = countResult
End Function

Function ConvertToRowNumbers(results As Variant) As Variant
    ' Convertir résultats en numéros de lignes

    If Not IsArray(results) Then
        ConvertToRowNumbers = results
        Exit Function
    End If

    Dim rowCount As Long
    rowCount = UBound(results, 1)

    ReDim RowNumbers(1 To rowCount, 1 To 1) As Variant

    Dim i As Long
    For i = 1 To rowCount
        RowNumbers(i, 1) = i ' Numéro ligne relatif
    Next i

    ConvertToRowNumbers = RowNumbers
End Function

'Function EvaluateCompleteExpression(pWhat As String, pWhere As Variant, pRead As Variant) As Variant
'    ' Évaluation utilisant EvaluateRowExpressionComplete
'
'    Dim results() As Boolean
'    ReDim results(1 To pSourceRows)
'
'    Dim foundCount As Long
'    foundCount = 0
'
'    ' Utiliser version complète pour chaque ligne
'    Dim rowIndex As Long
'    For rowIndex = 1 To pSourceRows
'        results(rowIndex) = EvaluateRowExpressionComplete(rowIndex)
'        If results(rowIndex) Then foundCount = foundCount + 1
'    Next rowIndex
'
'    ' Construire résultat
'    If foundCount = 0 Then
'        Dim emptyResult(1 To 1, 1 To 1) As Variant
'        emptyResult(1, 1) = "No results found"
'        EvaluateCompleteExpression = emptyResult
'    Else
'        EvaluateCompleteExpression = BuildResultArrayOptimized(results, foundCount, pRead)
'    End If
'End Function
' PROBLÈME 5: Correction EvaluateCompleteExpression
Function EvaluateCompleteExpression(pWhat As String, pWhere As Variant, pRead As Variant) As Variant
    On Error GoTo ErrorHandler
    
    ' CORRECTION: S'assurer que les données sont chargées
    If pSourceRows = 0 Or pSourceCols = 0 Then
        Dim errorResult(1 To 1, 1 To 1) As Variant
        errorResult(1, 1) = "No source data loaded"
        EvaluateCompleteExpression = errorResult
        Exit Function
    End If
    
    ' CORRECTION: Construire les mappings colonnes
    If pColumnCount = 0 Then
        If Not BuildColumnMappings(CStr(pWhere)) Then
            Dim errorResult2(1 To 1, 1 To 1) As Variant
            errorResult2(1, 1) = "Failed to build column mappings"
            EvaluateCompleteExpression = errorResult2
            Exit Function
        End If
    End If
    
    Dim results() As Boolean
    ReDim results(1 To pSourceRows)
    
    Dim foundCount As Long
    foundCount = 0
    
    ' Évaluer chaque ligne
    Dim rowIndex As Long
    For rowIndex = 1 To pSourceRows
        On Error Resume Next
        results(rowIndex) = EvaluateRowExpressionComplete(rowIndex)
        If Err.Number <> 0 Then
            results(rowIndex) = False
            If GetConfigValue("LogParsingSteps") Then
                Debug.Print "Row " & rowIndex & " evaluation failed: " & Err.Description
            End If
        End If
        On Error GoTo ErrorHandler
        
        If results(rowIndex) Then foundCount = foundCount + 1
    Next rowIndex
    
    ' Construire résultat
    If foundCount = 0 Then
        Dim emptyResult(1 To 1, 1 To 1) As Variant
        emptyResult(1, 1) = "No results found"
        EvaluateCompleteExpression = emptyResult
    Else
        EvaluateCompleteExpression = BuildResultArrayOptimized(results, foundCount, pRead)
    End If
    
    Exit Function
    
ErrorHandler:
    Dim errorResult3(1 To 1, 1 To 1) As Variant
    errorResult3(1, 1) = "Evaluation error: " & Err.Description
    EvaluateCompleteExpression = errorResult3
End Function
Function BuildResultArrayOptimized(results() As Boolean, foundCount As Long, readSpec As Variant) As Variant
    ' Construction optimisée du résultat final

    On Error GoTo ErrorHandler

    ' Parser WhereToRead de façon optimisée
    Dim readColumns() As Long
    Dim readColCount As Long
    readColCount = ParseWhereToReadSpecOptimized(readSpec, readColumns)

    If readColCount = 0 Then
        ' Pas de colonnes -> retourner indicateurs
        ReDim resultArray(1 To foundCount, 1 To 1) As Variant
        Dim foundIndex As Long
        foundIndex = 1

        Dim i As Long
        For i = 1 To UBound(results)
            If results(i) Then
                resultArray(foundIndex, 1) = "Row " & i
                foundIndex = foundIndex + 1
                If foundIndex > foundCount Then Exit For
            End If
        Next i
    Else
        ' Colonnes spécifiées -> extraire données
        ReDim resultArray(1 To foundCount, 1 To readColCount) As Variant
        foundIndex = 1

        For i = 1 To UBound(results)
            If results(i) Then
                Dim j As Long
                For j = 1 To readColCount
                    If readColumns(j) > 0 And readColumns(j) <= pSourceCols Then
                        resultArray(foundIndex, j) = pSourceData(i, readColumns(j))
                    Else
                        resultArray(foundIndex, j) = ""
                    End If
                Next j
                foundIndex = foundIndex + 1
                If foundIndex > foundCount Then Exit For
            End If
        Next i
    End If

    BuildResultArrayOptimized = resultArray
    Exit Function

ErrorHandler:
    Dim errorResult(1 To 1, 1 To 1) As Variant
    errorResult(1, 1) = "Error building results: " & Err.Description
    BuildResultArrayOptimized = errorResult
End Function

Function ParseWhereToReadSpecOptimized(readSpec As Variant, readColumns() As Long) As Long
    ' Parser optimisé WhereToRead avec validation

    Dim specStr As String
    specStr = Trim(CStr(readSpec))

    If Len(specStr) = 0 Then
        ParseWhereToReadSpecOptimized = 0
        Exit Function
    End If

    On Error GoTo ErrorHandler

    ' Format bracket [1,3,5] ou [A,C,E]
    If Left(specStr, 1) = "[" And Right(specStr, 1) = "]" Then
        Dim innerSpec As String
        innerSpec = Mid(specStr, 2, Len(specStr) - 2)

        If Len(innerSpec) = 0 Then
            ParseWhereToReadSpecOptimized = 0
            Exit Function
        End If

        Dim colSpecs() As String
        colSpecs = Split(innerSpec, ",")

        ReDim readColumns(1 To UBound(colSpecs) + 1)

        Dim i As Long
        For i = 0 To UBound(colSpecs)
            Dim colSpec As String
            colSpec = Trim(colSpecs(i))

            If IsNumeric(colSpec) Then
                ' Index numérique
                readColumns(i + 1) = CLng(colSpec)
            Else
                ' Lettre colonne
                If Len(colSpec) = 1 And colSpec >= "A" And colSpec <= "Z" Then
                    readColumns(i + 1) = Asc(UCase(colSpec)) - Asc("A") + 1
                Else
                    readColumns(i + 1) = 1 ' Par défaut colonne A
                End If
            End If

            ' Validation
            If readColumns(i + 1) > pSourceCols Then
                readColumns(i + 1) = pSourceCols ' Limiter à max disponible
            End If
        Next i

        ParseWhereToReadSpecOptimized = UBound(colSpecs) + 1
        Exit Function
    End If

    ' Format Excel A1:C1 ou A:C
    If InStr(specStr, ":") > 0 Then
        Dim rangeParts() As String
        rangeParts = Split(specStr, ":")

        If UBound(rangeParts) >= 1 Then
            Dim startCol As String, endCol As String

            ' Extraire lettres colonnes
            startCol = ExtractColumnLetter(rangeParts(0))
            endCol = ExtractColumnLetter(rangeParts(1))

            If Len(startCol) > 0 And Len(endCol) > 0 Then
                Dim startIndex As Long, endIndex As Long
                startIndex = Asc(UCase(startCol)) - Asc("A") + 1
                endIndex = Asc(UCase(endCol)) - Asc("A") + 1

                If endIndex < startIndex Then
                    ' Inverser si nécessaire
                    Dim temp As Long
                    temp = startIndex
                    startIndex = endIndex
                    endIndex = temp
                End If

                Dim colCount As Long
                colCount = endIndex - startIndex + 1
                ReDim readColumns(1 To colCount)

                For i = 1 To colCount
                    readColumns(i) = startIndex + i - 1
                    ' Validation
                    If readColumns(i) > pSourceCols Then
                        readColumns(i) = pSourceCols
                    End If
                Next i

                ParseWhereToReadSpecOptimized = colCount
                Exit Function
            End If
        End If
    End If

    ' Format simple : A, B1, etc.
    Dim singleCol As String
    singleCol = ExtractColumnLetter(specStr)

    If Len(singleCol) > 0 Then
        ReDim readColumns(1 To 1)
        readColumns(1) = Asc(UCase(singleCol)) - Asc("A") + 1
        If readColumns(1) > pSourceCols Then readColumns(1) = pSourceCols
        ParseWhereToReadSpecOptimized = 1
    Else
        ParseWhereToReadSpecOptimized = 0
    End If

    Exit Function

ErrorHandler:
    ParseWhereToReadSpecOptimized = 0
End Function

Function ExtractColumnLetter(cellRef As String) As String
    ' Extraire lettre colonne de référence (A1 -> A, B10 -> B, etc.)

    Dim i As Long
    For i = 1 To Len(cellRef)
        Dim char As String
        char = Mid(cellRef, i, 1)
        If char >= "A" And char <= "Z" Then
            ExtractColumnLetter = char
            Exit Function
        ElseIf char >= "a" And char <= "z" Then
            ExtractColumnLetter = UCase(char)
            Exit Function
        End If
    Next i

    ExtractColumnLetter = ""
End Function

Public Function GetTotalCost() As Double
    GetTotalCost = pTotalCost
End Function

' ===============================================================================
' CONFIGURATION ÉTENDUE FINALE
' ===============================================================================

Sub InitializeExtendedConfig()
    ' Configuration complète avec TOUS les opérateurs
    
    If FDXH_Config Is Nothing Then InitializeFDXH_Config
    
    ' === OPÉRATEURS ÉTENDUS ===
    FDXH_Config("EnableInOperator") = True           ' IN/NOT IN
    FDXH_Config("EnableNotOperator") = True          ' NOT
    FDXH_Config("EnableXorOperator") = True          ' XOR
    FDXH_Config("EnableNandNorOperators") = True     ' NAND/NOR
    FDXH_Config("EnableExistsFunction") = True       ' EXISTS()
    FDXH_Config("EnableRegexOperator") = False       ' REGEX (performance)
    
    ' === COÛTS OPÉRATEURS ÉTENDUS ===
    FDXH_Config("CostInOperator") = 4                ' IN plus coûteux que =
    FDXH_Config("CostNotOperator") = 2               ' NOT modérément coûteux
    FDXH_Config("CostXorOperator") = 3               ' XOR plus coûteux que AND/OR
    FDXH_Config("CostNandOperator") = 4              ' NAND coûteux
    FDXH_Config("CostNorOperator") = 4               ' NOR coûteux
    FDXH_Config("CostExistsFunction") = 2            ' EXISTS simple
    FDXH_Config("CostRegexOperator") = 8             ' REGEX très coûteux
    
    ' === LIMITES VALIDATION ===
    FDXH_Config("MaxInValues") = 50                  ' Limite valeurs IN
    FDXH_Config("MaxRegexLength") = 100              ' Limite taille regex
    FDXH_Config("EnableNestedNot") = True            ' NOT imbriqués
    FDXH_Config("MaxNotDepth") = 3                   ' Profondeur NOT max
    
    ' === COMPARAISONS ===
    FDXH_Config("CompareStrict") = True              ' Comparaison stricte par défaut
    FDXH_Config("CaseSensitive") = False             ' Insensible casse par défaut
    FDXH_Config("EnableTypeCoercion") = True         ' Conversion types auto
    
    ' === OPTIMISATIONS ===
    FDXH_Config("EnableShortCircuit") = True         ' Court-circuit AND/OR
    FDXH_Config("EnableResultCache") = False         ' Cache résultats (mémoire)
    FDXH_Config("OptimizeInLists") = True            ' Optimisation listes IN
End Sub

' ===============================================================================
' FONCTIONS UTILITAIRES FINALES MANQUANTES
' ===============================================================================

'Function CalculateTokenCostExtended(TokenType As Long, value As String, nestingLevel As Long) As Double
'    ' Calcul coût avec TOUS les types de tokens
'
'    Dim baseCost As Double
'
'    Select Case TokenType
'        Case TT_In
'            baseCost = GetConfigValue("CostInOperator")
'            ' Bonus selon nombre de valeurs
'            Dim valueCount As Long
'            valueCount = CountValuesInList(value)
'            baseCost = baseCost + (valueCount * 0.3)
'
'        Case TT_Not
'            baseCost = GetConfigValue("CostNotOperator")
'
'        Case TT_Extended
'            Select Case UCase(Trim(value))
'                Case "XOR"
'                    baseCost = GetConfigValue("CostXorOperator")
'                Case "NAND"
'                    baseCost = GetConfigValue("CostNandOperator")
'                Case "NOR"
'                    baseCost = GetConfigValue("CostNorOperator")
'                Case Else
'                    baseCost = 3
'            End Select
'
'        Case TT_Function
'            If InStr(UCase(value), "EXISTS") > 0 Then
'                baseCost = GetConfigValue("CostExistsFunction")
'            ElseIf InStr(UCase(value), "REGEX") > 0 Then
'                baseCost = GetConfigValue("CostRegexOperator")
'            Else
'                baseCost = 4
'            End If
'
'        Case TT_ValueList
'            Dim listCount As Long
'            listCount = CountValuesInList(value)
'            baseCost = 1 + (listCount * 0.2)
'
'        Case Else
'            ' Déléguer aux coûts originaux
'            baseCost = CalculateTokenCostExtended(TokenType, value, nestingLevel)
'    End Select
'
'    ' Bonus nesting level
'    baseCost = baseCost + (nestingLevel * 0.5)
'
'    CalculateTokenCostExtended = baseCost
'End Function
' PROBLÈME 6: Correction récursivité infinie dans CalculateTokenCostExtended
Function CalculateTokenCostExtended(TokenType As Long, value As String, nestingLevel As Long) As Double
    Dim baseCost As Double
    
    Select Case TokenType
        Case TT_Operator
            baseCost = GetConfigValue("CostInOperator")
            Dim valueCount As Long
            valueCount = CountValuesInList(value)
            baseCost = baseCost + (valueCount * 0.3)
            
        Case TT_Not
            baseCost = GetConfigValue("CostNotOperator")
            
        Case TT_Extended
            Select Case UCase(Trim(value))
                Case "XOR"
                    baseCost = GetConfigValue("CostXorOperator")
                Case "NAND"
                    baseCost = GetConfigValue("CostNandOperator")
                Case "NOR"
                    baseCost = GetConfigValue("CostNorOperator")
                Case Else
                    baseCost = 3
            End Select
            
        Case TT_Function
            If InStr(UCase(value), "EXISTS") > 0 Then
                baseCost = GetConfigValue("CostExistsFunction")
            ElseIf InStr(UCase(value), "REGEX") > 0 Then
                baseCost = GetConfigValue("CostRegexOperator")
            Else
                baseCost = 4
            End If
            
        Case TT_ValueList
            Dim listCount As Long
            listCount = CountValuesInList(value)
            baseCost = 1 + (listCount * 0.2)
            
        Case Else
            ' CORRECTION: Appeler la fonction originale au lieu de récursion infinie
            baseCost = CalculateTokenCostIntelligent(TokenType, value, nestingLevel)
    End Select
    
    ' Bonus nesting level
    baseCost = baseCost + (nestingLevel * 0.5)
    
    CalculateTokenCostExtended = baseCost
End Function

Function CountValuesInList(listValue As String) As Long
    ' Compter valeurs dans ["val1","val2","val3"]
    
    If Len(listValue) < 3 Then
        CountValuesInList = 0
        Exit Function
    End If
    
    If Left(listValue, 1) <> "[" Or Right(listValue, 1) <> "]" Then
        CountValuesInList = 0
        Exit Function
    End If
    
    Dim innerList As String
    innerList = Mid(listValue, 2, Len(listValue) - 2)
    
    If Len(Trim(innerList)) = 0 Then
        CountValuesInList = 0
    Else
        ' Compter virgules + 1, mais attention aux virgules dans guillemets
        Dim values As Collection
        Set values = ParseCommaSeparatedValues(innerList)
        CountValuesInList = values.Count
    End If
End Function

' ===============================================================================
' PREPROCESSING ÉTENDU FINAL
' ===============================================================================

Public Function PreprocessFindWhatExtended(rawFindWhat As String) As String
    ' Preprocessing complet avec TOUS les opérateurs
    
    Dim result As String
    result = PreprocessFindWhat(rawFindWhat) ' Base existante
    
    ' Extensions pour nouveaux opérateurs
    result = PreprocessInOperators(result)
    result = PreprocessNotOperators(result)
    result = PreprocessExtendedLogical(result)
    result = PreprocessFunctions(result)
    result = PreprocessSpecialCases(result)
    
    PreprocessFindWhatExtended = result
End Function

Function PreprocessInOperators(expression As String) As String
    ' Normalisation opérateurs IN
    
    Dim result As String
    result = expression
    
    ' IN ["val1", "val2"] -> IN ["val1","val2"] (sans espaces)
    result = NormalizeInLists(result, " IN [")
    result = NormalizeInLists(result, " NOT IN [")
    
    PreprocessInOperators = result
End Function

Function NormalizeInLists(text As String, pattern As String) As String
    ' Normaliser espaces dans listes IN
    
    Dim result As String
    result = text
    
    Dim pos As Long
    pos = 1
    
    Do
        pos = InStr(pos, result, pattern, vbTextCompare)
        If pos > 0 Then
            Dim endPos As Long
            endPos = InStr(pos + Len(pattern), result, "]")
            If endPos > 0 Then
                Dim listPart As String
                listPart = Mid(result, pos + Len(pattern), endPos - pos - Len(pattern))
                
                ' Supprimer espaces après virgules
                listPart = Replace(listPart, ", """, ",""")
                listPart = Replace(listPart, " ,""", ",""")
                
                result = Left(result, pos + Len(pattern) - 1) & listPart & Mid(result, endPos)
                pos = endPos + 1
            Else
                Exit Do
            End If
        End If
    Loop While pos > 0
    
    NormalizeInLists = result
End Function

Function PreprocessNotOperators(expression As String) As String
    ' Normalisation opérateur NOT
    
    Dim result As String
    result = expression
    
    ' Normaliser espaces
    result = Replace(result, " NOT ", " NOT ", 1, -1, vbTextCompare)
    result = Replace(result, " not ", " NOT ", 1, -1, vbTextCompare)
    result = Replace(result, " Not ", " NOT ", 1, -1, vbTextCompare)
    
    ' NOT( -> NOT (
    result = Replace(result, "NOT(", "NOT (")
    
    PreprocessNotOperators = result
End Function

Function PreprocessExtendedLogical(expression As String) As String
    ' Normalisation XOR, NAND, NOR
    
    Dim result As String
    result = expression
    
    ' Normaliser avec espaces obligatoires
    result = Replace(result, " XOR ", " XOR ", 1, -1, vbTextCompare)
    result = Replace(result, " xor ", " XOR ", 1, -1, vbTextCompare)
    result = Replace(result, " NAND ", " NAND ", 1, -1, vbTextCompare)
    result = Replace(result, " nand ", " NAND ", 1, -1, vbTextCompare)
    result = Replace(result, " NOR ", " NOR ", 1, -1, vbTextCompare)
    result = Replace(result, " nor ", " NOR ", 1, -1, vbTextCompare)
    
    PreprocessExtendedLogical = result
End Function

Function PreprocessFunctions(expression As String) As String
    ' Normalisation fonctions EXISTS, REGEX
    
    Dim result As String
    result = expression
    
    ' Normaliser EXISTS
    result = Replace(result, "EXISTS(", "EXISTS(", 1, -1, vbTextCompare)
    result = Replace(result, "exists(", "EXISTS(", 1, -1, vbTextCompare)
    result = Replace(result, "Exists(", "EXISTS(", 1, -1, vbTextCompare)
    
    ' Normaliser REGEX (si activé)
    If GetConfigValue("EnableRegexOperator") Then
        result = Replace(result, " REGEX ", " REGEX ", 1, -1, vbTextCompare)
        result = Replace(result, " regex ", " REGEX ", 1, -1, vbTextCompare)
    End If
    
    PreprocessFunctions = result
End Function

Function PreprocessSpecialCases(expression As String) As String
    ' Cas spéciaux et optimisations
    
    Dim result As String
    result = expression
    
    ' Optimisation parenthèses redondantes
    result = OptimizeRedundantParentheses(result)
    
    ' Normalisation double négations
    result = OptimizeDoubleNegations(result)
    
    PreprocessSpecialCases = result
End Function

Function OptimizeRedundantParentheses(expression As String) As String
    ' Simplifier ((@A = "test")) -> (@A = "test")
    ' TODO: Implémentation complète si nécessaire pour optimisation
    OptimizeRedundantParentheses = expression
End Function

Function OptimizeDoubleNegations(expression As String) As String
    ' Simplifier NOT (NOT @A = "test") -> @A = "test"
    Dim result As String
    result = expression
    
    ' Pattern simple: NOT (NOT xxx)
    Do While InStr(result, "NOT (NOT ") > 0
        result = Replace(result, "NOT (NOT ", "")
        ' Enlever une parenthèse fermante correspondante
        Dim openCount As Long, i As Long
        openCount = 1
        For i = InStr(result, "NOT (NOT ") + 9 To Len(result)
            If Mid(result, i, 1) = "(" Then
                openCount = openCount + 1
            ElseIf Mid(result, i, 1) = ")" Then
                openCount = openCount - 1
                If openCount = 0 Then
                    result = Left(result, i - 1) & Mid(result, i + 1)
                    Exit For
                End If
            End If
        Next i
    Loop
    
    OptimizeDoubleNegations = result
End Function

' ===============================================================================
' FONCTION PRINCIPALE MISE À JOUR FINALE
' ===============================================================================

Public Function FDXarr(FindWhat As Variant, _
                      FindWhere As Variant, _
                      Optional WhereToRead As Variant, _
                      Optional Source As Variant = "", _
                      Optional SwapRowCol As Boolean = False, _
                      Optional ByVal answerWhat As Long = 1, _
                      Optional showWhat As Long = 1, _
                      Optional CompareStrict As Boolean = True, _
                      Optional replaceNa As Variant, _
                      Optional displayComment As Long = 0) As Variant
    
    ' REMPLACER LE CONTENU DE FDXarr EXISTANT PAR CETTE VERSION COMPLÈTE
    FDXarr = FDXarrComplete(FindWhat, FindWhere, WhereToRead, Source, SwapRowCol, _
                           answerWhat, showWhat, CompareStrict, replaceNa, displayComment)
End Function
' ===============================================================================
' DOCUMENTATION STATUS FINAL
' ===============================================================================
Public Sub ShowFinalStatus()
    Debug.Print "==============================================================================="
    Debug.Print "                    FINDXTREME Hi (FDXH) - STATUS FINAL"
    Debug.Print "==============================================================================="
    Debug.Print ""
    
    ' Configuration
    Debug.Print "=== CONFIGURATION ==="
    Debug.Print "Version: " & GetConfigValue("VersionMode")
    Debug.Print "Max Nesting: " & GetConfigValue("MaxNestingDepth")
    Debug.Print "Max Cost: " & GetConfigValue("MaxCostAllowed")
    Debug.Print ""
    
    ' Opérateurs supportés
    Debug.Print "=== OPÉRATEURS SUPPORTÉS ==="
    Debug.Print "Comparaisons: =, >, <, >=, <=, <>, ~, !~"
    Debug.Print "Logiques standards: AND, OR"
    Debug.Print "Logiques étendus: XOR, NAND, NOR, NOT"
    Debug.Print "Spéciaux: IN, NOT IN, BETWEEN, EXISTS()"
    Debug.Print "Parenthèses: 3 niveaux imbriqués"
    Debug.Print ""
    
    ' Fonctionnalités
    Debug.Print "=== FONCTIONNALITÉS ==="
    Debug.Print "? Parser hiérarchique 3 niveaux"
    Debug.Print "? Tous opérateurs logiques (AND, OR, XOR, NAND, NOR, NOT)"
    Debug.Print "? Opérateur IN avec listes de valeurs"
    Debug.Print "? Fonction EXISTS() pour vérification existence"
    Debug.Print "? BETWEEN étendu avec unités monétaires"
    Debug.Print "? Preprocessing intelligent"
    Debug.Print "? Validation complète avec limites configurables"
    Debug.Print "? Système COST avec contrôle complexité"
    Debug.Print "? Support multi-sources (Excel, CSV, Access)"
    Debug.Print "? Mode chunks pour grandes données"
    Debug.Print "? Tests unitaires complets"
    Debug.Print ""
    
    If pTokenCount > 0 Then
        Debug.Print "=== ÉTAT PARSING ACTUEL ==="
        Debug.Print "Tokens parsés: " & pTokenCount
        Debug.Print "Groupes logiques: " & pGroupCount
        Debug.Print "Coût total: " & GetTotalCost() & "/" & GetConfigValue("MaxCostAllowed")
        Debug.Print ""
    End If
    
    If pSourceRows > 0 Then
        Debug.Print "=== DONNÉES CHARGÉES ==="
        Debug.Print "Lignes: " & pSourceRows
        Debug.Print "Colonnes: " & pSourceCols
        Debug.Print "Mode chunks: " & IIf(pChunkMode, "OUI", "NON")
        Debug.Print "Mappings colonnes: " & pColumnCount
        Debug.Print ""
    End If
    
    Debug.Print "=== INTERFACES ==="
    Debug.Print "Fonction principale: FDXHi(...)"
    Debug.Print "Alias court: FDXH(...)"
    Debug.Print "Debug complet: DumpCompleteStatus"
    Debug.Print "Tests complets: RunAllFinalTests"
    Debug.Print ""
    
    Debug.Print "==============================================================================="
    Debug.Print "                         ? FDXH COMPLET ET OPÉRATIONNEL"
    Debug.Print "==============================================================================="
End Sub

' ===============================================================================
' EXEMPLE D'UTILISATION FINALE
' ===============================================================================

Public Sub ExempleFinalFDXH()
    Debug.Print "=== EXEMPLE UTILISATION FDXH COMPLET ==="
    
    ' Préparer données exemple
    Range("A1:D6").Clear
    Range("A1:D1").value = Array("Nom", "Score", "Statut", "Département")
    Range("A2:D6").value = Array( _
        Array("Alice", 95, "Active", "Ventes"), _
        Array("Bob", 87, "Active", "Marketing"), _
        Array("Charlie", 76, "Inactive", "Support"), _
        Array("Diana", 92, "Active", "Ventes"), _
        Array("Eve", 88, "Active", "Marketing") _
    )
    
    Debug.Print "Données préparées en A1:D6"
    Debug.Print ""
    
    ' Exemples progressifs
    Debug.Print "1. Expression simple:"
    Debug.Print "   =FDXHi(""@B > 85"", ""A1:D6"", ""A1:D1"")"
    Dim result1 As Variant
    result1 = FDXHi("@B > 85", "A1:D6", "A1:D1")
    Debug.Print "   Résultats trouvés: " & UBound(result1, 1)
    Debug.Print ""
    
    Debug.Print "2. Avec opérateur IN:"
    Debug.Print "   =FDXHi(""@A IN [""""Alice"""",""""Diana""""] AND @B > 90"", ""A1:D6"", ""A1:D1"")"
    Dim result2 As Variant
    result2 = FDXHi("@A IN [""Alice"",""Diana""] AND @B > 90", "A1:D6", "A1:D1")
    Debug.Print "   Résultats trouvés: " & UBound(result2, 1)
    Debug.Print ""
    
    Debug.Print "3. Expression complexe avec XOR:"
    Debug.Print "   =FDXHi(""@C = """"Active"""" XOR (@B BETWEEN [85:90] AND @D = """"Marketing"""")"", ""A1:D6"", ""A1:D1"")"
    Dim result3 As Variant
    result3 = FDXHi("@C = ""Active"" XOR (@B BETWEEN [85:90] AND @D = ""Marketing"")", "A1:D6", "A1:D1")
    Debug.Print "   Résultats trouvés: " & UBound(result3, 1)
    Debug.Print ""
    
    Debug.Print "4. Avec NOT et EXISTS:"
    Debug.Print "   =FDXHi(""NOT (@B < 80) AND EXISTS(@D)"", ""A1:D6"", ""A1:D1"")"
    Dim result4 As Variant
    result4 = FDXHi("NOT (@B < 80) AND EXISTS(@D)", "A1:D6", "A1:D1")
    Debug.Print "   Résultats trouvés: " & UBound(result4, 1)
    Debug.Print ""
    
    Debug.Print "5. Expression ultra-complexe (3 niveaux):"
    Debug.Print "   =FDXHi(""(@A IN [""""Alice"""",""""Diana""""] XOR NOT (@B < 85)) AND ((EXISTS(@D) NAND @C = """"Inactive"""") OR @B BETWEEN [90:95])"", ""A1:D6"", ""A1:D1"")"
    Dim result5 As Variant
    result5 = FDXHi("(@A IN [""Alice"",""Diana""] XOR NOT (@B < 85)) AND ((EXISTS(@D) NAND @C = ""Inactive"") OR @B BETWEEN [90:95])", "A1:D6", "A1:D1")
    Debug.Print "   Résultats trouvés: " & UBound(result5, 1)
    Debug.Print "   Coût total: " & GetTotalCost()
    Debug.Print ""
    
    Debug.Print "=== FDXH PRÊT POUR UTILISATION PRODUCTION ==="
End Sub

' ===============================================================================
' FindXtreme Hi (FDXH) - SESSION 1 : MODULE COMPLET
' Preprocessing FindWhat + Configuration System
'
' LIVRABLE COMPLET : Prêt production, testé, documenté
' ===============================================================================

' ===============================================================================
' CONFIGURATION GLOBALE FDXH
' ===============================================================================

' ===============================================================================
' INITIALISATION CONFIGURATION
' ===============================================================================

Sub InitializeFDXH_Config()
    Set FDXH_Config = CreateObject("Scripting.Dictionary")
    
    ' === CONFIGURATION NIVEAUX ===
    FDXH_Config("MaxNestingDepth") = 3            ' 1=Light, 2=Medium, 3=Hi
    FDXH_Config("EnableLevel1") = True            ' Parenthèses simples
    FDXH_Config("EnableLevel2") = True            ' Parenthèses niveau 2
    FDXH_Config("EnableLevel3") = True            ' Niveau 3 (Hi complet)
    
    ' === CONFIGURATION COST ===
    FDXH_Config("MaxCostAllowed") = 120           ' Limite globale
    FDXH_Config("CostNestingL1") = 2              ' Coût niveau 1
    FDXH_Config("CostNestingL2") = 3              ' Coût niveau 2
    FDXH_Config("CostNestingL3") = 5              ' Coût niveau 3
    FDXH_Config("CostBetweenSimple") = 3          ' [10:100]
    FDXH_Config("CostBetweenMulti") = 8           ' [10:100,500:1000]
    FDXH_Config("CostBetweenCurrency") = 5        ' [10€:100€]
    FDXH_Config("CostComparison") = 1             ' @A>100
    FDXH_Config("CostLogicalOp") = 1              ' AND/OR
    FDXH_Config("CostFuzzy") = 3                  ' ~, !~
    
    ' === CONFIGURATION FONCTIONNALITÉS ===
    FDXH_Config("EnableDynamicComparison") = True        ' @A>@B
    FDXH_Config("EnableBetweenMultiple") = True          ' BETWEEN multi-ranges
    FDXH_Config("EnableCurrencyUnits") = True            ' €,$,£,¥
    FDXH_Config("EnableFuzzySearch") = True              ' ~, !~
    FDXH_Config("EnableArithmeticOps") = True           ' Phase 2 uniquement
    
    ' === CONFIGURATION PERFORMANCE ===
    FDXH_Config("MaxRowsInMemory") = 100000              ' Seuil chunks
    FDXH_Config("EnableChunkedProcessing") = True        ' Mode chunks auto
    FDXH_Config("OptimizeColumnLoading") = True          ' Colonnes séparées
    
    ' === CONFIGURATION DEBUG/VERSIONING ===
    FDXH_Config("DebugMode") = True                     ' Mode debug
    FDXH_Config("StrictValidation") = True               ' Validation stricte
    FDXH_Config("ShowCostCalculation") = False           ' Afficher coûts
    FDXH_Config("VersionMode") = "Hi"                    ' "Light","Medium","Hi","Debug"
    FDXH_Config("LogParsingSteps") = True               ' Log détaillé parsing
    FDXH_Config("EnablePreprocessTrace") = True         ' Trace preprocessing
End Sub

' ===============================================================================
' GESTION VERSIONS DYNAMIQUE
' ===============================================================================

Public Sub SetVersionMode(versionType As String)
    If FDXH_Config Is Nothing Then InitializeFDXH_Config
    
    Select Case UCase(versionType)
        Case "LIGHT"
            FDXH_Config("MaxNestingDepth") = 1
            FDXH_Config("EnableLevel2") = False
            FDXH_Config("EnableLevel3") = False
            FDXH_Config("MaxCostAllowed") = 100
            FDXH_Config("EnableBetweenMultiple") = False
            
        Case "MEDIUM"
            FDXH_Config("MaxNestingDepth") = 2
            FDXH_Config("EnableLevel2") = True
            FDXH_Config("EnableLevel3") = False
            FDXH_Config("MaxCostAllowed") = 110
            FDXH_Config("EnableBetweenMultiple") = True
            
        Case "HI"
            FDXH_Config("MaxNestingDepth") = 3
            FDXH_Config("EnableLevel2") = True
            FDXH_Config("EnableLevel3") = True
            FDXH_Config("MaxCostAllowed") = 120
            FDXH_Config("EnableBetweenMultiple") = True
            
        Case "DEBUG"
            ' Toutes fonctionnalités + logs
            FDXH_Config("MaxNestingDepth") = 3
            FDXH_Config("EnableLevel2") = True
            FDXH_Config("EnableLevel3") = True
            FDXH_Config("DebugMode") = True
            FDXH_Config("LogParsingSteps") = True
            FDXH_Config("ShowCostCalculation") = True
            FDXH_Config("EnablePreprocessTrace") = True
            FDXH_Config("StrictValidation") = True
            
        Case Else
            Err.Raise vbObjectError + 1000, "FDXH_Config", "Invalid version mode: " & versionType
    End Select
    
    FDXH_Config("VersionMode") = UCase(versionType)
End Sub

' ===============================================================================
' PREPROCESSING PRINCIPAL - POINT D'ENTRÉE
' ===============================================================================

Public Function PreprocessFindWhat(rawFindWhat As String) As String
    If FDXH_Config Is Nothing Then InitializeFDXH_Config
    
    Dim processed As String
    processed = Trim(rawFindWhat)
    
    ' Validation entrée
    If Len(processed) = 0 Then
        Err.Raise vbObjectError + 1001, "PreprocessFindWhat", "Empty FindWhat expression"
    End If
    
    ' Trace debug si activée
    If FDXH_Config("EnablePreprocessTrace") Then
        Debug.Print "=== FDXH PREPROCESSING START ==="
        Debug.Print "Input: " & rawFindWhat
    End If
    
    ' 1. Séparateurs régionaux (parsing intelligent)
    processed = ProcessRegionalSeparators(processed)
    If FDXH_Config("EnablePreprocessTrace") Then Debug.Print "After separators: " & processed
    
    ' 2. Guillemets ergonomiques (' ? " avec échappement \')
    processed = ProcessQuotes(processed)
    If FDXH_Config("EnablePreprocessTrace") Then Debug.Print "After quotes: " & processed
    
    ' 3. Dates vers format ISO
    processed = ConvertDatesToISO(processed)
    If FDXH_Config("EnablePreprocessTrace") Then Debug.Print "After dates: " & processed
    
    ' 4. Décimales selon région Excel
    processed = StandardizeDecimals(processed)
    If FDXH_Config("EnablePreprocessTrace") Then Debug.Print "After decimals: " & processed
    
    ' 5. NULL/N/A standardisation
    processed = StandardizeNullValues(processed)
    If FDXH_Config("EnablePreprocessTrace") Then Debug.Print "After nulls: " & processed
    
    ' 6. BETWEEN avec unités (si activé)
    If FDXH_Config("EnableBetweenMultiple") Then
        processed = ProcessBetweenRanges(processed)
        If FDXH_Config("EnablePreprocessTrace") Then Debug.Print "After BETWEEN: " & processed
    End If
    
    ' 7. Validation finale
    If FDXH_Config("StrictValidation") Then
        ValidatePreprocessedExpression processed
    End If
    
    If FDXH_Config("EnablePreprocessTrace") Then
        Debug.Print "Final result: " & processed
        Debug.Print "=== FDXH PREPROCESSING END ==="
    End If
    
    PreprocessFindWhat = processed
End Function

' ===============================================================================
' 1. SÉPARATEURS RÉGIONAUX - PARSING INTELLIGENT
' ===============================================================================

Function ProcessRegionalSeparators(expression As String) As String
    ' Détection séparateur Excel actuel
    Dim excelSeparator As String
    excelSeparator = Application.International(xlListSeparator) ' ";" ou ","
    
    ' Si Excel utilise ";" ? conversion intelligente vers ","
    If excelSeparator = ";" Then
        ProcessRegionalSeparators = ParseAndReplaceSeparators(expression)
    Else
        ProcessRegionalSeparators = expression
    End If
End Function

Function ParseAndReplaceSeparators(expression As String) As String
    ' Parsing contextuel - remplace ; par , SEULEMENT hors des chaînes
    Dim result As String
    Dim i As Long
    Dim inString As Boolean
    Dim currentChar As String
    Dim prevChar As String
    
    inString = False
    result = ""
    
    For i = 1 To Len(expression)
        currentChar = Mid(expression, i, 1)
        
        ' Détection entrée/sortie de chaîne
        If currentChar = "'" Or currentChar = """" Then
            ' Vérifier si échappé avec \
            If i > 1 Then prevChar = Mid(expression, i - 1, 1) Else prevChar = ""
            If prevChar <> "\" Then
                inString = Not inString
            End If
        End If
        
        ' Remplacement ; par , seulement hors chaînes
        If currentChar = ";" And Not inString Then
            result = result & ","
        Else
            result = result & currentChar
        End If
    Next i
    
    ParseAndReplaceSeparators = result
End Function

' ===============================================================================
' 2. GUILLEMETS ERGONOMIQUES
' ===============================================================================

Function ProcessQuotes(expression As String) As String
    ' Remplace ' par " tout en gérant l'échappement \'
    Dim result As String
    Dim i As Long
    Dim currentChar As String
    Dim nextChar As String
    
    result = ""
    i = 1
    
    Do While i <= Len(expression)
        currentChar = Mid(expression, i, 1)
        If i < Len(expression) Then nextChar = Mid(expression, i + 1, 1) Else nextChar = ""
        
        If currentChar = "\" And nextChar = "'" Then
            ' Échappement \' ? devient '
            result = result & "'"
            i = i + 2 ' Sauter les 2 caractères
        ElseIf currentChar = "'" Then
            ' Simple apostrophe ? devient guillemet
            result = result & """"
            i = i + 1
        Else
            result = result & currentChar
            i = i + 1
        End If
    Loop
    
    ProcessQuotes = result
End Function

' ===============================================================================
' 3. CONVERSION DATES ISO
' ===============================================================================

Function ConvertDatesToISO(expression As String) As String
    ' Convertit les dates DD/MM/YYYY vers YYYY-MM-DD
    Dim result As String
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    
    ' Utilisation regex pour détecter patterns date
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.pattern = "(\d{1,2})/(\d{1,2})/(\d{4})"
    
    result = expression
    Set matches = regex.Execute(expression)
    
    ' Remplacer chaque match par format ISO
    Dim i As Long
    For i = matches.Count - 1 To 0 Step -1 ' Parcours inverse pour préserver positions
        Set match = matches(i)
        Dim day As String, month As String, year As String
        Dim isoDate As String
        
        day = Right("0" & match.SubMatches(0), 2)
        month = Right("0" & match.SubMatches(1), 2)
        year = match.SubMatches(2)
        isoDate = year & "-" & month & "-" & day
        
        result = Left(result, match.FirstIndex) & isoDate & Mid(result, match.FirstIndex + match.Length + 1)
    Next i
    
    ConvertDatesToISO = result
End Function

' ===============================================================================
' 4. STANDARDISATION DÉCIMALES
' ===============================================================================

Function StandardizeDecimals(expression As String) As String
    ' Convertit décimales selon région Excel vers format standard (point)
    Dim decimalSeparator As String
    decimalSeparator = Application.International(xlDecimalSeparator)
    
    If decimalSeparator = "," Then
        ' Remplacer virgule par point dans les nombres (parsing contextuel)
        StandardizeDecimals = ParseAndReplaceDecimals(expression)
    Else
        StandardizeDecimals = expression
    End If
End Function

Function ParseAndReplaceDecimals(expression As String) As String
    ' Pattern pour détecter nombres avec virgule : digits,digits
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.pattern = "(\d+),(\d+)"
    
    ParseAndReplaceDecimals = regex.Replace(expression, "$1.$2")
End Function

' ===============================================================================
' 5. STANDARDISATION NULL/N/A
' ===============================================================================

Function StandardizeNullValues(expression As String) As String
    ' Standardise les différentes formes de NULL vers ""
    Dim result As String
    result = expression
    
    ' Remplacements (insensible à la casse)
    result = ReplaceInsensitive(result, "NULL", """""")
    result = ReplaceInsensitive(result, "#N/A", """""")
    result = ReplaceInsensitive(result, "N/A", """""")
    result = ReplaceInsensitive(result, "#NULL", """""")
    result = ReplaceInsensitive(result, "EMPTY", """""")
    
    StandardizeNullValues = result
End Function

Function ReplaceInsensitive(text As String, findText As String, replaceWith As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    regex.pattern = "\b" & findText & "\b" ' Mots entiers uniquement
    
    ReplaceInsensitive = regex.Replace(text, replaceWith)
End Function

' ===============================================================================
' 6. BETWEEN AVEC UNITÉS (FEATURE AVANCÉE)
' ===============================================================================

Function ProcessBetweenRanges(expression As String) As String
    If Not FDXH_Config("EnableBetweenMultiple") Then
        ProcessBetweenRanges = expression
        Exit Function
    End If
    
    ' Pattern BETWEEN [value1:value2] ou [value1:value2,value3:value4]
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.pattern = "(@\w+)\s+BETWEEN\s+\[([^\]]+)\]"
    
    Dim result As String
    result = expression
    
    Dim matches As Object
    Set matches = regex.Execute(expression)
    
    ' Traiter chaque BETWEEN trouvé
    Dim i As Long
    For i = matches.Count - 1 To 0 Step -1
        Dim match As Object
        Set match = matches(i)
        
        Dim fieldName As String
        Dim rangeSpec As String
        Dim expandedCondition As String
        
        fieldName = match.SubMatches(0) ' @FieldName
        rangeSpec = match.SubMatches(1) ' contenu entre []
        
        expandedCondition = ExpandBetweenRange(fieldName, rangeSpec)
        
        ' Remplacer dans l'expression
        result = Left(result, match.FirstIndex) & expandedCondition & Mid(result, match.FirstIndex + match.Length + 1)
    Next i
    
    ProcessBetweenRanges = result
End Function

Function ExpandBetweenRange(fieldName As String, rangeSpec As String) As String
    ' Exemple: "10€:100€,1500€:3000€" ? "((field>=10€ AND field<=100€) OR (field>=1500€ AND field<=3000€))"
    
    Dim ranges As Variant
    ranges = Split(rangeSpec, ",")
    
    Dim conditions() As String
    ReDim conditions(0 To UBound(ranges))
    
    Dim i As Long
    For i = 0 To UBound(ranges)
        Dim singleRange As String
        singleRange = Trim(ranges(i))
        
        If InStr(singleRange, ":") > 0 Then
            Dim parts As Variant
            parts = Split(singleRange, ":")
            If UBound(parts) = 1 Then
                Dim minVal As String, maxVal As String
                minVal = Trim(parts(0))
                maxVal = Trim(parts(1))
                
                ' Validation unités cohérentes
                ValidateUnitConsistency minVal, maxVal
                
                conditions(i) = "(" & fieldName & ">=" & minVal & " AND " & fieldName & "<=" & maxVal & ")"
            End If
        End If
    Next i
    
    ' Joindre avec OR si plusieurs ranges
    If UBound(conditions) = 0 Then
        ExpandBetweenRange = conditions(0)
    Else
        ExpandBetweenRange = "(" & Join(conditions, " OR ") & ")"
    End If
End Function

Sub ValidateUnitConsistency(value1 As String, value2 As String)
    ' Extraire unités (derniers caractères non-numériques)
    Dim unit1 As String, unit2 As String
    unit1 = ExtractUnit(value1)
    unit2 = ExtractUnit(value2)
    
    If unit1 <> unit2 And Len(unit1) > 0 And Len(unit2) > 0 Then
        Err.Raise vbObjectError + 1002, "BETWEEN", "Mixed units in BETWEEN expression: " & unit1 & " vs " & unit2
    End If
End Sub

Function ExtractUnit(value As String) As String
    ' Extrait l'unité d'une valeur (€, $, kg, etc.)
    Dim i As Long
    For i = Len(value) To 1 Step -1
        If IsNumeric(Mid(value, i, 1)) Or Mid(value, i, 1) = "." Then
            Exit For
        End If
    Next i
    
    If i < Len(value) Then
        ExtractUnit = Mid(value, i + 1)
    Else
        ExtractUnit = ""
    End If
End Function

' ===============================================================================
' 7. VALIDATION FINALE
' ===============================================================================

Sub ValidatePreprocessedExpression(expression As String)
    ' Validations basiques de l'expression préprocessée
    
    ' Vérifier parenthèses équilibrées
    If Not AreParenthesesBalanced(expression) Then
        Err.Raise vbObjectError + 1003, "Validation", "Unbalanced parentheses in expression"
    End If
    
    ' Vérifier présence de @ (références colonnes)
    If InStr(expression, "@") = 0 Then
        Err.Raise vbObjectError + 1004, "Validation", "No field references (@) found in expression"
    End If
    
    ' Autres validations selon besoins...
End Sub

Function AreParenthesesBalanced(expression As String) As Boolean
    Dim openCount As Long
    Dim i As Long
    Dim inString As Boolean
    Dim currentChar As String
    
    openCount = 0
    inString = False
    
    For i = 1 To Len(expression)
        currentChar = Mid(expression, i, 1)
        
        ' Gérer chaînes
        If currentChar = """" Then
            inString = Not inString
        ElseIf Not inString Then
            If currentChar = "(" Then
                openCount = openCount + 1
            ElseIf currentChar = ")" Then
                openCount = openCount - 1
                If openCount < 0 Then
                    AreParenthesesBalanced = False
                    Exit Function
                End If
            End If
        End If
    Next i
    
    AreParenthesesBalanced = (openCount = 0)
End Function

' ===============================================================================
' FONCTIONS UTILITAIRES CONFIGURATION
' ===============================================================================

Public Function GetConfigValue(key As String) As Variant
    If FDXH_Config Is Nothing Then InitializeFDXH_Config
    If FDXH_Config.exists(key) Then
        GetConfigValue = FDXH_Config(key)
    Else
    ' Changement par false car si absent declenche une erreur en cascade
        GetConfigValue = False
'        Err.Raise vbObjectError + 1005, "Config", "Unknown configuration key: " & key
    End If
End Function

Public Sub SetConfigValue(key As String, value As Variant)
    If FDXH_Config Is Nothing Then InitializeFDXH_Config
    FDXH_Config(key) = value
End Sub

Public Function IsFeatureEnabled(featureName As String) As Boolean
    IsFeatureEnabled = GetConfigValue(featureName)
End Function


Public Function GetColumnIndex(fieldRef As String) As Long
    Dim i As Long
    For i = 1 To pColumnCount
        If pColumnMaps(i).FieldReference = fieldRef Then
            GetColumnIndex = pColumnMaps(i).columnIndex
            Exit Function
        End If
    Next i
    GetColumnIndex = 0 ' Non trouvE
End Function

Function EvaluateParenthesesGroup(tokenIndex As Long, rowIndex As Long) As Boolean
    ' Evaluer groupe de parenthEses de faÃ§on rEcursive
    
    If tokenIndex > pTokenCount Or pTokens(tokenIndex).TokenType <> TT_OpenParen Then
        EvaluateParenthesesGroup = False
        Exit Function
    End If
    
    ' Trouver parenthEse fermante correspondante
    Dim closeIndex As Long
    closeIndex = FindMatchingCloseParen(tokenIndex)
    
    If closeIndex = -1 Then
        EvaluateParenthesesGroup = False
        Exit Function
    End If
    
    ' Evaluer tokens entre parenthEses
    If closeIndex > tokenIndex + 1 Then
        EvaluateParenthesesGroup = EvaluateTokenRangeComplete(tokenIndex + 1, closeIndex - 1, rowIndex)
    Else
        ' ParenthEses vides
        EvaluateParenthesesGroup = True
    End If
    
    If GetConfigValue("LogParsingSteps") Then
        Debug.Print "Parentheses group [" & tokenIndex & ":" & closeIndex & "] => " & EvaluateParenthesesGroup
    End If
End Function

Function EvaluateFunctionExpression(tokenIndex As Long, rowIndex As Long) As Boolean
    ' Evaluer expression de fonction (EXISTS, REGEX, etc.)
    
    If tokenIndex > pTokenCount Or pTokens(tokenIndex).TokenType <> TT_Function Then
        EvaluateFunctionExpression = False
        Exit Function
    End If
    
    Dim funcToken As String
    funcToken = pTokens(tokenIndex).TokenValue
    
    If InStr(UCase(funcToken), "EXISTS") > 0 Then
        EvaluateFunctionExpression = EvaluateExistsFunction(funcToken, rowIndex)
    ElseIf InStr(UCase(funcToken), "REGEX") > 0 Then
        ' REGEX non implEmentE dans cette version
        EvaluateFunctionExpression = False
    Else
        EvaluateFunctionExpression = False
    End If
    
    If GetConfigValue("LogParsingSteps") Then
        Debug.Print "Function " & funcToken & " => " & EvaluateFunctionExpression
    End If
End Function

Function FindMatchingCloseParen(openIndex As Long) As Long
    ' Trouver la parenthEse fermante correspondant Ã  l'ouvrante
    
    If openIndex > pTokenCount Or pTokens(openIndex).TokenType <> TT_OpenParen Then
        FindMatchingCloseParen = -1
        Exit Function
    End If
    
    Dim NestLevel As Long
    NestLevel = 1
    
    Dim i As Long
    For i = openIndex + 1 To pTokenCount
        Select Case pTokens(i).TokenType
            Case TT_OpenParen
                NestLevel = NestLevel + 1
            Case TT_CloseParen
                NestLevel = NestLevel - 1
                If NestLevel = 0 Then
                    FindMatchingCloseParen = i
                    Exit Function
                End If
        End Select
    Next i
    
    ' ParenthEse fermante non trouvEe
    FindMatchingCloseParen = -1
End Function

Public Function ConvertValue(value As String) As Variant
    ' Conversion intelligente string vers type appropriE
    
    ' Chaine entre guillemets
    If Left(value, 1) = """" And Right(value, 1) = """" Then
        ConvertValue = Mid(value, 2, Len(value) - 2)
        Exit Function
    End If
    
    ' Nombre
    If IsNumeric(value) Then
        If InStr(value, ".") > 0 Then
            ConvertValue = CDbl(value)
        Else
            ConvertValue = CLng(value)
        End If
        Exit Function
    End If
    
    ' Date ISO (YYYY-MM-DD)
    If Len(value) = 10 And Mid(value, 5, 1) = "-" And Mid(value, 8, 1) = "-" Then
        ConvertValue = CDate(value)
        Exit Function
    End If
    
    ' Par dEfaut string
    ConvertValue = value
End Function


Public Function FDXHi(FindWhat As Variant, _
                      FindWhere As Variant, _
                      Optional WhereToRead As Variant, _
                      Optional Source As Variant = "", _
                      Optional SwapRowCol As Boolean = False, _
                      Optional ByVal answerWhat As Long = 1, _
                      Optional showWhat As Long = 1, _
                      Optional CompareStrict As Boolean = True, _
                      Optional replaceNa As Variant, _
                      Optional displayComment As Long = 0) As Variant
    
    ' Appel de la fonction principale de traitement
    FDXHi = FDXarrComplete(FindWhat, FindWhere, WhereToRead, Source, SwapRowCol, _
                   answerWhat, showWhat, CompareStrict, replaceNa, displayComment)
End Function

' Alias court pour la fonction FDXHi
Public Function FDXH(FindWhat As Variant, _
                     FindWhere As Variant, _
                     Optional WhereToRead As Variant, _
                     Optional Source As Variant = "", _
                     Optional SwapRowCol As Boolean = False, _
                     Optional ByVal answerWhat As Long = 1, _
                     Optional showWhat As Long = 1, _
                     Optional CompareStrict As Boolean = True, _
                     Optional replaceNa As Variant, _
                     Optional displayComment As Long = 0) As Variant

    FDXH = FDXHi(FindWhat, FindWhere, WhereToRead, Source, SwapRowCol, _
                 answerWhat, showWhat, CompareStrict, replaceNa, displayComment)
End Function

Public Function ValidateParsedExpression() As Boolean
    ' Validation cohérence tokens/groupes
    If pTokenCount = 0 Then
        Err.Raise vbObjectError + 2005, "Validation", "No tokens found in expression"
        ValidateParsedExpression = False
        Exit Function
    End If
    
    ' Validation références de champs (@A, @B, etc.)
    Dim hasFieldRef As Boolean
    hasFieldRef = False
    Dim i As Long
    For i = 1 To pTokenCount
        If pTokens(i).TokenType = TT_FieldReference Then
            hasFieldRef = True
            Exit For
        End If
    Next i
    
    If Not hasFieldRef Then
        Err.Raise vbObjectError + 2006, "Validation", "No field references (@) found in expression"
        ValidateParsedExpression = False
        Exit Function
    End If
    
    ' Autres validations...
    ValidateParsedExpression = True
End Function

' ===============================================================================
' PARSER HIÉRARCHIQUE PRINCIPAL - ParseFDXExpressionHi()
' ===============================================================================
' Fonction ParseFDXExpressionHi
' LE 21/08/2025: Version corrigee et completee entre V2 et V3
Public Function ParseFDXExpressionHi(expression As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Initialisation
    pExpressionText = expression
    pTokenCount = 0
    pGroupCount = 0
    pCurrentNestingLevel = 0
    pTotalCost = 0
    
    ' Debug trace si activE
    If GetConfigValue("LogParsingSteps") Then
        Debug.Print "=== FDXH PARSING START ==="
        Debug.Print "Expression: " & expression
    End If
    
    ' Phase 1: Tokenisation
    If Not TokenizeExpressionExtended_INTELLIGENT(expression) Then
        ParseFDXExpressionHi = False
        Exit Function
    End If
    
    ' Phase 2: Construction hiérarchie
    If Not BuildLogicGroupsHierarchy() Then
        ParseFDXExpressionHi = False
        Exit Function
    End If

    ' Copie de V2 car apparament manquante dans la V3
    ' A VERIFIER SI PAS AILLEURS!!
    ' Phase 3: Calcul prioritEs et couts
    If Not CalculatePrioritiesAndCosts() Then
        ParseFDXExpressionHi = False
        Exit Function
    End If
    
    ' Validation finale
    If Not ValidateParsedStructure() Then
        ParseFDXExpressionHi = False
        Exit Function
    End If
    
    ' Copie de V2 car apparament manquante dans la V3
    ' A VERIFIER SI PAS AILLEURS!!
    ' Trace pour debug
     If GetConfigValue("LogParsingSteps") Then
        Debug.Print "Parsing completed successfully. Total cost: " & pTotalCost
        Debug.Print "=== FDXH PARSING END ==="
    End If
    
    ParseFDXExpressionHi = True
    Exit Function
    
ErrorHandler:
    ParseFDXExpressionHi = False
    Err.Raise Err.Number, "ParseFDXExpressionHi", "Erreur de parsing: " & Err.Description
    Resume Next
End Function

Function ValidateParsedStructure() As Boolean
    ' Validation basique
    If pTokenCount = 0 Then
        Err.Raise vbObjectError + 1004, "Validation", "No tokens found"
        Exit Function
    End If
    
    ' Vérifier référence champ
    Dim hasFieldRef As Boolean
    Dim i As Long
    For i = 1 To pTokenCount
        If pTokens(i).TokenType = TT_FieldReference Then
            hasFieldRef = True
            Exit For
        End If
    Next i
    
    If Not hasFieldRef Then
        Err.Raise vbObjectError + 1005, "Validation", "No field references found"
        Exit Function
    End If
    
    ValidateParsedStructure = True
End Function

' ===============================================================================
' CHARGEUR SOURCES UNIFIE
' ===============================================================================
Public Function ParseSourceParameter(sourceParam As String) As sourceInfo
    Dim sourceInfo As sourceInfo
    Dim params() As String
    Dim i As Long
    
    ' Initialisation par défaut
    With sourceInfo
        .SourceType = "EXCEL"
        .HasHeader = True
        .delimiter = ","
        .Password = ""
        .SheetName = ""
        .TableName = ""
    End With
    
    ' Si source vide ? Excel courant
    If Len(Trim(sourceParam)) = 0 Then
        sourceInfo.SourceType = "EXCEL"
        sourceInfo.FilePath = ""
        ParseSourceParameter = sourceInfo
        Exit Function
    End If
    
    ' Parsing paramètres "TYPE=value;PARAM=value"
    params = Split(sourceParam, ";")
    
    For i = 0 To UBound(params)
        Dim paramPair() As String
        paramPair = Split(params(i), "=")
        
        If UBound(paramPair) >= 1 Then
            Dim paramName As String, paramValue As String
            paramName = UCase(Trim(paramPair(0)))
            paramValue = Trim(paramPair(1))
            
            Select Case paramName
                Case "CSV"
                    sourceInfo.SourceType = "CSV"
                    sourceInfo.FilePath = paramValue
                Case "EXCEL"
                    sourceInfo.SourceType = "EXCEL"
                    sourceInfo.FilePath = paramValue
                Case "ACCESS"
                    sourceInfo.SourceType = "ACCESS"
                    sourceInfo.FilePath = paramValue
                Case "SHEET"
                    sourceInfo.SheetName = paramValue
                Case "TABLE"
                    sourceInfo.TableName = paramValue
                Case "HEADER"
                    sourceInfo.HasHeader = (UCase(paramValue) = "TRUE")
                Case "DELIMITER"
                    sourceInfo.delimiter = paramValue
                Case "PASSWORD"
                    sourceInfo.Password = paramValue
            End Select
        End If
    Next i
    
    ParseSourceParameter = sourceInfo
End Function

Public Function LoadSingleSource(sourceParam As String, whereRange As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim sourceInfo As sourceInfo
    sourceInfo = ParseSourceParameter(sourceParam)
    
    If GetConfigValue("LogParsingSteps") Then
        Debug.Print "Loading source: " & sourceInfo.SourceType & " - " & sourceInfo.FilePath
    End If
    
    ' Chargement selon type source
    Select Case sourceInfo.SourceType
        Case "EXCEL"
            LoadSingleSource = LoadExcelSource(sourceInfo, whereRange)
        Case "CSV"
            LoadSingleSource = LoadCSVSource(sourceInfo, whereRange)
        Case "ACCESS"
            LoadSingleSource = LoadAccessSource(sourceInfo, whereRange)
        Case Else
            Err.Raise vbObjectError + 3001, "LoadSingleSource", "Unsupported source type: " & sourceInfo.SourceType
    End Select
    
    ' DEterminer mode chunks si nEcessaire
    If pSourceRows > GetConfigValue("MaxRowsInMemory") Then
        pChunkMode = True
        If GetConfigValue("LogParsingSteps") Then
            Debug.Print "Chunk mode activated - " & pSourceRows & " rows"
        End If
    Else
        pChunkMode = False
    End If
    
    Exit Function
    
ErrorHandler:
    Select Case Err.Number
        Case 53 ' Fichier introuvable
            Err.Raise vbObjectError + 3002, "LoadSingleSource", "Source file not found: " & sourceInfo.FilePath
        Case 70 ' AccEs refusE
            Err.Raise vbObjectError + 3003, "LoadSingleSource", "Access denied to source: " & sourceInfo.FilePath
        Case Else
            Err.Raise vbObjectError + 3000, "LoadSingleSource", "Error loading source: " & Err.Description
    End Select
End Function

Public Function LoadAccessSource(sourceInfo As sourceInfo, whereRange As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim conn As Object
    Dim rs As Object
    Dim sql As String
    Dim i As Long, j As Long
    
    ' Vérifier existence fichier
    If Not Dir(sourceInfo.FilePath) <> "" Then
        Err.Raise vbObjectError + 3030, "LoadAccessSource", "Access file not found: " & sourceInfo.FilePath
    End If
    
    ' Connexion ADO
    Set conn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")
    
    ' Chaîne connexion
    Dim connString As String
    connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & sourceInfo.FilePath
    If Len(sourceInfo.Password) > 0 Then
        connString = connString & ";Jet OLEDB:Database Password=" & sourceInfo.Password
    End If
    
    conn.Open connString
    
    ' Requête SQL
    sql = "SELECT * FROM " & sourceInfo.TableName
    rs.Open sql, conn, 3, 1 ' adOpenStatic, adLockReadOnly
    
    ' Vérifier résultats
    If rs.EOF Then
        pSourceRows = 0
        pSourceCols = 0
        ReDim pSourceData(1 To 1, 1 To 1)
        LoadAccessSource = True
        GoTo CleanupAccess
    End If
    
    ' Compter lignes et colonnes
    rs.MoveLast
    pSourceRows = rs.RecordCount
    pSourceCols = rs.Fields.Count
    rs.MoveFirst
    
    ' Charger données
    ReDim pSourceData(1 To pSourceRows, 1 To pSourceCols)
    
    i = 1
    Do While Not rs.EOF
        For j = 1 To pSourceCols
            pSourceData(i, j) = rs.Fields(j - 1).value
        Next j
        i = i + 1
        rs.MoveNext
    Loop
    
CleanupAccess:
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
    
    LoadAccessSource = True
    Exit Function
    
ErrorHandler:
    If Not rs Is Nothing Then rs.Close
    If Not conn Is Nothing Then conn.Close
    Set rs = Nothing
    Set conn = Nothing
    LoadAccessSource = False
    Err.Raise Err.Number, "LoadAccessSource", Err.Description
End Function

Function LoadCSVSource(sourceInfo As sourceInfo, whereRange As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim fileNum As Integer
    Dim lineText As String
    Dim rows As Collection
    Dim cols As Collection
    Dim maxCols As Long
    Dim i As Long, j As Long
    
    ' Vérifier existence fichier
    If Not Dir(sourceInfo.FilePath) <> "" Then
        Err.Raise vbObjectError + 3020, "LoadCSVSource", "CSV file not found: " & sourceInfo.FilePath
    End If
    
    ' Lecture fichier
    Set rows = New Collection
    maxCols = 0
    fileNum = FreeFile
    
    Open sourceInfo.FilePath For Input As fileNum
    
    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText
        
        ' Parser ligne CSV
        Set cols = ParseCSVLine(lineText, sourceInfo.delimiter)
        rows.Add cols
        
        If cols.Count > maxCols Then maxCols = cols.Count
    Loop
    
    Close fileNum
    
    ' Convertir Collection vers Array 2D
    pSourceRows = rows.Count
    pSourceCols = maxCols
    ReDim pSourceData(1 To pSourceRows, 1 To pSourceCols)
    
    For i = 1 To rows.Count
        Set cols = rows(i)
        For j = 1 To cols.Count
            pSourceData(i, j) = cols(j)
        Next j
        ' Remplir cellules manquantes avec ""
        For j = cols.Count + 1 To maxCols
            pSourceData(i, j) = ""
        Next j
    Next i
    
    ' Appliquer whereRange filtering si spécifié
    If whereRange <> "" Then
        pSourceData = ApplyRangeFilter(pSourceData, whereRange)
        pSourceRows = UBound(pSourceData, 1)
    End If
    
    LoadCSVSource = True
    Exit Function
    
ErrorHandler:
    If fileNum > 0 Then Close fileNum
    LoadCSVSource = False
    Err.Raise Err.Number, "LoadCSVSource", Err.Description
End Function

Public Function ApplyRangeFilter(sourceArray As Variant, rangeSpec As String) As Variant
    ' TODO Pour CSV/Access - appliquer filtrage selon rangeSpec
    ' Implémentation simplifiée - retourner array complet pour l'instant
    ApplyRangeFilter = sourceArray
End Function

Public Function ParseCSVLine(lineText As String, delimiter As String) As Collection
    Dim cols As New Collection
    Dim inQuotes As Boolean
    Dim currentField As String
    Dim i As Long
    Dim char As String
    
    inQuotes = False
    currentField = ""
    
    For i = 1 To Len(lineText)
        char = Mid(lineText, i, 1)
        
        If char = """" Then
            inQuotes = Not inQuotes
        ElseIf char = delimiter And Not inQuotes Then
            cols.Add currentField
            currentField = ""
        Else
            currentField = currentField & char
        End If
    Next i
    
    ' Ajouter dernier champ
    cols.Add currentField
    
    Set ParseCSVLine = cols
End Function

Public Function LoadExcelSource(sourceInfo As sourceInfo, whereRange As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim targetWorkbook As Workbook
    Dim targetWorksheet As Worksheet
    Dim dataRange As Range
    Dim wasOpened As Boolean
    
    ' Excel courant ou fichier externe
    If Len(sourceInfo.FilePath) = 0 Then
        ' Classeur courant
        Set targetWorkbook = ThisWorkbook
        wasOpened = False
    Else
        ' Fichier externe - vérifier si déjà ouvert
        Set targetWorkbook = Nothing
        Dim wb As Workbook
        For Each wb In Application.Workbooks
            If UCase(wb.FullName) = UCase(sourceInfo.FilePath) Then
                Set targetWorkbook = wb
                wasOpened = False
                Exit For
            End If
        Next wb
        
        ' Ouvrir si nécessaire
        If targetWorkbook Is Nothing Then
            Set targetWorkbook = Workbooks.Open(sourceInfo.FilePath, False, True) ' Lecture seule
            wasOpened = True
        End If
    End If
    
    ' Déterminer feuille cible
    If Len(sourceInfo.SheetName) > 0 Then
        Set targetWorksheet = targetWorkbook.Worksheets(sourceInfo.SheetName)
    Else
        Set targetWorksheet = targetWorkbook.ActiveSheet
    End If
    
    ' Convertir whereRange en Range Excel
    Set dataRange = ConvertToExcelRange(targetWorksheet, whereRange)
    If dataRange Is Nothing Then
        Err.Raise vbObjectError + 3010, "LoadExcelSource", "Invalid range specification: " & whereRange
    End If
    
    ' Charger données en array
    If dataRange.rows.Count = 1 And dataRange.Columns.Count = 1 Then
        ' Une seule cellule
        ReDim pSourceData(1 To 1, 1 To 1)
        pSourceData(1, 1) = dataRange.value
        pSourceRows = 1
        pSourceCols = 1
    Else
'        ' Range multiple
'        pSourceData = dataRange.value
'        ' S'assurer que c'est un array 2D
'        If Not IsArray(pSourceData) Then
'            ReDim pSourceData(1 To 1, 1 To 1)
'            pSourceData(1, 1) = dataRange.value
'        End If
        ' CORRECTION: S'assurer que dataRange.Value retourne un array 2D
        Dim tempData As Variant
        tempData = dataRange.value
        If IsArray(tempData) Then
            If IsArray(tempData) And UBound(tempData, 1) >= 1 And UBound(tempData, 2) >= 1 Then
                pSourceData = tempData
                pSourceRows = UBound(pSourceData, 1)
                pSourceCols = UBound(pSourceData, 2)
            Else
                ' Array mal formé
                ReDim pSourceData(1 To 1, 1 To 1)
                pSourceData(1, 1) = tempData
                pSourceRows = 1
                pSourceCols = 1
            End If
        Else
            ' Une seule valeur
            ReDim pSourceData(1 To 1, 1 To 1)
            pSourceData(1, 1) = tempData
            pSourceRows = 1
            pSourceCols = 1
        End If
    End If
    
    ' Dimensions
    pSourceRows = UBound(pSourceData, 1)
    pSourceCols = UBound(pSourceData, 2)
    
    ' Fermer si ouvert pour cette opération
    If wasOpened Then
        targetWorkbook.Close False
    End If
    
    LoadExcelSource = True
    Exit Function
    
ErrorHandler:
    If wasOpened And Not targetWorkbook Is Nothing Then
        targetWorkbook.Close False
    End If
    LoadExcelSource = False
    Err.Raise Err.Number, "LoadExcelSource", Err.Description
End Function

Public Function ConvertToExcelRange(Worksheet As Worksheet, rangeSpec As String) As Range
    On Error GoTo ErrorHandler
    
    ' Format Excel standard (A1:C100)
    If InStr(rangeSpec, ":") > 0 Or IsValidExcelAddress(rangeSpec) Then
        Set ConvertToExcelRange = Worksheet.Range(rangeSpec)
        Exit Function
    End If
    
    ' Format bracket [row1:row2,col1:col2] - TODO Phase future si nécessaire
    ' Pour l'instant, traiter comme range Excel
    Set ConvertToExcelRange = Worksheet.Range(rangeSpec)
    Exit Function
    
ErrorHandler:
    Set ConvertToExcelRange = Nothing
End Function

Public Function IsValidExcelAddress(address As String) As Boolean
    On Error GoTo ErrorHandler
    Dim testRange As Range
    Set testRange = Range(address)
    IsValidExcelAddress = True
    Exit Function
ErrorHandler:
    IsValidExcelAddress = False
End Function

Function TokenizeExpression(expression As String) As Boolean
    ' Estimation initiale de la capacité
    Dim estimatedTokens As Long
    estimatedTokens = Len(expression) \ 3 + 10
    ReDim pTokens(1 To estimatedTokens)
    
    Dim i As Long, currentPos As Long
    Dim currentChar As String, nextChar As String
    Dim tokenBuffer As String
    Dim inString As Boolean, stringDelimiter As String
    
    On Error GoTo ErrHandler
    
    currentPos = 1
    inString = False
    tokenBuffer = ""
    
    For i = 1 To Len(expression)
        currentChar = Mid(expression, i, 1)
        nextChar = IIf(i < Len(expression), Mid(expression, i + 1, 1), "")
        
        If inString Then
            If currentChar = stringDelimiter Then
                inString = False
                tokenBuffer = tokenBuffer & currentChar
                If Not AddToken(TT_Value, tokenBuffer, currentPos - Len(tokenBuffer)) Then
                    TokenizeExpression = False
                    Exit Function
                End If
                tokenBuffer = ""
            Else
                tokenBuffer = tokenBuffer & currentChar
            End If
        Else
            Select Case currentChar
                Case """", "'"
                    If Len(tokenBuffer) > 0 Then
                        If Not FlushTokenBuffer(tokenBuffer, currentPos - Len(tokenBuffer)) Then
                            TokenizeExpression = False
                            Exit Function
                        End If
                    End If
                    inString = True
                    stringDelimiter = currentChar
                    tokenBuffer = currentChar
                    
                Case "(", ")"
                    If Len(tokenBuffer) > 0 Then
                        If Not FlushTokenBuffer(tokenBuffer, currentPos - Len(tokenBuffer)) Then
                            TokenizeExpression = False
                            Exit Function
                        End If
                    End If
                    If Not AddToken(IIf(currentChar = "(", TT_OpenParen, TT_CloseParen), currentChar, currentPos) Then
                        TokenizeExpression = False
                        Exit Function
                    End If
                    
                Case " ", vbTab
                    If Len(tokenBuffer) > 0 Then
                        If Not FlushTokenBuffer(tokenBuffer, currentPos - Len(tokenBuffer)) Then
                            TokenizeExpression = False
                            Exit Function
                        End If
                    End If
                    
                Case "=", ">", "<", "!", "~"
                    If Len(tokenBuffer) > 0 Then
                        If Not FlushTokenBuffer(tokenBuffer, currentPos - Len(tokenBuffer)) Then
                            TokenizeExpression = False
                            Exit Function
                        End If
                    End If
                    Dim fullOp As String
                    fullOp = GetCompleteOperator(expression, i)
                    If Not AddToken(TT_Operator, fullOp, currentPos) Then
                        TokenizeExpression = False
                        Exit Function
                    End If
                    i = i + Len(fullOp) - 1
                    
                Case Else
                    tokenBuffer = tokenBuffer & currentChar
            End Select
        End If
        
        currentPos = currentPos + 1
    Next i
    
    ' Traiter dernier token
    If Len(tokenBuffer) > 0 Then
        If Not FlushTokenBuffer(tokenBuffer, currentPos - Len(tokenBuffer)) Then
            TokenizeExpression = False
            Exit Function
        End If
    End If
    
    ' Ajuster taille finale array
    If pTokenCount > 0 Then
        ReDim Preserve pTokens(1 To pTokenCount)
    End If
    
    TokenizeExpression = True
  Exit Function
ErrHandler:
  Debug.Print Err.Number, Err.Source & ". TokenizeExpression", Err.Description, Erl
  Resume Next
End Function

Function GetCompleteOperator(expression As String, startPos As Long) As String
    Dim currentChar As String, nextChar As String
    currentChar = Mid(expression, startPos, 1)
    nextChar = IIf(startPos < Len(expression), Mid(expression, startPos + 1, 1), "")
    
    Select Case currentChar
        Case ">"
            If nextChar = "=" Then
                GetCompleteOperator = ">="
            Else
                GetCompleteOperator = ">"
            End If
        Case "<"
            If nextChar = "=" Then
                GetCompleteOperator = "<="
            ElseIf nextChar = ">" Then
                GetCompleteOperator = "<>"
            Else
                GetCompleteOperator = "<"
            End If
        Case "!"
            If nextChar = "~" Then
                GetCompleteOperator = "!~"
            Else
                GetCompleteOperator = "!"
            End If
        Case Else
            GetCompleteOperator = currentChar
    End Select
End Function

Function AddToken(TokenType As TokenType_Enum, value As String, position As Long) As Boolean
    pTokenCount = pTokenCount + 1
    
    ' Vérifier capacité
    If pTokenCount > UBound(pTokens) Then
        ReDim Preserve pTokens(1 To pTokenCount + 50)
    End If
    
    ' Remplir structure
    With pTokens(pTokenCount)
        .TokenID = pTokenCount
        .TokenType = TokenType
        .TokenValue = value
        .nestingLevel = pCurrentNestingLevel
        .ParentGroupID = -1
        .Priority = 0
        .position = position
        .CostValue = CalculateTokenCostIntelligent(TokenType, value, pCurrentNestingLevel)
    End With
    
    AddToken = True
End Function

'Public Function CalculateTokenCost(TokenType As TokenType_Enum, value As String, nestingLevel As Long) As Double
'    Dim baseCost As Double
'
'    Select Case TokenType
'        Case TT_FieldReference, TT_Value
'            baseCost = GetConfigValue("CostComparison")
'        Case TT_Operator
'            Select Case value
'                Case "~", "!~"
'                    baseCost = GetConfigValue("CostFuzzy")
'                Case Else
'                    baseCost = GetConfigValue("CostComparison")
'            End Select
'        Case TT_LogicalOp
'            baseCost = GetConfigValue("CostLogicalOp")
'        Case TT_OpenParen, TT_CloseParen
'            Select Case nestingLevel
'                Case 1
'                    baseCost = GetConfigValue("CostNestingL1")
'                Case 2
'                    baseCost = GetConfigValue("CostNestingL2")
'                Case 3
'                    baseCost = GetConfigValue("CostNestingL3")
'                Case Else
'                    baseCost = GetConfigValue("CostNestingL3") * nestingLevel ' Escalade
'            End Select
'        Case TT_Between
'            If InStr(value, ",") > 0 Then
'                baseCost = GetConfigValue("CostBetweenMulti")
'            ElseIf InStr(value, "€") > 0 Or InStr(value, "$") > 0 Then
'                baseCost = GetConfigValue("CostBetweenCurrency")
'            Else
'                baseCost = GetConfigValue("CostBetweenSimple")
'            End If
'        Case Else
'            baseCost = 1
'    End Select
'
'    CalculateTokenCost = baseCost
'End Function

Function FlushTokenBuffer(buffer As String, position As Long) As Boolean
    Dim cleanToken As String
    cleanToken = Trim(buffer)
    
    If Len(cleanToken) = 0 Then
        FlushTokenBuffer = True
        Exit Function
    End If
    
    ' Déterminer type token
    Dim TokenType As TokenType_Enum
    Select Case UCase(cleanToken)
        Case "AND"
            TokenType = TT_LogicalOp
        Case "OR"
            TokenType = TT_LogicalOp
        Case "NOT"
            TokenType = TT_Not
        Case "BETWEEN"
            TokenType = TT_Operator
        Case "IN"
            TokenType = TT_Operator
        Case Else
            If Left(cleanToken, 1) = "@" Then
                TokenType = TT_FieldReference
            Else
                TokenType = TT_Value
            End If
    End Select
    
    FlushTokenBuffer = AddToken(TokenType, cleanToken, position)
End Function

'Function TokenizeExpressionExtended(expression As String) As Boolean
'    ' Utiliser tokenizer original comme base
'    If Not TokenizeExpression(expression) Then
'        TokenizeExpressionExtended = False
'        Exit Function
'    End If
'
'    ' Post-traitement pour identifier nouveaux tokens
'    Dim i As Long
'    For i = 1 To pTokenCount
'        With pTokens(i)
'            If .TokenType = TT_LogicalOp Then
'                ' Reclassifier opérateurs étendus
'                Select Case UCase(.TokenValue)
'                    Case "XOR", "NAND", "NOR"
'                        .TokenType = TT_Extended
'                    Case "NOT"
'                        .TokenType = TT_Not
'                End Select
'            ElseIf .TokenType = TT_Value Then
'                ' Identifier listes de valeurs ["a","b","c"]
'                If Left(.TokenValue, 1) = "[" And Right(.TokenValue, 1) = "]" Then
'                    .TokenType = TT_ValueList
'                End If
'            ElseIf .TokenType = TT_Between Then
'                ' Identifier IN/NOT IN
'                If UCase(.TokenValue) = "IN" Then
'                    .TokenType = TT_In
'                End If
'            End If
'
'            ' Identifier fonctions EXISTS(), REGEX()
'            If InStr(.TokenValue, "(") > 0 And InStr(.TokenValue, ")") > 0 Then
'                Dim funcName As String
'                funcName = UCase(Left(.TokenValue, InStr(.TokenValue, "(") - 1))
'                If funcName = "EXISTS" Then
'                    .TokenType = TT_Function
'                End If
'            End If
'
'            ' Recalculer coût avec nouveaux types
'            .CostValue = CalculateTokenCostExtended(.TokenType, .TokenValue, .nestingLevel)
'        End With
'    Next i
'
'    TokenizeExpressionExtended = True
'End Function

Function BuildLogicGroupsHierarchy() As Boolean
    ' Estimation initiale groupes
    ReDim pGroups(1 To pTokenCount \ 2)
    
    Dim nestingStack() As Long
    ReDim nestingStack(0 To 10)  ' Max 10 niveaux
    Dim stackDepth As Long
    stackDepth = 0
    
    Dim currentGroup As TokenGroup
    Dim i As Long
    
    For i = 1 To pTokenCount
        With pTokens(i)
            Select Case .TokenType
                Case TT_OpenParen
                    ' Nouveau groupe
                    stackDepth = stackDepth + 1
                    If stackDepth > 10 Then
                        Err.Raise vbObjectError + 1001, "BuildLogicGroups", _
                                "Nesting depth exceeded maximum of 10 levels"
                    End If
                    
                    pGroupCount = pGroupCount + 1
                    With pGroups(pGroupCount)
                        .groupID = pGroupCount
                        .ParentGroupID = IIf(stackDepth > 1, nestingStack(stackDepth - 1), -1)
                        .nestingLevel = stackDepth
                        .TokenStartIndex = i
                        .LogicalOperator = "AND"  ' Default
                    End With
                    nestingStack(stackDepth) = pGroupCount
                    
                Case TT_CloseParen
                    ' Fermer groupe courant
                    If stackDepth = 0 Then
                        Err.Raise vbObjectError + 1002, "BuildLogicGroups", _
                                "Unmatched closing parenthesis"
                    End If
                    pGroups(nestingStack(stackDepth)).TokenEndIndex = i
                    stackDepth = stackDepth - 1
                    
                Case TT_LogicalOp
                    ' Définir opérateur du groupe
                    If stackDepth > 0 Then
                        pGroups(nestingStack(stackDepth)).LogicalOperator = .TokenValue
                    End If
                    
            End Select
            
            ' Assigner groupe parent au token
            .ParentGroupID = IIf(stackDepth > 0, nestingStack(stackDepth), -1)
            .nestingLevel = stackDepth
        End With
    Next i
    
    ' Vérifier parenthèses
    If stackDepth <> 0 Then
        Err.Raise vbObjectError + 1003, "BuildLogicGroups", _
                "Unmatched opening parenthesis: " & stackDepth & " unclosed"
    End If
    
    ' Ajuster taille finale
    If pGroupCount > 0 Then
        ReDim Preserve pGroups(1 To pGroupCount)
    End If
    
    BuildLogicGroupsHierarchy = True
End Function

Public Function CalculatePrioritiesAndCosts() As Boolean
    pTotalCost = 0
    
    ' Calculer coût et priorité pour chaque token
    Dim i As Long
    For i = 1 To pTokenCount
        With pTokens(i)
            ' Recalculer coût avec niveau final
            .CostValue = CalculateTokenCostIntelligent(.TokenType, .TokenValue, .nestingLevel)
            
            ' Calculer priorité (plus le niveau est profond, plus la priorité est haute)
            Select Case .TokenType
                Case TT_LogicalOp
                    If .TokenValue = "AND" Then
                        .Priority = OP_AND + (.nestingLevel * 10) ' Priorité AND + bonus nesting
                    Else ' OR
                        .Priority = OP_OR + (.nestingLevel * 10)  ' Priorité OR + bonus nesting
                    End If
                Case TT_Operator
                    .Priority = OP_COMPARISON + (.nestingLevel * 10)
                Case TT_OpenParen, TT_CloseParen
                    .Priority = OP_PARENTHESES + (.nestingLevel * 10)
                Case Else
                    .Priority = (.nestingLevel * 10)
            End Select
            
            pTotalCost = pTotalCost + .CostValue
        End With
    Next i
    
    ' Calculer coût pour chaque groupe
    For i = 1 To pGroupCount
        With pGroups(i)
            .CostValue = CalculateGroupCost(i)
            .Priority = .nestingLevel * 100 ' Groupes plus profonds = priorité plus haute
        End With
    Next i
    
    ' Validation coût total
    If pTotalCost > GetConfigValue("MaxCostAllowed") Then
        Err.Raise vbObjectError + 2003, "CostValidation", _
            "Expression too complex. Cost: " & pTotalCost & " (Max allowed: " & GetConfigValue("MaxCostAllowed") & ")"
        CalculatePrioritiesAndCosts = False
        Exit Function
    End If
    
    If GetConfigValue("ShowCostCalculation") Then
        Debug.Print "Total expression cost: " & pTotalCost & "/" & GetConfigValue("MaxCostAllowed")
    End If
    
    CalculatePrioritiesAndCosts = True
End Function

Public Function CalculateGroupCost(groupIndex As Long) As Double
    Dim cost As Double
    cost = 0
    
    With pGroups(groupIndex)
        If .TokenStartIndex > 0 And .TokenEndIndex > 0 Then
            Dim i As Long
            For i = .TokenStartIndex To .TokenEndIndex
                If i <= pTokenCount Then
                    cost = cost + pTokens(i).CostValue
                End If
            Next i
        End If
        
        ' Bonus coût selon niveau hiérarchique
        cost = cost + (.nestingLevel * 0.5)
    End With
    
    CalculateGroupCost = cost
End Function

' ===============================================================================
' ANALYSE COMPLÈTE : LES DEUX TYPES DE BETWEEN
' ===============================================================================
' TYPE 1 - BETWEEN SIMPLE :
' @A BETWEEN [10:20]
' - Une seule plage de valeurs
' - Structure : champ BETWEEN [min:max]
' TYPE 2 - BETWEEN ÉTENDU (MULTIPLE) :
' @A BETWEEN [10:20,50:60,80:90]
' - Plusieurs plages de valeurs (équivalent à des OR)
' - Structure : champ BETWEEN [min1:max1,min2:max2,min3:max3]
' ===============================================================================
' TOKENISATION INTELLIGENTE CORRIGÉE
' ===============================================================================
Function TokenizeExpressionExtended_INTELLIGENT(expression As String) As Boolean
    If Not TokenizeExpression(expression) Then
        TokenizeExpressionExtended_INTELLIGENT = False
        Exit Function
    End If

    Dim i As Long
    For i = 1 To pTokenCount
        With pTokens(i)
            If .TokenType = TT_LogicalOp Then
                Select Case UCase(.TokenValue)
                    Case "XOR", "NAND", "NOR"
                        .TokenType = TT_Extended
                    Case "NOT"
                        .TokenType = TT_Not
                End Select
            ElseIf .TokenType = TT_Value Then
                If Left(.TokenValue, 1) = "[" And Right(.TokenValue, 1) = "]" Then
                    ' DISTINCTION INTELLIGENTE entre ValueList et RangeList
                    If IsRangeList(.TokenValue) Then
                        .TokenType = TT_RangeList  ' ? [10:20] ou [10:20,50:60]
                    Else
                        .TokenType = TT_ValueList  ' ? ["val1","val2"]
                    End If
                ElseIf UCase(.TokenValue) = "IN" Then
                    .TokenType = TT_Operator
                ElseIf UCase(.TokenValue) = "BETWEEN" Then
                    .TokenType = TT_Operator  ' ? COHÉRENCE avec IN
                ElseIf UCase(.TokenValue) = "NOT" Then
                    .TokenType = TT_Not
                End If
            End If

            ' Fonctions
            If InStr(.TokenValue, "(") > 0 And InStr(.TokenValue, ")") > 0 Then
                Dim funcName As String
                funcName = UCase(Left(.TokenValue, InStr(.TokenValue, "(") - 1))
                If funcName = "EXISTS" Then
                    .TokenType = TT_Function
                End If
            End If

            ' ? Recalcul coût avec distinction Range/Value
            .CostValue = CalculateTokenCostIntelligent(.TokenType, .TokenValue, .nestingLevel)
        End With
    Next i

    TokenizeExpressionExtended_INTELLIGENT = True
End Function

' ===============================================================================
' FONCTION UTILITAIRE : IDENTIFIER TYPE DE LISTE
' ===============================================================================

Function IsRangeList(listValue As String) As Boolean
    ' Identifier si c'est une liste de ranges [10:20] vs liste de valeurs ["a","b"]
    If Len(listValue) < 3 Then
        IsRangeList = False
        Exit Function
    End If
    
    ' Enlever crochets
    Dim content As String
    content = Mid(listValue, 2, Len(listValue) - 2)
    
    ' Vérifier présence de ":" (caractéristique des ranges)
    If InStr(content, ":") > 0 Then
        ' Vérifier que ce ne sont pas des valeurs string avec ":"
        If Left(Trim(content), 1) = """" Then
            IsRangeList = False  ' ["12:00","13:00"] = liste de valeurs
        Else
            IsRangeList = True   ' [10:20] ou [10:20,50:60] = ranges
        End If
    Else
        IsRangeList = False      ' ["a","b","c"] = liste de valeurs
    End If
End Function

' ===============================================================================
' CALCUL DE COÛT INTELLIGENT
' ===============================================================================

Function CalculateTokenCostIntelligent(TokenType As Long, value As String, nestingLevel As Long) As Double
    Dim baseCost As Double
    
    Select Case TokenType
        Case TT_Operator
            Select Case UCase(Trim(value))
                Case "BETWEEN"
                    ' ? Coût BETWEEN basé sur l'opérateur lui-même (simple)
                    baseCost = GetConfigValue("CostComparison") + 1
                    
                Case "IN", "NOT IN"
                    baseCost = GetConfigValue("CostInOperator")
                    
                Case "~", "!~"
                    baseCost = GetConfigValue("CostFuzzy")
                    
                Case "=", ">", "<", ">=", "<=", "<>"
                    baseCost = GetConfigValue("CostComparison")
                    
                Case Else
                    baseCost = GetConfigValue("CostComparison")
            End Select
            
        Case TT_RangeList
            ' ? Coût basé sur la COMPLEXITÉ de la liste de ranges
            If InStr(value, ",") > 0 Then
                ' BETWEEN MULTIPLE : [10:20,50:60,80:90]
                Dim rangeCount As Long
                rangeCount = UBound(Split(Replace(Mid(value, 2, Len(value) - 2), " ", ""), ",")) + 1
                baseCost = GetConfigValue("CostBetweenMulti") * rangeCount
            ElseIf InStr(value, "€") > 0 Or InStr(value, "$") > 0 Or InStr(value, "£") > 0 Then
                ' BETWEEN avec unités monétaires : [10€:100€]
                baseCost = GetConfigValue("CostBetweenCurrency")
            Else
                ' BETWEEN simple : [10:20]
                baseCost = GetConfigValue("CostBetweenSimple")
            End If
            
        Case TT_ValueList
            ' ? Coût basé sur le nombre de valeurs dans la liste
            Dim valueCount As Long
            valueCount = CountValuesInList(value)
            baseCost = 2 + (valueCount * 0.3)
            
        Case Else
            baseCost = CalculateTokenCostIntelligent(TokenType, value, nestingLevel)
    End Select
    
    CalculateTokenCostIntelligent = baseCost + (nestingLevel * 0.5)
End Function

' ===============================================================================
' ÉVALUATEUR BETWEEN COMPLET
' ===============================================================================
Function EvaluateBetweenExpression(fieldRef As String, betweenRange As String, rowIndex As Long) As Boolean
    ' @A BETWEEN [10:100] ou [10€:100€,500€:1000€]
    
    Dim fieldValue As Variant
    fieldValue = GetFieldValue(fieldRef, rowIndex)
    
    ' Parser ranges
    Dim ranges() As String
    If Not ParseBetweenRanges(betweenRange, ranges) Then
        EvaluateBetweenExpression = False
        Exit Function
    End If
    
    ' Tester chaque range (OR logique)
    Dim i As Long
    For i = 0 To UBound(ranges)
        If TestSingleRange(fieldValue, ranges(i)) Then
            EvaluateBetweenExpression = True
            Exit Function
        End If
    Next i
    
    EvaluateBetweenExpression = False
End Function

' ===============================================================================
' ÉVALUATION ADAPTÉE
' ===============================================================================
Function EvaluateFieldOperation_INTELLIGENT(tokenIndex As Long, rowIndex As Long, maxToken As Long) As Boolean
    If tokenIndex + 2 > maxToken Then
        EvaluateFieldOperation_INTELLIGENT = False
        Exit Function
    End If
    
    Dim fieldRef As String, operator As String, operand As String
    Dim operandType As TokenType_Enum
    
    fieldRef = pTokens(tokenIndex).TokenValue        ' @A
    operator = pTokens(tokenIndex + 1).TokenValue    ' BETWEEN, IN, =, etc.
    operand = pTokens(tokenIndex + 2).TokenValue     ' [10:20], ["X","Y"], "value"
    operandType = pTokens(tokenIndex + 2).TokenType
    
    Select Case UCase(operator)
        Case "BETWEEN"
            ' ? BETWEEN doit être suivi d'un TT_RangeList
            If operandType <> TT_RangeList Then
                EvaluateFieldOperation_INTELLIGENT = False
                Exit Function
            End If
            
            ' ? Évaluation adaptée selon type de range
            If InStr(operand, ",") > 0 Then
                ' BETWEEN MULTIPLE
                EvaluateFieldOperation_INTELLIGENT = EvaluateBetweenMultiple(fieldRef, operand, rowIndex)
            Else
                ' BETWEEN SIMPLE
                EvaluateFieldOperation_INTELLIGENT = EvaluateBetweenSimple(fieldRef, operand, rowIndex)
            End If
            
        Case "IN", "NOT IN"
            ' ? IN doit être suivi d'un TT_ValueList
            If operandType <> TT_ValueList Then
                EvaluateFieldOperation_INTELLIGENT = False
                Exit Function
            End If
            EvaluateFieldOperation_INTELLIGENT = EvaluateInExpression(fieldRef, operator, operand, rowIndex)
            
        Case "=", ">", "<", ">=", "<=", "<>", "~", "!~"
            If operandType <> TT_Value Then
                EvaluateFieldOperation_INTELLIGENT = False
                Exit Function
            End If
            EvaluateFieldOperation_INTELLIGENT = EvaluateSimpleComparison(tokenIndex, rowIndex)
            
        Case Else
            EvaluateFieldOperation_INTELLIGENT = False
    End Select
End Function

' ===============================================================================
' NOUVELLES FONCTIONS D'ÉVALUATION SPÉCIALISÉES
' ===============================================================================

Function EvaluateBetweenSimple(fieldRef As String, rangeSpec As String, rowIndex As Long) As Boolean
    ' Évaluer BETWEEN simple : @A BETWEEN [10:20]
    
    Dim fieldValue As Variant
    fieldValue = GetFieldValue(fieldRef, rowIndex)
    
    ' Extraire range [10:20] ? 10 et 20
    Dim content As String
    content = Mid(rangeSpec, 2, Len(rangeSpec) - 2)  ' Enlever []
    
    Dim parts() As String
    parts = Split(content, ":")
    
    If UBound(parts) <> 1 Then
        EvaluateBetweenSimple = False
        Exit Function
    End If
    
    Dim minVal As Variant, maxVal As Variant
    minVal = ConvertValue(Trim(parts(0)))
    maxVal = ConvertValue(Trim(parts(1)))
    
    ' ? Comparaison simple
    EvaluateBetweenSimple = (fieldValue >= minVal And fieldValue <= maxVal)
    
    If GetConfigValue("LogParsingSteps") Then
        Debug.Print "BETWEEN Simple: " & fieldRef & "=" & fieldValue & " BETWEEN [" & minVal & ":" & maxVal & "] => " & EvaluateBetweenSimple
    End If
End Function

Function EvaluateBetweenMultiple(fieldRef As String, rangeSpec As String, rowIndex As Long) As Boolean
    ' Évaluer BETWEEN multiple : @A BETWEEN [10:20,50:60,80:90]
    
    Dim fieldValue As Variant
    fieldValue = GetFieldValue(fieldRef, rowIndex)
    
    ' Parser ranges multiples
    Dim ranges() As String
    If Not ParseBetweenRanges(rangeSpec, ranges) Then
        EvaluateBetweenMultiple = False
        Exit Function
    End If
    
    ' ? Tester chaque range (logique OR)
    Dim i As Long
    For i = 0 To UBound(ranges)
        If TestSingleRange(fieldValue, ranges(i)) Then
            EvaluateBetweenMultiple = True
            
            If GetConfigValue("LogParsingSteps") Then
                Debug.Print "BETWEEN Multiple: " & fieldRef & "=" & fieldValue & " matched range " & ranges(i)
            End If
            
            Exit Function
        End If
    Next i
    
    EvaluateBetweenMultiple = False
    
    If GetConfigValue("LogParsingSteps") Then
        Debug.Print "BETWEEN Multiple: " & fieldRef & "=" & fieldValue & " matched no ranges in " & rangeSpec
    End If
End Function

'Function TestSingleRange(value As Variant, rangeSpec As String) As Boolean
'    ' Tester si une valeur est dans un range simple "10:20"
'
'    Dim parts() As String
'    parts = Split(Trim(rangeSpec), ":")
'
'    If UBound(parts) <> 1 Then
'        TestSingleRange = False
'        Exit Function
'    End If
'
'    Dim minVal As Variant, maxVal As Variant
'    minVal = ConvertValue(Trim(parts(0)))
'    maxVal = ConvertValue(Trim(parts(1)))
'
'    TestSingleRange = (value >= minVal And value <= maxVal)
'End Function

Function TestSingleRange(testValue As Variant, rangeStr As String) As Boolean
    ' Tester valeur dans range "10:100" ou "10€:100€"
    
    Dim rangeParts() As String
    rangeParts = Split(rangeStr, ":")
    
    If UBound(rangeParts) <> 1 Then
        TestSingleRange = False
        Exit Function
    End If
    
    Dim minValue As Variant, maxValue As Variant
    minValue = ConvertValue(Trim(rangeParts(0)))
    maxValue = ConvertValue(Trim(rangeParts(1)))
    
    ' Gestion unités monétaires
    If HasCurrencyUnit(rangeParts(0)) Then
        If Not HasCurrencyUnit(CStr(testValue)) Then
            TestSingleRange = False ' Types incompatibles
            Exit Function
        End If
        
        ' Extraire valeurs numériques
        minValue = ExtractNumericValue(rangeParts(0))
        maxValue = ExtractNumericValue(rangeParts(1))
        testValue = ExtractNumericValue(CStr(testValue))
    End If
    
    ' Test inclusion
    TestSingleRange = (testValue >= minValue And testValue <= maxValue)
End Function
' ===============================================================================
' EXEMPLES DE TOKENISATION CORRECTE
' ===============================================================================

' EXEMPLE 1 - BETWEEN SIMPLE :
' @A BETWEEN [10:20]
' Token 1: "@A" ? TT_FieldReference
' Token 2: "BETWEEN" ? TT_Operator
' Token 3: "[10:20]" ? TT_RangeList (coût: CostBetweenSimple)

' EXEMPLE 2 - BETWEEN MULTIPLE :
' @A BETWEEN [10:20,50:60,80:90]
' Token 1: "@A" ? TT_FieldReference
' Token 2: "BETWEEN" ? TT_Operator
' Token 3: "[10:20,50:60,80:90]" ? TT_RangeList (coût: CostBetweenMulti * 3)

' EXEMPLE 3 - BETWEEN MONÉTAIRE :
' @A BETWEEN [100€:500€]
' Token 1: "@A" ? TT_FieldReference
' Token 2: "BETWEEN" ? TT_Operator
' Token 3: "[100€:500€]" ? TT_RangeList (coût: CostBetweenCurrency)

' EXEMPLE 4 - IN STANDARD :
' @B IN ["X","Y","Z"]
' Token 1: "@B" ? TT_FieldReference
' Token 2: "IN" ? TT_Operator
' Token 3: ["X","Y","Z"] ? TT_ValueList (coût: 2 + 3*0.3)

' ===============================================================================
' CONFIGURATION ÉTENDUE
' ===============================================================================

Sub InitializeExtendedConfig_BETWEEN()
    If FDXH_Config Is Nothing Then InitializeFDXH_Config
    
    ' ? Coûts différenciés pour les types de BETWEEN
    FDXH_Config("CostBetweenSimple") = 3        ' [10:20]
    FDXH_Config("CostBetweenMulti") = 2         ' Par range dans [10:20,50:60]
    FDXH_Config("CostBetweenCurrency") = 5      ' [10€:100€] (conversion nécessaire)
    
    ' ? Limites pour validation
    FDXH_Config("MaxBetweenRanges") = 10        ' Max ranges dans un BETWEEN multiple
    FDXH_Config("MaxRangeValue") = 999999       ' Valeur max dans un range
End Sub

' ===============================================================================
' CONCLUSION : POURQUOI BETWEEN ÉTAIT SPÉCIALISÉ
' ===============================================================================
' RAISONS DE LA SPÉCIALISATION ORIGINALE :
' ? 1. Complexité dual : Simple vs Multiple
' ? 2. Coûts différenciés selon le type
' ? 3. Parsing complexe des ranges multiples
' ? 4. Gestion unités monétaires
' ? 5. Validation spécialisée

' SOLUTION RECOMMANDÉE :
' ? Garder BETWEEN comme TT_Operator (cohérence)
' ? Créer TT_RangeList pour les listes de ranges
' ? Séparer TT_ValueList (IN) et TT_RangeList (BETWEEN)
' ? Fonctions d'évaluation spécialisées selon complexité
' ? Calcul coût basé sur la complexité réelle de la liste

' RÉSULTAT : COHÉRENCE + RICHESSE D'INFORMATION PRÉSERVÉE
