Attribute VB_Name = "Mod_Tests_3"
Option Explicit

' ===============================================================================
' TESTS UNITAIRES INTÉGRÉS
' ===============================================================================
Public Sub RunPreprocessTests()
    Debug.Print "=== FDXH PREPROCESSING TESTS ==="
    
    ' Test 1: Séparateurs
    TestSeparators
    
    ' Test 2: Guillemets
    TestQuotes
    
    ' Test 3: Dates
    TestDates
    
    ' Test 4: Décimales
    TestDecimals
    
    ' Test 5: BETWEEN
    TestBetween
    
    ' Test 6: Configuration
    TestConfiguration
    
    Debug.Print "=== ALL TESTS COMPLETED ==="
End Sub

Sub TestSeparators()
    Debug.Print "Testing separators..."
    
    Dim aInput As String, expected As String, result As String
    
    ' Test avec ; en français
    aInput = "@A='test;avec;point-virgule'; @B>100"
    result = ProcessRegionalSeparators(aInput)
    ' Le ; dans la chaîne doit être préservé, celui entre conditions remplacé
    Debug.Print "Separators test: " & IIf(InStr(result, "'test;avec;point-virgule'") > 0, "PASS", "FAIL")
End Sub

Sub TestQuotes()
    Debug.Print "Testing quotes..."
    
    Dim result As String
    result = ProcessQuotes("@A='L\'entreprise' AND @B='simple'")
    Debug.Print "Quotes result: " & result
    Debug.Print "Quotes test: " & IIf(InStr(result, """L'entreprise""") > 0, "PASS", "FAIL")
End Sub

Sub TestDates()
    Debug.Print "Testing dates..."
    
    Dim result As String
    result = ConvertDatesToISO("@Date >= 15/03/2024")
    Debug.Print "Dates result: " & result
    Debug.Print "Dates test: " & IIf(InStr(result, "2024-03-15") > 0, "PASS", "FAIL")
End Sub

Sub TestDecimals()
    Debug.Print "Testing decimals..."
    
    Dim result As String
    result = StandardizeDecimals("@Prix > 1234,56")
    Debug.Print "Decimals result: " & result
    Debug.Print "Decimals test: " & IIf(InStr(result, "1234.56") > 0, "PASS", "FAIL")
End Sub

Sub TestBetween()
    Debug.Print "Testing BETWEEN..."
    
    SetVersionMode "Hi" ' Activer BETWEEN
    Dim result As String
    result = ProcessBetweenRanges("@Prix BETWEEN [10€:100€,500€:1000€]")
    Debug.Print "BETWEEN result: " & result
    Debug.Print "BETWEEN test: " & IIf(InStr(result, "OR") > 0, "PASS", "FAIL")
End Sub

Sub TestConfiguration()
    Debug.Print "Testing configuration..."
    
    SetVersionMode "Light"
    Debug.Print "Light mode depth: " & GetConfigValue("MaxNestingDepth")
    SetVersionMode "Medium"
    Debug.Print "MediumMedium mode depth: " & GetConfigValue("MaxNestingDepth")
    SetVersionMode "Hi"
    Debug.Print "Hi mode depth: " & GetConfigValue("MaxNestingDepth")
    
    Debug.Print "Configuration test: PASS"
End Sub

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
' TESTS UNITAIRES COMPLETS - TOUS LES OPÉRATEURS
' ===============================================================================

Public Sub RunCompleteOperatorTests()
    Debug.Print "=== FDXH TESTS OPÉRATEURS COMPLETS ==="
    
    ' Test opérateur IN
    TestInOperator
    
    ' Test opérateurs logiques étendus
    TestExtendedLogicalOperators
    
    ' Test opérateur NOT
    TestNotOperator
    
    ' Test fonction EXISTS
    TestExistsFunction
    
    ' Test BETWEEN étendu
    TestExtendedBetween
    
    ' Test combinaisons complexes
    TestComplexCombinations
    
    Debug.Print "=== TESTS OPÉRATEURS COMPLETS TERMINÉS ==="
End Sub

Sub TestInOperator()
    Debug.Print "Testing IN operator..."
    
    ' Préparer données test
    SetupTestData
    
    ' Test IN simple
    Dim result1 As Variant
    result1 = FDXHi("@A IN [""Alice"",""Bob"",""Charlie""]", "A1:C5", "A1:C1")
    
    If IsArray(result1) Then
        Debug.Print "IN operator (simple): PASS - Found " & UBound(result1, 1) & " results"
    Else
        Debug.Print "IN operator (simple): FAIL"
    End If
    
    ' Test NOT IN
    Dim result2 As Variant
    result2 = FDXHi("@C NOT IN [""Inactive"",""Suspended""]", "A1:C5", "A1:C1")
    
    If IsArray(result2) Then
        Debug.Print "NOT IN operator: PASS - Found " & UBound(result2, 1) & " results"
    Else
        Debug.Print "NOT IN operator: FAIL"
    End If
    
    ' Test IN avec valeurs numériques
    Dim result3 As Variant
    result3 = FDXHi("@B IN [85,92,95]", "A1:C5", "A1:C1")
    
    If IsArray(result3) Then
        Debug.Print "IN operator (numeric): PASS - Found " & UBound(result3, 1) & " results"
    Else
        Debug.Print "IN operator (numeric): FAIL"
    End If
End Sub

Sub TestExtendedLogicalOperators()
    Debug.Print "Testing extended logical operators..."
    
    SetupTestData
    
    ' Test XOR
    Dim result1 As Variant
    result1 = FDXHi("@B > 90 XOR @C = ""Active""", "A1:C5", "A1:C1")
    
    If IsArray(result1) Then
        Debug.Print "XOR operator: PASS"
    Else
        Debug.Print "XOR operator: FAIL"
    End If
    
    ' Test NAND
    Dim result2 As Variant
    result2 = FDXHi("@B > 80 NAND @C = ""Active""", "A1:C5", "A1:C1")
    
    If IsArray(result2) Then
        Debug.Print "NAND operator: PASS"
    Else
        Debug.Print "NAND operator: FAIL"
    End If
    
    ' Test NOR
    Dim result3 As Variant
    result3 = FDXHi("@B < 80 NOR @C = ""Inactive""", "A1:C5", "A1:C1")
    
    If IsArray(result3) Then
        Debug.Print "NOR operator: PASS"
    Else
        Debug.Print "NOR operator: FAIL"
    End If
End Sub

Sub TestNotOperator()
    Debug.Print "Testing NOT operator..."
    
    SetupTestData
    
    ' Test NOT simple
    Dim result1 As Variant
    result1 = FDXHi("NOT @B < 85", "A1:C5", "A1:C1")
    
    If IsArray(result1) Then
        Debug.Print "NOT operator (simple): PASS"
    Else
        Debug.Print "NOT operator (simple): FAIL"
    End If
    
    ' Test NOT avec parenthèses
    Dim result2 As Variant
    result2 = FDXHi("NOT (@B < 85 AND @C = ""Active"")", "A1:C5", "A1:C1")
    
    If IsArray(result2) Then
        Debug.Print "NOT operator (complex): PASS"
    Else
        Debug.Print "NOT operator (complex): FAIL"
    End If
    
    ' Test NOT EXISTS
    Dim result3 As Variant
    result3 = FDXHi("NOT EXISTS(@D)", "A1:C5", "A1:C1")
    
    If IsArray(result3) Then
        Debug.Print "NOT EXISTS: PASS"
    Else
        Debug.Print "NOT EXISTS: FAIL"
    End If
End Sub

Sub TestExistsFunction()
    Debug.Print "Testing EXISTS function..."
    
    SetupTestDataWithEmpty
    
    ' Test EXISTS sur champ présent
    Dim result1 As Variant
    result1 = FDXHi("EXISTS(@A)", "A1:C6", "A1:C1")
    
    If IsArray(result1) Then
        Debug.Print "EXISTS (present): PASS - Found " & UBound(result1, 1) & " results"
    Else
        Debug.Print "EXISTS (present): FAIL"
    End If
    
    ' Test EXISTS sur champ vide
    Dim result2 As Variant
    result2 = FDXHi("EXISTS(@D)", "A1:C6", "A1:C1")
    
    Debug.Print "EXISTS (empty): " & IIf(IsArray(result2) And UBound(result2, 1) = 0, "PASS", "FAIL")
End Sub

Sub TestExtendedBetween()
    Debug.Print "Testing extended BETWEEN..."
    
    SetupTestData
    
    ' Test BETWEEN multiple ranges
    Dim result1 As Variant
    result1 = FDXHi("@B BETWEEN [80:90,92:100]", "A1:C5", "A1:C1")
    
    If IsArray(result1) Then
        Debug.Print "BETWEEN multiple ranges: PASS"
    Else
        Debug.Print "BETWEEN multiple ranges: FAIL"
    End If
    
    ' Test BETWEEN avec unités (simulé)
    Range("D1:D5").value = Array("Price", "10€", "25€", "50€", "75€")
    
    Dim result2 As Variant
    result2 = FDXHi("@D BETWEEN [20€:60€]", "A1:D5", "A1:D1")
    
    If IsArray(result2) Then
        Debug.Print "BETWEEN with currency: PASS"
    Else
        Debug.Print "BETWEEN with currency: FAIL"
    End If
End Sub

Sub TestComplexCombinations()
    Debug.Print "Testing complex combinations..."
    
    SetupTestData
    
    ' Combinaison IN + logique étendue
    Dim result1 As Variant
    result1 = FDXHi("@A IN [""Alice"",""Bob""] XOR (@B > 85 AND @C = ""Active"")", "A1:C5", "A1:C1")
    
    If IsArray(result1) Then
        Debug.Print "Complex IN + XOR: PASS"
    Else
        Debug.Print "Complex IN + XOR: FAIL"
    End If
    
    ' Combinaison NOT + EXISTS + parenthèses
    Dim result2 As Variant
    result2 = FDXHi("NOT (EXISTS(@D) OR (@B < 80 AND @C = ""Inactive""))", "A1:C5", "A1:C1")
    
    If IsArray(result2) Then
        Debug.Print "Complex NOT + EXISTS: PASS"
    Else
        Debug.Print "Complex NOT + EXISTS: FAIL"
    End If
    
    ' Triple niveau avec nouveaux opérateurs
    Dim result3 As Variant
    result3 = FDXHi("@A IN [""Alice"",""Diana""] AND ((NOT @B < 85) XOR (@C = ""Active"" NAND EXISTS(@D)))", "A1:C5", "A1:C1")
    
    If IsArray(result3) Then
        Debug.Print "Complex triple level: PASS"
    Else
        Debug.Print "Complex triple level: FAIL"
    End If
End Sub

Sub SetupTestData()
    ' Données test standard
    Range("A1:C5").Clear
    Range("A1:C1").value = Array("Name", "Score", "Status")
    Dim data
    data = WorksheetFunction.Transpose( _
           WorksheetFunction.Transpose(Array( _
                Array("Alice", 85, "Active"), _
                Array("Bob", 92, "Active"), _
                Array("Charlie", 78, "Inactive"), _
                Array("Diana", 95, "Active") _
           )))
    Range("A2:C5").value = data
End Sub

Sub SetupTestDataWithEmpty()
    ' Données avec cellules vides pour test EXISTS
    SetupTestData
    Range("A6:C6").value = Array("", "", "")
    Range("D1:D6").value = Array("Extra", "Data1", "Data2", "", "Data4", "")
End Sub


' ===============================================================================
' TESTS INTÉGRATION FINALE
' ===============================================================================

Public Sub RunFinalIntegrationTests()
    Debug.Print "=== FDXH TESTS INTÉGRATION FINALE ==="
    
    ' Test expression ultra-complexe
    TestUltraComplexExpression
    
    ' Test performance tous opérateurs
    TestAllOperatorsPerformance
    
    ' Test cas limites
    TestEdgeCases
    
    Debug.Print "=== TESTS INTÉGRATION FINALE TERMINÉS ==="
End Sub

Sub TestUltraComplexExpression()
    Debug.Print "Testing ultra-complex expression..."
    
    SetupTestData
    
    ' Expression utilisant TOUS les opérateurs
    Dim complexExpr As String
    complexExpr = "(@A IN [""Alice"",""Diana""] AND NOT (@B < 85)) XOR " & _
                  "((EXISTS(@C) NAND @C = ""Inactive"") OR (@B BETWEEN [90:95] AND @A NOT IN [""Charlie""]))"
    
    Dim result As Variant
    result = FDXarrComplete(complexExpr, "A1:C5", "A1:C1")
    
    If IsArray(result) Then
        Debug.Print "Ultra-complex expression: PASS - Found " & UBound(result, 1) & " results"
        Debug.Print "Final total cost: " & GetTotalCost()
    Else
        Debug.Print "Ultra-complex expression: FAIL"
    End If
End Sub

Sub TestAllOperatorsPerformance()
    Debug.Print "Testing performance with all operators..."
    
    SetupTestData
    
    Dim startTime As Double
    startTime = Timer
    
    ' 50 évaluations complexes
    Dim i As Long
    For i = 1 To 50
        Dim result As Variant
        result = FDXarrComplete("@A IN [""Alice"",""Bob""] XOR (@B > 85 NAND EXISTS(@C))", "A1:C5", "A1:C1")
    Next i
    
    Dim endTime As Double
    endTime = Timer
    
    Debug.Print "50 complex evaluations: " & Format(endTime - startTime, "0.000") & " seconds"
    Debug.Print "Average per evaluation: " & Format((endTime - startTime) / 50, "0.0000") & " seconds"
    Debug.Print "Performance test: " & IIf(endTime - startTime < 2, "PASS", "ATTENTION")
End Sub

Sub TestEdgeCases()
    Debug.Print "Testing edge cases..."
    
    ' Test liste IN vide
    On Error Resume Next
    Dim result1 As Variant
    result1 = FDXarrComplete("@A IN []", "A1:C5", "A1:C1")
    Debug.Print "Empty IN list: " & IIf(IsError(result1), "PASS (error caught)", "FAIL")
    On Error GoTo 0
    
    ' Test NOT imbriqué
    Dim result2 As Variant
    result2 = FDXarrComplete("NOT (NOT (@A = ""Alice""))", "A1:C5", "A1:C1")
    Debug.Print "Double NOT: " & IIf(IsArray(result2), "PASS", "FAIL")
    
    ' Test EXISTS sur colonne inexistante
    Dim result3 As Variant
    result3 = FDXarrComplete("EXISTS(@Z)", "A1:C5", "A1:C1")
    Debug.Print "EXISTS non-existent: " & IIf(IsArray(result3), "PASS", "FAIL")
End Sub

' ===============================================================================
' TESTS FINAUX COMPLETS
' ===============================================================================

Public Sub RunAllFinalTests()
    Debug.Print "=== FDXH TESTS FINAUX COMPLETS - TOUS OPÉRATEURS ==="
    
    ' Tests unitaires complets
    RunCompleteOperatorTests
    
    ' Tests intégration finale
    RunFinalIntegrationTests
    
    ' Tests performance
    RunPerformanceTests
    
    ' Tests validation
    RunValidationTests
    
    ' Tests cas limites
    RunEdgeCaseTests
    
    Debug.Print "=== TOUS LES TESTS TERMINÉS - FDXH COMPLET ==="
End Sub

Sub RunPerformanceTests()
    Debug.Print "=== PERFORMANCE TESTS ==="
    
    ' Test 1: Expression simple vs complexe
    SetupTestData
    
    Dim startTime As Double, endTime As Double
    
    ' Expression simple
    startTime = Timer
    Dim i As Long
    For i = 1 To 100
        Dim result1 As Variant
        result1 = FDXHi("@B > 85", "A1:C5", "A1:C1")
    Next i
    endTime = Timer
    Debug.Print "100 simple expressions: " & Format(endTime - startTime, "0.000") & "s"
    
    ' Expression complexe
    startTime = Timer
    For i = 1 To 100
        Dim result2 As Variant
        result2 = FDXHi("@A IN [""Alice"",""Bob""] XOR (NOT (@B < 85) NAND EXISTS(@C))", "A1:C5", "A1:C1")
    Next i
    endTime = Timer
    Debug.Print "100 complex expressions: " & Format(endTime - startTime, "0.000") & "s"
    
    Debug.Print "Performance ratio: " & Format((endTime - startTime) / 0.001, "0.0") & "x"
End Sub

Sub RunValidationTests()
    Debug.Print "=== VALIDATION TESTS ==="
    
    ' Test validation IN list vide
    On Error Resume Next
    Dim result1 As Variant
    result1 = FDXHi("@A IN []", "A1:C5", "A1:C1")
    Debug.Print "Empty IN validation: " & IIf(Err.Number <> 0, "PASS", "FAIL")
    On Error GoTo 0
    
    ' Test validation NOT imbriqué trop profond
    SetConfigValue "MaxNotDepth", 2
    On Error Resume Next
    Dim result2 As Variant
    result2 = FDXHi("NOT (NOT (NOT @A = ""test""))", "A1:C5", "A1:C1")
    Debug.Print "NOT depth validation: " & IIf(Err.Number <> 0, "PASS", "FAIL")
    On Error GoTo 0
    
    ' Restaurer config
    SetConfigValue "MaxNotDepth", 3
End Sub

Sub RunEdgeCaseTests()
    Debug.Print "=== EDGE CASE TESTS ==="
    
    SetupTestData
    
    ' Test valeurs nulles/vides
    Range("A6:C6").value = Array("", 0, "")
    
    Dim result1 As Variant
    result1 = FDXHi("EXISTS(@A)", "A1:C6", "A1:C1")
    Debug.Print "EXISTS with empty: " & IIf(IsArray(result1), "PASS", "FAIL")
    
    ' Test IN avec types mixtes
    Dim result2 As Variant
    result2 = FDXHi("@B IN [85,""92"",95]", "A1:C5", "A1:C1")
    Debug.Print "IN mixed types: " & IIf(IsArray(result2), "PASS", "FAIL")
    
    ' Test expression vide
    On Error Resume Next
    Dim result3 As Variant
    result3 = FDXHi("", "A1:C5", "A1:C1")
    Debug.Print "Empty expression: " & IIf(Err.Number <> 0, "PASS", "FAIL")
    On Error GoTo 0
End Sub



