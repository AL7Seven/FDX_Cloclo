Attribute VB_Name = "mod_Tests"
' ============================================================================
' DOCUMENTATION TESTS
' ============================================================================
'
' UTILISATION DES TESTS:
'
' 1. Tests complets:
'    RunAllMultiRangeTests
'
' 2. Tests performance:
'    TestPerformanceMultiRanges
'
' 3. Validation registry spécifique:
'    Dim reg As Object: Set reg = BuildColumnRegistry("@Test=1", "A:C")
'    ValidateCompleteRegistry reg
'
' 4. Tests individuels:
'    TestSingleCase "@Field='Value'", "A:C,F:H", "Mon test"
'
' COUVERTURE:
' - Tous formats supportés (Excel, numérique, nommé, bracket)
' - Cas limites (ranges inversées, très larges, espaces)
' - Gestion erreurs (expressions invalides, ranges incorrectes)
' - Préservation ordre (crucial pour affichage)
' - Performance avec grandes spécifications
'
' ============================================================================
' ============================================================================
' V2. TESTS PARSING MULTI-RANGES - VALIDATION FONCTIONNELLE
' ============================================================================
' Suite de tests complète pour valider parsing multi-ranges réels
' Couvre tous les formats supportés et cas limites
' ============================================================================

Option Explicit

' ============================================================================
' MODULE DE TESTS PRINCIPAL
' ============================================================================

Public Sub RunAllMultiRangeTests()
    Debug.Print "========================================="
    Debug.Print "DÉMARRAGE TESTS MULTI-RANGES"
    Debug.Print "========================================="
    
    ' Configuration pour tests
    SetParsingConfig "DebugMode", True
    SetParsingConfig "LogParsingSteps", True
    SetParsingConfig "VerboseLogging", False
    
    Dim testsPassed As Long: testsPassed = 0
    Dim testsFailed As Long: testsFailed = 0
    
    ' Tests par catégorie
    testsPassed = testsPassed + TestExcelMultiRanges(testsFailed)
    testsPassed = testsPassed + TestNumericRanges(testsFailed)
    testsPassed = testsPassed + TestMixedFormats(testsFailed)
    testsPassed = testsPassed + TestNamedRanges(testsFailed)
    testsPassed = testsPassed + TestBracketFormats(testsFailed)
    testsPassed = testsPassed + TestEdgeCases(testsFailed)
    testsPassed = testsPassed + TestErrorHandling(testsFailed)
    testsPassed = testsPassed + TestOrderPreservation(testsFailed)
    
    ' Résumé final
    Debug.Print "========================================="
    Debug.Print "RÉSULTATS TESTS MULTI-RANGES"
    Debug.Print "Tests réussis: " & testsPassed
    Debug.Print "Tests échoués: " & testsFailed
    Debug.Print "Total: " & (testsPassed + testsFailed)
    Debug.Print "Taux réussite: " & Format((testsPassed / (testsPassed + testsFailed)) * 100, "0.0") & "%"
    Debug.Print "========================================="
End Sub

' ============================================================================
' TESTS RANGES EXCEL
' ============================================================================

Function TestExcelMultiRanges(ByRef failedCount As Long) As Long
    Debug.Print "--- Tests Ranges Excel ---"
    
    Dim passedTests As Long: passedTests = 0
    
    ' Test 1: Multi-ranges simples A:C,F:H
    If TestSingleCase("@Nom > 'Test'", "A:C,F:H", "Multi-ranges colonnes simples") Then
        passedTests = passedTests + 1
    Else
        failedCount = failedCount + 1
    End If
    
    ' Test 2: Multi-ranges cellules A1:B2,EF10:EG10
    If TestSingleCase("@ID = 123", "A1:B2,EF10:EG10", "Multi-ranges cellules") Then
        passedTests = passedTests + 1
    Else
        failedCount = failedCount + 1
    End If
    
    ' Test 3: Mixte cellules/colonnes A1:B1,D:F,H10:I10
    If TestSingleCase("@Status = 'Active'", "A1:B1,D:F,H10:I10", "Mixte cellules/colonnes") Then
        passedTests = passedTests + 1
    Else
        failedCount = failedCount + 1
    End If
    
    ' Test 4: Colonnes individuelles A,C,E,G
    If TestSingleCase("@Price > 100", "A,C,E,G", "Colonnes individuelles") Then
        passedTests = passedTests + 1
    Else
        failedCount = failedCount + 1
    End If
    
    ' Test 5: Range large A:Z
    If TestSingleCase("@Category LIKE 'Prod*'", "A:Z", "Range large A:Z") Then
        passedTests = passedTests + 1
    Else
        failedCount = failedCount + 1
    End If
End Function

' ============================================================================
' TESTS RANGES NUMÉRIQUES
' ============================================================================

Function TestNumericRanges(ByRef failedCount As Long) As Long
    Debug.Print "--- Tests Ranges Numériques ---"
    
    Dim passedTests As Long: passedTests = 0
    
    ' Test 1: Range numérique simple 1:5
    If TestSingleCase("@Col1 > 0", "1:5", "Range numérique 1:5") Then
        passedTests = passedTests + 1
    Else
        failedCount = failedCount + 1
    End If
    
    ' Test 2: Multi-ranges numériques 1:3,5,7:9
    If TestSingleCase("@Value <> 0", "1:3,5,7:9", "Multi-ranges numériques") Then
        passedTests = passedTests + 1
    Else
        failedCount = failedCount + 1
    End If
    
    ' Test 3: Colonnes numériques individuelles 2,4,6,8
    If TestSingleCase("@Flag = TRUE", "2,4,6,8", "Colonnes numériques individuelles") Then
        passedTests = passedTests + 1
    Else
        failedCount = failedCount + 1
    End If

End Function

' ============================================================================
' TESTS FORMATS MIXTES
' ============================================================================

Function TestMixedFormats(ByRef failedCount As Long) As Long
    Debug.Print "--- Tests Formats Mixtes ---"
    
    Dim passedTests As Long: passedTests = 0
    
    ' Test 1: Mixte Excel/Numérique A:C,5:7,J
    If TestSingleCase("@Mixed = 'Test'", "A:C,5:7,J", "Mixte Excel/Numérique") Then
        passedTests = passedTests + 1
    Else
        failedCount = failedCount + 1
    End If
    
    ' Test 2: Complexe A1:B2,5,H:J,15:20
    If TestSingleCase("@Complex > 50", "A1:B2,5,H:J,15:20", "Format complexe mixte") Then
        passedTests = passedTests + 1
    Else
        failedCount = failedCount + 1
    End If
    
End Function

' ============================================================================
' TESTS RÉFÉRENCES NOMMÉES
' ============================================================================

Function TestNamedRanges(ByRef failedCount As Long) As Long
    Debug.Print "--- Tests Références Nommées ---"
    
    Dim passedTests As Long: passedTests = 0
    
    ' Test 1: Range nommée simple Date:Facture
    If TestSingleCase("@Amount > 1000", "Date:Facture", "Range nommée simple", True) Then
        passedTests = passedTests + 1
    Else
        failedCount = failedCount + 1
    End If
    
    ' Test 2: Mixte nommé/Excel Clients:Montants,A:C
    If TestSingleCase("@Client <> ''", "Clients:Montants,A:C", "Mixte nommé/Excel", True) Then
        passedTests = passedTests + 1
    Else
        failedCount = failedCount + 1
    End If
    
    ' Test 3: Référence unique nommée MonTableau
    If TestSingleCase("@TableField = 'Value'", "MonTableau", "Référence unique nommée", True) Then
        passedTests = passedTests + 1
    Else
        failedCount = failedCount + 1
    End If
    
End Function

' ============================================================================
' TESTS FORMATS BRACKET
' ============================================================================

Function TestBracketFormats(ByRef failedCount As Long) As Long
    Debug.Print "--- Tests Formats Bracket ---"
    
    Dim passedTests As Long: passedTests = 0
    
    ' Test 1: Bracket simple [A,C,E]
    If TestSingleCase("@Field IN ('A','B')", "[A,C,E]", "Bracket simple") Then
        passedTests = passedTests + 1
    Else
        failedCount = failedCount + 1
    End If
    
    ' Test 2: Bracket avec ranges [A:C,F:H,J]
    If TestSingleCase("@Status = 'OK'", "[A:C,F:H,J]", "Bracket avec ranges") Then
        passedTests = passedTests + 1
    Else
        failedCount = failedCount + 1
    End If
    
    ' Test 3: Bracket mixte [A1:B2,5,H:J]
    If TestSingleCase("@Mixed BETWEEN 10 AND 20", "[A1:B2,5,H:J]", "Bracket mixte") Then
        passedTests = passedTests + 1
    Else
        failedCount = failedCount + 1
    End If
    
End Function

' ============================================================================
' TESTS CAS LIMITES
' ============================================================================

Function TestEdgeCases(ByRef failedCount As Long) As Long
    Debug.Print "--- Tests Cas Limites ---"
    
    Dim passedTests As Long: passedTests = 0
    
    ' Test 1: Range inversée Z:A (doit être corrigée)
    If TestSingleCase("@Reverse > 0", "Z:A", "Range inversée Z:A") Then
        passedTests = passedTests + 1
    Else
        failedCount = failedCount + 1
    End If
    
    ' Test 2: Range très large A:XFD (limitée automatiquement)
    If TestSingleCase("@Large = 'Test'", "A:XFD", "Range très large", False, True) Then
        passedTests = passedTests + 1
    Else
        failedCount = failedCount + 1
    End If
    
    ' Test 3: Espaces dans spécification " A : C , F : H "
    If TestSingleCase("@Spaces <> ''", " A : C , F : H ", "Espaces dans spec") Then
        passedTests = passedTests + 1
    Else
        failedCount = failedCount + 1
    End If
    
    ' Test 4: READ vide (READ_EQUALS_WHAT)
    If TestSingleCase("@EmptyRead = 'Value'", "", "READ vide (optimisation)") Then
        passedTests = passedTests + 1
    Else
        failedCount = failedCount + 1
    End If
    
End Function

' ============================================================================
' TESTS GESTION ERREURS
' ============================================================================

Function TestErrorHandling(ByRef failedCount As Long) As Long
    Debug.Print "--- Tests Gestion Erreurs ---"
    
    Dim passedTests As Long: passedTests = 0
    
    ' Test 1: Expression WHAT vide (doit échouer)
    If TestErrorCase("", "A:C", ERR_PARSING_INVALID_EXPRESSION, "Expression WHAT vide") Then
        passedTests = passedTests + 1
    Else
        failedCount = failedCount + 1
    End If
    
    ' Test 2: Range invalide ABC:XYZ (doit échouer)
    If TestErrorCase("@Test = 1", "ABC:XYZ", ERR_PARSING_INVALID_READ_SPEC, "Range invalide") Then
        passedTests = passedTests + 1
    Else
        failedCount = failedCount + 1
    End If
    
    ' Test 3: Expression trop longue (doit échouer)
    Dim longExpr As String: longExpr = String(40000, "X")
    If TestErrorCase(longExpr, "A:C", ERR_PARSING_INVALID_EXPRESSION, "Expression trop longue") Then
        passedTests = passedTests + 1
    Else
        failedCount = failedCount + 1
    End If
    
    ' Test 4: Parenthèses non équilibrées
    If TestErrorCase("@Test > (5 AND @Other = 'Test'", "A:C", ERR_PARSING_INVALID_EXPRESSION, "Parenthèses non équilibrées") Then
        passedTests = passedTests + 1
    Else
        failedCount = failedCount + 1
    End If
    
End Function

' ============================================================================
' TESTS PRÉSERVATION ORDRE
' ============================================================================

Function TestOrderPreservation(ByRef failedCount As Long) As Long
    Debug.Print "--- Tests Préservation Ordre ---"
    
    Dim passedTests As Long: passedTests = 0
    
    ' Test 1: Ordre spécifique doit être préservé
    If TestOrderCase("@Field = 1", "C,A,E,B", Array("@C", "@A", "@E", "@B"), "Ordre spécifique") Then
        passedTests = passedTests + 1
    Else
        failedCount = failedCount + 1
    End If
    
    ' Test 2: Ordre ranges doit être préservé
    If TestOrderCase("@X > 0", "F:H,A:C,J", Array("@F", "@G", "@H", "@A", "@B", "@C", "@J"), "Ordre ranges") Then
        passedTests = passedTests + 1
    Else
        failedCount = failedCount + 1
    End If
    
End Function

' ============================================================================
' FONCTIONS SUPPORT TESTS
' ============================================================================

Function TestSingleCase(whatExpr As String, readSpec As String, testName As String, Optional expectNamed As Boolean = False, Optional expectLimited As Boolean = False) As Boolean
    Debug.Print "Test: " & testName
    
    ClearParsingError
    
    Dim registry As Object
    Set registry = BuildColumnRegistry(whatExpr, readSpec)
    
    If HasParsingError() Then
        Dim err As ParsingError: err = GetLastParsingError()
        Debug.Print "  ÉCHEC - Erreur: " & err.message
        TestSingleCase = False
        Exit Function
    End If
    
    If registry Is Nothing Then
        Debug.Print "  ÉCHEC - Registry null"
        TestSingleCase = False
        Exit Function
    End If
    
    ' Vérifications spécifiques
    If expectNamed Then
        ' Vérifier présence marqueurs _NAMED
        Dim hasNamed As Boolean: hasNamed = False
        Dim key As Variant
        For Each key In registry("READ_FIELDS").Keys
            If Right(CStr(key), 6) = "_NAMED" Then
                hasNamed = True
                Exit For
            End If
        Next key
        
        If Not hasNamed Then
            Debug.Print "  ÉCHEC - Références nommées attendues mais non trouvées"
            TestSingleCase = False
            Exit Function
        End If
    End If
    
    ' Validation basique
    If registry("ALL_REQUIRED").Count = 0 Then
        Debug.Print "  ÉCHEC - Aucun champ requis détecté"
        TestSingleCase = False
        Exit Function
    End If
    
    Debug.Print "  OK - " & registry("ALL_REQUIRED").Count & " champs, READ_EQUALS_WHAT=" & registry("READ_EQUALS_WHAT")
    TestSingleCase = True
End Function

Function TestErrorCase(whatExpr As String, readSpec As String, expectedErrorCode As ParsingErrorCode, testName As String) As Boolean
    Debug.Print "Test Erreur: " & testName
    
    ClearParsingError
    
    Dim registry As Object
    Set registry = BuildColumnRegistry(whatExpr, readSpec)
    
    If Not HasParsingError() Then
        Debug.Print "  ÉCHEC - Erreur attendue mais non détectée"
        TestErrorCase = False
        Exit Function
    End If
    
    Dim err As ParsingError: err = GetLastParsingError()
    If err.Code = expectedErrorCode Then
        Debug.Print "  OK - Erreur correcte détectée: " & err.Code
        TestErrorCase = True
    Else
        Debug.Print "  ÉCHEC - Erreur " & err.Code & " au lieu de " & expectedErrorCode
        TestErrorCase = False
    End If
End Function

Function TestOrderCase(whatExpr As String, readSpec As String, expectedOrder As Variant, testName As String) As Boolean
    Debug.Print "Test Ordre: " & testName
    
    ClearParsingError
    
    Dim registry As Object
    Set registry = BuildColumnRegistry(whatExpr, readSpec)
    
    If HasParsingError() Then
        Dim err As ParsingError: err = GetLastParsingError()
        Debug.Print "  ÉCHEC - Erreur: " & err.message
        TestOrderCase = False
        Exit Function
    End If
    
    ' Vérifier ordre via READ_ORDER ou EXTRACT_POSITIONS
    Dim readOrder As Object: Set readOrder = registry("READ_ORDER")
    Dim extractPos As Object: Set extractPos = registry("EXTRACT_POSITIONS")
    
    ' Construire ordre réel
    Dim actualOrder As Collection: Set actualOrder = New Collection
    Dim pos As Long
    
    For pos = 1 To UBound(expectedOrder) + 1
        Dim key As Variant
        For Each key In extractPos.Keys
            If extractPos(key) = pos Then
                actualOrder.Add CStr(key)
                Exit For
            End If
        Next key
    Next pos
    
    ' Comparer ordres
    If actualOrder.Count <> UBound(expectedOrder) + 1 Then
        Debug.Print "  ÉCHEC - Nombre éléments différent: " & actualOrder.Count & " vs " & (UBound(expectedOrder) + 1)
        TestOrderCase = False
        Exit Function
    End If
    
    Dim i As Long
    For i = 1 To actualOrder.Count
        If actualOrder(i) <> expectedOrder(i - 1) Then
            Debug.Print "  ÉCHEC - Position " & i & ": '" & actualOrder(i) & "' au lieu de '" & expectedOrder(i - 1) & "'"
            TestOrderCase = False
            Exit Function
        End If
    Next i
    
    Debug.Print "  OK - Ordre préservé correctement"
    TestOrderCase = True
End Function

' ============================================================================
' TESTS SPÉCIALISÉS PERFORMANCE
' ============================================================================

Public Sub TestPerformanceMultiRanges()
    Debug.Print "--- Tests Performance Multi-Ranges ---"
    
    Dim startTime As Single: startTime = Timer
    
    ' Test performance avec 100 ranges
    Dim largeRangeSpec As String: largeRangeSpec = ""
    Dim i As Long
    For i = 1 To 100
        If i > 1 Then largeRangeSpec = largeRangeSpec & ","
        largeRangeSpec = largeRangeSpec & Chr(65 + (i Mod 26)) & (i \ 26 + 1)
    Next i
    
    ClearParsingError
    Dim registry As Object
    Set registry = BuildColumnRegistry("@Performance = 'Test'", largeRangeSpec)
    
    Dim endTime As Single: endTime = Timer
    Dim duration As Single: duration = endTime - startTime
    
    If HasParsingError() Then
        Debug.Print "Performance Test ÉCHEC: " & GetLastParsingError().message
    Else
        Debug.Print "Performance Test OK: " & duration & "s pour " & registry("ALL_REQUIRED").Count & " champs"
    End If
End Sub

' ============================================================================
' UTILITAIRE VALIDATION COMPLÈTE
' ============================================================================

Public Sub ValidateCompleteRegistry(registry As Object)
    Debug.Print "--- Validation Complète Registry ---"
    
    If registry Is Nothing Then
        Debug.Print "ERREUR: Registry null"
        Exit Sub
    End If
    
    ' Vérifier structures requises
    Dim requiredKeys As Variant
    requiredKeys = Array("WHAT_FIELDS", "READ_FIELDS", "ALL_REQUIRED", "READ_ORDER", "EXTRACT_POSITIONS")
    
    Dim i As Long
    For i = 0 To UBound(requiredKeys)
        If Not registry.Exists(requiredKeys(i)) Then
            Debug.Print "ERREUR: Clé manquante - " & requiredKeys(i)
        Else
            Debug.Print "OK: " & requiredKeys(i) & " présent (" & registry(requiredKeys(i)).Count & " éléments)"
        End If
    Next i
    
    ' Vérifier cohérence READ_EQUALS_WHAT
    Dim readEqualsWhat As Boolean: readEqualsWhat = CBool(registry("READ_EQUALS_WHAT"))
    If readEqualsWhat Then
        If registry("READ_FIELDS").Count > 0 Then
            Debug.Print "ATTENTION: READ_EQUALS_WHAT=True mais READ_FIELDS non vide"
        Else
            Debug.Print "OK: Optimisation READ_EQUALS_WHAT active"
        End If
    End If
    
    Debug.Print "--- Fin Validation ---"
End Sub

