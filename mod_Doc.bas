Attribute VB_Name = "mod_Doc"
'# Guide d'Impl�mentation - Module de Parsing Complet
'
'## Vue d'Ensemble
'
'Ce module fournit une solution compl�te pour analyser et r�soudre les r�f�rences de colonnes dans les expressions de filtrage et sp�cifications de lecture, avec support avanc� pour Excel, CSV et SGBD.
'
'## Architecture
'
'Le syst�me utilise un **Registry centralis�** (Dictionary) contenant tous les mappings et m�tadonn�es n�cessaires pour optimiser l'acc�s aux donn�es.
'
'---
'
'## Fonctions d'Entr�e Principales
'
'### 1. `BuildColumnRegistry()` - FONCTION MA�TRE
'
'**Signature**
'```vb
'Function BuildColumnRegistry(whatExpression As String, readColumns As String) As Object
'```
'
'**Utilit�**
'- Point d'entr�e unique pour analyser WHAT + READ
'- Construit tous les mappings n�cessaires en une passe
'- Retourne un Registry complet pr�t � l'emploi
'
'**Param�tres**
'- `whatExpression` : Expression de filtrage ("@A > 5 AND @C < 10")
'- `readColumns` : Sp�cification colonnes � lire ("A:E", "[A,C,E]", "A1:B2,EF10:EG10", etc.)
'
'**Exemple d'usage**
'```vb
'Dim registry As Object
'Set registry = BuildColumnRegistry("@Nom LIKE 'Dupont*' AND @Age > 25", "A:F")
'```
'
'**Retour**
'Registry Dictionary avec toutes les structures :
'- `WHAT_FIELDS` : Champs dans expression
'- `READ_FIELDS` : Champs � lire
'- `ALL_REQUIRED` : Union des deux
'- `READ_ORDER` : Ordre d'affichage pr�serv�
'- `COMPARISON_FIELDS` : Champs utilis�s dans comparaisons
'- Mappings de positions (source/extract)
'
'---
'
'### 2. `ResolveExcelDynamicReferences()` - R�SOLUTION DYNAMIQUE
'
'**Signature**
'```vb
'Function ResolveExcelDynamicReferences(registry As Object, Optional context As ExcelResolutionContext) As Boolean
'```
'
'**Utilit�**
'- R�sout les plages nomm�es Excel en adresses r�elles
'- Support tableaux structur�s et en-t�tes colonnes
'- Gestion workbook ouvert/ferm�
'
'**Quand l'utiliser**
'- Apr�s `BuildColumnRegistry()` si r�f�rences nomm�es d�tect�es
'- Avant traitement des donn�es Excel
'
'**Exemple d'usage**
'```vb
'Dim registry As Object
'Set registry = BuildColumnRegistry("@Facture > 1000", "Clients:Montants")
'
'If ResolveExcelDynamicReferences(registry) Then
'    ' R�f�rences r�solues, continuer traitement
'Else
'    ' Erreur r�solution, g�rer fallback
'End If
'```
'
'---
'
'### 3. `BuildAdvancedPositionMappings()` - MAPPINGS AVANC�S
'
'**Signature**
'```vb
'Sub BuildAdvancedPositionMappings(registry As Object, sourceDataArray As Variant)
'```
'
'**Utilit�**
'- Construit mappings positions r�elles depuis donn�es source
'- Optimise ordre d'extraction pour performance
'- Cr�e index inverses pour acc�s O(1)
'
'**Quand l'utiliser**
'- Apr�s chargement des donn�es source
'- Avant traitement/filtrage des donn�es
'
'**Exemple d'usage**
'```vb
'Dim sourceData As Variant
'sourceData = range("A1:Z1000").value ' Donn�es Excel
'
'Dim registry As Object
'Set registry = BuildColumnRegistry("@Prix > 100", "A:E")
'
'' Construire mappings depuis donn�es r�elles
'BuildAdvancedPositionMappings registry, sourceData
'```
'
'---
'
'## Fonctions de Consultation (API Publique)
'
'### Acc�s aux Colonnes
'
'#### `GetWhatColumns(registry)` ? Object
'Retourne Dictionary des champs dans expression WHAT
'
'#### `GetReadColumns(registry)` ? Object
'Retourne Dictionary des champs � lire (optimis� READ_EQUALS_WHAT)
'
'#### `GetAllRequiredColumns(registry)` ? Object
'Retourne Dictionary union de tous les champs n�cessaires
'
'#### `GetComparisonFields(registry)` ? Object
'Retourne Dictionary des champs utilis�s dans comparaisons
'
'**Exemple d'usage**
'```vb
'Dim whatCols As Object: Set whatCols = GetWhatColumns(registry)
'Dim readCols As Object: Set readCols = GetReadColumns(registry)
'
'Debug.Print "Champs WHAT: " & whatCols.Count
'Debug.Print "Champs READ: " & readCols.Count
'```
'
'### Mappings de Position
'
'#### `GetExtractPosition(registry, fieldRef)` ? Long
'Position du champ dans tableau extrait ordonn�
'
'#### `GetFieldAtExtractPosition(registry, position)` ? String
'Champ � la position donn�e dans tableau extrait
'
'#### `GetSourcePosition(registry, fieldRef)` ? Long
'Position du champ dans donn�es source
'
'#### `SetSourcePosition(registry, fieldRef, sourcePos)`
'D�finir position source (appel� lors extraction)
'
'**Exemple d'usage**
'```vb
'' Conna�tre ordre d'extraction
'Dim pos As Long: pos = GetExtractPosition(registry, "@Nom")
'Debug.Print "@Nom sera en position " & pos & " dans r�sultats"
'
'' Mapping inverse
'Dim field As String: field = GetFieldAtExtractPosition(registry, 1)
'Debug.Print "Premi�re colonne r�sultat: " & field
'```
'
'### Utilitaires
'
'#### `HasColumn(registry, fieldRef)` ? Boolean
'Teste existence d'un champ dans le registry
'
'#### `GetColumnCount(registry)` ? Long
'Nombre total de colonnes requises
'
'#### `GetColumnIndex(registry, fieldRef)` ? Long
'Index num�rique du champ (A=1, B=2, etc.)
'
'---
'
'## Fonctions d'Analyse Avanc�e
'
'### `AnalyzeComparisonContexts(expression)` ? Collection
'
'**Utilit�**
'Analyse sophistiqu�e des comparaisons dans expression
'
'**Retour**
'Collection de `ComparisonContext` avec :
'- `FieldName` : Champ compar�
'- `Operator` : Op�rateur (=, >, LIKE, IN, etc.)
'- `ComparedValue` : Valeur de comparaison
'- `ContextType` : Type (FILTER, JOIN, SUBQUERY, etc.)
'
'**Exemple d'usage**
'```vb
'Dim contexts As Collection
'Set contexts = AnalyzeComparisonContexts("@Age BETWEEN 18 AND 65 AND @Nom LIKE 'Dup*'")
'
'Dim i As Long
'For i = 1 To contexts.Count
'    Dim ctx As ComparisonContext: ctx = contexts(i)
'    Debug.Print ctx.fieldName & " " & ctx.Operator & " " & ctx.ComparedValue
'Next i
'```
'
'### `GetOptimizationRecommendations(registry)` ? Collection
'
'**Utilit�**
'Fournit recommandations d'optimisation bas�es sur l'analyse
'
'**Exemple d'usage**
'```vb
'Dim recommendations As Collection
'Set recommendations = GetOptimizationRecommendations(registry)
'
'Dim i As Long
'For i = 1 To recommendations.Count
'    Debug.Print "Recommandation: " & recommendations(i)
'Next i
'```
'
'---
'
'## Workflow d'Impl�mentation Complet
'
'### �tape 1 : Analyse Initiale
'```vb
'' 1. Construire registry depuis sp�cifications utilisateur
'Dim registry As Object
'Set registry = BuildColumnRegistry(userWhatExpression, userReadColumns)
'
'' 2. V�rifier besoin r�solution dynamique (si Excel)
'If HasNamedReferences(registry) Then
'    If Not ResolveExcelDynamicReferences(registry) Then
'        ' G�rer �chec r�solution
'        MsgBox "Impossible de r�soudre toutes les r�f�rences nomm�es"
'    End If
'End If
'```
'
'### �tape 2 : Chargement Donn�es
'```vb
'' 3. Charger donn�es source (exemple Excel)
'Dim sourceRange As range
'Set sourceRange = Workbooks("MonFichier.xlsx").Worksheets("Donn�es").UsedRange
'Dim sourceData As Variant: sourceData = sourceRange.value
'
'' 4. Construire mappings positions r�elles
'BuildAdvancedPositionMappings registry, sourceData
'```
'
'### �tape 3 : Optimisation Extraction
'```vb
'' 5. Obtenir ordre optimal d'acc�s
'Dim optimalOrder As Collection
'Set optimalOrder = GetOptimalAccessOrder(registry)
'
'' 6. Pr�parer tableau r�sultat dans bon ordre
'Dim resultData() As Variant
'ReDim resultData(1 To UBound(sourceData, 1), 1 To GetColumnCount(registry))
'```
'
'### �tape 4 : Extraction S�lective
'```vb
'' 7. Extraire seulement colonnes n�cessaires
'Dim allRequired As Object: Set allRequired = GetAllRequiredColumns(registry)
'Dim sourceCol As Long, extractCol As Long
'
'For Each fieldRef In allRequired.Keys
'    sourceCol = GetSourcePosition(registry, fieldRef)
'    extractCol = GetExtractPosition(registry, fieldRef)
'
'    ' Copier colonne source ? position extract
'    Dim row As Long
'    For row = 1 To UBound(sourceData, 1)
'        resultData(row, extractCol) = sourceData(row, sourceCol)
'    Next row
'Next fieldRef
'```
'
'---
'
'## Gestion des Formats Support�s
'
'### Formats READ Support�s
'
'| Format | Exemple | Description |
'|--------|---------|-------------|
'| Range Excel | `A:E` ou `A1:E10` | Range contigu� |
'| Multi-ranges | `A1:B2,EF10:EG10` | Ranges multiples (ordre pr�serv�) |
'| Liste bracket | `[A,C,E]` | Colonnes sp�cifiques |
'| Num�rique | `1:5,8,10:12` | Index num�riques |
'| Nomm� Excel | `Clients:Montants` | R�solution dynamique |
'
'### Formats WHAT Support�s
'
'| Format | Exemple | Description |
'|--------|---------|-------------|
'| Comparaisons | `@A > 5 AND @B < 10` | Op�rateurs standard |
'| LIKE | `@Nom LIKE 'Dup*'` | Correspondance patterns |
'| IN | `@Code IN ('A','B','C')` | Valeurs multiples |
'| BETWEEN | `@Age BETWEEN 18 AND 65` | Ranges de valeurs |
'| Subqueries | `@ID IN (SELECT...)` | Requ�tes imbriqu�es |
'| Jointures | `@ID1 = @ID2` | Comparaisons inter-champs |
'
'---
'
'## Diagnostics et Debug
'
'### Mode Debug
'```vb
'' Activer debug dans configuration
'FDXH_Config("DebugMode") = True
'FDXH_Config("LogParsingSteps") = True
'
'' Diagnostic complet
'DiagnoseRegistry registry
'```
'
'### Validation
'```vb
'' V�rifier coh�rence mappings
'If Not ValidateMappingConsistency(registry) Then
'    MsgBox "Incoh�rence d�tect�e dans les mappings"
'End If
'
'' V�rifier r�solution compl�te
'If Not AreAllReferencesResolved(registry) Then
'    MsgBox "Certaines r�f�rences n'ont pas pu �tre r�solues"
'End If
'```
'
'---
'
'## Bonnes Pratiques
'
'### Performance
'- Utilisez `READ_EQUALS_WHAT` quand possible (�conomie m�moire)
'- Priorisez champs de comparaison dans ordre extraction
'- Validez mappings avant traitement massif
'
'### Robustesse
'- Toujours tester `ResolveExcelDynamicReferences()` avant usage
'- G�rer cas �chec r�solution avec fallbacks appropri�s
'- Valider coh�rence mappings en mode debug
'
'### Maintenance
'- Utilisez fonctions diagnostics pour identifier optimisations
'- Documentez r�f�rences nomm�es Excel utilis�es
'- Testez avec diff�rents formats d'entr�e
'
'---
'
'## D�pendances
'
'### Modules Requis
'- `mod_Global` : Variables globales et types
'- `FDXH_Config` : Configuration syst�me
'
'### R�f�rences Excel
'- Microsoft Excel Object Library (si r�solution dynamique)
'- Microsoft Scripting Runtime (Dictionary)
'
'### Configuration Minimale
'```vb
'' Dans mod_Global
'Public FDXH_Config As Object
'
'' Initialisation
'Sub InitializeExtendedConfig()
'    Set FDXH_Config = CreateObject("Scripting.Dictionary")
'    FDXH_Config("DebugMode") = False
'    FDXH_Config("LogParsingSteps") = False
'    FDXH_Config("MaxRowsInMemory") = 100000
'    FDXH_Config("MaxInValues") = 50
'    FDXH_Config("EnableShortCircuit") = True
'End Sub
'```
'
'Cette documentation couvre l'ensemble des fonctionnalit�s. Le module est con�u pour �tre extensible et maintenir la compatibilit� avec votre architecture existante.
