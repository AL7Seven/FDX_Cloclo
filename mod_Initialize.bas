Attribute VB_Name = "mod_Initialize"
Option Explicit

' ============================================================================
' DOCUMENTATION INTÉGRATION
' ============================================================================
'
' INTÉGRATION DANS VOTRE APPLICATION EXISTANTE:
'
' 1. Ajouter ces éléments à votre mod_Global existant
'
' 2. Appeler InitializeExtendedConfig() au démarrage de votre application
'    (par exemple dans Sub Main() ou Workbook_Open)
'
' 3. Appeler CleanupParsingResources() à la fermeture
'    (par exemple dans Workbook_BeforeClose)
'
' 4. Configurer selon vos besoins:
'    SetParsingConfig "DebugMode", True  ' Pour développement
'    SetParsingConfig "MaxRowsInMemory", 50000  ' Selon vos ressources
'
' 5. Le module de parsing utilisera automatiquement ces configurations
'    via GetParsingConfig() au lieu de l'ancien FDXH_Config()
'
' ============================================================================
' ============================================================================
' INITIALISATION CONFIGURATION (APPELÉE AU DÉMARRAGE APPLICATION)
' ============================================================================
Public Sub InitializeExtendedConfig()
    ' Créer objet configuration s'il n'existe pas
    If FDXH_Config Is Nothing Then
        Set FDXH_Config = CreateObject("Scripting.Dictionary")
    End If
    
    ' Configuration par défaut
    With FDXH_Config
        ' Debug et logging
        .Item("DebugMode") = False
        .Item("LogParsingSteps") = False
        .Item("VerboseLogging") = False
        
        ' Limites mémoire et performance
        .Item("MaxRowsInMemory") = 100000
        .Item("MaxInValues") = 50
        .Item("MaxRangeSize") = PARSING_MAX_RANGE_SIZE
        .Item("ResolutionTimeout") = PARSING_DEFAULT_TIMEOUT
        
        ' Optimisations
        .Item("EnableShortCircuit") = True
        .Item("CacheResolutions") = True
        .Item("OptimizeExtractionOrder") = False  ' Respecter ordre READ
        .Item("PrioritizeComparisonFields") = False  ' Respecter ordre READ
        
        ' Gestion erreurs
        .Item("ThrowOnResolutionFail") = False
        .Item("FallbackToDefaults") = True
        .Item("ValidateMappings") = True
    End With
End Sub

' ============================================================================
' FONCTIONS UTILITAIRES CONFIGURATION
' ============================================================================
Public Function GetParsingConfig(configKey As String) As Variant
    ' Wrapper sécurisé pour accès configuration
    On Error GoTo DefaultValue
    
    If FDXH_Config Is Nothing Then InitializeExtendedConfig
    
    If FDXH_Config.Exists(configKey) Then
        GetParsingConfig = FDXH_Config(configKey)
    Else
        GoTo DefaultValue
    End If
    
    Exit Function
    
DefaultValue:
    ' Valeurs par défaut sécurisées si clé manquante
    Select Case UCase(configKey)
        Case "DEBUGMODE": GetParsingConfig = False
        Case "LOGPARSINGSTEPS": GetParsingConfig = False
        Case "MAXROWSINMEMORY": GetParsingConfig = 100000
        Case "MAXINVALUES": GetParsingConfig = 50
        Case "ENABLESHORTCIRCUIT": GetParsingConfig = True
        Case "FALLBACKTODEFAULTS": GetParsingConfig = True
        Case "VALIDATEMAPPINGS": GetParsingConfig = True
        Case Else: GetParsingConfig = Null
    End Select
End Function

Public Sub SetParsingConfig(configKey As String, configValue As Variant)
    ' Définir valeur configuration avec validation
    If FDXH_Config Is Nothing Then InitializeExtendedConfig
    
    ' Validation selon clé
    Select Case UCase(configKey)
        Case "MAXROWSINMEMORY"
            If IsNumeric(configValue) And CLng(configValue) > 0 Then
                FDXH_Config(configKey) = CLng(configValue)
            End If
        Case "DEBUGMODE", "LOGPARSINGSTEPS", "ENABLESHORTCIRCUIT"
            FDXH_Config(configKey) = CBool(configValue)
        Case Else
            FDXH_Config(configKey) = configValue
    End Select
End Sub

Public Function GetLastParsingError() As ParsingError
    GetLastParsingError = m_LastParsingError
End Function

Public Sub SetParsingError(errorCode As ParsingErrorCode, message As String, Optional context As String = "")
    With m_LastParsingError
        .Code = errorCode
        .message = message
        .context = context
        .timestamp = Now
    End With
    
    ' Log si mode debug
    If GetParsingConfig("DebugMode") Then
        Debug.Print "PARSING ERROR [" & errorCode & "] " & message & IIf(context <> "", " | Context: " & context, "")
    End If
End Sub

Public Function HasParsingError() As Boolean
    HasParsingError = (m_LastParsingError.Code <> ERR_PARSING_SUCCESS)
End Function

Public Sub ClearParsingError()
    With m_LastParsingError
        .Code = ERR_PARSING_SUCCESS
        .message = ""
        .context = ""
        .timestamp = 0
    End With
End Sub

' ============================================================================
' VALIDATION ENVIRONNEMENT
' ============================================================================

Public Function ValidateParsingEnvironment() As Boolean
    ' Vérifier que l'environnement permet le parsing
    ValidateParsingEnvironment = True
    
    On Error GoTo EnvironmentError
    
    ' Test création Dictionary
    Dim testDict As Object
    Set testDict = CreateObject("Scripting.Dictionary")
    testDict("Test") = "OK"
    
    ' Test accès Excel si disponible
    If Not Application Is Nothing Then
        ' Excel disponible
    End If
    
    Exit Function
    
EnvironmentError:
    ValidateParsingEnvironment = False
    SetParsingError ERR_PARSING_UNKNOWN, "Environnement non compatible: " & err.Description
End Function

' ============================================================================
' NETTOYAGE RESSOURCES
' ============================================================================

Public Sub CleanupParsingResources()
    ' Nettoyage lors fermeture application
    Set FDXH_Config = Nothing
    ClearParsingError
End Sub


