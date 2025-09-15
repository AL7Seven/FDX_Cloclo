Attribute VB_Name = "mod_Global_9"
Option Explicit
Option Base 1
Option Compare Text
' =======================================================================
' FindXtreme Hi-Light - Final Version 2 et 3
' mod_Global
' =======================================================================
' This module contains global variables and constants used throughout the application.
' =======================================================================

' ===============================================================================
' ENUMERATIONS FDXH V3
' ===============================================================================
' NOUVELLE APPROCHE - GARDER LA RICHESSE D'INFORMATION :
Public Enum TokenType_Enum
    TT_FieldReference = 1    ' @A, @B, @C
    TT_Operator = 2          ' =, >, <, >=, <=, <>, ~, !~, IN, NOT IN, BETWEEN
    TT_Value = 3             ' Valeurs littérales simples
    TT_LogicalOp = 4         ' AND, OR
    TT_Not = 5
    TT_OpenParen = 6         ' (
    TT_CloseParen = 7        ' )
    TT_Function = 8          ' EXISTS(), REGEX()
    TT_Extended = 9          ' XOR, NAND, NOR
    TT_ValueList = 10        ' ["val1","val2"] - listes IN
    TT_RangeList = 11        ' [10:20] ou [10:20,50:60] - listes BETWEEN
End Enum

Public Enum OperatorPriority_Enum
    OP_OR = 1               ' Priorité la plus basse
    OP_AND = 2              ' Priorité standard
    OP_NOT = 3              ' Priorité haute
    OP_COMPARISON = 4       ' =, >, < etc.
    OP_PARENTHESES = 5      ' Priorité maximale
End Enum

' Structure optimisée TokenInfo V3
Public Type TokenInfo
    TokenID As Long              ' ID unique
    tokenType As TokenType_Enum  ' Type de token
    tokenValue As String         ' Valeur string
    nestingLevel As Long         ' Niveau hiérarchique
    ParentGroupID As Long        ' Groupe parent (-1 si racine)
    Priority As Long             ' Priorité évaluation
    Position As Long            ' Position dans expression
    CostValue As Double         ' Coût individuel
End Type

Public Type sourceInfo
    SourceType As String        ' "EXCEL", "CSV", "ACCESS"
    filePath As String         ' Chemin complet fichier
    sheetName As String        ' Nom feuille (Excel)
    tableName As String        ' Nom table (Access)
    HasHeader As Boolean       ' PremiEre ligne = en-tÃªtes
    delimiter As String        ' SEparateur CSV
    Password As String         ' Mot de passe (Access)
End Type

Public Type ColumnMapping
    FieldReference As String   ' "@A", "@B", "@C"
    columnIndex As Long       ' Index colonne dans donnEes (1-based)
    ColumnLetter As String    ' Lettre Excel (A, B, C)
    ColumnNumber As Long      ' N° de position de la colonne
    IsRequired As Boolean     ' Necessaire pour Evaluation
    NbLignes  As Long         ' Nombre de lignes
End Type

' Structure TokenGroup avec performance O(1)
Public Type TokenGroup
    groupID As Long             ' ID unique
    ParentGroupID As Long       ' ID parent (-1 si racine)
    nestingLevel As Long        ' Niveau profondeur
    TokenStartIndex As Long     ' Index début dans pTokens()
    TokenEndIndex As Long       ' Index fin dans pTokens()
    Priority As Long            ' Priorité calculée
    LogicalOperator As String   ' "AND" ou "OR"
    CostValue As Double         ' Coût du groupe
    IsEvaluated As Boolean      ' Flag évaluation
End Type

Public FDXH_Config As Object

Public pTokens() As TokenInfo          ' Array plat de tokens
Public pTokenCount As Long             ' Nombre de tokens
Public pGroups() As TokenGroup         ' Array plat de groupes
Public pGroupCount As Long             ' Nombre de groupes
Public pCurrentNestingLevel As Long    ' Niveau actuel
Public pTotalCost As Double            ' Coût total
Public pExpressionText As String       ' Expression en cours

' Variable de la V2 du 13/08
' Variables globales donnEes chargEes (mod_PaserHiMain ??)
Public pSourceData() As Variant      ' Array 2D donnEes source complEtes
Public pSourceRows As Long           ' Nombre lignes chargEes
Public pSourceCols As Long           ' Nombre colonnes chargEes
Public pColumnMaps() As ColumnMapping ' Mapping colonnes
Public pColumnCount As Long          ' Nombre mappings colonnes
Public pChunkMode As Boolean         ' Mode chunks actif
Public pCurrentChunk As Long         ' Chunk actuel en traitement

' ==== Save infos from "Application.Caller" : ====
' Utilisé pour sauvegarder la formule saisie dans la barre de formule
' Eventuellemment utile pour des traitements complémentaires
Public AppCaller As clsAppCaller  'Will be initiate in FDX function

' ============================================================================
' GESTION ERREURS PARSING
' ============================================================================
Public m_LastParsingError As ParsingError

' ============================================================================
' TYPES REQUIS POUR RÉSOLUTION DYNAMIQUE
' ============================================================================
Public Type ExcelResolutionContext
    workbookPath As String
    WorksheetName As String
    IsWorkbookOpen As Boolean
    HasNamedRanges As Boolean
    HasStructuredTables As Boolean
End Type
Public Type ComparisonContext
    fieldName As String
    Operator As String
    comparedValue As String
    contextType As String  ' "FILTER", "JOIN", "SUBQUERY", etc.
    Position As Long
End Type

' ============================================================================
' ÉNUMÉRATIONS POUR CODES ERREUR SPÉCIFIQUES
' ============================================================================
Public Enum ParsingErrorCode
    ERR_PARSING_SUCCESS = 0
    ERR_PARSING_INVALID_EXPRESSION = 1001
    ERR_PARSING_INVALID_READ_SPEC = 1002
    ERR_PARSING_RESOLUTION_FAILED = 1003
    ERR_PARSING_MAPPING_INCONSISTENT = 1004
    ERR_PARSING_MEMORY_EXCEEDED = 1005
    ERR_PARSING_EXCEL_ACCESS_FAILED = 1006
    ERR_PARSING_UNKNOWN = 1999
End Enum

Public Type ParsingError
    Code As ParsingErrorCode
    message As String
    context As String
    timestamp As Date
End Type

' ============================================================================
' CONSTANTES CONFIGURATION
' ============================================================================
Public Const PARSING_MAX_COLUMNS As Long = 16384      ' Limite Excel XFD
Public Const PARSING_MAX_RANGE_SIZE As Long = 1000    ' Sécurité ranges
Public Const PARSING_MAX_EXPRESSION_LENGTH As Long = 32767  ' Limite VBA String
Public Const PARSING_DEFAULT_TIMEOUT As Long = 30     ' Secondes pour résolution


