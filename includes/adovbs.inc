<%
' ---------------------------------------------------
'            The full complete constantfile
'                        for
'            ActiveX Data Object Version 2.1
'
'                     Written by
'                    Andreas Kviby
'                Devphiles AB, Sweden
'
'               http://www.devphiles.com
'
'                 Send your comments to
'              andreas.kviby@devphiles.com
' ---------------------------------------------------'
' ADODB.CursorOptionEnum
Const adAddNew = 16778240
Const adApproxPosition = 16384
Const adBookmark = 8192
Const adDelete = 16779264
Const adFind = 524288
Const adHoldRecords = 256
Const adIndex = 8388608
Const adMovePrevious = 512
Const adNotify = 262144
Const adResync = 131072
Const adSeek = 4194304
Const adUpdate = 16809984
Const adUpdateBatch = 65536

' ADODB.CursorLocationEnum
Const adUseClient = 3
Const adUseServer = 2

' ADODB.CursorTypeEnum
Const adOpenDynamic = 2
Const adOpenForwardOnly = 0
Const adOpenKeyset = 1
Const adOpenStatic = 3

' ADODB.AffectEnum
Const adAffectAllChapters = 4
Const adAffectCurrent = 1
Const adAffectGroup = 2

' ADODB.ConnectOptionEnum
Const adAsyncConnect = 16

' ADODB.ConnectModeEnum
Const adModeRead = 1
Const adModeReadWrite = 3
Const adModeShareDenyNone = 16 
Const adModeShareDenyRead = 4
Const adModeShareDenyWrite = 8
Const adModeShareExclusive = 12
Const adModeUnknown = 0
Const adModeWrite = 2

' ADODB.ConnectPromptEnum
Const adPromptAlways = 1
Const adPromptComplete = 2
Const adPromptCompleteRequired = 3
Const adPromptNever = 4

' ADODB.ExecuteOptionEnum
Const adAsyncExecute = 16
Const adAsyncFetch = 32
Const adAsyncFetchNonBlocking = 64
Const adExecuteNoRecords = 128

' ADODB.DataTypeEnum
Const adBigInt = 20
Const adBinary = 128
Const adBoolean = 11
Const adBSTR = 8
Const adChapter = 136
Const adChar = 129
Const adCurrency = 6
Const adDate = 7
Const adDBDate = 133
Const adDBFileTime = 137
Const adDBTime = 134
Const adDBTimeStamp = 135
Const adDecimal = 14
Const adDouble = 5
Const adEmpty = 0
Const adError = 10
Const adFileTime = 64
Const adGUID = 72
Const adIDispatch = 9
Const adInteger = 3
Const adIUnknown = 13
Const adLongVarBinary = 205
Const adLongVarChar = 201
Const adLongVarWChar = 203 
Const adNumeric = 131
Const adPropVariant = 138 
Const adSingle = 4
Const adSmallInt = 2
Const adTinyInt = 16
Const adUnsignedBigInt = 21
Const adUnsignedInt = 19
Const adUnsignedSmallInt = 18
Const adUnsignedTinyInt = 17
Const adUserDefined = 132
Const adVarBinary = 204
Const adVarChar = 200
Const adVariant = 12
Const adVarNumeric = 139
Const adVarWChar = 202
Const adWChar = 130

' ADODB.BookmarkEnum
Const adBookmarkCurrent = 0
Const adBookmarkFirst = 1
Const adBookmarkLast = 2

' ADODB.StringFormatEnum
Const adClipString = 2

' ADODB.CommandTypeEnum
Const adCmdFile = 256
Const adCmdStoredProc = 4
Const adCmdTable = 2
Const adCmdTableDirect = 512
Const adCmdText = 1
Const adCmdUnknown = 8

' ADODB.CompareEnum
Const adCompareEqual = 1
Const adCompareGreaterThan = 2
Const adCompareLessThan = 0
Const adCompareNotComparable = 4
Const adCompareNotEqual = 3

' ADODB.ADCPROP_UPDATECRITERIA_ENUM
Const adCriteriaAllCols = 1
Const adCriteriaKey = 0
Const adCriteriaTimeStamp = 3
Const adCriteriaUpdCols = 2

' ADODB.ADCPROP_ASYNCTHREADPRIORITY_ENUM
Const adPriorityAboveNormal = 4
Const adPriorityBelowNormal = 2
Const adPriorityHighest = 5
Const adPriorityLowest = 1
Const adPriorityNormal = 3

' ADODB.EditModeEnum
Const adEditAdd = 2
Const adEditDelete = 4
Const adEditInProgress = 1
Const adEditNone = 0

' ADODB.ErrorValueEnum
Const adErrBoundToCommand = 3707
Const adErrDataConversion = 3421
Const adErrFeatureNotAvailable = 3251
Const adErrIllegalOperation = 3219
Const adErrInTransaction = 3246
Const adErrInvalidArgument = 3001
Const adErrInvalidConnection = 3709
Const adErrInvalidParamInfo = 3708
Const adErrItemNotFound = 3265
Const adErrNoCurrentRecord = 3021
Const adErrNotReentrant = 3710
Const adErrObjectClosed = 3704
Const adErrObjectInCollection = 3367
Const adErrObjectNotSet = 3420
Const adErrObjectOpen = 3705
Const adErrOperationCancelled = 3712
Const adErrProviderNotFound = 3706
Const adErrStillConnecting = 3713
Const adErrStillExecuting = 3711
Const adErrUnsafeOperation = 3716

' ADODB.FilterGroupEnum
Const adFilterAffectedRecords = 2
Const adFilterConflictingRecords = 5
Const adFilterFetchedRecords = 3
Const adFilterNone = 0
Const adFilterPendingRecords = 1

' ADODB.FieldAttributeEnum
Const adFldCacheDeferred = 4096
Const adFldFixed = 16
Const adFldIsNullable = 32
Const adFldKeyColumn = 32768
Const adFldLong = 128
Const adFldMayBeNull = 64
Const adFldMayDefer = 2
Const adFldNegativeScale = 16384
Const adFldRowID = 256
Const adFldRowVersion = 512
Const adFldUnknownUpdatable = 8
Const adFldUpdatable = 4

' ADODB.GetRowsOptionEnum
Const adGetRowsRest = -1

' ADODB.LockTypeEnum
Const adLockBatchOptimistic = 4
Const adLockOptimistic = 3
Const adLockPessimistic = 2
Const adLockReadOnly = 1

' ADODB.MarschalOptionsEnum
Const adMarshalAll = 0
Const adMarshalModifiedOnly = 1

' ADODB.ParameterDirectionEnum
Const adParamInput = 1
Const adParamInputOutput = 3
Const adParamOutput = 2
Const adParamReturnValue = 4
Const adParamUnknown = 0

' ADODB.ParameterAttributesEnum
Const adParamLong = 128
Const adParamNullable = 64
Const adParamSigned = 16

' ADODB.PersistFormatEnum
Const adPersistADTG = 0
Const adPersistXML = 1

' ADODB.PositionEnum
Const adPosBOF = -2
Const adPosEOF = -3
Const adPosUnknown = -1

' ADODB.PropertyAttributesEnum
Const adPropNotSupported = 0
Const adPropOptional = 2
Const adPropRead = 512
Const adPropRequired = 1
Const adPropWrite = 1024

' ADODB.ADCPROP_AUTORECALC_ENUM
Const adRecalcAlways = 1
Const adRecalcUpFront = 0

' ADODB.RecordStatusEnum
Const adRecCanceled = 256
Const adRecCantRelease = 1024
Const adRecConcurrencyViolation = 2048
Const adRecDBDeleted = 262144
Const adRecDeleted = 4
Const adReIntegrityViolation = 4096
Const adRecInvalid = 16
Const adRecMaxChangesExceeded = 8192
Const adRecModified = 2
Const adRecMultipleChanges = 64
Const adRecNew = 1
Const adRecObjectOpen = 16384
Const adRecOK = 0
Const adRecOutOfMemory = 32768
Const adRecPendingChanges = 128
Const adRecPermissionDenied = 65536
Const adRecSchemaViolation = 131072
Const adRecUnmodified = 8

' ADODB.CEResyncEnum
Const adResyncAll = 15
Const adResyncAutoIncrement = 1
Const adResyncConflicts = 2
Const adResyncInserts = 8
Const adResyncNone = 0
Const adResyncUpdates = 4

' ADODB.ResyncEnum
Const adResyncAllValues = 2
Const adResyncUnderlyingValues = 1

' ADODB.EventReasonEnum
Const adRsnAddNew = 1
Const adRsnClose = 9
Const adRsnDelete = 2
Const adRsnFirstChange = 11
Const adRsnMove = 10
Const adRsnMoveFirst = 12
Const adRsnMoveLast = 15
Const adRsnMoveNext = 13
Const adRsnMovePrevious = 14
Const adRsnRequery = 7
Const adRsnResynch = 8
Const adRsnUndoAddNew = 5
Const adRsnUndoDelete = 6
Const adRsnUndoUpdate = 4
Const adRsnUpdate = 3

' ADODB.SchemaEnum
Const adSchemaAsserts = 0
Const adSchemaCatalogs = 1
Const adSchemaCharacterSets = 2
Const adSchemaCheckConstraints = 5
Const adSchemaCollations = 3
Const adSchemaColumnPrivileges = 13
Const adSchemaColumns = 4
Const adSchemaColumnsDomainUsage = 11
Const adSchemaConstraintColumnUsage = 6
Const adSchemaConstraintTableUsage = 7
Const adSchemaCubes = 32
Const adSchemaDBInfoKeywords = 30
Const adSchemaDBInfoLiterals = 31
Const adSchemaDimensions = 33
Const adSchemaForeignKeys = 27
Const adSchemaHierarchies = 34 
Const adSchemaIndexes = 12
Const adSchemaKeyColumnUsage = 8
Const adSchemaLevels = 35
Const adSchemaMeasures = 36
Const adSchemaMembers = 38
Const adSchemaPrimaryKeys = 28
Const adSchemaProcedureColumns = 29
Const adSchemaProcedureParameters = 26
Const adSchemaProcedures = 16
Const adSchemaProperties = 37
Const adSchemaProviderSpecific = -1
Const adSchemaProviderTypes = 22
Const adSchemaReferentialConstraints = 9
Const adSchemaSchemata = 17
Const adSchemaSQLLanguages = 18
Const adSchemaStatistics = 19
Const adSchemaTableConstraints = 10
Const adSchemaTablePrivileges = 14
Const adSchemaTables = 20
Const adSchemaTranslations = 21
Const adSchemaTrustees = 39
Const adSchemaUsagePrivileges = 15
Const adSchemaViewColumnUsage = 24
Const adSchemaViews = 23
Const adSchemaViewTableUsage = 25

' ADODB.SearchDirectionEnum
Const adSearchBackward = -1 
Const adSearchForward = 1

' ADODB.SeekEnum
Const adSeekAfter = 8
Const adSeekAfterEQ = 4
Const adSeekBefore = 32
Const adSeekBeforeEQ = 16
Const adSeekFirstEQ = 1
Const adSeekLastEQ = 2

' ADODB.ObjectStateEnum
Const adStateClosed = 0
Const adStateConnecting = 2
Const adStateExecuting = 4
Const adStateFetching = 8
Const adStateOpen = 1

' ADODB.EventStatusEnum
Const adStatusCancel = 4
Const adStatusCantDeny = 3
Const adStatusErrorsOccurred = 2
Const adStatusOK = 1
Const adStatusUnwantedEvent = 5

' ADODB.XactAttributeEnum
Const adXactAbortRetaining = 262144
Const adXactCommitRetaining = 131072

' ADODB.IsolationLevelEnum
Const adXactBrowse = 256
Const adXactChaos = 16
Const adXactCursorStability = 4096
Const adXactIsolated = 1048576
Const adXactReadCommitted = 4096
Const adXactReadUncommitted = 256
Const adXactRepeatableRead = 65536
Const adXactSerializable = 1048576
Const adXactUnspecified = -1
%>

