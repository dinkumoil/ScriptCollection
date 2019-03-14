' WbemImpersonationLevel
Const wbemImpersonationLevelAnonymous       = 1
Const wbemImpersonationLevelIdentify        = 2
Const wbemImpersonationLevelImpersonate     = 3
Const wbemImpersonationLevelDelegate        = 4

' WbemAuthenticationLevel
Const wbemAuthenticationLevelDefault        = 0
Const wbemAuthenticationLevelNone           = 1
Const wbemAuthenticationLevelConnect        = 2
Const wbemAuthenticationLevelCall           = 3
Const wbemAuthenticationLevelPkt            = 4
Const wbemAuthenticationLevelPktIntegrity   = 5
Const wbemAuthenticationLevelPktPrivacy     = 6

' WbemPrivilege
Const wbemPrivilegeCreateToken              = 1
Const wbemPrivilegePrimaryToken             = 2
Const wbemPrivilegeLockMemory               = 3
Const wbemPrivilegeIncreaseQuota            = 4
Const wbemPrivilegeMachineAccount           = 5
Const wbemPrivilegeTcb                      = 6
Const wbemPrivilegeSecurity                 = 7
Const wbemPrivilegeTakeOwnership            = 8
Const wbemPrivilegeLoadDriver               = 9
Const wbemPrivilegeSystemProfile            = 10
Const wbemPrivilegeSystemtime               = 11
Const wbemPrivilegeProfileSingleProcess     = 12
Const wbemPrivilegeIncreaseBasePriority     = 13
Const wbemPrivilegeCreatePagefile           = 14
Const wbemPrivilegeCreatePermanent          = 15
Const wbemPrivilegeBackup                   = 16
Const wbemPrivilegeRestore                  = 17
Const wbemPrivilegeShutdown                 = 18
Const wbemPrivilegeDebug                    = 19
Const wbemPrivilegeAudit                    = 20
Const wbemPrivilegeSystemEnvironment        = 21
Const wbemPrivilegeChangeNotify             = 22
Const wbemPrivilegeRemoteShutdown           = 23
Const wbemPrivilegeUndock                   = 24
Const wbemPrivilegeSyncAgent                = 25
Const wbemPrivilegeEnableDelegation         = 26
Const wbemPrivilegeManageVolume             = 27

' WbemCimtype
Const wbemCimtypeSint16                     = 2
Const wbemCimtypeSint32                     = 3
Const wbemCimtypeReal32                     = 4
Const wbemCimtypeReal64                     = 5
Const wbemCimtypeString                     = 8
Const wbemCimtypeBoolean                    = 11
Const wbemCimtypeObject                     = 13
Const wbemCimtypeSint8                      = 16
Const wbemCimtypeUint8                      = 17
Const wbemCimtypeUint16                     = 18
Const wbemCimtypeUint32                     = 19
Const wbemCimtypeSint64                     = 20
Const wbemCimtypeUint64                     = 21
Const wbemCimtypeDatetime                   = 101
Const wbemCimtypeReference                  = 102
Const wbemCimtypeChar16                     = 103

' Error codes
Const wbemNoErr                             = 0            '&h0
Const wbemErrFailed                         = -2147217407  '&h80041001
Const wbemErrNotFound                       = -2147217406  '&h80041002
Const wbemErrAccessDenied                   = -2147217405  '&h80041003
Const wbemErrProviderFailure                = -2147217404  '&h80041004
Const wbemErrTypeMismatch                   = -2147217403  '&h80041005
Const wbemErrOutOfMemory                    = -2147217402  '&h80041006
Const wbemErrInvalidContext                 = -2147217401  '&h80041007
Const wbemErrInvalidParameter               = -2147217400  '&h80041008
Const wbemErrNotAvailable                   = -2147217399  '&h80041009
Const wbemErrCriticalError                  = -2147217398  '&h8004100A
Const wbemErrInvalidStream                  = -2147217397  '&h8004100B
Const wbemErrNotSupported                   = -2147217396  '&h8004100C
Const wbemErrInvalidSuperclass              = -2147217395  '&h8004100D
Const wbemErrInvalidNamespace               = -2147217394  '&h8004100E
Const wbemErrInvalidObject                  = -2147217393  '&h8004100F
Const wbemErrInvalidClass                   = -2147217392  '&h80041010
Const wbemErrProviderNotFound               = -2147217391  '&h80041011
Const wbemErrInvalidProviderRegistration    = -2147217390  '&h80041012
Const wbemErrProviderLoadFailure            = -2147217389  '&h80041013
Const wbemErrInitializationFailure          = -2147217388  '&h80041014
Const wbemErrTransportFailure               = -2147217387  '&h80041015
Const wbemErrInvalidOperation               = -2147217386  '&h80041016
Const wbemErrInvalidQuery                   = -2147217385  '&h80041017
Const wbemErrInvalidQueryType               = -2147217384  '&h80041018
Const wbemErrAlreadyExists                  = -2147217383  '&h80041019
Const wbemErrOverrideNotAllowed             = -2147217382  '&h8004101A
Const wbemErrPropagatedQualifier            = -2147217381  '&h8004101B
Const wbemErrPropagatedProperty             = -2147217380  '&h8004101C
Const wbemErrUnexpected                     = -2147217379  '&h8004101D
Const wbemErrIllegalOperation               = -2147217378  '&h8004101E
Const wbemErrCannotBeKey                    = -2147217377  '&h8004101F
Const wbemErrIncompleteClass                = -2147217376  '&h80041020
Const wbemErrInvalidSyntax                  = -2147217375  '&h80041021
Const wbemErrNondecoratedObject             = -2147217374  '&h80041022
Const wbemErrReadOnly                       = -2147217373  '&h80041023
Const wbemErrProviderNotCapable             = -2147217372  '&h80041024
Const wbemErrClassHasChildren               = -2147217371  '&h80041025
Const wbemErrClassHasInstances              = -2147217370  '&h80041026
Const wbemErrQueryNotImplemented            = -2147217369  '&h80041027
Const wbemErrIllegalNull                    = -2147217368  '&h80041028
Const wbemErrInvalidQualifierType           = -2147217367  '&h80041029
Const wbemErrInvalidPropertyType            = -2147217366  '&h8004102A
Const wbemErrValueOutOfRange                = -2147217365  '&h8004102B
Const wbemErrCannotBeSingleton              = -2147217364  '&h8004102C
Const wbemErrInvalidCimType                 = -2147217363  '&h8004102D
Const wbemErrInvalidMethod                  = -2147217362  '&h8004102E
Const wbemErrInvalidMethodParameters        = -2147217361  '&h8004102F
Const wbemErrSystemProperty                 = -2147217360  '&h80041030
Const wbemErrInvalidProperty                = -2147217359  '&h80041031
Const wbemErrCallCancelled                  = -2147217358  '&h80041032
Const wbemErrShuttingDown                   = -2147217357  '&h80041033
Const wbemErrPropagatedMethod               = -2147217356  '&h80041034
Const wbemErrUnsupportedParameter           = -2147217355  '&h80041035
Const wbemErrMissingParameter               = -2147217354  '&h80041036
Const wbemErrInvalidParameterId             = -2147217353  '&h80041037
Const wbemErrNonConsecutiveParameterIds     = -2147217352  '&h80041038
Const wbemErrParameterIdOnRetval            = -2147217351  '&h80041039
Const wbemErrInvalidObjectPath              = -2147217350  '&h8004103A
Const wbemErrOutOfDiskSpace                 = -2147217349  '&h8004103B
Const wbemErrBufferTooSmall                 = -2147217348  '&h8004103C
Const wbemErrUnsupportedPutExtension        = -2147217347  '&h8004103D
Const wbemErrUnknownObjectType              = -2147217346  '&h8004103E
Const wbemErrUnknownPacketType              = -2147217345  '&h8004103F
Const wbemErrMarshalVersionMismatch         = -2147217344  '&h80041040
Const wbemErrMarshalInvalidSignature        = -2147217343  '&h80041041
Const wbemErrInvalidQualifier               = -2147217342  '&h80041042
Const wbemErrInvalidDuplicateParameter      = -2147217341  '&h80041043
Const wbemErrTooMuchData                    = -2147217340  '&h80041044
Const wbemErrServerTooBusy                  = -2147217339  '&h80041045
Const wbemErrInvalidFlavor                  = -2147217338  '&h80041046
Const wbemErrCircularReference              = -2147217337  '&h80041047
Const wbemErrUnsupportedClassUpdate         = -2147217336  '&h80041048
Const wbemErrCannotChangeKeyInheritance     = -2147217335  '&h80041049
Const wbemErrCannotChangeIndexInheritance   = -2147217328  '&h80041050
Const wbemErrTooManyProperties              = -2147217327  '&h80041051
Const wbemErrUpdateTypeMismatch             = -2147217326  '&h80041052
Const wbemErrUpdateOverrideNotAllowed       = -2147217325  '&h80041053
Const wbemErrUpdatePropagatedMethod         = -2147217324  '&h80041054
Const wbemErrMethodNotImplemented           = -2147217323  '&h80041055
Const wbemErrMethodDisabled                 = -2147217322  '&h80041056
Const wbemErrRefresherBusy                  = -2147217321  '&h80041057
Const wbemErrUnparsableQuery                = -2147217320  '&h80041058
Const wbemErrNotEventClass                  = -2147217319  '&h80041059
Const wbemErrMissingGroupWithin             = -2147217318  '&h8004105A
Const wbemErrMissingAggregationList         = -2147217317  '&h8004105B
Const wbemErrPropertyNotAnObject            = -2147217316  '&h8004105C
Const wbemErrAggregatingByObject            = -2147217315  '&h8004105D
Const wbemErrUninterpretableProviderQuery   = -2147217313  '&h8004105F
Const wbemErrBackupRestoreWinmgmtRunning    = -2147217312  '&h80041060
Const wbemErrQueueOverflow                  = -2147217311  '&h80041061
Const wbemErrPrivilegeNotHeld               = -2147217310  '&h80041062
Const wbemErrInvalidOperator                = -2147217309  '&h80041063
Const wbemErrLocalCredentials               = -2147217308  '&h80041064
Const wbemErrCannotBeAbstract               = -2147217307  '&h80041065
Const wbemErrAmendedObject                  = -2147217306  '&h80041066
Const wbemErrClientTooSlow                  = -2147217305  '&h80041067
Const wbemErrNullSecurityDescriptor         = -2147217304  '&h80041068
Const wbemErrTimeout                        = -2147217303  '&h80041069
Const wbemErrInvalidAssociation             = -2147217302  '&h8004106A
Const wbemErrAmbiguousOperation             = -2147217301  '&h8004106B
Const wbemErrQuotaViolation                 = -2147217300  '&h8004106C
Const wbemErrTransactionConflict            = -2147217299  '&h8004106D
Const wbemErrForcedRollback                 = -2147217298  '&h8004106E
Const wbemErrUnsupportedLocale              = -2147217297  '&h8004106F
Const wbemErrHandleOutOfDate                = -2147217296  '&h80041070
Const wbemErrConnectionFailed               = -2147217295  '&h80041071
Const wbemErrInvalidHandleRequest           = -2147217294  '&h80041072
Const wbemErrPropertyNameTooWide            = -2147217293  '&h80041073
Const wbemErrClassNameTooWide               = -2147217292  '&h80041074
Const wbemErrMethodNameTooWide              = -2147217291  '&h80041075
Const wbemErrQualifierNameTooWide           = -2147217290  '&h80041076
Const wbemErrRerunCommand                   = -2147217289  '&h80041077
Const wbemErrDatabaseVerMismatch            = -2147217288  '&h80041078
Const wbemErrVetoPut                        = -2147217287  '&h80041079
Const wbemErrVetoDelete                     = -2147217286  '&h8004107A
Const wbemErrInvalidLocale                  = -2147217280  '&h80041080
Const wbemErrProviderSuspended              = -2147217279  '&h80041081
Const wbemErrSynchronizationRequired        = -2147217278  '&h80041082
Const wbemErrNoSchema                       = -2147217277  '&h80041083
Const wbemErrProviderAlreadyRegistered      = -2147217276  '&h80041084
Const wbemErrProviderNotRegistered          = -2147217275  '&h80041085
Const wbemErrFatalTransportError            = -2147217274  '&h80041086
Const wbemErrEncryptedConnectionRequired    = -2147217273  '&h80041087
Const wbemErrRegistrationTooBroad           = -2147213311  '&h80042001
Const wbemErrRegistrationTooPrecise         = -2147213310  '&h80042002
Const wbemErrTimedout                       = -2147209215  '&h80043001
Const wbemErrResetToDefault                 = -2147209214  '&h80043002

' WbemObjectTextFormat
Const wbemObjectTextFormatCIMDTD20          = 1
Const wbemObjectTextFormatWMIDTD20          = 2

' WbemFlag
Const wbemFlagReturnWhenComplete            = &h000000
Const wbemFlagReturnImmediately             = &h000010
Const wbemFlagBidirectional                 = &h000000
Const wbemFlagForwardOnly                   = &h000020
Const wbemFlagNoErrorObject                 = &h000040
Const wbemFlagReturnErrorObject             = &h000000
Const wbemFlagDontSendStatus                = &h000000
Const wbemFlagSendStatus                    = &h000080
Const wbemFlagEnsureLocatable               = &h000100
Const wbemFlagDirectRead                    = &h000200
Const wbemFlagSendOnlySelected              = &h000000
Const wbemFlagUseAmendedQualifiers          = &h020000
Const wbemFlagGetDefault                    = &h000000
Const wbemFlagSpawnInstance                 = &h000001
Const wbemFlagUseCurrentTime                = &h000001

' WbemChangeFlag
Const wbemChangeFlagCreateOrUpdate          = &h00000000
Const wbemChangeFlagUpdateOnly              = &h00000001
Const wbemChangeFlagCreateOnly              = &h00000002
Const wbemChangeFlagUpdateCompatible        = &h00000000
Const wbemChangeFlagUpdateSafeMode          = &h00000020
Const wbemChangeFlagUpdateForceMode         = &h00000040
Const wbemChangeFlagStrongValidation        = &h00000080
Const wbemChangeFlagAdvisory                = &h00010000

' WbemComparisonFlag
Const wbemComparisonFlagIncludeAll          = 0
Const wbemComparisonFlagIgnoreQualifiers    = 1
Const wbemComparisonFlagIgnoreObjectSource  = 2
Const wbemComparisonFlagIgnoreDefaultValues = 4
Const wbemComparisonFlagIgnoreClass         = 8
Const wbemComparisonFlagIgnoreCase          = 16
Const wbemComparisonFlagIgnoreFlavor        = 32

' WbemQueryFlag
Const wbemQueryFlagDeep                     = 0
Const wbemQueryFlagShallow                  = 1
Const wbemQueryFlagPrototype                = 2

' WbemTextFlag
Const wbemTextFlagNoFlavors                 = 1

' WbemTimeout
Const wbemTimeoutInfinite                   = -1

' WbemConnectOptions
Const wbemConnectFlagUseMaxWait             = 128
