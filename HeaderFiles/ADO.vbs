'///////////////////////////////////////////////////////////////////////////////
'
' Header file for ADO constants
'
' Author: Andreas Heim
'
' http://msdn.microsoft.com/en-us/library/windows/desktop/ms678353(v=vs.85).aspx
'
' Required header files (to be included before): None
'
'///////////////////////////////////////////////////////////////////////////////



'Connect modes
Const adModeUnknown        = 0         'Default. Indicates that the permissions have not yet been set or cannot be determined.
Const adModeRead           = 1         'Indicates read-only permissions.
Const adModeWrite          = 2         'Indicates write-only permissions.
Const adModeReadWrite      = 3         'Indicates read/write permissions.
Const adModeShareDenyRead  = 4         'Prevents others from opening a connection with read permissions.
Const adModeShareDenyWrite = 8         'Prevents others from opening a connection with write permissions.
Const adModeShareExclusive = 12        'Prevents others from opening a connection.
Const adModeShareDenyNone  = 16        'Allows others to open a connection with any permissions. Neither read nor write access can be denied to others.
Const adModeRecursive      = &H400000  'Used in conjunction with the other *ShareDeny* values (adModeShareDenyNone, adModeShareDenyWrite, or adModeShareDenyRead) to propagate
                                       'sharing restrictions to all sub-records of the current Record. It has no affect if the Record does not have any children.
                                       'A run-time error is generated if it is used with adModeShareDenyNone only. However, it can be used with adModeShareDenyNone when combined
                                       'with other values. For example, you can use "adModeRead Or adModeShareDenyNone Or adModeRecursive".


'Connect options
Const adConnectUnspecified = -1  'Default. Opens the connection synchronously.
Const adAsyncConnect       = 16  'Opens the connection asynchronously. The ConnectComplete event may be used to determine when the connection is available.


'Connect prompt
Const adPromptAlways           = 1  'Prompts always.
Const adPromptComplete         = 2  'Prompts if more information is required.
Const adPromptCompleteRequired = 3  'Prompts if more information is required but optional parameters are not allowed.
Const adPromptNever            = 4  'Never prompts.


'Command types
Const adCmdUnspecified = -1   'Does not specify the command type argument.
Const adCmdText        = 1    'Evaluates CommandText as a textual definition of a command or stored procedure call.
Const adCmdTable       = 2    'Evaluates CommandText as a table name whose columns are all returned by an internally generated SQL query.
Const adCmdStoredProc  = 4    'Evaluates CommandText as a stored procedure name.
Const adCmdUnknown     = 8    'Default. Indicates that the type of command in the CommandText property is not known.
                              'When the type of command is not known, ADO will make several attempts to interpret the CommandText.
                              '  * CommandText is interpreted as a textual definition of a command or stored procedure call. This is the same behavior as adCmdText.
                              '  * CommandText is the name of a stored procedure. This is the same behavior as adCmdStoredProc.
                              '  * CommandText is interpreted as the name of a table. All columns are returned by an internally generated SQL query. This is the same behavior as adCmdTable.
Const adCmdFile        = 256  'Evaluates CommandText as the file name of a persistently stored Recordset. Used with Recordset.Open or Requery only.
Const adCmdTableDirect = 512  'Evaluates CommandText as a table name whose columns are all returned. Used with Recordset.Open or Requery only.
                              'To use the Seek method, the Recordset must be opened with adCmdTableDirect.
                              'This value cannot be combined with the ExecuteOption enum value adAsyncExecute.


'Command dialects for SQLOLEDB provider
Const adCmdDialectTSql        = "{C8B521FB-5CF3-11CE-ADE5-00AA0044773D}"  'Transact-SQL query
Const adCmdDialectXMLTemplate = "{5D531CB2-E6Ed-11D2-B252-00C04F681B71}"  'XML template query
Const adCmdDialectXPath       = "{EC2A4293-E898-11D2-B1B7-00C04F680C56}"  'XPath query


'Stream types
Const adTypeBinary = 1  'Indicates binary data.
Const adTypeText   = 2  'Default. Indicates text data, which is in the character set specified by Charset.


'Stream open options
Const adOpenStreamUnspecified = -1  'Default. Specifies opening the Stream object with default options.
Const adOpenStreamAsync       = 1   'Opens the Stream object in asynchronous mode.
Const adOpenStreamFromRecord  = 4   'Identifies the contents of the Source parameter to be an already open Record object. The default behavior is to treat Source as a URL that points directly to a node in a tree structure. The default stream associated with that node is opened.


'Stream read type
Const adReadAll  = -1  'Default. Reads all bytes from the stream, from the current position onwards to the EOS marker. This is the only valid value with binary streams (Type is adTypeBinary).
Const adReadLine = -2  'Reads the next line from the stream (designated by the LineSeparator property).


'Stream write type
Const adWriteChar = 0  'Default. Writes the specified text string (specified by the Data parameter) to the Stream object.
Const adWriteLine = 1  'Writes a text string and a line separator character to a Stream object. If the LineSeparator property is not defined, then this returns a run-time error.


'Line separators
Const adCRLF = -1  'Default. Indicates carriage return line feed.
Const adCR   = 13  'Indicates carriage return.
Const adLF   = 10  'Indicates line feed.


'Stream save options
Const adSaveCreateNotExist  = 1  'Default. Creates a new file if the file specified by the FileName parameter does not already exist.
Const adSaveCreateOverWrite = 2  'Overwrites the file with the data from the currently open Stream object, if the file specified by the Filename parameter already exists. If the file specified by the Filename parameter does not exist, a new file is created.


'States of Connection, Command, Stream, Recordset, Record
Const adStateClosed     = 0  'Indicates that the object is closed.
Const adStateOpen       = 1  'Indicates that the object is open.
Const adStateConnecting = 2  'Indicates that the object is connecting.
Const adStateExecuting  = 4  'Indicates that the object is executing a command.
Const adStateFetching   = 8  'Indicates that the rows of the object are being retrieved.


'Execute Options
Const adOptionUnspecified     = -1    'Indicates that the command is unspecified.
Const adAsyncExecute          = 16    'Indicates that the command should execute asynchronously.
                                      'This value cannot be combined with the CommandTypeEnum value adCmdTableDirect.
Const adAsyncFetch            = 32    'Indicates that the remaining rows after the initial quantity specified in the CacheSize property should be retrieved asynchronously.
Const adAsyncFetchNonBlocking = 64    'Indicates that the main thread never blocks while retrieving. If the requested row has not been retrieved, the current row automatically moves to the end of the file.
                                      'If you open a Recordset from a Stream containing a persistently stored Recordset, adAsyncFetchNonBlocking will not have an effect; the operation will be synchronous and blocking.
                                      'adAsynchFetchNonBlocking has no effect when the adCmdTableDirect option is used to open the Recordset.
Const adExecuteNoRecords      = 128   'Indicates that the command text is a command or stored procedure that does not return rows (for example, a command that only inserts data). If any rows are retrieved, they are discarded and not returned.
                                      'adExecuteNoRecords can only be passed as an optional parameter to the Command or Connection Execute method.
Const adExecuteStream         = 1024  'Indicates that the results of a command execution should be returned as a stream.
                                      'adExecuteStream can only be passed as an optional parameter to the Command Execute method.
Const adExecuteRecord         = 2048  'Indicates that the CommandText is a command or stored procedure that returns a single row which should be returned as a Record object.


'Data types of Field, Parameter and Property
Const adEmpty            = 0       'Specifies no value (DBTYPE_EMPTY).
Const adSmallInt         = 2       'Indicates a two-byte signed integer (DBTYPE_I2).
Const adInteger          = 3       'Indicates a four-byte signed integer (DBTYPE_I4).
Const adSingle           = 4       'Indicates a single-precision floating-point value (DBTYPE_R4).
Const adDouble           = 5       'Indicates a double-precision floating-point value (DBTYPE_R8).
Const adCurrency         = 6       'Indicates a currency value (DBTYPE_CY). Currency is a fixed-point number with four digits to the right of the decimal point. It is stored in an eight-byte signed integer scaled by 10,000.
Const adDate             = 7       'Indicates a date value (DBTYPE_DATE). A date is stored as a double, the whole part of which is the number of days since December 30, 1899, and the fractional part of which is the fraction of a day.
Const adBSTR             = 8       'Indicates a null-terminated character string (Unicode) (DBTYPE_BSTR).
Const adError            = 10      'Indicates a 32-bit error code (DBTYPE_ERROR).
Const adBoolean          = 11      'Indicates a Boolean value (DBTYPE_BOOL).
Const adDecimal          = 14      'Indicates an exact numeric value with a fixed precision and scale (DBTYPE_DECIMAL).
Const adTinyInt          = 16      'Indicates a one-byte signed integer (DBTYPE_I1).
Const adUnsignedTinyInt  = 17      'Indicates a one-byte unsigned integer (DBTYPE_UI1).
Const adUnsignedSmallInt = 18      'Indicates a two-byte unsigned integer (DBTYPE_UI2).
Const adUnsignedInt      = 19      'Indicates a four-byte unsigned integer (DBTYPE_UI4).
Const adBigInt           = 20      'Indicates an eight-byte signed integer (DBTYPE_I8).
Const adUnsignedBigInt   = 21      'Indicates an eight-byte unsigned integer (DBTYPE_UI8).
Const adFileTime         = 64      'Indicates a 64-bit value representing the number of 100-nanosecond intervals since January 1, 1601 (DBTYPE_FILETIME).
Const adGUID             = 72      'Indicates a globally unique identifier (GUID) (DBTYPE_GUID).
Const adBinary           = 128     'Indicates a binary value (DBTYPE_BYTES).
Const adChar             = 129     'Indicates a string value (DBTYPE_STR).
Const adWChar            = 130     'Indicates a null-terminated Unicode character string (DBTYPE_WSTR).
Const adNumeric          = 131     'Indicates an exact numeric value with a fixed precision and scale (DBTYPE_NUMERIC).
Const adUserDefined      = 132     'Indicates a user-defined variable (DBTYPE_UDT).
Const adDBDate           = 133     'Indicates a date value (yyyymmdd) (DBTYPE_DBDATE).
Const adDBTime           = 134     'Indicates a time value (hhmmss) (DBTYPE_DBTIME).
Const adDBTimeStamp      = 135     'Indicates a date/time stamp (yyyymmddhhmmss plus a fraction in billionths) (DBTYPE_DBTIMESTAMP).
Const adChapter          = 136     'Indicates a four-byte chapter value that identifies rows in a child rowset (DBTYPE_HCHAPTER).
Const adPropVariant      = 138     'Indicates an Automation PROPVARIANT (DBTYPE_PROP_VARIANT).
Const adVarNumeric       = 139     'Indicates a numeric value.
Const adVarChar          = 200     'Indicates a string value.
Const adLongVarChar      = 201     'Indicates a long string value.
Const adVarWChar         = 202     'Indicates a null-terminated Unicode character string.
Const adLongVarWChar     = 203     'Indicates a long null-terminated Unicode string value.
Const adVarBinary        = 204     'Indicates a binary value.
Const adLongVarBinary    = 205     'Indicates a long binary value.
Const adArray            = &H2000  'A flag value, always combined with another data type constant, that indicates an array of the other data type. Does not apply to ADOX.


'Parameter attributes
Const adParamSigned   = 16   'Indicates that the parameter accepts signed values.
Const adParamNullable = 64   'Indicates that the parameter accepts null values.
Const adParamLong     = 128  'Indicates that the parameter accepts long binary data.


'Parameter directions
Const adParamUnknown     = 0  'Indicates that the parameter direction is unknown.
Const adParamInput       = 1  'Default. Indicates that the parameter represents an input parameter.
Const adParamOutput      = 2  'Indicates that the parameter represents an output parameter.
Const adParamInputOutput = 3  'Indicates that the parameter represents both an input and output parameter.
Const adParamReturnValue = 4  'Indicates that the parameter represents a return value.


'Recordset persist format
Const adPersistADTG             = 0  'Indicates Microsoft Advanced Data TableGram (ADTG) format.
Const adPersistADO              = 1  'Indicates that ADO's own Extensible Markup Language (XML) format will be used. This value is the same as adPersistXML and is included for backwards compatibility.
Const adPersistXML              = 1  'Indicates Extensible Markup Language (XML) format.
Const adPersistProviderSpecific = 2  'Indicates that the provider will persist the Recordset using its own format.


'Recordset position
Const adPosUnknown = -1  'Indicates that the current record pointer is at BOF (that is, the BOF property is True).
Const adPosBOF     = -2  'Indicates that the current record pointer is at EOF (that is, the EOF property is True).
Const adPosEOF     = -3  'Indicates that the Recordset is empty, the current position is unknown, or the provider does not support the AbsolutePage or AbsolutePosition property.


'Recordset search direction
Const adSearchBackward = -1  'Searches backward, stopping at the beginning of the Recordset. If a match is not found, the record pointer is positioned at BOF.
Const adSearchForward  = 1   'Searches forward, stopping at the end of the Recordset. If a match is not found, the record pointer is positioned at EOF.


'Recordset seek type
Const adSeekFirstEQ  = 1   'Seeks the first key equal to KeyValues.
Const adSeekLastEQ   = 2   'Seeks the last key equal to KeyValues.
Const adSeekAfterEQ  = 4   'Seeks either a key equal to KeyValues or just after where that match would have occurred.
Const adSeekAfter    = 8   'Seeks a key just after where a match with KeyValues would have occurred.
Const adSeekBeforeEQ = 16  'Seeks either a key equal to KeyValuesor just before where that match would have occurred.
Const adSeekBefore   = 32  'Seeks a key just before where a match with KeyValues would have occurred.


'Cursor types
Const adOpenUnspecified = -1  'Does not specify the type of cursor.
Const adOpenForwardOnly = 0   'Default. Uses a forward-only cursor. Identical to a static cursor, except that you can only scroll forward through records. This improves performance when you need to make only one pass through a Recordset.
Const adOpenKeyset      = 1   'Uses a keyset cursor. Like a dynamic cursor, except that you can't see records that other users add, although records that other users delete are inaccessible from your Recordset. Data changes by other users are still visible.
Const adOpenDynamic     = 2   'Uses a dynamic cursor. Additions, changes, and deletions by other users are visible, and all types of movement through the Recordset are allowed, except for bookmarks, if the provider doesn't support them.
Const adOpenStatic      = 3   'Uses a static cursor, which is a static copy of a set of records that you can use to find data or generate reports. Additions, changes, or deletions by other users are not visible.


'Cursor options
Const adHoldRecords    = &H100      'Retrieves more records or changes the next position without committing all pending changes.
Const adMovePrevious   = &H200      'Supports the MoveFirst and MovePrevious methods, and Move or GetRows methods to move the current record position backward without requiring bookmarks.
Const adBookmark       = &H2000     'Supports the Bookmark property to gain access to specific records.
Const adApproxPosition = &H4000     'Supports the AbsolutePosition and AbsolutePage properties.
Const adUpdateBatch    = &H10000    'Supports batch updating (UpdateBatch and CancelBatch methods) to transmit groups of changes to the provider.
Const adResync         = &H20000    'Supports the Resync method to update the cursor with the data that is visible in the underlying database.
Const adNotify         = &H40000    'Indicates that the underlying data provider supports notifications (which determines whether Recordset events are supported).
Const adFind           = &H80000    'Supports the Find method to locate a row in a Recordset.
Const adIndex          = &H100000   'Supports the Index property to name an index.
Const adSeek           = &H200000   'Supports the Seek method to locate a row in a Recordset.
Const adAddNew         = &H1000400  'Supports the AddNew method to add new records.
Const adDelete         = &H1000800  'Supports the Delete method to delete records.
Const adUpdate         = &H1008000  'Supports the Update method to modify existing data.


'Cursor locations
Const adUseNone        = 1  'Does not use cursor services. (This constant is obsolete and appears solely for the sake of backward compatibility.)
Const adUseServer      = 2  'Default. Uses cursors supplied by the data provider or driver. These cursors are sometimes very flexible and allow for additional sensitivity to changes others make to the data source. However, some features of the Microsoft Cursor Service for OLE DB, such as disassociated Recordset objects, cannot be simulated with server-side cursors and these features will be unavailable with this setting.
Const adUseClient      = 3  'Uses client-side cursors supplied by a local cursor library. Local cursor services often will allow many features that driver-supplied cursors may not, so using this setting may provide an advantage with respect to features that will be enabled. For backward compatibility, the synonym adUseClientBatch is also supported.
Const adUseClientBatch = 3  'Same as adUseClient
