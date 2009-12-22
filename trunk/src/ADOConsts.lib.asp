<SCRIPT Runat="Server" Language="VBScript">

' Constants used by MS ADO.DB
'---- The Cellset State Values
Const adStateClosed = 0
Const adStateOpen   = 1

'---- CursorTypeEnum Values ----
Const adOpenForwardOnly = 0
Const adOpenKeyset = 1
Const adOpenDynamic = 2
Const adOpenStatic = 3

'---- LockTypeEnum Values ----
Const adLockReadOnly = 1
Const adLockPessimistic = 2
Const adLockOptimistic = 3
Const adLockBatchOptimistic = 4

'---- CursorLocationEnum Values ----
Const adUseServer = 2
Const adUseClient = 3

'---- SearchDirection Values ----
Const adSearchForward = 1
Const adSearchBackward = -1

'---- CommandTypeEnum Values ----
Const adCmdUnknown = &H0008
Const adCmdText = &H0001
Const adCmdTable = &H0002
Const adCmdStoredProc = &H0004

' For ADO.Stream.Type
Const adTypeBinary = 1
Const adTypeText   = 2

' the field types
Const adLongVarBinary = 205
Const adLongVarChar   = 201

</SCRIPT> 
