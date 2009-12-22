<%
' the MeASP global variable and consts 
'Dim gMeDBConnStr ' the main database connection string
Dim gErrorMessage ' the error message(xml format) if any
Dim gRaiseException ' enabled it to raise the exception for error.
Dim gMeSysPath
'Dim gMeDatabase  ' the main MeDatabase object
Dim gApplication ' the MeApplicaion object
'Dim gCookies ' the TMeCookies object
'Dim gCookieHash
Dim gServerRoot
Dim gCache

Dim SCRIPT_NAME

Dim vbLastCreatedErrorCode
Dim vbInvalidLibScriptError, vbExecuteLibScriptError, vbListDuplicateError, vbInvalidObjAppCacheError
Dim vbRegistryPermissionError

' constants for FileSystemObjects
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const TristateUseDefault = -2 '使用系统默认值打开文件。
Const TristateTrue  = -1  '以 Unicode 方式打开文件。
Const TristateFalse = 0   '以 ASCII 方式打开文件。

Const DictionaryObjectName = "Scripting.Dictionary"
Const FileSystemObjectName = "Scripting.FileSystemObject"
'Const XMLObjectName = "Msxml.DOMDocument"  ' DO NOT USE For ApplicationCaches Performance
Const XMLObjectName = "Msxml2.FreeThreadedDOMDocument"

Const CDateBase = #2006-07-17#

vbInvalidLibScriptError = vbObjectError + 1
vbExecuteLibScriptError = vbObjectError + 2
vbListDuplicateError = vbObjectError + 3
vbInvalidObjAppCacheError = vbObjectError + 4
vbRegistryPermissionError = vbObjectError + 5

'Update the vbLastCreatedErrorCode
vbLastCreatedErrorCode =  vbRegistryPermissionError

If (ScriptEngineMajorVersion < 5) Or (ScriptEngineMajorVersion = 5 And ScriptEngineMinorVersion < 5) Then
    Response.Write("<h2>Error: Missing VBScript v5.5</h2>")
    Response.Write("In order for this script to work correctly the component " _
                 & "VBScript v5.5 " _
                 & "or a higher version needs to be installed on the server. You can download this component from " _
                 & "<a href=""http://msdn.microsoft.com/scripting/"">http://msdn.microsoft.com/scripting/</a>.")
    Response.End
End If

gRaiseException = True
gMeSysPath = "/" 'you should override this for your path

SCRIPT_NAME = Request.ServerVariables("SCRIPT_NAME")
%>