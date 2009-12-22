<%
' the MeLib Class

' the lib file should must be head with <"Script>" and foot with <"/Script">

Dim Lib

Set Lib = New TMeLib

' the following libs are the core libs!!
' the core libs are always in the memory!! no nessary to load!!
'Lib.Require("MeConsts")
'Lib.Require("MeSysUtils")
'Lib.Require("MeList")
'Lib.RequireFile('ADOConsts')

'if you wanna use FDatabase lib feature:
'Lib.Require("MeDatabase")

' if you wanna encrypt lib, the "Security.Cipher" library should be included
'Lib.Require("Security.Cipher")

Const cDefaultLibFileExtName = ".lib.asp"
Const cDefaultLibRootPath = "/SYS/CODE/LIB/"

' the MeLib Options
Const optLoadFromFile = 1
Const optLoadFromDB   = 2
Const optCacheEnabled = 4

Const cLibCachePrefix = ".LIB:"

Const sqlSelectLibById = "Select category, crv_text From view__categories where cat_path=%CatPath% and cat_id = '%CatId%' and cat_language = 'en' and cat_type = 'SYS'"

Class TMeLib
    Private FLoadedLibs
    ' if the lib is encrypted if set
    Private FPassword
    Private FIncludeDirs
    ' you must assign it to use the lib in db.
    Private FDatabase
    ' the LibName field value must be a macro: %LibName% in the SQL, and the only select fields are libname and code!
    Public SQL
    
    ' the MeLib Options: optLoadFromFile, optLoadFromDB
    Public Options

    Private Sub Class_Initialize()
        Set FLoadedLibs = New TMeList
        Set FIncludeDirs = New TMeList
        Set FDatabase = Nothing
        FIncludeDirs.Delimiter = ";"
        FIncludeDirs.Duplicates = dupIgnore
        FLoadedLibs.Duplicates = dupIgnore
        SQL = "Select cat_id, crv_text From view__categories where cat_path= cat_id = '"&cDefaultLibRootPath&"%LibName%' and cat_language = 'en' and cat_type = 'SYS'"
        Options = optLoadFromFile + optLoadFromDB + optCacheEnabled
    End Sub

    Private Sub Class_Terminate()
        On Error Resume Next
        Set FLoadedLibs = Nothing
        Set FIncludeDirs = Nothing
        Set FDatabase = Nothing
        On Error Goto 0
    End Sub

    Public Property Set Database(pValue)
        Set FDatabase = pValue
    End Property

    Public Property Get Database()
        Set Database = FDatabase
    End Property

    Public Property Let Password(pValue)
        FPassword = pValue
    End Property

    Public Property Get Password()
        Password = FPassword
    End Property

    ' the libray folders; folder seperated by semicolon ";"
    Public Property Let IncludeDirs(pValue)
        FIncludeDirs.DelimitedText = pValue
        TrimInvalidPaths
    End Property

    Public Property Get IncludeDirs()
        IncludeDirs = FIncludeDirs.DelimitedText
    End Property

    Public Function AddIncludeDir(pDir)
        AddIncludeDir = -1
        if pDir <> "" then 
          if FolderExists(pDir) then AddIncludeDir = FIncludeDirs.Add(pDir)
        end if
    End Function

    Public Function RemoveIncludeDir(pDir)
        RemoveIncludeDir = FIncludeDirs.Remove(pDir)
    End Function

    ' first search the library in DB, if not found then search the libaray in the folders.
    ' 同一名称函数库只会被加载一次
    ' 采用类似Java的Package的方式装载
    ' 如： "Security.Hash.*" 将装载所有的散列类函数库。
    ' 函数库的使用： :Lib.Require("Security.Hash.MD5")
    Public Function Require(ByRef pLibName)
      Dim Result
      'WriteLn(pLibName)
      Result = FLoadedLibs.IndexOf(pLibName) >= 0
      if Result then Require = True: Exit Function

      if IsBitIn(Options, optCacheEnabled) then
        if Mid(pLibName, Len(pLibName), 1) <> "*" then 
          ' Check Cache:
          Result = iGetInCache(pLibName)
          if IsArray(Result) then
            Require = iExec(pLibName, Result(1), Result(0))
            Exit Function
          end if
        else
          if iExistsInCache(pLibName) then
            Result = iLoadLibsInCache(Mid(pLibName, 1, Len(pLibName)-1))
            if Result > 0 then Require = True: Exit Function
          end if
        end if
      end if


      if IsBitIn(Options, optLoadFromDB) then Result = (RequireInDB(pLibName) > 0) else Result = False
      if not Result and IsBitIn(Options, optLoadFromFile) then
        Result = (RequireFile(pLibName) > 0)
      end if
      Require = Result
    End Function

    Private Function TrimInvalidPaths()
        Dim Result, vFSO, i
        Set vFSO = Server.CreateObject(FileSystemObjectName)
        FIncludeDirs.Trim("")
        For i = FIncludeDirs.Count - 1 to 0 Step -1 
          if not vFSO.FolderExists(GetPhysicPath(FIncludeDirs.Items(i))) then FIncludeDirs.Delete(i)
        Next
        Set vFSO = Nothing
    End Function

    Private Function IsDBActive()
      Dim Result
      Result = not (FDatabase is Nothing)
      if Result then Result = FDatabase.Active
      IsDBActive = Result
    End Function

    ' load the lib in DB if not load
    ' 返回成功加载的函数库个数
    ' Return the count of successful loaded library
    Private Function RequireInDB(pLibName)
      Dim Result, vRS, vS
      Result = IsDBActive()
      if Result then
        vS = Replace(SQL, "%LibName%", pLibName)
        vS = Replace(vS, "*", FDatabase.Wildcard)
        Set vRS = FDatabase.OpenTable(vS, ForReading) 'Open table readonly
        Result = not (vRS is Nothing)
        if Result then Result = not vRS.Bof
        if Result then
          Result = 0
          Do
            vS = Mid(vRS.Fields(0).Value, Len(cDefaultLibRootPath)+1)
            vS = Replace(vS, "/", ".")
            if FLoadedLibs.IndexOf(vS) >= 0 then
              Result = Result + 1
            else
              if LoadLib(vS, vRS.Fields(1).Value) then Result = Result + 1
            end if
            vRS.MoveNext
          Loop Until vRS.Eof
          vRS.Close
        end if
        Set vRS = Nothing
      end if
      if Result = False then Result = 0
      if Result > 0 and Right(pLibName, 1) = "*" and IsBitIn(Options, optCacheEnabled) and isObject(gCache) then 
        if not (gCache is Nothing) then
          if not gCache.CacheExists(cLibCachePrefix+pLibName) then gCache.Values(cLibCachePrefix+pLibName) = Result
        end if
      end if

      RequireInDB = Result
    End Function

    '只要有一个文件装载成功就算True
    Private Function iRequireFiles(ByVal aLibName, ByRef pFolder)
      Dim vItem, vFileContent, Result, vFileStream, vLibName
      Dim i

      'writeln("FolderPath= " & pFolder.Path)
      Result = 0
      For Each vItem In pFolder.Files
        if LCase(Right(vItem.Name, Len(cDefaultLibFileExtName))) = cDefaultLibFileExtName then
          'writeln("tryFile= " & vItem.Name)
          vLibName = aLibName + Mid(vItem.Name, 1, Len(vItem.Name)-Len(cDefaultLibFileExtName))
          if FLoadedLibs.IndexOf(vLibName) < 0 then
            'writeln("tryFile= " & vLibName)

            ' Check Cache:
            vFileContent = iGetInCache(aLibName)
            if IsArray(vFileContent) then
              if iExec(aLibName, vFileContent(1), vFileContent(0)) then Result = Result + 1
            else
              Set vFileStream = vItem.OpenAsTextStream(ForReading)
              vFileContent = vFileStream.ReadAll
              vFileStream.Close
              Set vFileStream = Nothing
              if vFileContent <> "" then if LoadLib(vLibName, vFileContent) then Result = Result + 1
            end if
          end if
        end if
      Next

      For Each vItem In pFolder.SubFolders
        'writeln("SubFolder= " & aLibName + pFolder.Name)
        if iRequireFiles(aLibName + vItem.Name + ".", vItem) then Result = Result + 1
      Next

      iRequireFiles = Result
    End Function

    ' the library in fileSystem(Folders), Note: the aLibName MUST BE the defaut ext name ".lib.asp".
    ' Return the count of successful loaded library
    Private Function RequireFile(ByVal aLibName)
        Dim i, Result, vFolder, vFSO, vFileStream, vFileContent, vIncludesDirs, vLoadAllFilesInDir, aLibFileName

        'writeln(aLibName+":ddd")
        Result = 0
        aLibFileName = LCase(aLibName)
        aLibFileName = Replace(aLibFileName, ".", "/")
        'vFolder = ""
        if Mid(aLibFileName, Len(aLibFileName), 1) <> "*" then 
          'i = InStrRev(aLibFileName, "/")
          'if i > 0 then vFileContent = Mid(aLibFileName, 1, i): aLibFileName = Mid(aLibFileName, i+1)
          aLibFileName = aLibFileName + cDefaultLibFileExtName
          vLoadAllFilesInDir = False

        else
          vLoadAllFilesInDir = True
          aLibName = Mid(aLibName, 1, Len(aLibName)-1)
          aLibFileName = Mid(aLibFileName, 1, Len(aLibFileName)-1)
        end if

        'Find file in include dir
        Set vFSO = Server.CreateObject(FileSystemObjectName)

        Set vIncludesDirs = FIncludeDirs.Clone()
        if gMeSysPath <> "" then vIncludesDirs.Add(gMeSysPath)
        For i = 0 to vIncludesDirs.Count - 1
          'if vIncludesDirs(i) = "" then response.write "<br>Folder " & i & " empty"
          vFolder = IncludeTrailingPathDelimiter(vIncludesDirs.Items(i))
          if vFolder <> "" then
            'if vFolder <> "" then if vFSO.FolderExists(vFolder) then vIncludesDirs(i) = ""
            if vLoadAllFilesInDir then vFolder = vFolder + aLibFileName
            vFolder = GetPhysicPath(vFolder)
            if vFolder <> "" and vFSO.FolderExists(vFolder) then
              if vLoadAllFilesInDir then
                Set vFolder = vFSO.GetFolder(vFolder)
                Result = iRequireFiles(aLibName, vFolder)
                Set vFolder = Nothing
                'to cache it
                if Result > 0 and IsBitIn(Options, optCacheEnabled) and isObject(gCache) then 
                  if not (gCache is Nothing) then
                    if not gCache.CacheExists(cLibCachePrefix+aLibName+"*") then gCache.Values(cLibCachePrefix+aLibName+"*") = Result
                  end if
                end if

              elseif FLoadedLibs.IndexOf(aLibName) < 0 then
                'writeln("Folder " & vFolder.Name)
                'Set vFile = FindFile(vFolder, aLibFileName)
                aLibFileName = GetPhysicPath(aLibFileName)
                if vFSO.FileExists(aLibFileName) then
                  Set vFileStream = vFSO.OPenTextFile(aLibFileName, ForReading)
                  if not (vFileStream is Nothing) then
                    'writeln(aLibName + ":"+ aLibFileName)
                    'Set vFileStream = vFile.OpenAsTextStream(ForReading)
                    vFileContent = vFileStream.ReadAll
                    vFileStream.Close
                    Set vFileStream = Nothing
                    'Set vFile = Nothing
                    if vFileContent <> "" then if LoadLib(aLibName, vFileContent) then Result = 1
                  end if
                end if
              end if
              if Result > 0 then Exit For
            end if
          end if
        Next
        Set vFSO = Nothing
        'vIncludesDirs.Trim("")
        ' Now the vIncludesDirs are all the illegal path!
        ' remove these illegal path if any
        'For i = 0 to vIncludesDirs.Count - 1
         ' FIncludeDirs.Remove(vIncludesDirs(i))
        'Next
        Set vIncludesDirs = Nothing
        
       RequireFile = Result
    End Function

    Private Function iLoadLibsInCache(ByRef aLibName)
        Dim i, vLibs, vLib, Result
        Set vLibs = gCache.Contents(cLibCachePrefix + aLibName)
        Result = 0
        i = CInt(gCache.Values(cLibCachePrefix + aLibName + "*"))
        if i > vLibs.Count then gCache.Remove(cLibCachePrefix + aLibName + "*"): Exit Function
        For i = 0 to vLibs.Count - 1
          vLib = vLibs.Items(i)
          if IsArray(vLib) then
            if ubound(vLib) = 2 then if iExec(vLib(2), vLib(1), vLib(0)) then Result = Result + 1
            'writeln(aLibName&".C="&vLib(0))
          end if
        Next
        iLoadLibsInCache = Result
        'writeln(Result)
    End Function

    Private Function iExistsInCache(ByRef aLibName)
        iExistsInCache = False
        if IsBitIn(Options, optCacheEnabled) and IsObject(gCache) then 
          if not (gCache is Nothing) then
            iExistsInCache = gCache.CacheExists(cLibCachePrefix+aLibName)
          end if
        end if
    End Function

    Private Function iGetInCache(ByRef aLibName)
        ' Check Cache:
        if IsBitIn(Options, optCacheEnabled) and IsObject(gCache) then 
          if not (gCache is Nothing) then
            if gCache.CacheExists(cLibCachePrefix+aLibName) then
              iGetInCache = gCache.Values(cLibCachePrefix+aLibName)
            end if
          end if
        end if
    End Function

    Private Function iExec(ByRef pLibName, ByVal aScript, ByRef pScriptType)
        Dim vErrNum, Result
        if FPassword <> "" then
          On Error Resume Next
          vErrNum = 0
          aScript = DefaultDecrypt(aScript, FPassword)
          If Err.Number <> 0 then
            ' &H800a01f4: no defaultDecrypt function.
            if Err.Number <> &H800a01f4 then vErrNum = Err.Number: Result = Err.Desciption
          End if
          On Error Goto 0
          if vErrNum <> 0 then 
            Call RaiseError(vErrNum, "LoadLib: DefaultDecrypt function", Result)
            iExec = False
            Exit Function
          end if
        end if

        Result = ExecuteScript(pLibName, aScript, pScriptType)
        if Result then FLoadedLibs.Add(pLibName)
        iExec = Result
    End Function

    ' Note: internal used, common user just be careful.
    ' return false means can not load!
    Public Function LoadLib(ByRef pLibName, ByVal aScript)
        Dim Result, i, vFirstLine

        'writeln("Load= " & pLibName)
        Result = FLoadedLibs.IndexOf(pLibName) >= 0
        if Result then Exit Function
        if IsNull(aScript) or aScript = "" then Exit Function

        aScript = TrimChar(aScript, vbCRLF)
        i = InStr(1, aScript, vbCRLF, vbTextCompare)
        if i > 1 then
          vFirstLine = Left(aScript, i-1)
          aScript = Right(aScript, Len(aScript) - i - 1)     '-Len(vbCRLF)+1
          i = InStr(1, vFirstLine, "<"&"SCRIPT", vbTextCompare)
          if i > 0 then
            ' check the last
            i = InStrRev(aScript, "<"&"/SCRIPT>", -1, vbTextCompare)
            if i > 0 then
              'remove the last line
              aScript = Mid(aScript, 1, i-1)
              i = GetScriptType(vFirstLine)
              Result = iExec(pLibName, aScript, i)
            end if
          else 'check the %%
            i = InStr(1, vFirstLine, "<"&"%", vbTextCompare)
            if i > 0 then
              ' check the last
              'aScript = Mid(aScript, i+2)
              i = InStrRev(aScript, "%"&">", -1, vbTextCompare)
              if i > 0 then
                aScript = Left(aScript, i-1)
                Result = iExec(pLibName, aScript, stVBScriptType)
                i = stVBScriptType
              end if
           end if
          end if
        end if

        LoadLib = Result

        if Result then
          'to cache it
          if IsBitIn(Options, optCacheEnabled) and isObject(gCache) then 
            if not (gCache is Nothing) then
              Redim Result(2)
              Result(0) = i
              Result(1) = aScript
              Result(2) = pLibName
              if not gCache.CacheExists(cLibCachePrefix+pLibName) then gCache.Values(cLibCachePrefix+pLibName) = Result
            end if
          end if
        else
          Call RaiseError(vbInvalidLibScriptError, "TMeLib.LoadLib", pLibName +" Library is Invalid Lib Script Foramt.")
        end if
    End Function
End Class

%>