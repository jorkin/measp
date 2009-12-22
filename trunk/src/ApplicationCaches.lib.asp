<%
Set gCache = New TApplicationCaches

Const cObjCachePrefixName = ".OBJ:"

' Manage the Application Caches
'
' Properties:
'   Property CacheTime: LongInt r/w
'     the cache alived time(minutes)
'   Property Name: String r/w
'     the cache part prefix name in the application
'   Property Values[pCacheName]: Variant r/w
'     get/Set the the value of the cache name. the Cache has been alived only in the CacheTime.
'   Property Count: integer readonly
'     return the application cache count.
'   Sub SetValue(pCacheName, pValue, pNeverExpired);
'     Set the the value of the cache name. you can control the cache live temp or forever
'     U MUST use the SetValue method to set the object.
'   Sub RemoveAll;
'     Clear all the application caches.
'   Sub Remove(pCacheName);
'     Clear the application pCacheName cache.
'   Function CacheExists(pCacheName): Boolean;
'     Test the pCacheName whether Exists/Expired. if Expired or not exists return false. if Expired this method will remove it from application caches..
'   Sub BeginUpdate;
'   Sub EndUpdate;
'     the BeginUpdate/EndUpdate for batch update the cache data.
Class TApplicationCaches
    Private FName, FCacheData
    ' the Cache alive time(min)
    Private FCacheTime
    Private FLockCount

    Private Sub Class_Initialize()
      ' the default cache basename
      FName = "me_"
      FCacheTime = 60  ' 60 min = 1 hour
      FLockCount = 0
    End Sub

    Private Sub Class_Terminate()
    End Sub

    Public Property Get CacheTime()
      CacheTime = FCacheTime
    End Property

    Public Property Let CacheTime(pValue)
      FCacheTime = Clng(pValue)
    End Property

    Public Property Get Name()
      Name = FName
    End Property

    Public Property Let Name(pValue)
      FName = pValue
    End Property

    Public Property Get Contents(ByRef pCacheName)
      Dim Result, vKey, s
      Set Result = New TMeList
      Result.Duplicates = dupAccept
      For Each vKey In Application.Contents
        s = Mid(CStr(vkey), Len(FName)+1)
        if CacheExists(s) then
          if pCacheName = "" then
            Result.Add(Values(s))
          elseif Mid(s,1,Len(pCacheName)) = pCacheName then 
            'writeln(s)
            if Right(s,1) <> "*" then Result.Add(Values(s))
            'writeln("EE")
          end if
        end if
      Next
      Set Contents = Result
    End Property

    Public Property Get Values(ByRef pCacheName)
      FCacheData = Application.Contents(FName & pCacheName)
      if IsArray(FCacheData) then
        if IsObject(FCacheData(0)) then
          Set Values = FCacheData(0)
        else
          Values = FCacheData(0)
        end if
      else
        Values = ""
      end if
    End Property

    Public Property Let Values(ByRef pCacheName, ByRef pValue)
      Call SetValue(pCacheName, pValue, False)
    End Property

    Public Property Get Count()
        Count = Application.Contents.Count
    End Property

    Public Property Get Objects(ByRef pObjId)
      Dim vObjData, Result
      FCacheData = Application.Contents(FName & cObjCachePrefixName & pObjId)
      Set Result = Nothing
      if IsArray(FCacheData) then
        vObjData = FCacheData(0)
        if IsArray(vObjData) then
          On Error Resume Next
          Set Result = ArrayToObject(vObjData)
          On Error goto 0
        end if
      end if
      Set Objects = Result
    End Property

    Public Property Let Objects(pObjectId, pValue)
      Call SetObject(pObjectId, pValue, False)
    End Property

    Public Sub RemoveObject(ByRef pObjId)
      BeginUpdate
      On Error Resume Next
      Call Application.Contents.Remove(FName & cObjCachePrefixName & pObjId)
      EndUpdate
      On Error Goto 0
    End Sub

    Public Function ObjectExists(ByRef pObjId)
      ObjectExists = False
      FCacheData = Application(FName & cObjCachePrefixName & pObjId)
      If Not IsArray(FCacheData) Then Exit Function
      If Not IsDate(FCacheData(1)) Then Exit Function
      If DateDiff("n",CDate(FCacheData(1)),Now()) < FCacheTime Then
        ObjectExists = True
      Else
        ' the cache is expired.
        RemoveObject(pObjId)
      End If
    End Function

    Public Sub SetObject(pObjectId, pValue, pNeverExpired)
      Dim vErrNum, vErrDesc
      ReDim FCacheData(2)
      if IsObject(pValue) then
        FCacheData(0)=ObjectToArray(pValue)
        if pNeverExpired then FCacheData(1)=DateAdd("yyyy",5, Now()) else FCacheData(1)=Now()
        BeginUpdate
        On Error Resume Next
        Application.Contents(FName & cObjCachePrefixName & pObjectId) = FCacheData
        EndUpdate
        vErrNum  = Err.Number
        If vErrNum <> 0 Then
          vErrDesc = Err.Description
        End If
        On Error Goto 0
      else
        vErrNum  = vbInvalidObjAppCacheError
        vErrDesc = "the value is not valid object."
      end if
      If vErrNum <> 0 Then
        Call RaiseError(vErrNum, "ApplicationCache.SetValue", vErrDesc)
      End If
    End Sub

    Public Sub SetValue(pCacheName, pValue, pNeverExpired)
      Dim vErrNum, vErrDesc
      ReDim FCacheData(2)
      if IsObject(pValue) then
        Set FCacheData(0)=pValue
      else
        FCacheData(0)=pValue
      end if
      if pNeverExpired then FCacheData(1)=DateAdd("yyyy",5, Now()) else FCacheData(1)=Now()
      BeginUpdate
      On Error Resume Next
      Application.Contents(FName & pCacheName) = FCacheData
      EndUpdate
      vErrNum  = Err.Number
      If vErrNum <> 0 Then
        vErrDesc = Err.Description
        'Err.Clear
      End If
      On Error Goto 0
      If vErrNum <> 0 Then
        Call RaiseError(vErrNum, "ApplicationCache.SetValue", vErrDesc)
      End If
    End Sub

    Public Sub BeginUpdate
      if FLockCount <= 0 then Application.Lock
      FLockCount = FLockCount + 1
    End Sub

    Public Sub EndUpdate
      FLockCount = FLockCount - 1
      if FLockCount <= 0 then Application.UnLock
    End Sub

    Public Sub RemoveAll()
      BeginUpdate
      On Error Resume Next
      Call Application.Contents.RemoveAll()
      EndUpdate
      On Error Goto 0
    End Sub

    Public Sub Remove(ByRef pCacheName)
      BeginUpdate
      On Error Resume Next
      Call Application.Contents.Remove(FName & pCacheName)
      EndUpdate
      On Error Goto 0
    End Sub

  Public Function CacheExists(ByRef pCacheName)
    CacheExists = False
    FCacheData = Application(FName & pCacheName)
    If Not IsArray(FCacheData) Then Exit Function
    If Not IsDate(FCacheData(1)) Then Exit Function
    If DateDiff("n",CDate(FCacheData(1)),Now()) < FCacheTime Then
      CacheExists = True
    Else
      ' the cache is expired.
      Remove(pCacheName)
    End If
  End Function

End Class

%>
