<SCRIPT Runat="Server" Language="VBScript">

' ---------------------------------------------------------------------------
'      $Source: /home/cvs/MeCMS/src/MeCMS/Security/UserMgr.lib.asp,v $
'      $Revision: 1.10 $
'      $Author: riceball $
' ---------------------------------------------------------------------------

Lib.Require("Security.Hash.MD5")
Lib.Require("lang.object")

' the login Result constants
Const lsLoginFailed = 0 ' user or password wrong
Const lsLogined = 1
Const lsLoginDisabled = 2

Const cCurrentUserKey ="MeUser"

Private Const sqlSelectUserInfoById = "Select * From cms_users Where usr_name=%Id%"
Private Const sqlUpdateUserRetryCount =  "Update cms_users Set %Enabled% usr_last_retry_time = GetDate(), usr_last_retry_count = %LastRetryCount%"_ 
             & ", usr_retry_count = usr_retry_count + 1 Where usr_name=%Id%"
Private Const sqlUpdateLoginCount = "Update cms_users Set usr_login_count = usr_login_count + 1 Where usr_name=%Id%"
Private Const sqlUpdateLoginCount = "Update cms_users Set usr_login_count = usr_login_count + 1 Where usr_name=%Id%"

Class TMeUserMgr
    Private FCurrentUser

    Private Sub Class_Initialize
      if gApplicaion.Session.ObjectExists(cCurrentUserKey) then 
        Set FCurrentUser = gApplicaion.Session.Objects(cCurrentUserKey)
      else 
        Set FCurrentUser = Nothing
      end if
    End Sub

    Private Sub Class_Terminate
      Set FCurrentUser = Nothing
    End Sub

    Public Property Get Logined()
        Logined = not (FCurrentUser is Nothing)
    End Property

    Public Property Get CurrentUser()
        Set CurrentUser = FCurrentUser
    End Property

    'see the login Result constants
    Public Function Login(ByRef pUserName, ByVal pUserPwd)
        Dim Result
        Result = lsLoginFailed
        if Logined then Logout
        Set FCurrentUser = New TMeUserInfo
        if FCurrentUser.Fetch(pUserName) then
          With FCurrentUser
            if .Enabled then
              pUserPwd = MD5(pUserName + MD5(pUserPwd))
              if pUserPwd = .Password then 
                Result = lsLogined
                .UpdateLoginCount()
                gApplicaion.Session.Objects(cCurrentUserKey) = FCurrentUser
              else
                .UpdateRetryCount()
              end if
            else
              Result = lsLoginDisabled
            end if
          End With
        else
          Set FCurrentUser = Nothing
        end if
        Login = Result
    End Function

    Public Sub Logout()
        gApplicaion.Session.RemoveObject(cCurrentUserKey)
        Set FCurrentUser = Nothing
    End Sub

End Class

Class TMeUserInfo
    Private FObjectStatus
    Private FId
    Public FPassword
    Public Enabled, Language, Creator, CreationDate, UpdateDate, Description
    Public LoginCount, RetryCount, LastRetryCount, LastRetryTime

    Private Sub Class_Initialize
      FObjectStatus = osInit
      LoginCount = 0
      LastRetryCount = 0
      RetryCount = 0
      Enabled = False
      CreationDate = Now()
      UpdateDate = Now()
      if not (gApplicaion.Users.CurrentUser is Nothing) then Creator = gApplicaion.Users.CurrentUser.Id
    End Sub

    Private Sub Class_Terminate
    End Sub

    Public Property Get ClassName()
      ClassName = "TMeUserInfo"
    End Property

    Public Property Get ObjectId()
      ObjectId = Id
    End Property

    Public Property Get Id()
      Id = FId
    End Property

    Public Property Let Id(ByVal aValue)
      FId = TrimId(aValue)
    End Property

    Public Property Get Password()
      Password = FPassword
    End Property

    Public Property Let Password(ByVal aValue)
      FPassword = MD5(FId + MD5(aValue))
    End Property

    Public Property Get ObjectStatus()
      ObjectStatus = FObjectStatus
    End Property

    Public Property Let ObjectStatus(ByVal aValue)
      FObjectStatus = aValue
    End Property

     '现在返回 TMetaObject 对象了！
    Public Function GetMetaObject()
      Dim Result,v
      Set Result = New TMeMetaObject
      Result.ClassName = ClassName()
      v = "Id:ftString,Password:ftPassword,Enabled:ftBoolean"_
        + ",Language:ftString,Creator:ftString,CreationDate:ftDateTime,UpdateDate:ftDateTime,Description:ftMemo"_
        + ",LoginCount:ftInteger,RetryCount:ftInteger,LastRetryCount:ftInteger,LastRetryTime:ftDateTime"_
        + ",ObjectStatus:ftInteger"
      Result.AssignFieldsFromString(v)
      Set GetMetaObject = Result
    End Function

    Public Sub UpdateLoginCount()
        Dim vSQL
          if Id <> "" then
            With gApplication.Database
              vSQL = Replace(sqlUpdateLoginCount, "%Id%", .QuotedStr(Id))
              .Execute(vSQL)
            End With
          end if
      'GetClientIp()
    End Sub

    ' if retryCount >= gApplication.LoginRetryMaxCount dury the LoginRetryTimeInterval then it will block this user.
    Public Sub UpdateRetryCount
        Dim vEnabledStr, vSQL
        if Id <> "" then
          if IsDate(LastRetryTime) then
            if DateDiff("s",Now(), LastRetryTime) <= gApplication.LoginRetryTimeInterval then
              if LastRetryCount >= gApplication.LoginRetryMaxCount then
                vEnabledStr = "usr_enabled = 0, "
                LastRetryCount = Null
                Enabled = False
              else 
                vEnabledStr = ""
              end if
            else ' reset the monitor
              LastRetryCount = 0
            end if
           end if
          LastRetryTime = Now()
          if IsNull(LastRetryCount) then LastRetryCount = 0 else LastRetryCount = LastRetryCount + 1
          if IsNull(RetryCount) then RetryCount = 0 else RetryCount = RetryCount + 1
          vSQL = Replace(sqlUpdateUserRetryCount, "%Enabled%", vEnabledStr)
          vSQL = Replace(vSQL, "%LastRetryCount%", CStr(LastRetryCount))
          With gApplication.Database
            vSQL = Replace(vSQL, "%Id%", .QuotedStr(Id))
            .Execute(vSQL)
          End With
        end if
    End Sub

    'Note: the pUserId is the ObjectId!!
    Public Function Fetch(ByRef pUserId)
        Dim Result, vRS, vSQL
        Result = False
        With gApplication.Database
          vSQL = Replace(sqlSelectUserInfoById, "%Id%", .QuotedStr(TrimId(pUserId)))
          Set vRS = .OpenTable(vSQL,  ForReading)
        End With
        if not (vRS is Nothing) then
          Result = not vRS.BoF
          if Result then
            FId = vRS("usr_name").Value
            FPassword = vRS("usr_password").Value
            Enabled = vRS("usr_enabled").Value
            Language = vRS("usr_language").Value
            Creator = vRS("usr_creator").Value
            CreationDate = vRS("usr_creationdate").Value
            UpdateDate = vRS("usr_updatedate").Value
            Description = vRS("usr_description").Value
            LoginCount = vRS("usr_login_count").Value
            RetryCount = vRS("usr_retry_count").Value
            LastRetryCount = vRS("usr_last_retry_count").Value
            LastRetryTime  = vRS("usr_last_retry_time").Value
          end if
          vRS.Close
          Set vRS = Nothing
        end if
        Fetch = Result
        if Result then FObjectStatus = osLoaded
    End Function

    'save to database
    Public Function Save()
        Dim Result, vSQL, vRS
        Result = (FId <> "")
        if Result then
          With gApplication.Database
            vSQL = Replace(sqlSelectUserInfoById, "%Id%", .QuotedStr(FId))
            Set vRS = .OpenTable(vSQL,  ForWriting)
          End With
          Result = vRS.BOF
          Select Case FObjectStatus
            Case osInit
               if Result then vRS.Append
            Case osLoaded
               Result = not Result 'that should be true.
          End Select
          if Result then 'else someothers update this
            vRS("usr_name").Value = FId
            vRS("usr_password").Value = FPassword
            vRS("usr_enabled").Value = Enabled
            vRS("usr_language").Value = Language
            vRS("usr_creator").Value = Creator
            vRS("usr_creationdate").Value = CreationDate
            vRS("usr_updatedate").Value = UpdateDate
            vRS("usr_description").Value = Description
            vRS("usr_login_count").Value = LoginCount
            vRS("usr_retry_count").Value = RetryCount
            vRS("usr_last_retry_count").Value = LastRetryCount
            vRS("usr_last_retry_time").Value = LastRetryTime
          end if
          if Result then vRS.Update
          vRS.Close
          Set vRS = Nothing
        end if
        Save = Result
        if Result then FObjectStatus = osLoaded
    End Function

    Private Function TrimId(ByRef aValue)
      Dim Result
      Result = Replace(aValue, "%", "")
      Result = Replace(Result, "'", "")
      Result = Replace(Result, """", "")
      Result = Replace(Result, ";", "")
      TrimId = Result
    End Function

End Class
</SCRIPT> 
