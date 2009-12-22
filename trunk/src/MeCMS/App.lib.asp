<SCRIPT Runat="Server" Language="VBScript">

Public Const cActionRequestPostfix = "Request"
Private Const cLoginRequestAction = "LoginRequest"
Private Const cLoginAction = "Login"

Class TMeCMSApp
    Private FDatabase
    Private FConfigCollection
    Private FSession
    Private FUsers
    Private FActions
    Public LoginRetryTimeInterval  'seconds
    Public LoginRetryMaxCount

    Private Sub Class_Initialize()
        Dim SERVER_NAME, SERVER_PORT, SERVER_PORT_SECURE

        LoginRetryTimeInterval = 60
        LoginRetryMaxCount = 5

        Lib.Require("ApplicationCaches")
        Lib.Require("Security.Session")
        Lib.Require("MeDatabase")
        Lib.Require("MeCMS.Security.UserMgr")
        Lib.Require("MeCMS.ActionMgr")

        Set FDatabase    = New TMeDatabase
        Set FSession     = New TMeSession
        Set Lib.Database = FDatabase
        Set FUsers       = New TMeUserMgr
        Set FActions     = New TMeActionMgr
        Set FConfigCollection = Server.CreateObject(DictionaryObjectName)

        SERVER_NAME        = Request.ServerVariables("SERVER_NAME")
        SERVER_PORT        = Request.ServerVariables("SERVER_PORT")
        SERVER_PORT_SECURE = Request.ServerVariables("SERVER_PORT_SECURE")

        If SERVER_PORT_SECURE = 0 Then
            gServerRoot = "http://" & SERVER_NAME
        Else
            gServerRoot = "https://" & SERVER_NAME
        End If
        If SERVER_PORT <> 80 Then
            gServerRoot = gServerRoot & ":" & SERVER_PORT
        End If
        gServerRoot = gServerRoot & Left(SCRIPT_NAME, InStrRev(SCRIPT_NAME, "/"))

    End Sub

    Private Sub Class_Terminate()
        Set FSession  = Nothing
        Set FDatabase = Nothing
        Set FUsers    = Nothing
        Set FActions  = Nothing
        Set gCache    = Nothing
        Set FConfigCollection = Nothing
    End Sub

    Public Property Get Database()
        Set Database = FDatabase
    End Property

    Public Property Get Config()
        Set Config = FConfigCollection
    End Property

    Public Property Get Cache()
        On Error Resume Next
        Set Cache = gCache
        On Error Goto 0
    End Property

    Public Property Get Session()
        Set Session = FSession
    End Property

    Public Property Get Users()
        Set Users = FUsers
    End Property

    Public Property Get Actions()
        Set Actions = FActions
    End Property

    Public Sub Run()
      FDatabase.Open
      ' 运行前准备
      Lib.Require("MeCMS.Init") ' 动态装入运行初始化过程函数库在数据库中。也可以不存在

      ' * 分析参数，根据参数动态装入不同函数库并运行
      ProcessRequest()
      '运行后处理
      Lib.Require("MeCMS.Terminate")
    End Sub

    ' * 分析参数，根据参数动态装入不同函数库并运行
    Public Sub ProcessRequest()
      Dim vActionId, vCatId
      vActionId = Request(cActionIdURLParamName)
      if IsEmpty(vActionId) or vActionId = "" then 
        vActionId = "#"
      else
      end if
      vCatId = Request(cCatIdURLParamName)
      if IsEmpty(vCatId) or vCatId = "" then 
        vCatId = "#"
      else
      end if

      'if not (FUsers.Logined or (vActionId = cLoginRequestAction and vCatId = "#")) then
      'if Not FCurrentUser.HasPermission(vCatId, vActionId) then ' 可以放在执行action前测试！！
      if not FUsers.Logined then
        ' * 检查是否登录？没有则检查是否有Login Action 如有则要求用户登陆，否则就长驱直入。
        '如何区分该动作是显示界面还是返回数据？: 根据返回的类型 t=r表示request 返回数据, t=v表示该动作是界面。
        '放在actions中区分处理。
        if FActions.ActionExists("#", cLoginAction) then
          vCatId = "#"
          vActionId = cLoginAction
        end if
      end if
      'end if

      if FActions.ActionExists(vCatId, vActionId) then FActions.Execute(vCatId, vActionId)
    End Sub

    ' check whether has the permission for the action.!
    Public Function HasRunPermissionForAction(ByRef aCatId, ByRef aActionId);
      Dim Result
      Result = True
      HasRunPermissionForAction = Result
    End Function
End Class

</SCRIPT>

