<SCRIPT Runat="Server" Language="VBScript">

Lib.Require("util.RegExp")

' the action type consts: I USE the atViewTypeId ascii code.
'Public Const atViewType = 0
'Public Const atRequestType = 1

Public Const cActionIdURLParamName = "a"
Public Const cActionTypeURLParamName = "t"
Public Const cCatIdURLParamName = "c"

' use the ASCII number as ParamTypeId
Private Const atViewTypeId = "v"
Private Const atRequestTypeId = "r"

Public Const cActionObjectPrefix = ".act:"

Private Const sqlSelectActionById = "Select * From cms_actions Where cat_id=%CatId% and act_id=%ActId% and act_type=%ActType%"
Private Const sqlSelectActionParamsByActionId = "Select * From cms_action_params Where cat_id=%CatId% and act_id=%ActId% order by parm_order"
Private Const sqlSelectActionParamById = "Select * From cms_action_params Where cat_id=%CatId% and act_id=%ActId% and parm_id=%ParamId%"

Private Dim vbExecuteActionNoPermError

vbExecuteActionNoPermError = 50


Class TMeActionMgr
    Private FCurrentActionType

    Private Sub Class_Initialize()
      FCurrentActionType = Request(cActionTypeURLParamName)
      if IsEmpty(FCurrentActionType) or FCurrentActionType = "" then 
        FCurrentActionType = atViewTypeId 
      else 
        FCurrentActionType = LCase(FCurrentActionType)
      end if
      'FCurrentActionType = Asc(FCurrentActionType)
    End Sub

    Private Sub Class_Terminate()
    End Sub

    Public Default Property Get Action(ByRef aCatId, ByRef aActionId)
      Dim Result, vId
      Set Result = Nothing
      'check cache
      if not IsEmpty(gCache) then
          vId = cActionObjectPrefix + aCatId + ":" + aActionId + ":" + FCurrentActionType
          if gCache.ObjectExists(vId) then Set Result = gCache.Objects(vId)
      end if

      if Result is Nothing then 'check the database
        Set Result = GetActionFromDB(aCatId, aActionId)
      end if

      Set Action = Result
    End Property

    Public Function ActionExists(ByRef aCatId, ByRef aActionId)
      Dim Result, vId
      Result = False
      'check cache
      if not IsEmpty(gCache) then
          vId = cActionObjectPrefix + aCatId + ":" + aActionId + ":" + FCurrentActionType
          Result = gCache.ObjectExists(vId)
      end if
      if not Result and IsBitIn(Lib.Options, optLoadFromDB) then 'check the database
        Result = not (GetActionFromDB(aCatId, aActionId) is Nothing)
      end if
      'if not Result and IsBitIn(Lib.Options, optLoadFromFile) then 'check the file
      '  Result = not (GetActionFromFile(aCatId, aActionId) is Nothing)
      'end if

      ActionExists = Result
    End Function

    'if action exists and execute successful then return true.
    Public Function Execute(ByRef aCatId, ByRef aActionId)
      Dim vAction, Result
      'check the permission before run!!
      Result = gApplication.HasRunPermissionForAction(aCatId, aActionId)
      if Result then
        Set vAction = Action(aCatId, aActionId)
        Result = not (vAction is Nothing)
        if Result then
          Result = vAction.Run()
        end if
      else
        Call RaiseError(vbExecuteActionNoPermError, "Action Execute", "you have no permission to execute " & aCatId & ":" & aActionId)
      end if
      Execute = Result
    End Function

    Public Function RegisterAction(ByRef aAction)
      RegisterAction = aAction.Save()
    End Function

    Private Function GetActionFromDB(ByRef aCatId, ByRef aActionId)
        Set Result = New TMeAction
        Result.Id = aActionId
        Result.CatId = aCatId
        Result.ActionType = Asc(FCurrentActionType)
        if Result.Fetch("") then
          if not IsEmpty(gCache) then gCache.Objects(cActionObjectPrefix+Result.ObjectId) = Result
        else
          Set Result = Nothing
        end if
        Set GetActionFromDB = Result
    End Function
End Class

Class TMeAction
    Private FObjectStatus
    Private FId
    Public  CatId, Name, ActionType, ActionLib, ActionClass, ActionFunc
    Private FParams

    Private Sub Class_Initialize()
      FObjectStatus = osInit
      Set FParams = New TMeParams
    End Sub

    Private Sub Class_Terminate()
      Set FParams = Nothing
    End Sub

    Public Property Get ClassName()
      ClassName = "TMeAction"
    End Property

    Public Property Get ObjectId()
      ObjectId = CatId + ":" + Id + ":"+ Chr(ActionType)
    End Property

    Public Property Get Id()
      Id = FId
    End Property

    Public Property Let Id(ByRef aValue)
      FId = TrimURLParamName(aValue)
    End Property

    Public Property Get ObjectStatus()
      ObjectStatus = FObjectStatus
    End Property

    Public Property Let ObjectStatus(ByVal aValue)
      FObjectStatus = aValue
    End Property

    Public Function GetMetaObject()
      Dim Result,v
      Set Result = New TMeMetaObject
      Result.ClassName = ClassName()
      v = "Id:ftString,CatId:ftString,Name:ftString,ActionType:ftInteger"_
        + ",ActionLib:ftString,ActionClass:ftString,ActionFunc:ftString"_
        + ",Params:ftCollection"_
        + ",ObjectStatus:ftInteger"
      Result.AssignFieldsFromString(v)
      Set GetMetaObject = Result
    End Function

    Public Property Get Params()
        Set Params = FParams
    End Property

    Public Property Let Params(ByRef aValue)
        FParams.Assign(aValue)
    End Property

    Public Function Fetch(ByRef aObjectId)
      Dim vId, Result, vRS, vSQL
      Result = False
      if aObjectId <> "" then
        vId = Split(aObjectId, ":")
        if UBound(vId) = 2 then
          CatId = vId(0)
          Id = vId(1)
          ActionType = vId(2)
        end if
      end if
      if FId <> "" then
        With gApplication.Database
          vSQL = Replace(sqlSelectActionById, "%ActId%", .QuotedStr(Id))
          vSQL = Replace(vSQL, "%ActType%", ActionType)
          vSQL = Replace(vSQL, "%CatId%", .QuotedStr(CatId))
          Set vRS = .OpenTable(vSQL,  ForReading)
        End With
        if not (vRS is Nothing) then
          Result = not vRS.BoF
          if Result then
            Id = vRS("act_id").Value
            ActionType = vRS("act_type").Value
            CatId = vRS("cat_id").Value
            Name = vRS("act_name").Value
            ActionLib = vRS("act_lib_name").Value
            ActionClass = vRS("act_class_name").Value
            ActionFunc = vRS("act_sub_name").Value
          end if
          vRS.Close
          Set vRS = Nothing
        end if
      end if
      Fetch = Result
      if Result then 
        FParams.Fetch(CatId, Id)
        FObjectStatus = osLoaded
      end if
    End Function

    'save to database
    Public Function Save()
        Dim Result, vSQL, vRS
        Result = (FId <> "")
        if Result then
          With gApplication.Database
            vSQL = Replace(sqlSelectActionById, "%ActId%", .QuotedStr(Id))
            vSQL = Replace(vSQL, "%ActType%", ActionType)
            vSQL = Replace(vSQL, "%CatId%", .QuotedStr(CatId))
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
            vRS("act_id").Value = Id
            vRS("act_parentid").Value = 0
            vRS("act_type").Value = ActionType
            vRS("cat_id").Value = CatId
            vRS("act_name").Value = Name
            vRS("act_lib_name").Value = ActionLib
            vRS("act_class_name").Value = ActionClass
            vRS("act_sub_name").Value = ActionFunc
          end if
          if Result then vRS.Update
          vRS.Close
          Set vRS = Nothing
        end if
        Save = Result
        if Result then FParams.Save(CatId, FId):FObjectStatus = osLoaded
    End Function

    Public Function Run()
      Dim Result, vFunc
      Result = (ActionFunc <> "")
      if Result then
        if ActionLib <> "" then Result = Lib.Require(ActionLib)
        if Result then
          vFunc = ""
          if ActionClass <> "" then 
            vFunc = "g" + Mid(ActionClass, 2)
            if Eval("not IsObject("+vFunc+") or ("+vFunc+" is Nothing)") then ExecuteGlobal("Set " + vFunc + "=New "+ ActionClass)
            vFunc = vFunc + "."
          end if
          vFunc = vFunc + ActionFunc + "(" + GetParamsValueString + ")"
          Result = EVal(vFunc)
        end if
      end if
      Run = Result
    End Function

    Private Function GetParamsValueString()
      Dim i, Result
      Result = ""
      for i = 0 to FParams.Count - 1
        if Result <> "" then Result = Result + ","
        'Result = Result + QuotedString(FList(i).Value, """")
        Result = Result + "FParams(" +CStr(i) + ").Value"
      next 'i
      GetParamsValueString = Result
    End Function
End Class

Class TMeParams
    Private FList

    Private Sub Class_Initialize()
      Set FList = New TMeList
    End Sub

    Private Sub Class_Terminate()
      Set FList = Nothing
    End Sub

    Public Property Get ClassName()
      ClassName = "TMeParams"
    End Property

    Public Default Property Get Items(ByVal Index)
      Set Items = FList(Index)
    End Property

    Public Function Count()
      Count = FList.Count
    End Function

    Public Function Add(ByRef aItem)
      if IsClass(aItem, "TMeParam") then Add = FList.Add(aItem) else Add = -1
    End Function

    Public Function Fetch(ByRef aCatId, ByRef aActionId)
      Dim vParam, Result, vRS, vSQL
      FList.Clear()
      Result = (aActionId <> "")
      if Result then
        With gApplication.Database
          vSQL = Replace(sqlSelectActionParamsByActionId, "%ActId%", .QuotedStr(aActionId))
          vSQL = Replace(vSQL, "%CatId%", .QuotedStr(aCatId))
          Set vRS = .OpenTable(vSQL,  ForReading)
        End With
        if not (vRS is Nothing) then
          Result = not vRS.EoF
          Do While not vRS.EoF
            Set vParam = New TMeParam
            vParam.ParamType = vRS("parm_type").Value
            vParam.Id = vRS("parm_id").Value
            vParam.Name = vRS("parm_name").Value
            vParam.ObjectStatus = osLoaded
            FList.Add(vParam)
            vRS.MoveNext
          Loop
          vRS.Close
          Set vRS = Nothing
        end if
      end if
      Fetch = Result
    End Function

    'save to database
    Public Function Save(ByRef aCatId, ByRef aActionId)
        Dim Result, vSQL, vRS, i
        Result = (aActionId <> "")
        if Result then
          For i = 0 to Count -1
            FList(i).Save(aCatId, aActionId)
          Next
        end if
        Save = Result
    End Function

End Class

Class TMeParam
    Private FObjectStatus
    Private FId
    Public  Name, ParamType, Value

    Private Sub Class_Initialize()
      FObjectStatus = osInit
    End Sub

    Private Sub Class_Terminate()
    End Sub

    Public Property Get ClassName()
      ClassName = "TMeParam"
    End Property

    Public Property Get ObjectId()
      ObjectId = Id
    End Property

    Public Property Get Id()
      Id = FId
    End Property

    ' the ParamType MUST be set first.
    Public Property Let Id(ByRef aValue)
      FId = TrimURLParamName(aValue)
      Value = Request(FId)
      'Value = TypeCast(Request(FId), ParamType)
    End Property

    Public Property Get ObjectStatus()
      ObjectStatus = FObjectStatus
    End Property

    Public Property Let ObjectStatus(ByVal aValue)
      FObjectStatus = aValue
    End Property

    Public Function GetMetaObject()
      Dim Result,v
      Set Result = New TMeMetaObject
      Result.ClassName = ClassName()
      v = "Id:ftString,Name:ftString,ParamType:ftInteger"_
        + ",ObjectStatus:ftInteger"
      Result.AssignFieldsFromString(v)
      Set GetMetaObject = Result
    End Function

    'save to database
    Public Function Save(ByRef aCatId, ByRef aActionId)
        Dim Result, vSQL, vRS
        Result = (aActionId <> "") and (FId <> "")
        if Result then
          With gApplication.Database
            vSQL = Replace(sqlSelectActionParamById, "%ActId%", .QuotedStr(aActionId))
            vSQL = Replace(vSQL, "%CatId%", .QuotedStr(aCatId))
            vSQL = Replace(vSQL, "%ParamId%", .QuotedStr(FId))
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
            vRS("cat_id").Value = aCatId
            vRS("act_id").Value = aActionId
            vRS("parm_id").Value = vParam.Id
            vRS("parm_type").Value = vParam.ParamType
            vRS("parm_name").Value = vParam.Name
          end if
          if Result then vRS.Update
          vRS.Close
          Set vRS = Nothing
        end if
        Save = Result
        if Result then FObjectStatus = osLoaded
    End Function
End Class

Private Function TrimURLParamName(ByRef aValue)
      Dim Result
      Result = gRegExp.s(aValue, "(\w{1,10})", "$1", True, False)
      TrimURLParamName = Result
End Function
</SCRIPT>
