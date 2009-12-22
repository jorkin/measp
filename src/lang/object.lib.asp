<SCRIPT Runat="Server" Language="VBScript">

' ---------------------------------------------------------------------------
'      $Source: /home/cvs/MeCMS/src/lang/object.lib.asp,v $
'      $Revision: 1.4 $
'      $Author: riceball $
' ---------------------------------------------------------------------------

Public Const ftString     = &H10
Public Const ftPassword   = &H12
Public Const ftMemo       = &H13
Public Const ftMemoRTF    = &H14

Public Const ftInteger  = &H20
Public Const ftFloat    = &H21
Public Const ftCurreny  = &H22
Public Const ftDateTime = &H23
Public Const ftDate     = &H24
Public Const ftTime     = &H25
Public Const ftBoolean  = &H26

Public Const ftURL      = &H30
Public Const ftEmailURL = &H31
Public Const ftImageURL = &H32
Public Const ftFlashURL = &H33
Public Const ftSoundURL = &H34

Public Const ftObject     = &H40
Public Const ftCollection = &H41 ' the collection object MUST have ClassName, Items(i), Add(Value) and Count !! the first i is 0!

' the object status constants
Public Const osInit   = 0
Public Const osLoaded = 1 ' already fetched from database

Class TMeMetaObject
    ' 对象的类名，指示是哪一个对象类的 MetaInfo.
    Public  ClassName
    Private FFields

    Private Sub Class_Initialize()
        Set FFields = New TMeMetaFields
    End Sub

    Private Sub Class_Terminate()
        Set FFields = Nothing
    End Sub

    '属性列表对象
    Public Property Get Fields()
        Set Fields = FFields
    End Property

    '返回指定属性名的属性对象
    Public Default Property Get FieldByName(ByRef aFieldName)
        Set FieldByName = FFields.FieldByName(aFieldName)
    End Property

    '新增属性，返回新增属性对象
    Public Function AddField()
        Dim Result
        Set Result = New TMeMetaField
        if FFields.Add(Result) < 0 then
          Set Result = Nothing
        end if
        Set AddField = Result
    End Function

     '从字符串建立对象的所有属性。字符串格式： "属性名:属性类型[:Size:Validator:Constraints:Required:Confirmed:Filters],..." 
     '属性之间用逗号“,”分隔，属性名和它的相关类型之间用冒号“:”分隔，数据之间不能有空格,如果是字符串必须用引号。
     '示例： "Id:ftString:32:true:false:""the Validator"":""the Constraints"":"the filters",Password:ftPassword:32:true:true"
    Public Sub AssignFieldsFromString(ByRef aText)
      Dim i, vFields
      vFields = Split(aText, ",")
      FFields.Clear
      For i = 0 to UBound(vFields)
        FFields.Append(vFields[i])
      Next
    End Sub

End Class

'MetaInfo 的属性字段列表类
Class TMeMetaFields
    Private FList

    Private Sub Class_Initialize()
      Set FList = New TMeList
    End Sub

    Private Sub Class_Terminate()
      Set FList = Nothing
    End Sub

    '字段列表项，索引从0到(Count-1)
    Public Default Property Get Items(ByRef Index)
      Set Items = FList(Index)
    End Property

    '根据字段名返回字段属性对象
    Public Property Get FieldByName(ByRef aFieldName)
        Dim i, Result
        For i = 0 to FList.Count - 1
          Set Result = FList(i)
          if Result.Name = aFieldName then Set FieldByName = Result : Exit Property
        Next 'i
        Set FieldByName = Nothing
    End Property

    Public Function IndexOf(ByRef aFieldName)
        Dim i
        For i = 0 to FList.Count - 1
          if FList(i).Name = aFieldName then IndexOf = i : Exit Function
        Next 'i
        IndexOf = -1
    End Function

    '从字符串建立一个属性字段，字段名和它的相关类型之间用冒号“:”分隔，第一个是字段名，接着的则是字段的属性。
     '示例： "Id:ftString:32:true:false:""the Validator"":""the Constraints"":""the filters"""
    Public Function Append(ByRef aText)
        Dim Result, vTemp
        vTemp = Split(aText, ":")
        if IndexOf(vTemp(0)) < 0 then
          Set Result = New TMeMetaField
          Result.Assign(vTemp)
          FList.Add(Result)
        else
          Set Result = Nothing
        end if
        Set Append = Result
    End Function

    '属性个数
    Public Function Count()
      Count = FList.Count
    End Function

    '清除所有的属性
    Public Sub Clear()
      FList.Clear()
    End Sub
End Class

'对于comboBox 和ListBox的数据怎么处理？ 作为 ftCollection 类型处理。
Class TMeMetaField
    Private FItems(7)
    'Public Name, FieldType, Size
    'Public Required  'Boolean
    'Public Confirmed 'Boolean 真则该字段需要输入两遍
    'Public Validator ' the js condition script.
    'Public Constraints ' the js condition script.
    'Public Filters ' the js filter value script.

    Private Sub Class_Initialize()
        Set FFields = New TMeMetaFields
    End Sub

    Private Sub Class_Terminate()
        Set FFields = Nothing
    End Sub

    Public Property Get Name()
        Name = FItems(0)
    End Property

    Public Property Let Name(ByRef aValue)
        FItems(0) = aValue
    End Property

    Public Property Get FieldType()
        FieldType = FItems(1)
    End Property

    Public Property Let FieldType(ByRef aValue)
        FItems(1) = CByte(aValue)
    End Property

    Public Property Get Size()
        Size = FItems(2)
    End Property

    Public Property Let Size(ByRef aValue)
        FItems(2) = CLng(aValue)
    End Property

    Public Property Get Required()
        Required = FItems(3)
    End Property

    Public Property Let Required(ByRef aValue)
        FItems(3) = CBool(aValue)
    End Property

    Public Property Get Confirmed()
        Confirmed = FItems(4)
    End Property

    Public Property Let Confirmed(ByRef aValue)
        FItems(4) = CBool(aValue)
    End Property

    Public Property Get Validator()
        Validator = FItems(5)
    End Property

    Public Property Let Validator(ByRef aValue)
        FItems(5) = aValue
    End Property

    Public Property Get Constraints()
        Constraints = FItems(6)
    End Property

    Public Property Let Constraints(ByRef aValue)
        FItems(6) = aValue
    End Property

    Public Property Get Filters()
        Filters = FItems(7)
    End Property

    Public Property Let Filters(ByRef aValue)
        FItems(7) = aValue
    End Property

    Public Sub Assign(ByRef aArray)
      Dim i, vArraySize
      vArraySize = UBound(aArray)
      For i = 0 to UBound(FItems)
        if i <= vArraySize then
          FItems(i) = Eval(aArray(i))
        else
          FItems(i) = Empty
        end if
      Next
    End Sub
End Class

'将数组的值复制到对象中,注意数组的第一个(0)为类名！
Public Function ArrayToObject(ByRef pArray)
    Dim i, vMetaObject, Result
    set Result = Eval("New "+pArray(0))
    if IsObject(Result) then
      Set vMetaObject = pObject.GetMetaObject
      For i = 1 to UBound(pArray)
        Select Case vMetaObject.Fields(i).FieldType
          Case ftCollection
            Execute("Result."+vMetaObject.Fields(i).Name + "=ArrayToCollection(pArray("+CStr(i)+"))")
          Case ftObject
            Execute("Set Result."+vMetaObject.Fields(i).Name + "=ArrayToObject(pArray("+CStr(i)+"))")
          Case else
            Execute("Result."+vMetaObject.Fields(i).Name + "=pArray("+CStr(i)+")")
        End Select
      Next 'i
      Set vMetaObject = Nothing
    end if
    Set ArrayToObject = Result
End Function

Public Function ObjectToArray(ByRef pObject)
    Dim i, Result, vMetaObject
    Set vMetaObject = pObject.GetMetaObject()
    Redim Result(vMetaObject.Count)
    if IsArray(Result) then
      Result(0) = vMetaObject.ClassName
      For i = 1 to UBound(Result)
        'Execute("Result("+i+")" + "=pObject."+vProps(i))
        Select Case vMetaObject.Fields(i).FieldType
          Case ftCollection
            Result(i) = CollectionToArray(Eval("pObject."+vMetaObject.Fields(i).Name))
          Case ftObject
            Result(i) = ObjectToArray(Eval("pObject."+vMetaObject.Fields(i).Name))
          Case else 
            Result(i) = Eval("pObject."+vMetaObject.Fields(i).Name)
        End Select
      Next 'i
    end if
    Set vMetaObject = Nothing
    ObjectToArray = Result
End Function

Public Function CollectionToArray(ByRef pObject)
    Dim i, Result
    Redim Result(pObject.Count)
    Result(0) = pObject.ClassName
    for i =1 to UBound(Result)
        Select Case VarType(pObject.Items(i-1))
          Case vbObject
            Result(i) = ObjectToArray(pObject.Items(i-1))
          Case else 
            Result(i) = pObject.Items(i-1)
        End Select
    next 'i
    CollectionToArray = Result
End Function

Public Function ArrayToCollection(ByRef pArray)
    Dim i, Result, vItem, s
    set Result = Eval("New "+pArray(0))
    if IsObject(Result) then
      For i = 1 to UBound(pArray)
        vItem = pArray(i)
        Select Case VarType(vItem)
          Case vbArray
            if IsClassName(vItem(0)) then 
              Execute("Result.Add(ArrayToObject(vItem))")
            else
              Execute("Result.Add(vItem)")
            end if
          Case else
            Execute("Result.Add(vItem)")
        End Select
      Next 'i
      Set vMetaObject = Nothing
    end if
    Set ArrayToCollection = Result
End Function

Public Function IsClassName(ByRef aString)
  Dim Result
  Result = (VarType(aString) = vbString)
  if Result then Result = (Len(aString) > 1)
  if Result then Result = (Mid(aString, 1, 1) = "T")
End Function

Public Function IsClass(aObject, ByRef aClassName)
  Dim Result
  Result = (VarType(aObject) = vbObject)
  if Result then 
    Result = False
    On Error Resume Next
    Result = (aObject.ClassName = aClassName)
    On Error Goto 0
  end if
End Function

Function MakeGlobalObjectId(ByRef pObject)
  MakeGlobalObjectId = MakeGlobalObjectIdBy(pObject.ClassName, pObject.ObjectId)
End Function

Function MakeGlobalObjectIdBy(ByRef pClassName, ByRef pId)
  MakeGlobalObjectIdBy = pClassName + ":" + pId
End Function

Function TypeCast(ByRef aValue, ByVal aType)
  Dim Result
  if not IsEmpty(aValue) then
    Select Case aType
      Case ftString, ftURL, ftPassword, ftMemo, ftMemoRTF, ftEmailURL, ftImageURL, ftFlashURL, ftSoundURL
        Result = CStr(aValue)
      Case ftInteger
        Result = CLng(aValue)
      Case ftFloat
        Result = CDbl(aValue)
      Case ftCurreny
        Result = CCur(aValue)
      Case ftDate, ftTime, ftDateTime
        Result = CDate(aValue)
      Case ftBoolean
        Result = CBool(aValue)
    End Select
  end if
  TypeCast = Result
End Function

</SCRIPT> 
