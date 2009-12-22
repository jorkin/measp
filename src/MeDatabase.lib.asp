<SCRIPT Runat="Server" Language="VBScript">
  
Lib.Require("ADOConsts")

' possible values for DBType
Const DB_ACCESS      = 0
Const DB_SQLSERVER   = 1
Const DB_ORACLE      = 2
Const DB_MYSQL       = 3
Const DB_POSTGRESQL  = 4

'Note: sql 语句使用 MSSql 标准书写，然后自动转换成其它数据库SQL标准

Class TMeDatabase
    Private FConn, FDBType

    Private Sub Class_Initialize()
        FDBType = DB_ACCESS
        Set FConn = Server.CreateObject("ADODB.Connection")
    End Sub

    Private Sub Class_Terminate()
        On Error Resume Next
        Close ' close the connection first
        Set FConn = Nothing
        On Error Goto 0
    End Sub
    
    Public Property Get Conn()
        Set Conn = FConn
    End Property

    Public Property Get DBType()
        DBType = FDBType
    End Property

    Public Property Let DBType(ByRef pValue)
        FDBType = pValue
    End Property

    Public Property Get ConnectionString()
        ConnectionString = FConn.ConnectionString
    End Property

    'ConnectionString 属性在连接关闭时为读/写，在连接打开时为只读。
    Public Property Let ConnectionString(ByRef pValue)
        if not Active then FConn.ConnectionString = pValue
    End Property

    Public Property Get Active()
        Active = (FConn.State = adStateOpen)
    End Property

    Public Property Let Active(ByRef pValue)
        if pValue then 
          Open()
        else
          Close()
        end if
    End Property

    Public Property Get Delimiters()
        Select Case FDBType
          Case DB_ACCESS, DB_SQLSERVER
            Delimiters = "[]"
          Case Else
            Delimiters = """"""
        End Select
    End Property

    Public Property Get Quote()
        Select Case FDBType
          Case DB_ACCESS, DB_SQLSERVER
            Quote = "'"
          Case Else
            Quote = "'"
        End Select
    End Property

    Public Property Get Wildcard()
        Wildcard = "%"
    End Property

    ' Close the database 
    Public Sub Close()
        if FConn.State = adStateOpen then FConn.Close
    End Sub

    ' Open the database, return the connection. note: you must close the connection and assigned the ConnectionString before open. 
    Public Function Open()
        if FConn.State = adStateClosed then FConn.Open
        Set Open = FConn
    End Function

    ' Open the database, return the connection. note: you must close it before open. 
    Public Function iOpen(ByRef pConnStr)
        if FConn.State = adStateClosed then FConn.Open pConnStr
        Set iOpen = FConn
    End Function

    Public Sub BeginTrans()
        If FDBType <> DB_MYSQL and FConn.State = adStateOpen Then
            FConn.BeginTrans()
        End If
    End Sub

    Public Sub CommitTrans()
        If FDBType <> DB_MYSQL and FConn.State = adStateOpen Then
            FConn.CommitTrans()
        End If
    End Sub

    Public Sub RollbackTrans()
        If FDBType <> DB_MYSQL and FConn.State = adStateOpen Then
            FConn.RollbackTrans()
        End If
    End Sub

    Public Function Execute(ByRef pSql)
        if FConn.State = adStateOpen then Set Execute = FConn.Execute(TranalateSql(pSql)) else Set Execute = Nothing
    End Function

    ' the sql should like this "select count(*) from XXX"
    Public Function GetStatCountBy(ByRef pSql)
      Dim vRS

        GetStatCountBy = 0
        Set vRS = Execute(pSql)
        if not (vRS is Nothing) then
          if not vRS.BOF then GetStatCountBy = vRS.Fields(0)
          vRS.Close
          Set vRS = Nothing
        end if
    End Function

    Public Function CreateRecordSet()
        if FConn.State = adStateOpen then
          Set CreateRecordSet = Server.CreateObject("ADODB.RecordSet")
        else
          Set CreateRecordSet = Nothing
        end if
    End Function

    ' Open the recordset, 
    Public Function OpenRecordSet(ByRef aSource, ByRef aCursorType, ByRef aLockType)
        Dim vResult

        Set vResult = CreateRecordSet
        if not (vResult is Nothing) then
          vResult.Open TranalateSql(aSource), FConn, aCursorType, aLockType
        end if
        Set OpenRecordSet = vResult
    End Function

    ' Open the Table, if not readonly the default adLockOptimistic, adOpenForwardOnly
    Public Function OpenTable(ByRef aTableName, ByRef aReadOnly)
        Dim vLockType

        if aReadOnly = ForReading then vLockType = adLockReadOnly else vLockType = adLockOptimistic
        Set OpenTable = OpenRecordSet(aTableName, adOpenForwardOnly, vLockType)
    End Function

    Public Function FilterQuote(ByRef pText)
        FilterQuote = Replace(pText, Quote, "")
    End Function

    ' Returns the quoted version of a string.
    'Use QuotedStr to convert the string S to a quoted string. 
    ' A single quote is inserted at the beginning and end of S, and each single quote character in the string is repeated.
    Public Function QuotedStr(ByRef pText)
        Dim Result, vQuote
        vQuote = Quote
        Result = Replace(pText, vQuote, vQuote+vQuote)
        Result = vQuote + Result + vQuote
        QuotedStr = Result
    End Function

    Public Function TranalateSql(ByRef pSql)
        Dim Result
        Select Case FDBType
          Case DB_ACCESS
            Result = SqlToAccess(pSql)
          Case Else
            Result = pSql
        End Select
      TranalateSql = Result
    End Function

  Private Function SqlToAccess(ByVal Sql)
    Dim regEx, Matches, Match
    '创建正则对象
    Set regEx = New RegExp
    regEx.IgnoreCase = True
    regEx.Global = True
    regEx.MultiLine = True

    '转:GetDate()
    regEx.Pattern = "(?=[^']?)GETDATE\(\)(?=[^']?)"
    Sql = regEx.Replace(Sql,"NOW()")

    '转:UPPER()
    regEx.Pattern = "(?=[^']?)UPPER\([\s]?(.+?)[\s]?\)(?=[^']?)"
    Sql = regEx.Replace(Sql,"UCASE($1)")

    '转:日期表示方式
    '说明:时间格式必须是2004-23-23 11:11:10 标准格式
    regEx.Pattern = "'([\d]{4,4}\-[\d]{1,2}\-[\d]{1,2}(?:[\s][\d]{1,2}:[\d]{1,2}:[\d]{1,2})?)'"
    Sql = regEx.Replace(Sql,"#$1#")
    
    regEx.Pattern = "DATEDIFF\([\s]?(second|minute|hour|day|month|year)[\s]?\,[\s]?(.+?)[\s]?\,[\s]?(.+?)([\s]?\)[\s]?)"
    Set Matches = regEx.Execute(Sql)
    Dim temStr
    For Each Match In Matches
        temStr = "DATEDIFF("
        Select Case lcase(Match.SubMatches(0))
            Case "second" :
                temStr = temStr & "'s'"
            Case "minute" :
                temStr = temStr & "'n'"
            Case "hour" :
                temStr = temStr & "'h'"
            Case "day" :
                temStr = temStr & "'d'"
            Case "month" :
                temStr = temStr & "'m'"
            Case "year" :
                temStr = temStr & "'y'"
        End Select
        temStr = temStr & "," & Match.SubMatches(1) & "," &  Match.SubMatches(2) & Match.SubMatches(3)
        Sql = Replace(Sql,Match.Value,temStr,1,1)
    Next

    '转:Insert函数
    regEx.Pattern = "CHARINDEX\([\s]?'(.+?)'[\s]?,[\s]?'(.+?)'[\s]?\)[\s]?"
    Sql = regEx.Replace(Sql,"INSTR('$2','$1')")

    Set regEx = Nothing
    SqlToAccess = Sql
  End Function
End Class
  
</SCRIPT>
