<!-- Include file="Cipher.lib.asp" -->
<!-- Include virtual="/util/string.lib.asp" -->
<!-- Include file="Hash/Base64.lib.asp" -->
<!-- Include virtual="/lang/object.lib.asp" -->

<SCRIPT Runat="Server" Language="VBScript">

Const cMeObjectCookiePrefix = "@"

Class TMeSession
    Private FVisitPassword
    Private FCookieSupported
    Private FCookies
    Private FTimeout 'minutes
    Private FEnhancedSecure

    Private Sub Class_Initialize
      Lib.Require("lang.object")
      Lib.Require("util.string")
      Lib.Require("Security.Cipher")
      Lib.Require("Security.Hash.Base64")
      Randomize
      FTimeout = 20
      FEnhancedSecure = False
      FVisitPassword = "To22vB4dV39513Cu3nh2O$)"
      FCookieSupported = IsClientCookieSupported()
      if FCookieSupported then
        Lib.Require("Security.Cookies")
        Set FCookies = New TMeCookies
      end if
    End Sub

    Private Sub Class_Terminate
      Set FCookies = Nothing
    End Sub

    Public Property Get Timeout()
      Timeout = FTimeout
    End Property

    Public Property Let Timeout(ByRef pValue)
      FTimeout = CLng(pValue)
    End Property

    Public Property Get EnhancedSecure()
      EnhancedSecure = FEnhancedSecure
    End Property

    Public Property Let EnhancedSecure(ByRef pValue)
      FEnhancedSecure = CBool(pValue)
    End Property

    Public Property Get VisitPassword()
      VisitPassword = FVisitPassword
    End Property

    Public Property Let VisitPassword(ByRef pValue)
      FVisitPassword = CStr(pValue)
    End Property

    'get the non-object items
    Public Property Get Items(ByVal aKey)
      Dim Result, vExpired
      'writeln "safdsfsdf=" + aKey
      aKey = StringToAnsi(aKey)
      aKey = AnsiToBase64(EnDeCryptXOR(aKey, FVisitPassword))
      'writeln "safdsfsdf=" + aKey
      if not IsEmpty(FCookies) then
        vExpired = FCookies.SubItems(aKey, "D")
        if vExpired <> "" then
          'vExpired = EnDeCryptXOR(vExpired, FVisitPassword)
          if IsNumeric(vExpired) then
            'vExpired = CDate(vExpired)
            Result = DecryptValue(FCookies.SubItems(aKey, "V"), FVisitPassword + vExpired)
          end if
        end if
        'if Result 
      else ' TODO: no cookie support: use url parameters
      end if

      Items = Result
    End Property

    ' set the common value(not object)
    Public Property Let Items(ByVal aKey, ByVal aValue)
      Dim vExpired, vCookieExpired
      vExpired = DateAdd("n", FTimeOut, Now())
      vCookieExpired = vExpired
      vExpired = DateAdd("s", Int(59*Rnd()), vExpired)
      vExpired = CSng( vExpired - CDateBase)
      aKey   = EncryptValue(aKey, FVisitPassword)
      aValue = EncryptValue(aValue, FVisitPassword + CStr(vExpired))
      if not IsEmpty(FCookies) then
        FCookies.SubItems(aKey, "V") = aValue
        FCookies.SubItems(aKey, "D") = CStr(vExpired)
        if not FEnhancedSecure then FCookies.Expires(aKey) = vCookieExpired
      else ' TODO: no cookie support: use url parameters
      end if
    End Property

    'the Key should be the classname,and the value is ObjectId.
    Public Property Get Objects(ByRef pKey)
      Dim Result, vObjectId, vId
      vObjectId = Items(cMeObjectCookiePrefix+pKey)

      if vObjectId <> "" then
        if not IsEmpty(gCache) then
          vId = MakeGlobalObjectIdBy(pkey, vObjectId)
          if gCache.ObjectExists(vId) then Set Result = gCache.Objects(vId)
        end if
        if not IsObject(Result) then
          Set Result = Eval("New "+pKey)
          if IsObject(Result) then 
            if Result.Fetch(vObjectId) and not IsEmpty(gCache) then
               vId = MakeGlobalObjectId(Result)
               if not gCache.ObjectExists(vId) then gCache.Objects(vId) = Result
            end if
          end if
        end if
      end if

      if IsObject(Result) then Set Objects = Result else Set Objects = Nothing
    End Property

    ' set the common value(not object)
    Public Property Let Objects(ByRef pKey, ByRef pObject)
      Dim vObjectId, vId

      if not IsEmpty(gCache) then
        vId = MakeGlobalObjectId(pObject)
        if not gCache.ObjectExists(vId) then gCache.Objects(vId) = pObject
      end if

      Items(cMeObjectCookiePrefix+pKey) = pObject.ObjectId

    End Property

    Public Function ItemExists(ByVal aKey)
      ItemExists = False
      aKey   = EncryptValue(aKey, FVisitPassword)
      if not IsEmpty(FCookies) then
        ItemExists = FCookies.Items(aKey) <> ""
      else ' TODO: no cookie support: use url parameters
      end if
    End Function

    Public Function ObjectExists(ByRef aKey)
      ObjectExists = ItemExists(cMeObjectCookiePrefix+aKey)
    End Function

    Public Sub Clear()
      if not IsEmpty(FCookies) then
        FCookies.Clear()
      else ' TODO: no cookie support: use url parameters
      end if
    End Sub

    Public Sub RemoveAll(ByRef pPrefixName)
      if not IsEmpty(FCookies) then
        FCookies.RemoveAll(pPrefixName)
      else ' TODO: no cookie support: use url parameters
      end if
    End Sub

    Public Sub Remove(ByVal aKey)
      aKey   = StringToAnsi(aKey)
      aKey   = AnsiToBase64(EnDeCryptXOR(aKey, FVisitPassword))
      if not IsEmpty(FCookies) then
        FCookies.Remove(aKey)
      else ' TODO: no cookie support: use url parameters
      end if
    End Sub

    Public Sub RemoveObject(ByRef pKey)
      Remove(cMeObjectCookiePrefix + pKey)
    End Sub

    Public Function EncryptValue(ByRef aValue, ByRef aPassword)
      Dim Result
      Result   = StringToAnsi(aValue)
      Result   = AnsiToBase64(EnDeCryptXOR(Result, aPassword))
      EncryptValue = Result
    End Function

    Public Function DecryptValue(ByRef aValue, ByRef aPassword)
      Dim Result
      Result = Base64ToAnsi(aValue)
      Result = EnDeCryptXOR(Result, aPassword)
      Result = AnsiToString(Result)
      DecryptValue = Result
    End Function

End Class

</SCRIPT>

