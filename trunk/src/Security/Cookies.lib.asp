<SCRIPT Runat="Server" Language="VBScript">

Lib.Require("Security.Hash")

Class TMeCookies
    'Public DefaultSecure ' 如果为真那么必须是https协议才会发送接受cookies
    ' the Cookies prefix name
    Public Name

    Private Sub Class_Initialize
      'DefaultSecure = True

        if gCache.CacheExists(SCRIPT_NAME & "CookieHash") then
          Name = gCache.Values(SCRIPT_NAME & "CookieHash")
        else
          Name = "H" & SimpleHash(gServerRoot & SCRIPT_NAME)
          gCache.Values(SCRIPT_NAME & "CookieHash") = Name
        end if
    End Sub

    Private Sub Class_Terminate
    End Sub

    ' 如果为真那么必须是https协议才会发送接受cookies
    Public Property Let Secure(ByRef pKey, ByRef pValue)
      Response.Cookies(Name + pKey).Secure = pValue
    End Property
    'Public Property Get Secure(ByRef pKey)
    '  Secure = Request.Cookies(Name + pKey).Secure
    'End Property

    ' The domain for which the cookie is valid.
    Public Property Let Domain(ByRef pKey, ByRef pValue)
      Response.Cookies(Name + pKey).Domain = pValue
    End Property
    'Public Property Get Domain(ByRef pKey)
    '  Domain = Request.Cookies(Name + pKey).Domain
    'End Property

    ' This property specified the path for which the HTTP cookie is valid. 
    ' If no path is specified when defining a cookie (see Response.cookies) then the cookie is only valid for the path of the current request.
    Public Property Let Path(ByRef pKey, ByRef pValue)
      Response.Cookies(Name + pKey).Path = pValue
    End Property
    'Public Property Get Path(ByRef pKey)
    '  Path = Request.Cookies(Name + pKey).Path
    'End Property


    Public Property Get Items(ByRef pKey)
      Dim Result
      Result = Request.Cookies(Name + pKey)

      Items = Result
    End Property

    ' set the common value
    Public Property Let Items(ByRef pKey, ByRef pValue)
      Response.Cookies(Name + pKey) = pValue
      'Secure(pKey) = DefaultSecure
    End Property

    Public Property Get SubItems(ByRef pKey, ByRef pSubKey)
      Dim Result
      Result = Request.Cookies(Name + pKey)(pSubKey)

      SubItems = Result
    End Property

    ' set the common value
    Public Property Let SubItems(ByRef pKey, ByRef pSubKey, ByRef pValue)
      Response.Cookies(Name + pKey)(pSubKey) = pValue
      writeln(pSubKey+":"+pValue)
      'Secure(pKey) = DefaultSecure
    End Property

    ' the cookie Expired date
    ' If the expiration date is prior to the current time the cookie is deleted from the client.
    ' If no expiration date is specified the client deletes the cookie at the end of the session (for example, when the browser is closed).
    Public Property Let Expires(ByRef pKey, ByRef pValue)
      Response.Cookies(Name + pKey).Expires = pValue
    End Property
    'Public Property Get Expires(ByRef pKey)
    '  Expires = Request.Cookies(Name + pKey).Expires
    'End Property


    Public Function HasKeys(ByRef pKey)
      HasKeys = Request.Cookies(Name + pKey).HasKeys
    End Function

    Public Sub Clear()
      Dim vCookie
      For Each vCookie In Request.Cookies
        if Left(vCookie, Len(Name)) = Name then Response.Cookies(vCookie) = ""
      Next
    End Sub

    Public Sub RemoveAll(ByRef pPrefixName)
      Dim vCookie
      For Each vCookie In Request.Cookies
        if pPrefixName <> "" then
          if Left(vCookie, Len(pPrefixName)) = pPrefixName then 
            Response.Cookies(vCookie) = ""
            Response.Cookies(vCookie).Expires = Date() - 1000
          end if
        else
          Response.Cookies(vCookie) = ""
          Response.Cookies(vCookie).Expires = Date() - 1000
        end if
      Next
    End Sub

    Public Sub Remove(ByRef pKey)
      Response.Cookies(Name + pKey) = ""
      Response.Cookies(Name + pKey).Expires = Date() - 1000
    End Sub

End Class

</SCRIPT>

