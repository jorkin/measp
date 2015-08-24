MeASP Web Framework -- Development Class SDK
  * dynamic-lazy-loading modules(libraries).
  * the libraries(modules) codes can in the database.
  * the libraries can be encrypted via password.
  * the libraries can be cached to speed up.
  * can support the VB and JScript Libraries together.

MeDatabase.asp
  * TMeDatabase: the database class for ASP

MeLib.asp
  * TMeLib: implements the dynamic-lazy-loading modules(libraries)

The following libs are the core libs! The core libs are always in the memory!
```
' NO NESSARY to load these libraries:'
Lib.Require("MeConsts")
Lib.Require("MeSysUtils")
Lib.Require("MeList")
Lib.RequireFile('ADOConsts')
```

If you wanna use the Database lib feature:
```
Lib.Require("MeDatabase")
```

If you wanna encrypt lib, the "Security.Cipher" library should be included
```
Lib.Require("Security.Cipher")
```

The Moudle(Library) Organized and Named in the folder

  * All module file name should be ended by ".lib.asp".
  * Lib.AddIncludeDir("Some/Module/Path") to add a libraries path.
  * the "Security/Cipher/RC4.lib.asp" file: Lib.Require("Security.Cipher.RC4")
  * Lib.Require("Security.Cipher") will load all libraries in the "Security/Cipher/" folder.

MeList.asp
  * this is a core lib. NO NESSARY to load.
  * Implements a resizable List class.


ApplicationCaches.lib.asp
  * Lib.Require("ApplicationCaches")
  * Manage the Application Caches


lang\object.lib.asp
  * Lib.Require("lang.object")
  * the very simple Object Entity(ORM) supports.

Usages:
```

<%@Language="VBScript"%>
<%Option Explicit%>
<!--#Include file="MeAll.asp" -->
<!--.Include file="ApplicationCaches.lib.asp" -->

<%
Lib.Require("ApplicationCaches")
Lib.Require("MeCMS.App")

With Lib
  ' config the lib parameters: '
  '.Sql =""  ' u can change the sql to suite you code library database.
  '.Options = optLoadFromDB  'you can forbid the load lib from file system only load from db to speedup. 
End With

Set gApplication = New TMeCMSApp

With gApplication
  .Database.ConnectionString = "Driver={Microsoft Access Driver (*.mdb)};DBQ="&Server.Mappath("database/MeCMS.mdb")
  .Database.DBType = DB_ACCESS
  .Run
End With

Set gApplication = Nothing
%>

```


MeCMS Engine Development Platform:
> the MeCMS Core Development Platform, It's OpenSource and total free, even commercial!!
