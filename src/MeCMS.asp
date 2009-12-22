<%@Language="VBScript"%>
<%Option Explicit%>
<!--#Include file="MeCMSAll.asp" -->
<!--.Include file="ApplicationCaches.lib.asp" -->

<%
Lib.Require("ApplicationCaches")
Lib.Require("MeCMS.App")

'Set gCache = New TApplicationCaches

With Lib
  'config the lib parameters:
  '.Sql =""  ' u can change the sql to suite you database. 如有必要修改 SQL 语句以适合你自己的数据库
  '.Options = optLoadFromDB  'you can forbbide the load lib from file system only load from db to speedup. 为了加快速度你可以禁止从文件系统中加载函数库,只从数据库中加载
End With




Set gApplication = New TMeCMSApp

With gApplication
  .Database.ConnectionString = "Driver={Microsoft Access Driver (*.mdb)};DBQ="&Server.Mappath("database/MeCMS.mdb")
  .Database.DBType = DB_ACCESS
  .Run
End With

Set gApplication = Nothing
%>

