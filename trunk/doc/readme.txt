MeASP Development Class SDK
  * dynamic-lazy-loading units.
  * the modules codes in the database.

提供最基本核心 ASP 功能类:

MeDatabase.asp
  数据库功能类

Lib.RequireFile('ADOConsts')

TMeDatabase = Class
    Property DBType
    Readonly Property Conn
    ' Close the database 
    Public Sub Close()
    ' Open the database, note: you must close it before open. 
    Public Function Open(ByRef aConnStr)
    Public Sub BeginTrans()
    Public Sub CommitTrans()
    Public Sub RollbackTrans()
    Public Function Execute(ByRef pSql)
    Public Function CreateRecordSet()
    Public Function OpenRecordSet(ByRef aSource, ByRef aCursorType, ByRef aLockType)
    Public Function OpenTable(ByRef aTableName, ByRef aReadOnly)
End
  
MeLib.asp
 动态按需加载的功能函数库类

TMeLib = Class
  ' require the library: search in db first then search in the folders
  ' 同一名称函数库只会被加载一次
  ' 采用类似Java的Package的方式装载
  ' 如： "Security.Hash.*" 将装载所有的散列类函数库。
  ' 函数库的使用： :Lib.Require("Security.Hash.MD5")
  ' Lib.Require("Security.Hash")
  '   在文件系统下意味着装入 /Security/Hash.lib.asp
  Public Function Require(aLibName)
End

函数库的内容必须包裹在 '<SCRIPT Runat="Server" Language="VBScript">' 和 </SCRIPT> 之间或"<%" 和"%>" 之间。
注意如果是"<%" 和"%>" 之间那么这里面的默认语言是VB。

ASP函数库文件命名： 文件名必须以 ".lib.asp" 结尾！
  Security.Hash.MD5 : 表示该函数库位于Security/Hash/MD5.lib.asp 的文件中
  注意函数库文件文件必须保存为 Ansi格式，不要使用 unicode 格式保存！
  函数库的加载顺序：以 Security.* 为例，首先装载 Security 目录下的函数库文件，然后进入子目录并装载子目录下的函数库文件


对于要想使功能库放到数据库中，数据表必须符合如下要求：

目录命名：
/SYS/CODE/LIB/    /系统/代码/库 存放功能库的根目录
/SYS/CODE/LIB/MeCMS/MACRO/   /系统/代码/宏

目录类别： SYS
 目录子类别, CODE, TMLT(版式模板) 

MeCMS Engine Development Platform:
  the MeCMS Core Development Platform, It's OpenSource and total free, even commercial!!
  CMS核心开发平台。

Features:
  * 表现和逻辑完全分开
  * Multi-languages Supports. 多语支持
  * 修改数据结构，进一步增加可读性
  * 重新修改和扩充作为CMS核心开发平台。命名规范化，插件开发规范化。
    如： Resource(资源) = Content(内容); Slice（碎片） = Macro（宏）; Special(特性) = Special(专题); 频道(Channel) = 栏目、类别(Category)
    + 栏目类型 ChannelType
      + 如何自定义栏目类型，供开发者扩充？
      + 注册新的栏目类型向导编写规范

Action 规范
Action 可以是类中的方法，也可以是函数，注意Action必须是返回真假值的函数,返回为真表示执行成功。
  ActLib 该动作所在函数库名称
  ActClass 如果是类方法，那么则是该动作所在类名，否则为空
  ActName  该动作的函数名称
注意：
  如果 Action 是类中的方法，那么在该类的函数库中必须存在一个全局变量 gClassName(去掉类名的首字母"T"的)。
  如果 Action 有界面那么必定是类！而且必须遵照一定的规范，说明该界面可能的参数！
  界面在 Template/Controls 目录下面。

 <!--Action:[CatId.]CatName[(无参数可以不要括号)]-->
 将显示动作的界面。

宏(Macro) 规范：
  用在模板中，替换建立新的文字，本质上是一个替换函数。标签也是一种宏。
  Marco 的实质就是ASP 自定义函数，以函数名称Macro打头。存放在目录：
  /SYS/CODE/LIB/MeCMS/MACRO/ 下。

  标准宏：系统内定的宏。

 <!--Macro:[LibName.]MacroName[(无参数可以不要括号)]-->
 约定：如果希望将某段文字用指定的宏或模板替换，那么可以使用：
 <!--Macro:MacroName[()][:Begin]--> 如果有:Begin那么必然有：
 <!--Macro:MacroName[(参数可以省略)]:End-->


版式模板(Template)规范：
  版式模板是用来显示前台时所看到的网页的界面布局形式，如分栏、表格布局、图片和文字要显示的位置等样式，
  有时也称为版面设计、版面划分或版面布局。版式模板包括网站通用模板和各频道的首页、栏目页、内容页等页面的模板。

  版式模板支持简单循环语句(宏)，模板嵌套，模板参数（这样可以改变一些栏目名称或自由切换数据表，只要字段名相同）。
  版式模板的分类： 
    普通版式模板：普通皮肤文件
    控件版式模板 存放在Template/Controls子目录下，控件的版式模板类似于Delphi DFM 或 ASP.NET 的 ascx。控件版式模版不能通过“<!--Template:XXX-->”的方式调用，而是<!--Action:XXX-->的形式。
      与普通的版式模板不同，它的前头有一个多的参数，用于指名和某个Action配合（这个Action 类似于 ASP.NET 的 ascx.cs 源文件。）
      比如登录（Login） Action. 
      界面关键部分：
        <!--Login:Begin--> //这里会被替换成<Form>
          <!--Edit(&Name)-->    //&Name表明为参数，Edit(XXX)则是会被替换成相应的<input>。
          <!--Edit(&Passwd)-->
        <!--Login:End-->  //这里会被替换成</Form>
    容器模板：容器皮肤。

  <!--Template:TemplateName[(无参数可以不要括号)]-->


样式模板 规范
  样式模板可控制整个网站在前台显示时看到的的字体、风格、图片等样式，通常是用CSS 网页样式语句来进行设计和控制整个网站的风格样式。
  待定。可以参考DNN的方案，CSS和版式皮肤放在一起。


目录：
src\system\ 系统目录核心文件
src\database\ Access database
src\manager\: 管理后台
 使用绝对目录避免漏洞，如果修改了系统目录名称或位置，那么必须修改管理后台的这两个文件：
   src\manager\Default_inc.asp
   src\manager\Macro_inc.asp
 
MeCMS Platform
  MeCMS平台，在MeCMS SDK平台上开发（用户只需要会使用后台就可以，无需要开发编程）,
Features:

      + 默认自定义栏目类型： 
        单页面：该栏目为一个页面，单击该栏目即可编辑该页面，可以插入各种标签和宏。
        通用内容（文章、新闻、论坛）：内容的区分在模板中用不同的宏来区分？还是每一个作为一种栏目类型，这样的好处就是模板可以通用。最后我决定两种都用！！
          你可以用宏，也可以采用通用标记（根据栏目类型自动选择内容）。
        ：

  通用模板的构想：对于栏目下的模板总是有：
    索引页面：栏目标题，列表（是否含子目录列表）
    内容细节：标题，作者，时间，内容

目录：
src\program\ 存放扩充的程序

功能：
  * 表现和逻辑完全分开(完全类化)
  * ASP 单页面入口机制：更加规范，并且降低了程序暴露的风险，配置也更加容易
  * 动态按需加载函数类库机制
     * 解决单页面入口机制中把所有的功能都需要全部包括进来的问题: 现在动态按需要加载了
     * 形成函数类库，容易形成规模化，工程化管理。
     * 高度模块化，可以将函数类库放入数据库中，动态编写代码，联机添加模块和封杀模块
       一句话，你可以直接在后台写模块代码，直接挂接模块或删除模块。本系统本身就是模块化的产物。
  * 支持高速查找的无级分类（不过将受到Access数据库的字段长度限制）
  * 后台管理系统
     * Ajax 支持（我选择Dojo 作为俺的首选 Ajax）
     * 版本控制系统模块：类似于 CVS，保存每一次变动的内容或类别
       （只保存不同处，这样可以轻易取回以前的版本）
     * 系统类别管理模块（函数类库，宏库）
     * 用户管理模块
     * 完善的用户权限管理系统模块，同样，该模块存放于数据库
     * 多语管理模块
     * 模板管理模块
     * 文件系统管理模块
     * 内容管理模块
     * 类别（栏目）管理模块
     * 专题管理模块
     * 其它类别类型模块：单页类别模块，BBS类别模块，通用内容类别模块（新闻，文章....），产品类别模块.....
       并可任意扩充和挂接。


开发计划：
   核心开发：完成动态按需加载函数类库机制(已完成)
     已经开发完成的核心类或函数库：
      不能放入数据库中的:
       * TMeLib
       * TMeList
       * MeSysUtils.asp
       * TMeDebugger(可以不要，调试用)
      能放入数据库中的:
       * TMeDatabase
   核心模块开发：完成一个单用户单语言(English)用来应付我自己的需求先(开发中)
     * 数据库UML分析设计（已完成）
     * 用户管理模块（登陆验证，无权限控制）
     * 系统类别管理模块（函数类库，宏库）
     * 模板管理模块
     * 单页类别模块（用于产生类似于首页的单页面）,通用内容类别模块

完成后单页面入口 asp 就象这样：
[Code]
  see MeCMS.asp
[/Code]

VBScript 命名规范：
  类名以大写字母T打头，后面为概括描述该类功能的英语名词，如：
    TMeDatabase: 我的数据库类
  类的私有字段名称以大写的F打头: Private FUserName；
  属性名称必须是能概括描述该属性作用的英语名词: ConnectionString；
  方法或过程名称则必须是能概括描述该方法作用的英语动词或动名词结构；
  过程的局部变量以小写字母v打头，循环变量i,j例外；
  全局变量以字母g打头。

注：单词之间用首字母大写的形式进行区分：ConnectionString

序列化脚本对象类

用户自定义的VB脚本对象类还称不上是真正的对象(没有继承)，而且它的生存期只能是该ASP页面的周期， Script 对象是无法直接缓存的！为了能够缓存对象的属性以及在Ajax中使用的需要，我设计了如下的机制来序列化Script对象。

1、能序列化的脚本类的规范
凡是要想序列化的类，在该类中必须存在如下的属性和方法：

    Public Property Get ObjectId() ' 只读属性, 该对象的唯一ObjectId，调用 MakeGlobalObjectId(aObj) 产生唯一的全球ID(MakeGlobalObjectId 实际上简单返回 ClassName+":"+ObjectId )
    Public Property Get ClassName() ' 只读属性, 该对象的类名
    Public Function GetMetaObject()  '现在返回 TMetaObject 对象了！

    Public Function Fetch(ByRef pId) '根据传入的Id, 从数据库或somewhere取回该对象的所有信息, 返回真则表示成功

在创建对象前,你需要引入支持新的语言特性的"Lang.Object"库: Lib.Require('Lang.Object').

2. TMeMetaObject
TMeMetaObject 用来保存该对象类需要序列化的属性的MetaInfo：属性名和类型。
Class TMeMetaObject
    ' 对象的类名，指示是哪一个对象类的 MetaInfo.
    Public ClassName
    '属性列表对象
    Public Property Get Fields()   
    '返回指定属性名的属性对象
    Public Default Property Get FieldByName(ByRef aFieldName) 
    '新增属性，返回新增属性对象
    Public Function AddField()     
     '从字符串建立对象的所有属性。字符串格式： "属性名:属性类型[:Size:Required:Confirmed:Validator:Constraints:Filters],..." 
     '属性之间用逗号“,”分隔，属性名和它的相关类型之间用冒号“:”分隔，数据之间不能有空格,如果是字符串必须用引号。
     '示例： "Id:ftString:32:true:false:""the Validator"":""the Constraints"":""the filters"",Password:ftPassword:32:true:true"
    Public Sub AssignFieldsFromString(ByRef aText) 
End Class

'MetaInfo 的属性字段列表类
Class TMeMetaFields
    '字段列表项，索引从0到(Count-1)
    Public Default Property Get Items(ByRef Index)
    '根据字段名返回字段属性对象
    Public Property Get FieldByName(ByRef aFieldName)
    Public Function IndexOf(ByRef aFieldName)
    '从字符串建立一个属性字段，字段名和它的相关类型之间用冒号“:”分隔，第一个是字段名，接着的则是字段的属性。
     '示例： "Id:ftString:32:true:false:""the Validator"":""the Constraints"":""the filters"""
    Public Function Append(ByRef aText)
    '属性个数
    Public Function Count()
    '清除所有的属性
    Public Sub Clear()
End Class

'Meta 字段属性对象
Class TMeMetaField
    '字段名
    Public Name
    '字段类型
    Public FieldType
    '字段大小
    Public Size
    '是否必填项
    Public Required  'Boolean
    '是否需要确认，真则需要输入两遍
    Public Confirmed 'Boolean 
    '有效性
    Public Validator ' the js condition script.
    ' 约束
    Public Constraints ' the js condition script.
    ' 过滤结果
    Public Filters ' the js filter value script.
End Class

符合规范的类事例：
[code]
Class TMeUserInfo
    Public Id, Name, Password, '.....

    Private Sub Class_Initialize
    End Sub

    Private Sub Class_Terminate
    End Sub

    Public Property Get ClassName()
      ClassName = "TMeUserInfo"
    End Property

    Public Property Get ObjectId()
      ObjectId = Id
    End Property

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

    Public Function Fetch(ByRef pUserId)
        Dim Result, vRS, vSQL
        Result = False
        With gApplication.Database
          vSQL = Replace(sqlGetUserInfo, "%Id%", .QuotedStr(pUserId))
          Set vRS = .OpenTable(vSQL,  ForReading)
        End With
        if not (vRS is Nothing) then
          Result = not vRS.BoF
          if Result then
            Id = vRS("usr_name").Value
            '.....
          end if
          vRS.Close
          Set vRS = Nothing
        end if
        Fetch = Result
    End Function
End Class

Set CurrentUser = New TMeUserInfo

CurrentUser.Fetch("aUser")

[/code]


使用方法：

将对象存放到缓存中：
[code]
  gCache.Objects(MakeGlobalObjectId(CurrentUser)) = CurrentUser
[/code]
检查对象是否存在以及从缓存中取出对象：
[code]
  if gCache.ObjectExists("TMeUserInfo:aUser") then Set CurrentUser = gCache.Objects("TMeUserInfo:aUser") 
[/code]

从缓存中删除对象:
[code]
  gCache.RemoveObject("TMeUserInfo:aUser")
[/code]

实现原理：
  很简单通过将对象属性转换成数组就能够存放到TApplicationCache（俺写的类包裹了系统的Application）上了。
  
  Id 存放在Cookies中，其它数据则存放在ApplicationCache上。
  Set myObj = TApplicationCaches.Obejcts("ObjId")
  TApplicationCaches.Obejcts("ObjId") = aObj
  TApplicationCaches.ObejctExists(pObjId)
  
  存放在AppCache 上的对象格式：
  Redim vObjData(PropertyCount)
   vObjData(0) = ClassName
   for i = 1 to PropertyCount
     vObjData(i) =PropertyValue

MeSession 机制
编写MeSession 的初衷：ASP Session 一点都不好用，消耗资源大不说，而且不能保存内建对象，并且只使用 Cookies 进行存放。其实如果它能保存内建对象那么我也能忍了！但是！！所以创建 TMeSession 类来替代 ASP Session 类.

编写自己的方式 TMeSession,来代替 ASP Session, 我设想的 TMeSession 可以保存内建对象，并且可以根据客户端是否支持Cookies 自动切换，不过现目前只支持 Cookies.
  为了能够支持保存内建对象我不得不对想要保存到 Session 中的对象作一些规范：
  该对象必须有 
    Public ClassName: string;
    Public Function GetPropertyNames: string; //需要保存的属性列表，属性之间用逗号分隔
  这样实际上我是将它转换成属性存放在Cookies中:
实现原理：
  规定每一个对象实例有且只有一个唯一的对象ID(ObjectId).
  首先将对象需要保存的属性存放的到Application Cache 中，通过ObjectId联系；然后将ObjectId放入Cookies中。
  这样从客户端只需要取回ObjectId即可，极大减小数据的传送，同时提高了安全性能，减小敏感数据被窃听到的机会（在客户端上只有Id,不会保存用户名和密码等敏感数据）。

那么如何实现安全的objectId 呢？
为了防止cookies被伪造，一旦我在客户端存放Id后，就会认为用户是登录用户。因此如果该cookie 的id值被截获那么危险！！！
必须构造安全Id,使得该id的值在客户端的表现每次都应该不一样，并且设置期限限制（该id只有在一定的时间内才有效）。
解决方案：
  将Id加密, 将到期时间插入到值，
      vExpired = DateAdd("n", FTimeOut, Now())
      aKey     = EnDeCryptXOR(aKey, FVisitKey)
      aValue   = EnDeCryptXOR(aValue, FVisitKey + vExpired)
        FCookies.Items(aKey)("V") = aValue
        FCookies.Items(aKey)("D") = EnDeCryptXOR(vExpired, FVisitKey)
        FCookies.Expires = vExpired

用户权限控制系统模块
权限控制：
权限列表： 列举本系统的权限 cms_permissions
权限(prm_id)的命名： 
  SYS.ModuleGroup.Module.List: 

默认的权限（动作）： List(View), Edit, Del, Add, Audit, Clear, Exec(执行标准宏函数),SExec(SuperExec执行一切函数), Gen(Generate the files), Hide(能够看到隐藏的), Login, Logout
角色列表： cms_roles
通过角色你可以实现：普通会员，VIP会员，金牌代理 等不同权限级别

想法将专题纳入通用的权限管理中，专题列表是放在类别中统一管理，但是由于专题没有专门的权限id，那么应该怎么办？
虽然专题id没有放入 cms_permissions ，但是我可以在程序中增加一段代码，查找完  cms_permissions 数据表后，接着
查找专题id不就行了！！ 手工构造权限 “SPEC.专题id”！

继承权力： 首先是找自己，如果没有，并且该项（类别或内容）是允许继承父类，那么则找自己是否有父类的权限。

prm_id: 权限ID: 如果权限类型是内容：SYS.ModuleGroup.Module.List；如果是专题则是专题ID;
  如果是函数则是Lib.[FunctioName|*]，不过函数有必要分这么细致么？

prm_type: 权限类型: 内容，专题，函数 
prm_visible: 权限是否可见


模版调用(宏)
CMS:LIST 主要用于制作首页(索引页)；
CMS:CONTENT 主要用于对特定文章的包含性调用，比如在A文章的内容页调用B文章的内容，其主要用途并不是制作内容页；
  CONTENT.Record(aIndexId [, aTemplateName] [, aCached] [aFieldList] [, aTableNameList]): 调用指定的单个记录
  使用举例：
    <!--Macro:MeContent.Record("68"):Begin-->
      <!--调用文章IndexID为68的记录-->
      标题： [$Rec.Title]
      简介： [$Rec.Intro] 
    <!--Macro:MeContent.Record("68"):End-->
  CONTENT.LIST(aIndexId [, aObjectName] [, aTemplateName] [, aCached] [aFieldList] [, aTableNameList]): 调用多个记录集合
    aIndexId: 为要返回的字段Id列表，id之间用逗号分隔。也可以为字段或变量。变量名：VarName, 字段名[$FieldName].
    aObjectName: 当前记录对象的名称，供模板内使用，默认为 "var"
    aTemplateName: 版式模板名，如果为空，表示使用Begin...End之间的作为模板
    aCached: 是否缓存该列表，默认为False
    aFieldList: 返回的字段名列表，默认为全部；限制返回的字段可以提高性能。
    aTableName: 参与的数据表，如果没有则和调用者是同一数据表。
  使用举例：例如：比如一篇文章里，有个推荐阅读的书籍的列表，同一个表中的其他
   记录，可将RecomandedBook字段设为“其它结点内容”，然后模版里这样调用：

<!--Macro:MeContent.List("[$RecomandedBook]"):Begin-->
  <ul>
   <LOOP> <!--如果没有loop那么所有的都参与循环--> 
    <li><a href="[$var.URL]">[$var.Title]</a></li>
   </LOOP>
  </ul>
<!--Macro:MeContent.List("[$var.RecomandedBook]"):End-->

CMS:NODELIST和CMS:NODE 主要用于制作各种导航条；
CMS:SEARCH 主要用于相关文章、相关软件等一篇内容的相关内容的调用；
CMS:COMMENT 主要用于调用评论列表和内容；
CMS:COUNT 主要用于各种计算统计，如统计软件下载站的软件总数、今日更新、下载总次数、新闻文章总数等；
CMS:SQL 主要是对数据库的直接查询调用，比如您可以用它来调用您的论坛的公告、各个版块的内容让他们显示在您网站的首页等任意位置；
Cache参数 和 returnKey参数是调用辅助标签，可以提升系统性能。 
内容页变量标签 主要用于制作内容页。 

内容页变量标签调用
说明 
内容页变量包括所有公共变量和内容模型变量(比如新闻系统模型、下载系统模型)。

比如你做新闻站点，你的新闻内容页模版就可以调用公共标签＋新闻系统标签 。
同理你做下载站点，你的下载内容页模版就可以调用公共标签＋下载系统标签 。

内容模型中定义的字段名，都可以在内容页模版中使用[$字段名]来调用显示


内容模型标签表(公共变量)
说明 
以下的内容模型标签表是MeCMS官方默认定义的标签表，MeCMS内容管理系统手册中的所有例子均基于官方定义的内容模型，
最终用户对内容模型的二次修改(增加、修改字段)不包括在内。


公共标签表 
标签名 描述 
IndexID 索引ID 
ContentID 内容ID 
NodeID 归属结点ID 
URL 内容URL地址 
PublishDate 内容发布时间戳 
top 置顶权重 
Sort 排序权重 
pink 精华权重 
Hits_Total 总点击数 
Hits_Today 本日点击数 
Hits_Week 本周点击数 
Hits_Month 本月点击数 
Hits_Date 最新访问时间戳 
CommentNum 评论数 




新闻系统模型标签表 
标签名 标识 
Title 标题 
TitleColor 标题颜色 
SubTitle 副标题 
Author 作者 
Photo 新闻图片 
Editor 责任编辑 
FromSite 来源网站 
Keywords 关键字 
Content 新闻内容 
CustomLinks 自定义相关文章 
Intro 简介 

 


下载系统模型标签表 
标签名 标识 
SoftName 软件名称 
SoftType 软件类别 
Softsize 软件大小 
Language 软件语言 
Environment 运行环境 
Star 软件评级 
SoftKeywords 软件关键字 
Developer 开 发 商 
Intro 软件介绍 
Download 下载地址 
Localload 本地上传下载地址 
Photo 界面预览 

