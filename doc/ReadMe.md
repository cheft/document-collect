# 归档服务程序部署文档

******

## ASP.NET Server端部署
> 要求服务器环境为windows Server 2003，装有安装版Office 2003而非绿色版，安装好IIS 6.0

### IIS 部署
> 打开IIS管理器
![Alt ](/images/01.jpg)

> 右键添加网站
![Alt ](/images/02.jpg)

> 为网站取名，例如：CollectService
![Alt ](/images/03.jpg)

> 绑定端口，例如：12333 
![Alt ](/images/04.jpg)

> 将网站主目录指向归档服务程序的物理目录(归档服务程序由开发人员提供)
![Alt ](/images/05.jpg)

> 设定网站访问权限
![Alt ](/images/06.jpg)

> 左键属性(Properties)查看此网站所用的应用程序池是什么，这里为 ___DefaultAppPool___
![Alt ](/images/07.jpg)
![Alt ](/images/08.jpg)

> 那么就去查看DefaultAppPool应用程序池的标识是否为 ___网络服务(NetworkService)___，若不是，请改成它
![Alt ](/images/09.jpg)
![Alt ](/images/10.jpg)


### 权限配置

* 因为IIS调用Office com组件需要配置权限，否则就会引发错误：检索 COM 类工厂中 CLSID 为 {000209FF-0000-0000-C000-000000000046} 的组件时失败，原因是出现以下错误: 80070005。 请按以下方法配置权限：

> win + R快捷键 打开运行命令框，敲入 mmc -32，进入控制台
![Alt ](/images/11.jpg)

> 点击 增加/删除管理单元
![Alt ](/images/12.jpg)

> 增加 组件服务
![Alt ](/images/13.jpg)

> 依次双击"组件服务"->"计算机"->"我的电脑"->"DCOM配置" 找到与word相关的组件，右键属性
![Alt ](/images/14.jpg)

> ___常规___ 中可以确认本地路径是以WINWORD.EXE 结尾的
![Alt ](/images/15.jpg)

> 点击 ___标识___ 标签,选择 ___下列用户___，填写本机用户及密码
![Alt ](/images/16.jpg)

> 点击 ___安全(Security)___ 标签,在 ___启动和激活权限(Launch and Activation Permissions)___ 上点击 ___自定义(Customerize)___ ,
![Alt ](/images/17.jpg)

> 点击对应的___编辑(Edit)___按钮,在弹出的 ___安全性___ 对话框中点击添加
![Alt ](/images/18.jpg)

> 注意 ___查找位置___ 要选择本计算机，填入 ___NETWORK SERVICE___ 点击确定增加此用户
 ![Alt ](/images/19.jpg)


> 赋予 NETWORK SERVICE ___本地启动（Local Launch)___ 等所有权限；
同时在 ___访问权限(Access Permissions)___ 、___配置权限(Configuration Permissions)___ 选项上也要做同样操作，点击 ___自定义___ ,然后点击"编辑",在弹出的"安全性"对话框中也填加一个"NETWORK SERVICE"用户,也赋予所有权限.
![Alt ](/images/20.jpg)

> 刚刚仅把word相关组件配置了权限，另外还有Excel、PPT相关的组件需要设置，请依次找到进行相同的操作

> 下图标记出Excel和Word组件，在所有的组件中没有找到PPT的，于是在花时间在后面的UUID的组件中，才找到关于PPT的组件，这里需要根据本地路径来寻找， 以POWERPNT.EXE结尾。(根据安装的office和windows不同会有不同情况，需要特别注意一下，万一找不到就需要去UUID的组件中找路径相关的，Word组件是以 WINWORD.EXE结尾，Excel组件是以 EXCEL.EXE结尾，PPT组件是以POWERPNT结尾 )
![Alt ](/images/21.jpg)
![Alt ](/images/22.jpg)


### 归档服务程序配置

> 创建归档服务需要的临时目录，可以为任意目录，比如选择 ___D:\\topway\\collect___ 下创建这四个文件夹 ___document, request, log, temp___ ，右键 ___D:\\topway\\collect___ 目录 ___属性___ ，选择 ___安全___ 标签，点击 ___编辑___ ,在弹出的对话框中添加 ___NETWORK SERVICE___ 用户 然后赋予 ___完全控制___ 的权限.
![Alt ](/images/23.jpg)
![Alt ](/images/24.jpg)
![Alt ](/images/25.jpg)

> 修改归档服务程序中的web.config 配置文件： 找到 _QueueDir_ 的value，及 _LogFileAppender file_ 的value，修改如下图：

```xml
<appSettings>
    <add key="QueueDir" value="D:\topway\collect"/> <!-- 队列目录 -->
    <add key="QueueInterval" value="3000"/>         <!-- 队列循环间隔时间 -->
    <add key="IsSheet" value="1"/>                  <!-- Excel是否切割Sheet页 -->
    <add key="IsImage" value="0"/>            <!-- Word, Excel是否图片模式，需要office2007及以上 -->
	<add key="JavaServerOn" value="0"/>     <!--是否向客户端传回转换日志-->
    <add key="JavaServerUrl" value="http://172.18.97.60:8080/Project/ServletUrl"/>  <!--客户端传回转换日志请求地址-->
</appSettings>
<log4net>
	<!--定义日志输出到文件中-->
	<appender name="LogFileAppender" type="log4net.Appender.FileAppender">
	  <!--定义日志文件存放位置-->
	  <file value="D:\topway\collect\log\DocumentService.log"/>
	  <appendToFile value="true"/>
	  <rollingStyle value="Date"/>

<!-- 省略 -->
</log4net>
```

##IIS配置注意

	部署好CollectService服务，右键 DocumentWebServcie.asmx 页面浏览，若浏览器没有显示或显示出错，请按下图配置
![Alt ](/images/31.png)

	若操作系统为 server 2008 64位，请按下图配置 启动32位应用程序
![Alt ](/images/32.jpg)
	
	若操作系统为 server 2003 64位，请安以下步骤配置
	
	*单击“开始”，单击“运行”，键入 cmd，然后单击“确定”。
	
	*键入以下命令启用 32 位模式：
	cscript %SYSTEMDRIVE%\inetpub\adminscripts\adsutil.vbs SET W3SVC/AppPools/Enable32bitAppOnWin64 1
	
	*键入以下命令，安装 ASP.NET 2.0（32 位）版本并在 IIS 根目录下安装脚本映射：
	%SYSTEMROOT%\Microsoft.NET\Framework\v2.0.40607\aspnet_regiis.exe -i

	*确保在 Internet 信息服务管理器的 Web 服务扩展列表中，将 ASP.NET 版本 2.0.40607（32 位）的状态设置为允许。
![Alt ](/images/31.png)


