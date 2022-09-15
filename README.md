# DrawWeChat
抓取weixin好友朋友圈内容


**环境准备：**  
**1.PC端安装[android studio](https://developer.android.com/studio)、[Appium Server](https://github.com/appium/appium-desktop/releases)、[Appium Inspector](https://github.com/appium/appium-inspector/releases)、[JDK 1.8](https://www.oracle.com/java/technologies/downloads/#jdk18-windows)、[PyCharm Community](https://www.jetbrains.com/zh-cn/pycharm/download/#section=windows)（Pycharm配置依赖库还是挺方便的）**  

**2.配置软件和环境变量**  
**配置android studio sdk：**  
在SDK Platforms中安装与手机或模拟器对应版本的Android API，在SDK Tools中安装Android Emulator和Android SDK Platforms Tools（默认在安装android studio时已经配置）  

**PC配置系统环境变量（仅供参考）：**
```
ANDROID_HOME  C:\Users\Administrator\AppData\Local\Android\Sdk  
JAVA_HOME C:\Program Files\Java\jdk1.8.0_333  
CLASSPATH .;%JAVA_HOME%\lib;%JAVA_HOME%\lib\tools.jar;  
Path  .;%JAVA_HOME%\bin;%JAVA_HOME%\jre\bin;%ANDROID_HOME%\platform-tools;%ANDROID_HOME%\tools;  
```

**3.手机打开USB调试、USB安装、USB调试（安全设置）（如有），并连接到PC。也可以使用安卓模拟器。**

**4.Pycharm创建新项目，安装依赖包，将main.py和config.py文件放入项目**  
```
Appium-Python-Client、pandas、openpyxl、numpy
```
**5.运行**  
- 手机连接PC，并登录weixin
- 打开Appium Server GUI，点击*startServer*运行
- 命令行运行*adb devices -l*命令获得手机设备名称
- 根据需要修改config.py文件
- 运行main.py

**说明**  
程序通过元素定位来模拟点击操作，如果发现程序不能正确找到对应元素，可能是weixin改变了元素名，可以通过Appium Inspector来检查元素。
```
Remote Path:    /wd/hub  
Desired Capability：
{
  "platformName": "Android",
  "appium:deviceName": "GM1900",
  "appium:appPackage": "com.tencent.mm",
  "appium:appActivity": ".ui.LauncherUI",
  "appium:noReset": "True"
}
```
