<div align="center">
<img alt="LOGO" src="https://fastly.jsdelivr.net/gh/pk5ls20/txlbot_server@master/txmbotlogo.png" width="256" height="256" />

# txlbot_server
<hr>
<div>
    <img alt="Python" src="https://img.shields.io/badge/python-3.9.11 Pass-%2300599C?logo=python">
</div>
<div>
    <img alt="platform" src="https://img.shields.io/badge/platform-Windows-blueviolet">
</div>
<div>
    <img alt="license" src="https://img.shields.io/github/license/pk5ls20/txlbot_server">
    <img alt="commit" src="https://img.shields.io/github/commit-activity/m/pk5ls20/txlbot_server?color=%23ff69b4">
    <img alt="stars" src="https://img.shields.io/github/stars/pk5ls20/txlbot_server?style=social">
</div>
<br>
一个适合部署在Windows的腾讯会议全自动挂机软件！  

基于项目[Yewandou7/WemeetSignIn](https://github.com/Yewandou7/WemeetSignIn)构建并加强其功能
<br>
</div>

### 功能介绍
- 全自动挂机
- 运行日志输出
- 挂机失败重试：大幅提高挂机成功概率
- 进入会议时间宽限：实现更有弹性的挂机
- 状态实时推送（基于[Server酱](https://sct.ftqq.com/sendkey)）
### 原理介绍
- 通过模拟点击/输入来入会
- 通过判断屏幕中是否存在入会后显示的状态栏来判断入会是否成功
### 食用方法（保姆教学向）
1. 将该软件加载到一个**可以在你需要的时间内一直运行的电脑**</br>该电脑可以是不用的电脑，也可以是云主机
2. 配置`config.xlsx`
> `config.xlsx`内存储挂机课程的全部信息

| 项名         | 格式                             | 示例                               |
|------------|--------------------------------|----------------------------------|
| lesson     | 正常输入课程名                        | 这是一个测试课堂                         |
| start_time | 以**24小时制**输入                   | 11:14                            |
| meeting_id | xxx xxx xxxx</br>或 xxx xxx xxx | 114 514 1919 </br> 或 810 114 514 |
| password   | 有密码的会议输入密码，无密码的会议输入`xxxxxx`    | /                                |
| day        | 以**阿拉伯数字**的形式输入课在周几上（1-7）      | 6                                |
  
3.配置`dconfig.json`
> `dconfig.json`存储软件参数

| 键名      | 值                                                                                           |
|---------|---------------------------------------------------------------------------------------------|
| xmlpath | 填写[config.xlsx](https://github.com/pk5ls20/txlbot_server/blob/master/config.xlsx)文件所在路径                                                                     |
| pushapi | 填写你的Server酱的[api](https://sct.ftqq.com/sendkey)                                             |
| tmpath  | 填写腾讯会议软件所在路径                                                                                |
| moepath | 填写[listenbot.png](https://github.com/pk5ls20/txlbot_server/blob/master/listenbot.png)文件所在路径 |

json文件示例：
```angular2html
{
  "xmlpath" : "C:\\Users\\txlbot_server\\Desktop\\config.xlsx",
  "pushapi" : "https://sctapi.ftqq.com/NSYGYGTKJAAAAA.send",
  "tmpath" : "C:\\Program Files (x86)\\Tencent\\WeMeet\\wemeetapp.exe",
  "moepath" : "C:\\Users\\txlbot_server\\Desktop\\22222.png"
}
```
4. 运行`main_ucloud.exe`
5. 听课酱开始帮您听课啦~
### 本地部署
>在`txlbot_server`文件夹下终端执行`pip install -r requirement.txt`即可
### 已知Bug
1. 运行于云主机时关闭远程桌面连接导致截屏失效从而导致txlbot无法使用（参见[这个issue](https://github.com/python-pillow/Pillow/issues/2631)）</br>
```angular2html
不完美解决方案:
使用两台云主机（A/B)，A运行txlbot，B用于远程桌面连接A，主机连接B的远程桌面
```
### 下回更新预告
- 使用VoovMeeting入会