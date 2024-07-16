# popx 使用pop3协议查看，删除，下载所有邮件，下载单邮件的所有附件

- **缘起:** 现在的网络环境对小开发者极其不友好，没有一个可以直链存放文件的地方，有也会因为某种原因访问困难，或者要备案，既然如此，那就不直链了，使用一个邮箱来进行存储，标题就是文件名，文件放到附件中

- **说明:** 目前使用go语言开发出了原型，使用smtp协议发送邮件 (https://github.com/linpinger/smtpx) ，使用pop3协议 (https://github.com/linpinger/popx) 查看标题下载并自动提取附件,删除邮件等操作

- **妄想:** 由于go的特性,可以跨平台使用,目前pc,安卓手机都已实现基本目标,后续只需要简化操作流程,或使用imap协议进行更高级的管理文件,目标是自动分布式存储个人所有文件,只要邮箱够多^_^

- 由于反垃圾邮件系统存在，用smtpx发的邮件大概率被拦截，根本发不出去，尤其标题不是test，木有正文，只有附件

- **编译:** 参见 (http://linpinger.olsoul.com/usr/2017-06-12_golang.html)  下的一般编译方法
  - `go mod init popx`
  - `go mod tidy`
  - `go build -ldflags "-s -w" -buildvcs=false`

- 本项目主要依赖: (https://github.com/knadh/go-pop3) 和 (https://github.com/emersion/go-message)

- 已移除: 主要改进了标题的解码: `emlSubjectDecode()`  ，目前只支持简中，英文的 `?B?` `?Q?` 编码，说多了都是泪

## 日志

- 2024-07-12: 第二版，移除了自己写的大部分转码脚本，原来引用的库里自带了，修改参数

- 2023-07-07: 第一版，基本能用

