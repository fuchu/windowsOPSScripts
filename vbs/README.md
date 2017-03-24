# LogCleaner.vbs
## 作用
用于公司服务器日志文件夹清理
实现对指定的windows目录下的时间超过30天的文件进行删除操作，将递归匹配,支持多个目录
## 使用方法
计划任务[本脚本] [路径1] [路径2]
## Tips:
 1. 根据https://www.iis.net/learn/manage/provisioning-and-managing-iis/managing-iis-log-file-storage#02改写，
 2. 根据的是文件的创建时间