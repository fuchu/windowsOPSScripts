
'用于公司服务器日志文件夹清理
'实现对指定的windows目录下的时间超过30天的文件进行删除操作，将递归匹配,支持多个目录
'Tips:根据https://www.iis.net/learn/manage/provisioning-and-managing-iis/managing-iis-log-file-storage#02改写，
'根据的是文件的创建时间
'使用方法计划任务[本脚本] [路径1] [路径2]
Set oArgs = WScript.Arguments
For Each s In oArgs
    sLogFolder = s
    iMaxAge = 30   'in days
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    set colFolder = objFSO.GetFolder(sLogFolder)
    For Each colSubfolder in colFolder.SubFolders
            Set objFolder = objFSO.GetFolder(colSubfolder.Path)
            Set colFiles = objFolder.Files
            For Each objFile in colFiles
                    iFileAge = now-objFile.DateCreated
                    if iFileAge > (iMaxAge+1)  then
                            objFSO.deletefile objFile, True
                    end if
            Next
    Next
    set colFiles = colFolder.Files
    For Each objFilePar in colFiles
            iFileAge = now-objFilePar.DateCreated
            if iFileAge > (iMaxAge+1)  then
                    objFSO.deletefile objFilePar, True
            end if
    Next
Next
Set oArgs = Nothing