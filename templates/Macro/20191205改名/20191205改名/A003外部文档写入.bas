Attribute VB_Name = "A003"
Public Function 界面读取(运参)

    Const ForReading = 1, ForWriting = 2, ForAppending = 3
    Const TristateUseDefault = -2, TristateTrue = -1, TristateFasle = 0
    Dim fs, f, Nfile
    Set fs = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    Set f = fs.GetFile("D:\\Solidworksavedeta11.tmp")
    If f Then
        Set Nfile = f.openastextStream(ForReading, TristateUseDefault)
        '====================================================================
        loaded_机台名字 = Nfile.ReadLine
        loaded_项目编号 = Nfile.ReadLine
        loaded_设计 = Nfile.ReadLine
        loaded_审核 = Nfile.ReadLine
        loaded_批准 = Nfile.ReadLine
        loaded_保存DWG文档位置 = Nfile.ReadLine
        loaded_工程图模板位置 = Nfile.ReadLine
        loaded_左边 = Nfile.ReadLine
        loaded_上边 = Nfile.ReadLine

        Nfile.Close
    End If

        机台名字 = loaded_机台名字
        项目编号 = loaded_项目编号
        设计 = loaded_设计
        审核 = loaded_审核
        批准 = loaded_批准
        保存DWG文档位置 = loaded_保存DWG文档位置
        工程图模板位置 = loaded_工程图模板位置
        左边 = loaded_左边
        上边 = loaded_上边
                
运参.List(0) = 机台名字
运参.List(1) = 项目编号
运参.List(2) = 设计
运参.List(3) = 审核
运参.List(4) = 批准
运参.List(5) = 左边
运参.List(6) = 上边
运参.List(7) = 保存DWG文档位置
运参.List(8) = Data
运参.List(9) = 工程图模板位置
    
End Function

Public Function 外部文档写入界面(运参, 机台名字, 项目编号, 设计, 审核, 批准, 左边, 上边, 保存DWG文档位置, 工程图模板位置, 当前日期)

    '外部文件载入
    Const ForReading = 1, ForWriting = 2, ForAppending = 3
    Const TristateUseDefault = -2, TristateTrue = -1, TristateFasle = 0
    Dim fs, f, Nfile
    Set fs = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    Set f = fs.GetFile("D:\\Solidworksavedeta11.tmp")
    If f Then
        Set Nfile = f.openastextStream(ForReading, TristateUseDefault)
        loaded_机台名字 = Nfile.ReadLine
        loaded_项目编号 = Nfile.ReadLine
        loaded_设计 = Nfile.ReadLine
        loaded_审核 = Nfile.ReadLine
        loaded_批准 = Nfile.ReadLine
        loaded_保存DWG文档位置 = Nfile.ReadLine
        loaded_工程图模板位置 = Nfile.ReadLine
        loaded_左边 = Nfile.ReadLine
        loaded_上边 = Nfile.ReadLine
        
         
        Nfile.Close
    End If

    '添加自定义属性
    Set swapp = GetObject(, "SldWorks.Application")
    Set swpart = swapp.ActiveDoc
    Set swSheet = swpart.GetCurrentSheet
    Set swActiveView = swpart.ActiveDrawingView
    Dim retval As Variant

        机台名字.Text = loaded_机台名字
        项目编号.Text = loaded_项目编号
        设计.Text = loaded_设计
        审核.Text = loaded_审核
        批准.Text = loaded_批准
        保存DWG文档位置.Text = loaded_保存DWG文档位置
        工程图模板位置.Text = loaded_工程图模板位置
        左边.Text = loaded_左边
        上边.Text = loaded_上边
        
        
当前日期.Text = Date
机台名字.Text = 运参.List(0)
项目编号.Text = 运参.List(1)
设计.Text = 运参.List(2)
审核.Text = 运参.List(3)
批准.Text = 运参.List(4)
左边.Text = 运参.List(5)
上边.Text = 运参.List(6)
保存DWG文档位置.Text = 运参.List(7)
运参.List(8) = 当前日期
工程图模板位置 = 运参.List(9)
        
  '==========================================================================
End Function


Public Function 界面录入保存(机台名字, 项目编号, 设计, 审核, 批准, 左边, 上边, 保存DWG文档位置, 工程图模板位置, 显参)

    '创建并写入外部文件===============================================================
    Const ForReading = 1, ForWriting = 2, ForAppending = 3
    Const TristateUseDefault = -2, TristateTrue = -1, TristateFasle = 0
    Dim fs, f, Nfile
    Set fs = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    Set f = fs.CreateTextFile("D:\\Solidworksavedeta11.tmp")
    If f Then
        Set Nfile = f.openastextStream(ForReading, TristateUseDefault)
        f.WriteLine (机台名字)
        f.WriteLine (项目编号)
        f.WriteLine (设计)
        f.WriteLine (审核)
        f.WriteLine (批准)
        f.WriteLine (保存DWG文档位置)
        f.WriteLine (工程图模板位置)
        f.WriteLine (左边)
        f.WriteLine (上边)
        
        Nfile.Close
    End If
    
显参.List(0) = "正确"
显参.List(1) = "参数录入完成"
     
End Function


