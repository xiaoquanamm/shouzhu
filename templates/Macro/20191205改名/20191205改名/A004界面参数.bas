Attribute VB_Name = "A004"
Public Function 界面参数(材料, 处理方式, 零件名字, 加工件类别, 改名选择, 输出文件选择)

    材料.AddItem "S45C"
    材料.AddItem "SKD11"
    材料.AddItem "A6061"
    材料.AddItem "SUS304"
    材料.AddItem "SPCC"
    材料.AddItem "POM"
    材料.AddItem "PC"
    材料.AddItem "SUJ2"
    材料.AddItem "ASP23"
    材料.AddItem "优力胶"
    材料.AddItem "亚克力"
    材料.AddItem "--"
    材料.ListIndex = 0

    处理方式.AddItem "Hard Cr"
    处理方式.AddItem "HRC58~62°Hard Cr"
    处理方式.AddItem "NI"
    处理方式.AddItem "HRC58~62°NI"
    处理方式.AddItem "HRC64~66"
    处理方式.AddItem "透明"
    处理方式.AddItem "橘纹RAL7035"
    处理方式.AddItem "喷砂，本色阳极"
    处理方式.AddItem "硬质阳极"
    处理方式.AddItem "红色烤漆"
    处理方式.AddItem "黑色"
    处理方式.AddItem "--"
    处理方式.ListIndex = 0

    
    零件名字.AddItem "安装块"
    零件名字.AddItem "圆柱"
    零件名字.AddItem "连接块"
    零件名字.AddItem "垫板"
    零件名字.AddItem "加强筋"
    零件名字.AddItem "滑块"
    零件名字.AddItem "挡块"
    零件名字.AddItem "站脚"
    零件名字.AddItem "刀"
    零件名字.AddItem "盖板"
    零件名字.AddItem "防护罩"
    零件名字.AddItem "机架"
    零件名字.AddItem "底板"
    零件名字.ListIndex = 0

       
    改名选择.AddItem "加工"
    改名选择.AddItem "装配1"
    改名选择.AddItem "装配2"
    改名选择.AddItem "标件"
    改名选择.AddItem "附件"
    改名选择.ListIndex = 0
       
       
    输出文件选择.AddItem "Dwg/Step"
    输出文件选择.AddItem "批量Bom"
    输出文件选择.AddItem "单个Bom"
    输出文件选择.AddItem "项目建立"
    输出文件选择.AddItem "属性清除"
    输出文件选择.AddItem "Part"
    输出文件选择.ListIndex = 0
    
End Function
