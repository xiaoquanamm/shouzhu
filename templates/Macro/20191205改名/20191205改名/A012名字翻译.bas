Attribute VB_Name = "A012"
Public Function 名字翻译(同名, 零件名字)


Dim strMMMD As String

strMMMD = 零件名字.Text

  Select Case strMMMD
  
  Case Is = "底板"
  
  同名.List(3) = "floor"
  
  Case Is = "机架"
         
   同名.List(3) = "Frame"
         
  Case Is = "防护罩"
  
  同名.List(3) = "Shield"
         
  Case Is = "垫板"
         
   同名.List(3) = "Padding block "
         
  Case Is = "站脚"
  
   同名.List(3) = "Stand foot"
         
  Case Is = "加强筋"
  
   同名.List(3) = "Reinforced tendons"
      
  Case Is = "滑块"
  
   同名.List(3) = "Slide"
      
  Case Is = "挡板"
  
   同名.List(3) = "Block"
             
  Case Is = "盖板"
  
   同名.List(3) = "Coverplate"
   
  Case Is = "连接块"
  
   同名.List(3) = "Connection Block"
   
     Case Is = "刀"
  
   同名.List(3) = "Knife"
   
     Case Is = "针"
  
   同名.List(3) = "PIN"
   
     Case Is = "圆柱"
  
   同名.List(3) = "Pillar"
   
     Case Is = "安装块"
  
   同名.List(3) = "Install Block"
   
     Case Is = ""
  
   同名.List(3) = ""
        
  End Select
End Function
