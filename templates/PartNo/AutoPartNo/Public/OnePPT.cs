using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using SolidWorks.Interop.swconst;

 namespace Interop.Office.Core
{

     internal class OneSldPPt
    {
         internal OneSldPPt(string name)
        {
            this.name = name;
        }

        /// <summary>
        /// 属性名
        /// </summary>
         internal string name = "";
        /// <summary>
        /// 属性值
        /// </summary>
         internal string value = null;
        /// <summary>
        /// 解析值
        /// </summary>
         internal string solValue = null;
        /// <summary>
        /// 数据类型：1 文字 2:日期 3数字 4是否
        /// </summary>
         internal int iType = 1;
    }



    /// <summary>
    /// 记录一个属性的信息,保存在Combox的Tag中
    /// </summary>
    internal class OnePPt
    {
        /// <summary>
        /// 属性表中的一行，生成一个属性信息
        /// </summary>
        internal OnePPt(DataRow dr)
        {
            this.pName = this.pMDName = dr["属性名称"].ToString().Trim();
            if (dr.Table.Columns.Contains("MDName"))
            {
                object o = dr["MDName"];
                if (o != null && o.ToString() != "")
                {
                    this.pMDName = dr["MDName"].ToString().Trim();
                }
            }

            this.DefaultVal = dr["默认值"].ToString().Trim();
            this.strRef = dr["关联数据"].ToString();

            this.iType = (int)swCustomInfoType_e.swCustomInfoText;
            string pType = dr["数据类型"].ToString().Trim();
            if (pType == "日期") this.iType = (int)swCustomInfoType_e.swCustomInfoDate;
            else if (pType == "整数") this.iType = (int)swCustomInfoType_e.swCustomInfoNumber;
            else if (pType == "小数") this.iType = (int)swCustomInfoType_e.swCustomInfoDouble;
            else if (pType == "是或否") this.iType = (int)swCustomInfoType_e.swCustomInfoYesOrNo;

            if (dr.Table.Columns.Contains("InGeneral"))
            {
                object o = dr["InGeneral"];
                if (o != null && o.ToString() != "" && (bool)(dr["InGeneral"]) == true)
                {
                    this.bGeneralPPT = (bool)(dr["InGeneral"]);
                }
            }

            if (dr.Table.Columns.Contains("别名"))
            {
                object o = dr["别名"];
                if (o != null && o.ToString().Length > 0)
                {
                    this.pName2 = o.ToString();
                }
            }
            if (dr.Table.Columns.Contains("SaveLastVal"))
            {
                object o = dr["SaveLastVal"];
                if (o != null && o.ToString().Length > 0)
                {
                    this.bSaveLastVal = (bool)(o);
                }
            }
        }

        /// <summary>
        /// 有可能是表达式 //如: "\"SW-Mass@@" + cfgname2 + "@" + swModel.GetTitle() + "\"";
        /// </summary>
        internal string AttValue = "";
        /// <summary>
        /// 实际值
        /// </summary>
        internal string AttSolveValue = "";

        /// <summary>
        /// 迈迪固定的属性名称
        /// </summary>
        internal string pMDName = "";
        /// <summary>
        /// 用户可改的属性名称
        /// </summary>
        internal string pName = "";
        /// <summary>
        /// 别名,V5.0新添加的,但是还没有用到
        /// </summary>
        internal string pName2 = "";
        /// <summary>
        /// 默认值,数量的默认值是1,材料的默认值是SW-Material,质量的默认值是SW-Mass
        /// </summary>
        internal string DefaultVal = "";
        /// <summary>
        /// 是否记录上一个输入的值，并默认选中这个值
        /// </summary>
        internal bool bSaveLastVal = false;
        /// <summary>
        /// 数据类型
        /// </summary>
        internal int iType = 0;

        /// <summary>
        /// 关联数据
        /// </summary>
        internal string strRef = "";

        /// <summary>
        /// 是否原先就有这个属性
        /// </summary>
        internal bool bHas = false;
        /// <summary>
        /// 是否原先就有这个属性的别名,如果有pName1这个属性, bHas2=false
        /// </summary>
        internal bool bHas2 = false;

        /// <summary>
        /// 是自定义属性吗，默认否
        /// </summary>
        internal bool bGeneralPPT = false;


    }

}
