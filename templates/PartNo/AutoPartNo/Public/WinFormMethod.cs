using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Collections;

namespace Interop.Office.Core
{
    /// <summary>
    /// 有关控件操作的函数
    /// </summary>
    internal class WinFormMethod
    {
        /// <summary>
        /// 在两个ListBox之间移动选项,即把一个ListBox中的选中项移动到另一个ListBox中
        /// </summary>
        internal void RemoveListItemBetweenListBox(ListBox listA, ListBox listB)
        {
            int index = listA.SelectedIndex;
            if (index < 0) return;

            string txt = listA.Items[index].ToString();
            listA.Items.RemoveAt(index);
            if (listA.Items.Count > index + 1)
            {
                listA.SelectedIndex = index;
            }

            listB.Items.Add(txt);
        }


        /// <summary>
        /// 把ListA中的选中项复制一份,移动到ListB中,移动前先检查ListB中是否有了,
        /// 如果CanRepeat=true,则可以重复
        /// 如果DeleteRemoveItem=true,则移动后删除ListA中的项
        /// </summary>
        internal void RemoveListItemBetweenListBox(ListBox listA, ListBox listB, bool CanRepeat, bool DeleteRemoveItem)
        {
            int index = listA.SelectedIndex;
            if (index < 0) return;

            string txt = listA.Items[index].ToString();

            if (DeleteRemoveItem)//移动后删除原来的
            {
                listA.Items.RemoveAt(index);
                if (listA.Items.Count > index) listA.SelectedIndex = index;
                else listA.SelectedIndex = index - 1;
            }

            if (!CanRepeat)
            {
                for (int i = 0; i < listB.Items.Count; i++)
                {
                    string s = listB.Items[i].ToString();
                    if (s == txt) return;
                }
            }
            listB.Items.Add(txt);

        }

        /// <summary>
        /// 向listA中添加一个新项( CanRepeat =是否允许重复, index=插入的位置,-1标识最后)
        /// </summary>
        internal void AddListItem(ListBox listA, string strVal, bool CanRepeat, int index)
        {
            if (!CanRepeat)
            {
                for (int i = 0; i < listA.Items.Count; i++)
                {
                    string s = listA.Items[i].ToString();
                    if (s == strVal) return;
                }
            }

            if (index == -1)
            {
                listA.Items.Add(strVal);
            }
            else
            {
                if (listA.Items.Count <= index) index = listA.Items.Count - 1;

                listA.Items.Insert(index, strVal);
            }
        }
        /// <summary>
        /// 删除一项
        /// </summary>
        internal void DeleteListItem(ListBox listA, string strVal)
        {
            if (listA.Items.Contains(strVal))
            {
                listA.Items.Remove(strVal);
            }
            else
            {
                for (int i = listA.Items.Count - 1; i >= 0; i--)
                {
                    if (listA.Items[i].ToString() == strVal)
                    {
                        listA.Items.RemoveAt(i); break;
                    }
                }
            }
        }



        /// <summary>
        /// 移动ListItem项往上或者往下
        /// </summary>
        /// <param name="isRemoveUP"></param>
        internal void RemoveListItemUpOrDown(ListBox listA, bool isRemoveUP)
        {
            int index = listA.SelectedIndex;
            if (index >= 0)//有选中项
            {
                string txt = listA.Items[index].ToString();
                if (isRemoveUP)//向上移动
                {
                    if (index == 0) return;//已经是第0条了
                    listA.Items.RemoveAt(index);
                    index = index - 1;
                }
                else//向下移动
                {
                    if (index == listA.Items.Count - 1) return;//已经是最后一条了
                    listA.Items.RemoveAt(index);
                    index = index + 1;
                }
                listA.Items.Insert(index, txt);
                listA.SelectedIndex = index;
            }
        }


        /// <summary>
        /// 移动ListItem项往上或者往下//并返回当前选中的Index
        /// </summary>
        /// <param name="isRemoveUP"></param>
        internal int RemoveListItemUpOrDown2(ListBox listA, bool isRemoveUP)
        {
            int index = listA.SelectedIndex;
            if (index >= 0)//有选中项
            {
                string txt = listA.Items[index].ToString();
                if (isRemoveUP)//向上移动
                {
                    if (index == 0) return 0;//已经是第0条了
                    listA.Items.RemoveAt(index);
                    index = index - 1;
                }
                else//向下移动
                {
                    if (index == listA.Items.Count - 1) return index;//已经是最后一条了
                    listA.Items.RemoveAt(index);
                    index = index + 1;
                }
                listA.Items.Insert(index, txt);
               
                return index;
            }
            return -1;
        }


        /// <summary>
        /// 得到Combox中和strVal最接近的index,这里如果转换不成数字就有可能出错
        /// </summary>
        internal int getNearValue(ComboBox cmb, string strVal)
        {
            int indexmin = -1;
            double dbmin = 1000;

            double dL = Convert.ToDouble(strVal);
            for (int i = 0; i < cmb.Items.Count; i++)
            {
                try
                {
                    string str = cmb.Items[i].ToString().Trim();
                    if (str.Length == 0) continue;
                    double dba = Convert.ToDouble(str);

                    double dist = Math.Abs(dL - dba);
                    if (dist < dbmin)
                    {
                        dbmin = dist;
                        indexmin = i;
                    }
                }
                catch (Exception e)
                {
                    #if Debug
                    StringOperate.Alert("getNearValue Error:" + e.Message);
                    #endif
                }
            }

            return indexmin;
        }


        /// <summary>
        /// 获取datagridveiw数据区域的尺寸宽度，就是所有的列的宽度之和( dgv, 最大宽度-为0时不管，最小宽度-为0时不管）
        /// </summary>
        internal int getDGVDataAreaWidth(DataGridView dgv,int imax,int imin)
        {
            int a = dgv.RowHeadersWidth;
            foreach (DataGridViewColumn dc in dgv.Columns)
            {
                a += dc.Width;
                a++;
            }
            if (imax > 0) a = Math.Min(a, imax);
            if (imin > 0) a = Math.Max(a, imin);
            return a;
        }


        /// <summary>
        /// 获取datagridveiw数据区域的尺寸高度，就是行数* 行高( dgv, 最大高度-为0时不管，最小高度-为0时不管）
        /// </summary>
        internal int getDGVDataAreaHeight(DataGridView dgv, int imax, int imin)
        {
            int a = dgv.ColumnHeadersHeight;
            if (dgv.Rows.Count > 0)
            {
                a += dgv.Rows.Count * dgv.RowTemplate.Height;
            }
            if (imax > 0) a = Math.Min(a, imax);
            if (imin > 0) a = Math.Max(a, imin);
            return a;
        }


        /// <summary>
        /// 设置一个DataGridViewCell的值，自动转换类型,若成功，返回“”，否则返回错误提示
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="objValue"></param>
        /// <returns></returns>
        internal string SetDGVCellValue(DataGridViewCell cell, object objValue)
        {
            if (objValue == null || objValue.ToString().Length == 0)
            {
                cell.Value = null; return "";
            }
            Type T = cell.ValueType;
            try
            {
                if (T == typeof(System.String))
                    cell.Value = objValue.ToString();
                else if (T == typeof(System.Int32))
                    cell.Value = Convert.ToInt32(objValue);
                else if (T == typeof(System.Int16))
                    cell.Value = Convert.ToInt16(objValue);
                else if (T == typeof(System.Single))
                    cell.Value = Convert.ToSingle(objValue);
                else if (T == typeof(System.Boolean))
                    cell.Value = Convert.ToBoolean(objValue);
                else if (T == typeof(System.DateTime))
                    cell.Value = Convert.ToDateTime(objValue);
                else
                    cell.Value = objValue;
            }
            catch (Exception ea)
            {
                return "第[" + cell.RowIndex + "]行，第[" + cell.ColumnIndex + "]列赋值出错：" + ea.Message + System.Environment.NewLine;
            }
            return "";
        }


    }


    /// <summary>
    /// 向任何控件添加一个ToolTip
    /// </summary>
    internal class AddToolTips
    {
        /// <summary>
        /// 为控件添加ToolTipText
        /// </summary>
        public void AddToolTipTxt(Control ctl, string txt)
        {
            ToolTip tooltip = new ToolTip();

            tooltip.AutoPopDelay = 5000;
            tooltip.InitialDelay = 100;
            tooltip.ReshowDelay = 100;
            tooltip.ShowAlways = true;

            tooltip.SetToolTip(ctl, txt);
        }
    }


    /// <summary>
    /// 操作树视图的类
    /// </summary>
    internal class TreeViewMethod
    {

        /// <summary>
        /// 得到一个树的所有节点，不分级别
        /// </summary>
        internal TreeNode[] getAllTreeNodes(TreeView tv)
        {
            ArrayList ar = new ArrayList();
            foreach (TreeNode tn in tv.Nodes)
            {
                TraverseNode(tn, ref ar);
            }

            TreeNode[] tnArr = new TreeNode[ar.Count];
            for (int i = 0; i < ar.Count; i++)
            {
                tnArr[i] = (TreeNode)ar[i];
            }

            return tnArr;
        }
        /// <summary>
        /// 得到一个树的所有节点，不分级别,tn自身也加进去
        /// </summary>
        internal TreeNode[] getAllTreeNodes(TreeNode tn)
        {
            ArrayList ar = new ArrayList();

            TraverseNode(tn, ref ar);

            TreeNode[] tnArr = new TreeNode[ar.Count];
            for (int i = 0; i < ar.Count; i++)
            {
                tnArr[i] = (TreeNode)ar[i];
            }

            return tnArr;
        }
        private void TraverseNode(TreeNode tn,ref ArrayList ar)
        {
            ar.Add(tn);//至少有偶一个元素，不至于出错
            if (tn.Nodes.Count > 0)
            {
                foreach (TreeNode tnsub in tn.Nodes)
                {
                    TraverseNode(tnsub, ref ar);
                }
            }
        }
        /// <summary>
        /// 得到一个树的指定名称的节点,没有返回null
        /// </summary>
        private TreeNode GetNodeByName(string name, TreeView tv)
        {
            TreeNode[] tnArr = this.getAllTreeNodes(tv);
            foreach (TreeNode tn in tnArr)
            {
                if (tn.Name == name) return tn;
            }
            return null;
        }

    }


}
