using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using System.Collections;
using System.Runtime.Serialization.Formatters.Binary;
using System.Runtime.Serialization;
using System.IO;
using System.Runtime.InteropServices;
using System.Reflection;
using Interop.Office.Core;

namespace MDTV
{
    public class TVS
    {
        /// <summary>
        /// 遍历查找一个树节点
        /// </summary>
        public TreeNode GetNodeByFullPath(string fullPath, TreeNodeCollection tnArr)
        {
            if (fullPath.Length == 0)
            {
                return null;
            }
            TreeNode nodeByFullPath = null;
            foreach (TreeNode node2 in tnArr)
            {
                if (node2.FullPath == fullPath)
                {
                    return node2;
                }
                if ((nodeByFullPath == null) && (node2.Nodes.Count > 0))
                {
                    if ((node2.Nodes.Count == 1) && (node2.Nodes[0].Text.ToLower() == "temp"))
                    {
                        node2.Expand();
                    }
                    nodeByFullPath = this.GetNodeByFullPath(fullPath, node2.Nodes);
                }
                if (nodeByFullPath != null)
                {
                    return nodeByFullPath;
                }
            }
            return nodeByFullPath;
        }
        /// <summary>
        /// 根据全路径（key）得到树节点(效率高，没有返回空）(注意树节点都是小写）
        /// </summary>
        public TreeNode GetNodeByFullPath(string fullPath, TreeView tv)
        {
            char[] carr = new char[] { '\\' };
            fullPath = fullPath.ToLower();
            fullPath = fullPath.Replace("\\\\", "\\").TrimStart(carr).TrimEnd(carr);
            if (fullPath.ToLower().StartsWith("md3dparts"))
            {
                fullPath = fullPath.Substring(fullPath.IndexOf('\\') + 1);
            }
            string[] keys = fullPath.Split(carr);
            if (keys.Length == 0) return null;

            TreeNode tn = tv.Nodes[keys[0]];
            if (tn == null) return null;

            for (int i = 1; i < keys.Length; i++)
            {
                string s = keys[i];
                if (tn.Nodes.ContainsKey(s))
                {
                    if (i == keys.Length - 1)
                    {
                        return tn.Nodes[s];
                    }
                    else
                    {
                        tn = tn.Nodes[s];
                    }
                }
                else
                {
                    return null;
                }
            }
            return null;
        }
        /// <summary>
        /// 得到一个树节点的所有子节点数组（还有子节点的以“：”结尾），若这个树节点没有子节点返回null
        /// </summary>
        public string[] GetFileAndFolderNames(string fullPath, TreeView tv)
        {
            TreeNode tn = this.GetNodeByFullPath(fullPath, tv);
            if (tn == null || tn.Nodes.Count == 0) return null;

            string[] strArr = new string[tn.Nodes.Count];
            for (int i = 0; i < strArr.Length; i++)
            {
                string s = tn.Nodes[i].Name;
                if (tn.Nodes[i].Nodes.Count>0) s += ":";
                strArr[i] = s;
            }
            return strArr;
        }


        public Hashtable LoadTreeViewData(TreeNode tn, string path)
        {
            if (path.ToLower().EndsWith(".xml") || path.ToLower().EndsWith(".mdzip"))
            {
                XMLOperate xmlop = new XMLOperate();
                return  xmlop.LoadXmlToTreeNode(tn, path);
            }
            else
            {
                BinaryFormatter formatter = new BinaryFormatter();
                //如果没有这一句，无法实现不同版本的反序列化
                formatter.Binder = new UBinder();
                Stream serializationStream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
                ((TVD)formatter.Deserialize(serializationStream)).PopulateTree(tn);
                serializationStream.Close();
                return null;
            }
        }
        public void LoadTreeViewData(TreeView treeView, string path)
        {
            if (path.ToLower().EndsWith(".xml") || path.ToLower().EndsWith(".mdzip"))
            {
                XMLOperate xmlop = new XMLOperate();
                xmlop.LoadXmlToTreeView(treeView, path);
            }
            else
            {
                BinaryFormatter formatter = new BinaryFormatter();
                formatter.Binder = new UBinder();
                Stream serializationStream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
                ((TVD)formatter.Deserialize(serializationStream)).PopulateTree(treeView);
                serializationStream.Close();
            }
        }

        public void SaveTreeViewData(TreeNode tn, string path)
        {
            if (path.ToLower().EndsWith(".xml") || path.ToLower().EndsWith(".mdzip") )
            {
                XMLOperate xmlop = new XMLOperate();
                xmlop.SeavTreeViewToXml(tn.Nodes , path, "MD");
            }
            else
            {
                BinaryFormatter formatter = new BinaryFormatter();
                Stream serializationStream = new FileStream(path, FileMode.Create);
                formatter.Serialize(serializationStream, new TVD(tn));
                serializationStream.Close();
            }
        }
        public void SaveTreeViewData(TreeView tv, string path)
        {
            if (path.ToLower().EndsWith(".xml") || path.ToLower().EndsWith(".mdzip"))
            {
                XMLOperate xmlop = new XMLOperate();
                xmlop.SeavTreeViewToXml(tv.Nodes , path,"MD");
            }
            else
            {
                BinaryFormatter formatter = new BinaryFormatter();
                Stream serializationStream = new FileStream(path, FileMode.Create);
                formatter.Serialize(serializationStream, new TVD(tv));
                serializationStream.Close();
            }
        }


        /// <summary>
        /// TreeViewData
        /// </summary>
        [Serializable, StructLayout(LayoutKind.Sequential)]
        public struct TND
        {
            public string Nm;//name
            public string T;//Text
            public int Idx;//imageIndex
            public int sIdx;//selectedImageIndex
            public bool C; //checked
            public bool Ep;//Expaneded
            public object Tag;
            public TVS.TND[] nds;
            public TND(TreeNode TN)
            {
                this.Nm = TN.Name;
                this.T = TN.Text;
                this.Idx = TN.ImageIndex;
                this.sIdx = TN.SelectedImageIndex;
                this.C = TN.Checked;
                this.Ep = TN.IsExpanded;
                this.nds = new TVS.TND[TN.Nodes.Count];
                if ((TN.Tag != null) && TN.Tag.GetType().IsSerializable)
                {
                    this.Tag = TN.Tag;
                }
                else
                {
                    this.Tag = null;
                }
                if (TN.Nodes.Count != 0)
                {
                    for (int i = 0; i <= (TN.Nodes.Count - 1); i++)
                    {
                        this.nds[i] = new TVS.TND(TN.Nodes[i]);
                    }
                }
            }

            public TreeNode ToTreeNode()
            {
                TreeNode node = new TreeNode(this.T, this.Idx, this.sIdx);
                node.Name = this.Nm;
                node.Checked = this.C;
                node.Tag = this.Tag;
                if (this.Ep)
                {
                    node.Expand();
                }
                if ((this.nds == null) && (this.nds.Length == 0))
                {
                    return null;
                }
                if ((node == null) || (this.nds.Length != 0))
                {
                    for (int i = 0; i <= (this.nds.Length - 1); i++)
                    {
                        node.Nodes.Add(this.nds[i].ToTreeNode());
                    }
                }
                return node;
            }
        }

        /// <summary>
        /// TreeViewData
        /// </summary>
        [Serializable, StructLayout(LayoutKind.Sequential)]
        public struct TVD
        {
            public TVS.TND[] Nodes;
            public TVD(TreeView treeview)
            {
                this.Nodes = new TVS.TND[treeview.Nodes.Count];
                if (treeview.Nodes.Count != 0)
                {
                    for (int i = 0; i <= (treeview.Nodes.Count - 1); i++)
                    {
                        this.Nodes[i] = new TVS.TND(treeview.Nodes[i]);
                    }
                }
            }

            public TVD(TreeNode tn)
            {
                this.Nodes = new TVS.TND[tn.Nodes.Count];
                if (tn.Nodes.Count != 0)
                {
                    for (int i = 0; i <= (tn.Nodes.Count - 1); i++)
                    {
                        this.Nodes[i] = new TVS.TND(tn.Nodes[i]);
                    }
                }
            }

            public void PopulateTree(TreeView treeview)
            {
                if ((this.Nodes != null) && (this.Nodes.Length != 0))
                {
                    treeview.BeginUpdate();
                    for (int i = 0; i <= (this.Nodes.Length - 1); i++)
                    {
                        treeview.Nodes.Add(this.Nodes[i].ToTreeNode());
                    }
                    treeview.EndUpdate();
                }
            }

            public void PopulateTree(TreeNode tn)
            {
                if ((this.Nodes != null) && (this.Nodes.Length != 0))
                {
                    for (int i = 0; i <= (this.Nodes.Length - 1); i++)
                    {
                        tn.Nodes.Add(this.Nodes[i].ToTreeNode());
                    }
                }
            }
        }

    }

    //如果没有这一句，无法实现不同版本的反序列化
    public class UBinder : SerializationBinder
    {
        public override Type BindToType(string assemblyName, string typeName)
        {
            Assembly ass = Assembly.GetExecutingAssembly();
            return ass.GetType(typeName);
        }
    }

}


////解决了.NET反序列化时的“无法找到程序集＂问题，潸然泪下那
////http://www.kuqin.com/dotnet/20111108/314585.html
////1.将dll加入强名称，注册到全局程序集缓存中
////2.在反序列化使用的IFormatter 对象加入Binder 属性,使其获取要反序列化的对象所在的程序集，示例如下:
////public void DeSerialize( byte [] data, int offset)
////{
////     IFormatter formatter = new BinaryFormatter();
////     formatter.Binder = new UBinder();
////     MemoryStream stream = new MemoryStream(data, offset, stringlength);
////    this .m_bodyobject = ( object )formatter.Deserialize(stream);
////}
////public class UBinder:SerializationBinder
////{
////    public override Type BindToType( string assemblyName, string typeName)
////     {
////        Assembly ass = Assembly.GetExecutingAssembly();
////       return ass.GetType(typeName);
////     }
////}