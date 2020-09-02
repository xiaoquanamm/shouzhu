using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Windows.Forms;
using System.IO;
using System.Collections;

namespace Interop.Office.Core
{
    internal class XMLOperate
    {

        /// <summary>
        /// 读取一个XML文档到一个Treeview中,(根据AttributeName)
        /// </summary>
        internal void LoadXmlToTreeView(TreeView tv, string xmlFullName)
        {
            if (!File.Exists(xmlFullName)) return;
            XmlDocument doc = new XmlDocument();
            doc.Load(xmlFullName);

            //获取文档的跟节点
            XmlNode root = doc.DocumentElement;
            this.loadOneNode(tv ,null , root);
        }
        internal Hashtable LoadXmlToTreeNode(TreeNode tn, string xmlFullName)
        {
            if (!File.Exists(xmlFullName)) return null;
            XmlDocument doc = new XmlDocument();
            doc.Load(xmlFullName);

            //获取文档的跟节点
            XmlNode root = doc.DocumentElement;
            this.loadOneNode(null , tn , root);

            //获取父节点所有的属性
            Hashtable htAttr = new Hashtable();
            foreach (XmlAttribute attr in root.Attributes)
            {
                if (!htAttr.ContainsKey(attr.Name)) htAttr.Add(attr.Name, attr.Value);
            }
            return htAttr;
        }
        private void loadOneNode(TreeView tv,TreeNode tn, XmlNode xmlNode)
        {
            for (int i = 0; i < xmlNode.ChildNodes.Count; i++)
            {
                XmlNode subXmlNod = xmlNode.ChildNodes[i];

                TreeNode subTN = new TreeNode();
                bool bExpand = false;

                //获取所有的属性
                foreach (XmlAttribute att in subXmlNod.Attributes)
                {
                    if (att.Name == "Nm")
                    {
                        subTN.Name = att.Value;
                    }
                    else if (att.Name == "T")
                    {
                        subTN.Text = att.Value;
                    }
                    else if (att.Name == "Idx")
                    {
                        subTN.ImageIndex = Convert.ToInt16(att.Value);
                    }
                    else if (att.Name == "sIdx")
                    {
                        subTN.SelectedImageIndex = Convert.ToInt16(att.Value);
                    }
                    else if (att.Name == "Ep")
                    {
                        if (att.Value == "1") bExpand = true;
                    }
                    else if (att.Name == "C")
                    {
                        if (att.Value == "1") subTN.Checked = true;
                    }
                    else if (att.Name == "Tg")
                    {
                        subTN.Tag = att.Value;
                    }
                    else if (att.Name == "CompID")
                    {
                        subTN.ToolTipText = "CompID:" + att.Value;//企业ID
                    }
                    else if (att.Name == "MXID")
                    {
                        subTN.ToolTipText = "MXID:" + att.Value;//模型ID
                    }
                }
                //把树节点加入进去
                if (tv != null)
                {
                    tv.Nodes.Add(subTN);
                }
                else
                {
                    tn.Nodes.Add(subTN);
                }
               
                //遍历子节点
                loadOneNode(null,subTN , subXmlNod);

                if (bExpand)
                {
                    subTN.Expand();
                }
            }
        }


        /// <summary>
        /// 把一个树视图保存成XML文件，Tag属性只能保存String,其它格式不保存
        /// Nm;  //name
        /// T;   //Text
        /// Idx; //imageIndex
        /// sIdx;//selectedImageIndex
        /// C;   //checked   （1:选中）
        /// Ep;  //Expaneded (1:展开）
        /// Tg;  //只保存String类型的，Object类型的忽略
        /// </summary>
        internal void SeavTreeViewToXml(TreeNodeCollection tnc, string xmlFullName, string topTitle)
        {
            XmlDocument doc = new XmlDocument();
            //创建声明
            XmlNode decl = doc.CreateXmlDeclaration("1.0", "UTF-8", null);
            doc.AppendChild(decl);

            //加入一个根元素
            XmlElement topElem = doc.CreateElement("node");
            topElem.SetAttribute("name", topTitle);
            doc.AppendChild(topElem);

            //保存所有的元素
            foreach (TreeNode tn in tnc)
            {
                AddNodes(tn, doc, topElem);
            }

            doc.Save(xmlFullName);
        }
        private void AddNodes(TreeNode tn, XmlDocument doc, XmlElement topElem)
        {
            XmlElement subElem = doc.CreateElement("node");
            subElem.SetAttribute("Nm", tn.Name);
            subElem.SetAttribute("T", tn.Text);
            if (tn.ImageIndex != -1)
            {
                subElem.SetAttribute("Idx", tn.ImageIndex.ToString());
            }
            if (tn.SelectedImageIndex != -1)
            {
                subElem.SetAttribute("sIdx", tn.SelectedImageIndex.ToString());
            }
            if (tn.IsExpanded)
            {
                subElem.SetAttribute("Ep", "1");
            }
            if (tn.Checked)
            {
                subElem.SetAttribute("C", "1");
            }
            if (tn.Tag != null && tn.Tag.GetType() == typeof(System.String))
            {
                subElem.SetAttribute("Tg", tn.Tag.ToString());
            }
            if (tn.ToolTipText.StartsWith("CompID:"))
            {
                subElem.SetAttribute("CompID", tn.ToolTipText.Substring(tn.ToolTipText.IndexOf(':') + 1));
            }
            if (tn.ToolTipText.StartsWith("MXID:"))
            {
                subElem.SetAttribute("MXID", tn.ToolTipText.Substring(tn.ToolTipText.IndexOf(':') + 1));
            }

            foreach (TreeNode tnsub in tn.Nodes)
            {
                AddNodes(tnsub, doc, subElem);
            }
            topElem.AppendChild(subElem);
        }


        //创建xml文件方法二
        private void btn2_OnClick(object sender, EventArgs e)
        {
            //XmlDocument xmldoc = new XmlDocument(); //创建空的XML文档
            //xmldoc.LoadXml("<?xml version='1.0' encoding='gb2312'?>" +
            //"<bookstore>" +
            //"<book genre='fantasy' ISBN='2-3631-4'>" +
            //"<title>Oberon's Legacy</title>" +
            //"<author>Corets, Eva</author>" +
            //"<price>5.95</price>" +
            //"</book>" +
            //"</bookstore>");
            //xmldoc.Save(("bookstore2.xml")); //保存
        }



    }
}
