using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace HertZ_ExcelAddIn
{
    /// <summary>
    /// 配置读写基类
    /// </summary>
    public class ClsBaseConfig
    {
        #region 属性
        protected string ConfigName { get; set; } //配置文件名（要包含后缀名）
        protected string ConfigPath { get; set; } //配置文件的路径
        protected string RootNodeName { get; set; } //配置文件根节点名

        //当前程序配置文件路径
        protected string ConfigFullName
        {
            get { return Path.Combine(ConfigPath, ConfigName); }
        }
        #endregion

        #region 创建文档和节点
        /// <summary>
        /// 创建Config文件
        /// </summary>
        /// <returns>返回创建是否成功</returns>
        protected bool CreateConfig()
        {
            try
            {
                bool blnExists = File.Exists(this.ConfigFullName);
                if (blnExists) File.Delete(this.ConfigFullName);    //若文件存在，则删除

                //创建配置文件
                XmlDocument xmlDoc = new XmlDocument();
                //创建类型声明节点  
                XmlNode node = xmlDoc.CreateXmlDeclaration("1.0", "utf-8", "");
                xmlDoc.AppendChild(node);
                //创建根节点  
                XmlNode root = xmlDoc.CreateElement(RootNodeName);
                xmlDoc.AppendChild(root);
                //保存xml文件
                xmlDoc.Save(this.ConfigFullName);

                //销毁资源
                root = null;
                node = null;
                xmlDoc = null;
                return true;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message, "创建配置文件出错");
                return false;
            }
        }
        /// <summary>
        /// 创建子节点
        /// </summary>
        /// <param name="xmlDoc">xml文档</param>
        /// <param name="parentNode">父节点</param>
        /// <param name="strName">节点名</param>
        /// <param name="strValue">值</param>
        protected void CreateSubNode(XmlDocument xmlDoc, XmlNode parentNode, string strName, string strValue)
        {
            XmlNode node = xmlDoc.CreateNode(XmlNodeType.Element, strName, null);
            node.InnerText = strValue;
            parentNode.AppendChild(node);

            node = null;
        }
        #endregion

        #region 获取相关Xml对象
        /// <summary>
        /// 获取Xml配置文档
        /// </summary>
        /// <returns>返回Xml配置文档</returns>
        protected XmlDocument GetDocument()
        {
            bool blnExists = File.Exists(this.ConfigFullName);//判断文件是否存在
            //不存在则创建
            if (!blnExists)
            {
                if (!this.CreateConfig()) return null;
            }

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(this.ConfigFullName);                 //加载xml文件
            return xmlDoc;
        }
        /// <summary>
        /// 获取根节点，若没有则自动创建
        /// </summary>
        /// <param name="xmlDoc">Xml文档</param>
        /// <returns>返回根节点</returns>
        protected XmlNode GetRootNode(XmlDocument xmlDoc)
        {
            XmlNode xmlRoot = xmlDoc.SelectSingleNode(RootNodeName);//找到根节点
            //获取不到，则添加进去
            if (xmlRoot == null)
            {
                xmlRoot = xmlDoc.CreateElement(RootNodeName);
                xmlDoc.AppendChild(xmlRoot);
            }
            return xmlRoot;
        }
        /// <summary>
        /// 获取节点，若没有则自动创建
        /// </summary>
        /// <param name="xmlDoc">Xml文档</param>
        /// <param name="xmlParent">父节点</param>
        /// <param name="strName">节点名</param>
        /// <returns>返回节点</returns>
        protected XmlNode GetNode(XmlDocument xmlDoc, XmlNode xmlParent, string strName)
        {
            XmlNode xmlNode = xmlParent.SelectSingleNode(strName);
            //获取不到，则添加进去
            if (xmlNode == null)
            {
                xmlNode = xmlDoc.CreateElement(strName);
                xmlParent.AppendChild(xmlNode);
            }
            return xmlNode;
        }
        #endregion

        #region 读写方法
        /// <summary>
        /// 读取设置
        /// </summary>
        /// <typeparam name="T">值类型</typeparam>
        /// <param name="strParent">父节点</param>
        /// <param name="strName">参数名</param>
        /// <param name="defaultValue">默认值</param>
        /// <returns>返回读取结果，没有则返回默认值</returns>
        public T ReadConfig<T>(string strParent, string strName, T defaultValue)
        {
            try
            {
                XmlDocument xmlDoc = GetDocument();                     //加载xml文件
                XmlNode xmlRoot = GetRootNode(xmlDoc);                  //获取根节点
                XmlNode xmlParent = GetNode(xmlDoc, xmlRoot, strParent);//获取父节点
                XmlNode xmlNode = GetNode(xmlDoc, xmlParent, strName);  //获取该节点

                //判断是否有内容
                if (xmlNode.InnerText == "")
                {
                    xmlNode.InnerText = defaultValue.ToString();
                    xmlDoc.Save(this.ConfigFullName);
                }
                //返回内容
                string strText = xmlNode.InnerText;
                xmlNode = null;
                xmlParent = null;
                xmlRoot = null;
                xmlDoc = null;
                return (T)Convert.ChangeType(strText, typeof(T));

            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message, "写入配置错误");
                return defaultValue;
            }
        }
        /// <summary>
        /// 写入设置
        /// </summary>
        /// <param name="strParent">父节点</param>
        /// <param name="strName">参数名</param>
        /// <param name="strValue">参数值</param>
        /// <returns>返回布尔值，表示是否写入成功</returns>
        public bool WriteConfig(string strParent, string strName, string strValue)
        {
            try
            {
                XmlDocument xmlDoc = GetDocument();                     //加载xml文件
                XmlNode xmlRoot = GetRootNode(xmlDoc);                  //获取根节点
                XmlNode xmlParent = GetNode(xmlDoc, xmlRoot, strParent);//获取父节点
                XmlNode xmlNode = GetNode(xmlDoc, xmlParent, strName);  //获取该节点
                xmlNode.InnerText = strValue;                           //写值
                xmlDoc.Save(this.ConfigFullName);                       //保存文件

                xmlNode = null;
                xmlParent = null;
                xmlRoot = null;
                xmlDoc = null;
                return true;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message, "读取配置错误");
                return false;
            }
        }
        #endregion
    }
}
