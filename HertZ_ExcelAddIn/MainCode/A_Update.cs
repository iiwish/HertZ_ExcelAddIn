using Newtonsoft.Json;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace Update
{
    class Update
    {
        private readonly string jsonURL = "";//github release api

        ///<summary>
        ///读取json文件
        ///</summary>
        public string GetJson(string jsonURL)//传入网址
        {
            try
            {
                string pagejson = "";
                WebClient MyWebClient = new WebClient
                {
                    Credentials = CredentialCache.DefaultCredentials//获取或设置用于向Internet资源的请求进行身份验证的网络凭据
                };
                Byte[] pageData = MyWebClient.DownloadData(jsonURL); //从指定网站下载数据
                MemoryStream ms = new MemoryStream(pageData);
                using (StreamReader sr = new StreamReader(ms, Encoding.GetEncoding("GB2312")))
                {
                    pagejson = sr.ReadLine();
                }
                return pagejson;
            }
            catch
            {
                MessageBox.Show("获取更新失败，请检查网络连接");
                return null;
            }
            
        }

        /// <summary>
        /// Json 字符串 转换为 DataTable数据集合
        /// </summary>
        /// <param name="json"></param>
        /// <returns></returns>
        public static Dictionary<string,string> GetJsonDic(string json)
        {
            
            if (string.IsNullOrEmpty(json)) { return new Dictionary<string, string>(); }
            Dictionary<string, string> jsonDict = JsonConvert.DeserializeObject<Dictionary<string, string>>(json);
            return jsonDict;
        }

        ///<summary>
        ///检查更新
        ///</summary>
        public void CheckUpdate()
        {
            //读取版本json
            Dictionary<string, string> jsonDic = GetJsonDic(GetJson(jsonURL));
            
            //从我的文档读取配置
            string strPath = Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments);
            HertZ_ExcelAddIn.ClsThisAddinConfig clsConfig = new HertZ_ExcelAddIn.ClsThisAddinConfig(strPath);

            //从父节点Info中读取配置名为Vertion的值，该值为字符串
            string VerInfo = clsConfig.ReadConfig<string>("Info", "Vertion", "0.0.0.01");

            //设置更新文件夹

            try
            {
                if (jsonDic["name"] != VerInfo)
                {
                    using (WebClient web = new WebClient())
                    {
                        web.DownloadFile(jsonDic["url"], strPath);
                    }
                }
            }
            catch
            {
                MessageBox.Show("更新失败，请检查网络连接");
            }
        }

        /////////
        ///



    }
}
