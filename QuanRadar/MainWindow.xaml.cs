using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Web;
using Flurl.Http;
using QuanRadar.Helper;
using Newtonsoft.Json.Linq;
using System.Threading;
using Microsoft.Win32;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;

namespace QuanRadar
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {

        /* https://quan.qq.com/node/api2/getHomePageFeed/ */
        //临时保存爬取到的数据
        //private dic

        public MainWindow()
        {
            InitializeComponent();
            if (!Directory.Exists(AppDomain.CurrentDomain.BaseDirectory + "data/"))
            {
                Directory.CreateDirectory(AppDomain.CurrentDomain.BaseDirectory + "data/");
            }

        }
        //使用子线程获取所有数据，避免卡UI
        public delegate void GetAllDataHandler(string _UserID, string _userType);
        public delegate void GetAlldataByExcelFileHandler(string path);
        //爬取所有用户的数据
        private void GetAlldataByExcelFile(string path)
        {
            try
            {
                IWorkbook workbook = null;  //新建IWorkbook对象  
                FileStream fileStream = new FileStream(path, FileMode.Open, FileAccess.Read);
                if (path.IndexOf(".xlsx") > 0) // 2007版本  
                {
                    workbook = new XSSFWorkbook(fileStream);  //xlsx数据读入workbook  
                }
                else if (path.IndexOf(".xls") > 0) // 2003版本  
                {
                    workbook = new HSSFWorkbook(fileStream);  //xls数据读入workbook  
                }
                ISheet sheet = workbook.GetSheetAt(0);  //获取第一个工作表  
                IRow row;// = sheet.GetRow(0);            //新建当前工作表行数据  
                for (int i = 3; i < sheet.LastRowNum; i++)  //对工作表每一行  
                {
                    row = sheet.GetRow(i);   //row读入第i行数据  
                    if (row != null)
                    {
                        string cellValue = row.GetCell(2).ToString();
                        string _temp = "";

                        GetAllDataHandler handler = new GetAllDataHandler(GetAllData);
                        if (!string.IsNullOrWhiteSpace(cellValue))
                        {
                            if (cellValue.Substring(0, 3) == "101")
                            {
                                _temp = cellValue.Substring(4, cellValue.Length - 4);
                                GetAllData(_temp, "101");
                            }
                            else
                            {
                                _temp = cellValue.Substring(2, cellValue.Length - 2);
                                GetAllData(_temp, "2");
                            }
                        }
                    }
                }
                fileStream.Close();
                workbook.Close();
            }
            catch (Exception e)
            {

                MessageBox.Show(e.ToString());
            }
        }
        //获取所有数据
        private void GetAllData(string _UserID, string userType)
        {
            JObject Jobject = null;
            bool isEnd = false;
            string lasttime = null;
            string nickName = "";
            PageFeed pgFeed = new PageFeed();

            //JArray array = (JArray)_tempJObject["data"]["vHomePageFeed"];
            //o9GiTuCSz6w4uJRkPUvZeSNo-2_U
            //o9GiTuOaGdgD_-ZS4YWn5f-p8ZiE
            for (int i = 1; !isEnd; i++)
            {
                //第一次没有lasttime
                Jobject = GetQuanData(_UserID, i, userType);
                //101_c4707a5494b0db899f5c7d073ef6b1c3
                if (Jobject.ToString().Length > 150)
                {
                    if (i == 1)
                    {
                        //为lasttime赋值,以便获得后续数据
                        lasttime = Jobject["data"]["lLastTime"].ToString();
                        if (Jobject["data"]["vHomePageFeed"][0] != null)
                        {
                            nickName = (string)Jobject["data"]["vHomePageFeed"][0]["stPostSummary"]["postUser"]["nickname"];
                        }
                    }
                    else
                    {
                        Jobject = GetQuanData(_UserID, i, userType, lasttime);
                        lasttime = Jobject["data"]["lLastTime"].ToString();
                    }
                    //判断是否为最后一次
                    if (Jobject["data"]["lLastTime"].ToString() == "0")
                    {
                        isEnd = true;
                    }
                    else
                    {
                        //有效数据，进行分析
                        analysePage(Jobject, pgFeed);
                    }
                }
                else
                {
                    isEnd = true;
                }

            }
            pgFeed.Save2Excel(nickName);
        }
        //提交一次请求获取十条数据回来，如果有的话
        private JObject GetQuanData(string UserID, int Page, String accountType, string lasttime = "0")
        {
            //RquestHeader,一般不用修改，如果要修改直接在Chrome浏览器F12复制到这里就可以。
            string heads = @"Accept: application/json, text/javascript
                             Accept-Encoding: gzip, deflate, br
                             Accept-Language: zh-CN,zh;q=0.9
                             Connection: keep-alive
                             Content-Length: 155
                             Content-Type: application/x-www-form-urlencoded
                             Host: quan.qq.com
                             Origin: https://quan.qq.com
                             Referer: https://quan.qq.com/
                             Sec-Metadata: destination="", target=subresource, site=same-origin
                             User-Agent: Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.100 Safari/537.36
                             X-Requested-With: XMLHttpRequest";
            //api地址
            string url = "https://quan.qq.com/node/api2/getHomePageFeed/";
            HttpRequestHelper s = new HttpRequestHelper(true);
            //表单数据
            //string content = string.Format("userid={0}&accountType=2&start={1}", UserID, (Page*10).ToString());
            string content = "userId=" + UserID + "&accountType=" + accountType + "&start=" + Page * 10 + "&lastTime=" + lasttime;
            string response = s.httpPost(url, heads, content, Encoding.UTF8);
            //StreamWriter sw = File.AppendText(@"D:\\test.txt"); //保存到指定路径
            //sw.Write(response);
            //sw.Flush();
            //sw.Close();
            //第一页会返回html而不是json，所以加个判断。
            return JObject.Parse(response);
        }
        //分析数据
        private void analysePage(JObject jq, PageFeed pg)
        {
            foreach (var jb in jq["data"]["vHomePageFeed"])
            {
                //判断是否为文章
                if (jb["eFeedType"].ToString() == "1")
                {
                    //文章ID作为唯一值
                    string postID = jb["stPostSummary"]["postId"].ToString();
                    //添加一行数据
                    pg.allPages.Add(postID, new OnePage(postID));
                    pg.pageIDList.Add(postID);
                    //pg.allPages[postID]
                    pg.allPages[postID].nickName = jb["stPostSummary"]["postUser"]["nickname"].ToString();
                    pg.allPages[postID].PageID = jb["stPostSummary"]["postId"].ToString();
                    pg.allPages[postID].PageName = jb["stPostSummary"]["postField"]["title"].ToString();
                    pg.allPages[postID].Date = jb["stPostSummary"]["elapseTime"].ToString();
                    pg.allPages[postID].praiseNum = jb["stPostSummary"]["praiseNum"].ToString();
                    pg.allPages[postID].commentNum = jb["stPostSummary"]["commentNum"].ToString();
                    pg.allPages[postID].viewNum = jb["stPostSummary"]["viewNum"].ToString();
                    pg.allPages[postID].quanName = jb["stPostSummary"]["simpleInfo"]["circleName"].ToString();
                    pg.allPages[postID].quanID = jb["stPostSummary"]["simpleInfo"]["circleId"].ToString();
                    pg.allPages[postID].quanUrl = "https://quan.qq.com/circle/" + pg.allPages[postID].quanID;
                    pg.allPages[postID].pageUrl = "https://quan.qq.com/post/" + pg.allPages[postID].quanID + "/" + postID;
                }
            }
        }
        //开始爬取按钮
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            //删除data目录下所有文件
            if (!Directory.Exists(AppDomain.CurrentDomain.BaseDirectory + "data/"))
            {
                Directory.CreateDirectory(AppDomain.CurrentDomain.BaseDirectory + "data/");
            }
            else
            {
                try
                {
                    DirectoryInfo dir = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory + "data/");
                    FileSystemInfo[] fileinfo = dir.GetFileSystemInfos();  //返回目录中所有文件和子目录
                    foreach (FileSystemInfo i in fileinfo)
                    {
                        if (i is DirectoryInfo)            //判断是否文件夹
                        {
                            DirectoryInfo subdir = new DirectoryInfo(i.FullName);
                            subdir.Delete(true);          //删除子目录和文件
                        }
                        else
                        {
                            File.Delete(i.FullName);      //删除指定文件
                        }
                    }
                }
                catch (Exception a)
                {
                    MessageBox.Show(a.ToString());
                }
            }
            //xlsx or xls
            if (UserID.Text.Substring(UserID.Text.Length - 5, 5) == ".xlsx" | UserID.Text.Substring(UserID.Text.Length - 4, 4) == ".xls")
            {
                GetAlldataByExcelFileHandler handler = new GetAlldataByExcelFileHandler(GetAlldataByExcelFile);
                IAsyncResult result = handler.BeginInvoke(UserID.Text, new AsyncCallback(FileCallback), null);
            }
            //单个UserID
            else
            {
                GetAllDataHandler handler = new GetAllDataHandler(GetAllData);
                string _temp = "";


                if (UserID.Text.Substring(0, 3) == "101")
                {
                    _temp = UserID.Text.Substring(4, UserID.Text.Length - 4);
                    IAsyncResult result = handler.BeginInvoke(_temp, "101", new AsyncCallback(SingleIDCallback), null);
                }
                else
                {
                    _temp = UserID.Text.Substring(2, UserID.Text.Length - 2);
                    IAsyncResult result = handler.BeginInvoke(_temp, "2", new AsyncCallback(SingleIDCallback), null);
                }
            }


            UnEnabledButton();
            StartButton.Content = "爬取中……";
        }


        //打开配置文件
        private void FileOpenButton_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog()
            {
                Filter = "Excel文件|*.xls;*.xlsx"
            };
            var result = openFileDialog.ShowDialog();
            //
            if (result == true)
            {
                UserID.Text = openFileDialog.FileName;
            }
        }
        //合并所有data目录下的excel文件为一个，方便统计
        private void FileToOne_Click(object sender, RoutedEventArgs e)
        {
            UnEnabledButton();
            //保存所有数据
            PageFeed pgFeed = new PageFeed();
            string AppDomainPath = AppDomain.CurrentDomain.BaseDirectory;
            string DataPath = AppDomainPath + "data\\";
            DirectoryInfo root = new DirectoryInfo(DataPath);
            //遍历目录下所有文件
            foreach (FileInfo f in root.GetFiles())
            {
                if (f.Name.Split('_')[0] == "merge")
                {
                    f.Delete();
                }
                else
                {
                    try
                    {
                        IWorkbook workbook = null;  //新建IWorkbook对象  
                        FileStream fileStream = new FileStream(f.FullName, FileMode.Open, FileAccess.Read);
                        if (f.FullName.IndexOf(".xlsx") > 0) // 2007版本  
                        {
                            workbook = new XSSFWorkbook(fileStream);  //xlsx数据读入workbook  
                        }
                        else if (f.FullName.IndexOf(".xls") > 0) // 2003版本  
                        {
                            workbook = new HSSFWorkbook(fileStream);  //xls数据读入workbook  
                        }
                        ISheet sheet = workbook.GetSheetAt(0);  //获取第一个工作表  
                        IRow row;// = sheet.GetRow(0);            //新建当前工作表行数据  
                        for (int i = 1; i <= sheet.LastRowNum; i++)  //对工作表每一行  
                        {
                            row = sheet.GetRow(i);   //row读入第i行数据  
                            if (row != null)
                            {
                                //读取pageID
                                string temp = row.GetCell(0).ToString();
                                pgFeed.allPages.Add(temp, new OnePage(temp));
                                pgFeed.pageIDList.Add(temp);
                                for (int j = 1; j < 11; j++)
                                {
                                    pgFeed.allPages[temp].SetCell(j, row.GetCell(j).ToString());
                                }
                            }
                        }
                        fileStream.Close();
                        workbook.Close();
                    }
                    catch (Exception x)
                    {
                        MessageBox.Show(x.ToString());
                    }
                }
            }
            pgFeed.Save2Excel("merge_" + DateTime.Now.ToString("yyyyMMddHHmmssffff"));
            MessageBox.Show("合并完成");
            EnabledButton();
        }
        public void UnEnabledButton()
        {
            StartButton.IsEnabled = false;
            UserID.IsEnabled = false;
            FileOpenButton.IsEnabled = false;
            FileToOne.IsEnabled = false;
        }
        public void EnabledButton()
        {
            StartButton.IsEnabled = true;
            UserID.IsEnabled = true;
            FileOpenButton.IsEnabled = true;
            FileToOne.IsEnabled = true;
        }
        private void SingleIDCallback(IAsyncResult ar)
        {
            MessageBox.Show("爬取结束,文件处于运行程序目录下的data文件夹中");
            StartButton.Dispatcher.Invoke(new Action(delegate { StartButton.Content = "开始爬取"; }));
            StartButton.Dispatcher.Invoke(new Action(delegate { EnabledButton(); }));
        }
        //多用户爬取回调
        private void FileCallback(IAsyncResult ar)
        {
            MessageBox.Show("爬取结束,文件处于运行程序目录下的data文件夹中");
            StartButton.Dispatcher.Invoke(new Action(delegate { StartButton.Content = "开始爬取"; }));
            StartButton.Dispatcher.Invoke(new Action(delegate { EnabledButton(); }));
        }
    }
}
