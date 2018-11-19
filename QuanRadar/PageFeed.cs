using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using QuanRadar.Helper;
using System.Windows;
namespace QuanRadar
{
    //一条数据
    public class OnePage
    {
        public string index = "";
        //用户昵称 ["stPostSummary"]["postUser"]["nickname"]
        public string nickName = "";
        //文章ID stPostSummary postId
        public string PageID;

        //文章名称 stPostSummary postField title
        public string PageName;

        //文章发布时间 stPostSummary elapseTime
        public string Date = "";

        //文章获得的赞 stPostSummary praiseNum
        public string praiseNum = "";

        //文章阅读量 stPostSummary viewNum
        public string viewNum = "";

        //文章获得的评论 stPostSummary commentNum
        public string commentNum = "";

        //圈子名称 stPostSummary simpleInfo circleName
        public string quanName = "";

        //圈子ID   stPostSummary simpleInfo circleId
        public string quanID = "";

        //文章发布所在圈子网址
        public string quanUrl = "";

        //文章网址
        public string pageUrl = "";

        public OnePage(string _PageID)
        {
            PageID = _PageID;
        }
        public void SetCell(int i,string j)
        {
            switch (i)
            {
                case 2:
                    //返回文章ID
                    PageID = j;
                    break;
                case 0:
                    //返回用户ID
                    nickName = j;
                    break;
                case 1:
                    //文章名称
                    PageName = j;
                    break;
                case 3:
                    //发布时间
                    Date = j;
                    break;
                case 4:
                    //获得的赞
                    praiseNum = j;
                    break;
                case 5:
                    //获得的评论
                    commentNum = j;
                    break;
                case 6:
                    //阅读量
                    viewNum = j;
                    break;
                case 7:
                    //发布圈子ID
                    quanID = j;
                    break;
                case 8:
                    //发布圈子名称
                    quanName = j;
                    break;
                case 9:
                    //发布圈子地址
                    quanUrl = j;
                    break;
                case 10:
                    //发布文章地址
                    pageUrl = j;
                    break;
            }
        }
        public string GetCell(int i)
        {
            switch (i)
            {
                case 2:
                    //返回文章ID
                    return PageID;
                case 0:
                    //返回用户ID
                    return nickName;
                case 1:
                    //文章名称
                    return PageName;          
                case 3:
                    //发布时间
                    return Date;
                case 4:
                    //获得的赞
                    return praiseNum;
                case 5:
                    //获得的评论
                    return commentNum;
                case 6:
                    //阅读量
                    return viewNum;
                case 7:
                    //发布圈子ID
                    return quanID;
                case 8:
                    //发布圈子名称
                    return quanName;
                case 9:
                    //发布圈子地址
                    return quanUrl;
                case 10:
                    //发布文章地址
                    return pageUrl;                    
            }
            return "";
        }
    }
    //
    public class PageFeed
    {
        public PageFeed()
        {
            allPages.Clear();
            pageIDList.Clear();
        }
        //保存所有文章数据
        public Dictionary<string,OnePage> allPages = new Dictionary<string,OnePage>();
        public List<string> pageIDList = new List<string>();
        //使用ExcelHelper将临时存放处的数据保存到本地excel文件中
        public void Save2Excel(string userName)
        {
            if (!string.IsNullOrWhiteSpace(userName))
            {
                IWorkbook workbook = new XSSFWorkbook();
                ISheet sheet = workbook.CreateSheet("数据表");
                //创建行
                for (int i = 0; i < allPages.Count + 1; i++)
                {
                    IRow row = sheet.CreateRow(i); //i表示了创建行的索引，从0开始
                    if (i == 0)
                    {
                        for (int j = 0; j < 11; j++)
                        {
                            ICell cell = row.CreateCell(j, CellType.String);  //同时这个函数还有第二个重载，可以指定单元格存放数据的类型
                            cell.SetCellValue(GetCell(j));
                        }
                    }
                    else
                    {
                        string pageID = pageIDList[i - 1];
                        //创建单元格
                        for (int j = 0; j < 11; j++)
                        {
                            ICell cell = row.CreateCell(j, CellType.String);  //同时这个函数还有第二个重载，可以指定单元格存放数据的类型
                            cell.SetCellValue(allPages[pageID].GetCell(j));
                        }
                    }
                }
                //保存
                //创建一个文件流对象
                try
                {
                    if (!Directory.Exists(AppDomain.CurrentDomain.BaseDirectory + "data/"))
                    {
                        Directory.CreateDirectory(AppDomain.CurrentDomain.BaseDirectory + "data/");
                    }
                    
                    using (FileStream fs = File.Create(AppDomain.CurrentDomain.BaseDirectory + "data/" + MakeValidFileName(ref userName) + ".xlsx"))
                    {
                        workbook.Write(fs);
                        //最后记得关闭对象
                        workbook.Close();
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message.ToString());
                }
            }
            
            
        }
        //去掉不能为文件名的字符，避免报错
        public static string MakeValidFileName(ref string text, string replacement = "_")
        {
            StringBuilder str = new StringBuilder();
            var invalidFileNameChars = System.IO.Path.GetInvalidFileNameChars();
            foreach (var c in text)
            {
                if (invalidFileNameChars.Contains(c))
                {
                    str.Append(replacement ?? "");
                }
                else
                {
                    str.Append(c);
                }
            }

            return str.ToString();
        }
        //顶栏
        public string GetCell(int i)
        {
            switch (i)
            {
                case 2:
                    //返回文章ID
                    return "文章ID";
                case 0:
                    //返回用户ID
                    return "用户ID";
                case 1:
                    //文章名称
                    return "文章名称";
                case 3:
                    //发布时间
                    return "发布时间";
                case 4:
                    //获得的赞
                    return "获得的赞";
                case 5:
                    //获得的评论
                    return "获得的评论";
                case 6:
                    //阅读量
                    return "阅读量";
                case 7:
                    //发布圈子ID
                    return "发布圈子ID";
                case 8:
                    //发布圈子名称
                    return "发布圈子名称";
                case 9:
                    //发布圈子地址
                    return "发布圈子地址";
                case 10:
                    //发布文章地址
                    return "发布文章地址";
            }
            return "";
        }
    }
}
