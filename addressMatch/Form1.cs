using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Security.Permissions;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using System.Timers;
using System.Threading;
using System.Collections; // 导入命名空间
using System.IO;
using System.Reflection;
using System.Net;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Configuration;

namespace convertPointToPoint
{
    [PermissionSet(SecurityAction.Demand, Name = "FullTrust")]
    [System.Runtime.InteropServices.ComVisibleAttribute(true)]

    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        DataSet finished = new DataSet();//已完成名单
        public List<string> namelist = new List<string>();
        List<resdata> xydata = new List<resdata>();
        public int count = 0;
        public int startcount = 0;
        int searchNumber = Convert.ToInt16(ConfigurationManager.AppSettings["searchNumber"].ToString());
        DataSet orids = new DataSet();

        //坐标转换常量
        //定义一些常量
        double x_PI = 3.14159265358979324 * 3000.0 / 180.0;
        double PI = 3.1415926535897932384626;
        double a = 6378245.0;
        double ee = 0.00669342162296594323;
        double[] json;
        private void selectExcel_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Microsoft Excel files(*.xls)|*.xls;*.xlsx";//过滤一下，只要表格格式的
            ofd.InitialDirectory = "d:\\";
            ofd.RestoreDirectory = true;
            ofd.FilterIndex = 1;
            ofd.AddExtension = true;
            ofd.CheckFileExists = true;
            ofd.CheckPathExists = true;
            ofd.ShowHelp = true;//是否显示帮助按钮
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                string DBString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" + ofd.FileName + ";Extended Properties=Excel 12.0";
                excelLocation.Text = ofd.FileName;
                OleDbConnection con = new OleDbConnection(DBString);
                con.Open();
                System.Data.DataTable datatable = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                //获取表单，原始的是：Sheet1，Sheet2，Sheet3
                for (int i = 0; i < 1; i++)
                {
                    //获取表单的名字
                    String sheet = datatable.Rows[i][2].ToString().Trim();
                    OleDbDataAdapter ole = new OleDbDataAdapter("select * from [" + sheet + "]", con);

                    ole.Fill(orids);
                    //输出表格里面的内容，我这里就两列数据，如果数据列数不确定就需要写循环了：Rows.Count
                    for (int k = 20; k < 150; k++)
                    {

                        namelist.Add(orids.Tables[0].Rows[k][2].ToString());
                    }

                    //foreach (DataRow col in ds.Tables[0].Rows)
                    //{
                    //   // Console.WriteLine(col[0].ToString());
                    //  //  Console.WriteLine(col[1].ToString());if


                    //}
                }
                con.Close();
            }
        }


        private string HttpPost(string Url, string postDataStr)
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(Url);
            request.Method = "POST";
            request.ContentType = "application/x-www-form-urlencoded";
            request.ContentLength = Encoding.UTF8.GetByteCount(postDataStr);
            // request.CookieContainer = cookie;
            Stream myRequestStream = request.GetRequestStream();
            StreamWriter myStreamWriter = new StreamWriter(myRequestStream, Encoding.GetEncoding("gb2312"));
            myStreamWriter.Write(postDataStr);
            myStreamWriter.Close();

            HttpWebResponse response = (HttpWebResponse)request.GetResponse();

            // response.Cookies = cookie.GetCookies(response.ResponseUri);
            Stream myResponseStream = response.GetResponseStream();
            StreamReader myStreamReader = new StreamReader(myResponseStream, Encoding.GetEncoding("utf-8"));
            string retString = myStreamReader.ReadToEnd();
            myStreamReader.Close();
            myResponseStream.Close();
            //string jsonText = "[{'a':'aaa','b':'bbb','c':'ccc'},{'a':'aaa2','b':'bbb2','c':'ccc2'}]";
            //JArray ja = (JArray)JsonConvert.DeserializeObject(jsonText);
            //  resdata rsd= convertToObject(retString);
            return retString;
        }
        int searchindex = 0;

        private resdata convertToObject(string rstring)
        {
            JObject o = (JObject)JsonConvert.DeserializeObject(rstring);
            JToken record;
            resdata rd = new resdata();
            rd.x = "";
            rd.y = "";
            rd.addr = "";
            //   rd.name = namelist[count];
            if (o.Property("status").ToString() != "302")
            {

                try
                {
                    string result = o.Property("results").ToString();

                    string results = o["results"].ToString();
                    double baidu_x = 0.0;
                    double baidu_y = 0.0;
                    if (results != "[]")
                    {
                        record = o["results"][0];

                        if (record != null)
                        {
                            foreach (JProperty jp in record)
                            {
                                string a = (String)jp.Name;
                                switch (a)
                                {
                                    case "name":
                                        string b = (String)jp.Value;
                                        rd.name = b;
                                        break;
                                    case "location":
                                        JToken xy = jp.Value;
                                        foreach (JProperty jxy in xy)
                                        {
                                            switch ((String)jxy.Name)
                                            {
                                                case "lng":
                                                    rd.x = (String)jxy.Value;
                                                    baidu_x = (double)jxy.Value;
                                                    break;
                                                case "lat":
                                                    rd.y = (String)jxy.Value;
                                                    baidu_y = (double)jxy.Value; ;
                                                    break;
                                                default:
                                                    break;
                                            }
                                        }
                                        break;
                                    case "address":
                                        rd.addr = (String)jp.Value;
                                        break;
                                    default:
                                        break;
                                }

                            }

                        }

                    }
                    else
                    {

                        rd.x = "";
                        rd.y = "";
                        rd.addr = "";
                    }
                    if (baidu_x != 0.0)
                    {
                        double[] gcj_xy = bd09togcj02(baidu_x, baidu_y);
                        // console.log(gcj_xy[0], xygcj_xy1]);
                        double[] wgs_xy = gcj02towgs84(gcj_xy[0], gcj_xy[1]);
                        rd.x = wgs_xy[0].ToString();
                        rd.y = wgs_xy[1].ToString();
                    }



                    //   Console.WriteLine(rd.x + "," + rd.y);
                    xydata.Add(rd);
                }
                catch (Exception Err)              //Exception可以针对不同的异常改为不同的异常类,这是异常的基类；
                {
                    //  MessageBox.Show(Err.Message);
                    MessageBox.Show(o.Property("message").ToString());
                    System.Environment.Exit(0);
                    //  MessageBox.show(Err.Message);  
                }
            }
            else
            {
                rd.x = "";
                rd.y = "";
                rd.addr = "";
                MessageBox.Show(o.Property("message").ToString());
                System.Environment.Exit(0);
            }
            return rd;
        }


        public void addxytooracle()
        {

        }
   
        public bool hasSearchFinish(string _id)
        {
            if (Convert.ToInt32(_id) != count)
            {
                count++;
                return true;
            }
            else
            {
                System.Threading.Thread.Sleep(2000);
                hasSearchFinish(_id);
            }
            return false;
        }
    

       

     
       
      

        private void load_Click(object sender, EventArgs e)
        {
            string pathALL = ConfigurationManager.AppSettings["loadPath"].ToString();
            // string pathALL = "../../data/temp.xls";
            finished = ExcelToDS(pathALL);
            for (int k = 0; k < finished.Tables[0].Rows.Count; k++)
            {
                if (finished.Tables[0].Rows[k]["state"].ToString() == "")
                {
                    startcount = k;
                    break;

                }

            }
            //  Console.WriteLine(startcount);
            namelist.Clear();
            for (int m = startcount; m < (startcount + searchNumber); m++)
            {
                namelist.Add(finished.Tables[0].Rows[m][2].ToString());
            }
        //    button1.Enabled = true;

        }
        public DataSet ExcelToDS(string Path)
        {
            DataSet ds = null;
            //try { 
            string strConn = "";
            // strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + Path + ";" + "Extended Properties=Excel 8.0";
            strConn = "Provider=Microsoft.Ace.OleDb.12.0;" + "data source=" + Path + ";Extended Properties='Excel 12.0; HDR=Yes; IMEX=1'"; //此连接可以操作.xls与.xlsx文件 (支持Excel2003 和 Excel2007 的连接字符串)


            OleDbConnection conn = new OleDbConnection(strConn);
            conn.Open();
            string strExcel = "";
            OleDbDataAdapter myCommand = null;

            strExcel = "select * from [sheet1$]";
            myCommand = new OleDbDataAdapter(strExcel, strConn);
            ds = new DataSet();
            myCommand.Fill(ds, "table1");
            //}
            //catch
            //{
            //    MessageBox.Show("文件中表不存在！");
            //}
            return ds;
        }
        /*坐标转换*/
        public double[] bd09togcj02(double bd_lon, double bd_lat)
        {
            double x_pi = 3.14159265358979324 * 3000.0 / 180.0;
            double x = bd_lon - 0.0065;
            double y = bd_lat - 0.006;
            double z = Math.Sqrt(x * x + y * y) - 0.00002 * Math.Sin(y * x_pi);
            double theta = Math.Atan2(y, x) - 0.000003 * Math.Cos(x * x_pi);
            double gg_lng = z * Math.Cos(theta);
            double gg_lat = z * Math.Sin(theta);
            double[] h = new double[2];
            h[0] = gg_lng;
            h[1] = gg_lat;
            return h;
        }
        /**
 * GCJ02 转换为 WGS84
 * @param lng
 * @param lat
 * @returns {*[]}
 */
        public double[] gcj02towgs84(double lng, double lat)
        {
            double[] r = new double[2];
            r[0] = lng;
            r[1] = lat;

            if (out_of_china(lng, lat))
            {
                return r;
            }
            else
            {
                double dlat = transformlat(lng - 105.0, lat - 35.0);
                double dlng = transformlng(lng - 105.0, lat - 35.0);
                double radlat = lat / 180.0 * PI;
                double magic = Math.Sin(radlat);
                magic = 1 - ee * magic * magic;
                double sqrtmagic = Math.Sqrt(magic);
                dlat = (dlat * 180.0) / ((a * (1 - ee)) / (magic * sqrtmagic) * PI);
                dlng = (dlng * 180.0) / (a / sqrtmagic * Math.Cos(radlat) * PI);
                double mglat = lat + dlat;
                double mglng = lng + dlng;
                r[0] = lng * 2 - mglng;
                r[1] = lat * 2 - mglat;
                return r;
            }
        }
        public double transformlat(double lng, double lat)
        {
            double ret = -100.0 + 2.0 * lng + 3.0 * lat + 0.2 * lat * lat + 0.1 * lng * lat + 0.2 * Math.Sqrt(Math.Abs(lng));
            ret += (20.0 * Math.Sin(6.0 * lng * PI) + 20.0 * Math.Sin(2.0 * lng * PI)) * 2.0 / 3.0;
            ret += (20.0 * Math.Sin(lat * PI) + 40.0 * Math.Sin(lat / 3.0 * PI)) * 2.0 / 3.0;
            ret += (160.0 * Math.Sin(lat / 12.0 * PI) + 320 * Math.Sin(lat * PI / 30.0)) * 2.0 / 3.0;
            return ret;
        }

        public double transformlng(double lng, double lat)
        {
            var ret = 300.0 + lng + 2.0 * lat + 0.1 * lng * lng + 0.1 * lng * lat + 0.1 * Math.Sqrt(Math.Abs(lng));
            ret += (20.0 * Math.Sin(6.0 * lng * PI) + 20.0 * Math.Sin(2.0 * lng * PI)) * 2.0 / 3.0;
            ret += (20.0 * Math.Sin(lng * PI) + 40.0 * Math.Sin(lng / 3.0 * PI)) * 2.0 / 3.0;
            ret += (150.0 * Math.Sin(lng / 12.0 * PI) + 300.0 * Math.Sin(lng / 30.0 * PI)) * 2.0 / 3.0;
            return ret;
        }
        /**
 * 判断是否在国内，不在国内则不做偏移
 * @param lng
 * @param lat
 * @returns {boolean}
 */
        public bool out_of_china(double lng, double lat)
        {
            return (lng < 72.004 || lng > 137.8347) || ((lat < 0.8293 || lat > 55.8271) || false);
        }
        private static string strConn = ConfigurationSettings.AppSettings["orclCon"];
        private void conn_Click(object sender, EventArgs e)
        {
            OleDbConnection conn = new OleDbConnection(strConn);
            conn.Open();
            string tablename = "zzjgtmp";
            string strComm = "select * from " + tablename;
            OleDbCommand comm = new OleDbCommand(strComm, conn);
            OleDbDataReader sdr = comm.ExecuteReader();
            while (sdr.Read())
            {
                if (sdr["bmatch"].ToString() == "")
                {
                    string para = "";
                    string dwname = sdr["dwmc"].ToString();
                    string key = ConfigurationManager.AppSettings["baiduKey"].ToString();
                    string url = "http://api.map.baidu.com/place/v2/search?q=" + dwname + "&region=杭州&output=json&ak=" + key;
                    resdata updata = convertToObject(HttpPost(url, para));
                    var showstr = dwname + "," + updata.name + "," + updata.addr + "," + updata.x + "," + updata.y;
                    Console.WriteLine(showstr);
                    //richTextBox1.Text = showstr;
                    // convert.Document.Write(showstr);
                    //convert.DocumentText = "<div>"+showstr+"</div>";
                    if (updata.x != "" || updata.x != null)
                    {
                        //continue;
                        if (updata.addr.Length > 50 || updata.addr == null)
                        {
                            updata.addr = "";
                        }
                        //
                        string updatasql = " update " + tablename + " j set x='" + updata.x + "', y='" + updata.y + "',bmatch='f',address='" + updata.addr + "' where dwmc='" + dwname + "'";
                        OleDbCommand update = new OleDbCommand(updatasql, conn);
                        update.CommandTimeout = 20;
                        //  continue;
                        update.ExecuteNonQuery();
                    }
                    else
                    {
                        string updatasql = " update " + tablename + " j set bmatch='f' where dwmc='" + dwname + "'";
                        OleDbCommand update = new OleDbCommand(updatasql, conn);
                        update.CommandTimeout = 20;
                        update.ExecuteNonQuery();
                    }
                }
                else
                {
                    continue;
                }


            }


            if (!sdr.HasRows)
            {
                MessageBox.Show("查询无结果");
            }


            sdr.Close();

            conn.Close();
            conn.Dispose();
        }

        private void connAccess_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "(*.mdb)|*.mdb";//过滤一下，只要access格式的
            ofd.InitialDirectory = "d:\\";
            ofd.RestoreDirectory = true;
            ofd.FilterIndex = 1;
            ofd.AddExtension = true;
            ofd.CheckFileExists = true;
            ofd.CheckPathExists = true;
            ofd.ShowHelp = true;//是否显示帮助按钮
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                string pPath = ofd.FileName;
                OleDbConnection conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\"" + pPath + "\""); //Jet OLEDB:Database Password=
                OleDbCommand cmd = conn.CreateCommand();
                string tablename = "address";
                cmd.CommandText = "select * from " + tablename;
                conn.Open();

                OleDbDataReader sdr = cmd.ExecuteReader();
                runstate.Text = "程序运行中...";
                while (sdr.Read())
                {
                    if (sdr["bmatch"].ToString() == "")
                    {
                        string para = "";
                        string dwname = sdr["dwmc"].ToString();
                        string key = ConfigurationManager.AppSettings["baiduKey"].ToString();
                        string url = "http://api.map.baidu.com/place/v2/search?q=" + dwname + "&region=杭州&output=json&ak=" + key;
                        resdata updata = convertToObject(HttpPost(url, para));
                        var showstr = dwname + "," + updata.name + "," + updata.addr + "," + updata.x + "," + updata.y;
                        Console.WriteLine(showstr);
                        //richTextBox1.Text = showstr;
                        //  convert.DocumentText = showstr;
                        if (updata.x != "" || updata.x != null)
                        {
                            //continue;
                            if (updata.addr.Length > 50 || updata.addr == null)
                            {
                                updata.addr = "";
                            }
                            //
                            string updatasql = " update " + tablename + " j set x='" + updata.x + "', y='" + updata.y + "',bmatch='f',address='" + updata.addr + "' where dwmc='" + dwname + "'";
                            OleDbCommand update = new OleDbCommand(updatasql, conn);
                            update.CommandTimeout = 20;
                            //  continue;
                            update.ExecuteNonQuery();
                        }
                        else
                        {
                            string updatasql = " update " + tablename + " j set bmatch='f' where dwmc='" + dwname + "'";
                            OleDbCommand update = new OleDbCommand(updatasql, conn);
                            update.CommandTimeout = 20;
                            update.ExecuteNonQuery();
                        }
                    }
                    else
                    {
                        continue;
                    }


                }


                if (!sdr.HasRows)
                {
                    MessageBox.Show("查询无结果");
                }

                runstate.Text = "匹配结束";
                sdr.Close();

                conn.Close();
                conn.Dispose();
                //System.Data.DataTable dt = new System.Data.DataTable();
                //if (dr.HasRows)
                //{
                //    for (int i = 0; i < dr.FieldCount; i++)
                //    {
                //        dt.Columns.Add(dr.GetName(i));
                //    }
                //    dt.Rows.Clear();
                //}
                //while (dr.Read())
                //{
                //    DataRow row = dt.NewRow();
                //    for (int i = 0; i < dr.FieldCount; i++)
                //    {
                //        row[i] = dr[i];
                //    }
                //    dt.Rows.Add(row);
                //}
                //cmd.Dispose();
                //conn.Close();
                //dataGridView1.DataSource = dt;
            }
        }

        private void localsearch_Click(object sender, EventArgs e)
        {
          //  try
        //    {
                OleDbConnection conn = new OleDbConnection(strConn);
                conn.Open();
                string tablename = "zzjgtmp";
                string strComm = "select * from " + tablename + " where lmatch is  null and x is null";
                OleDbCommand comm = new OleDbCommand(strComm, conn);
                OleDbDataReader sdr = comm.ExecuteReader();
                while (sdr.Read())
                {
                    //if (sdr["lmatch"].ToString() == "")
                    //{
                    string para = "";
                    string dwname = sdr["dwmc"].ToString();
                  //  string key = ConfigurationManager.AppSettings["baiduKey"].ToString();
                    string url = "http://126.10.32.23/Geocoding/LiquidGIS/ContainsAddress.gis?key=" + dwname + "&FORMAT=json&SRID=4326";
                    resdata updata = strToObj(HttpPost(url, para));
                    var showstr = dwname + "," + updata.name + "," + updata.addr + "," + updata.x + "," + updata.y;
                    Console.WriteLine(showstr);
                    //richTextBox1.Text = showstr;
                    // convert.Document.Write(showstr);
                    //convert.DocumentText = "<div>"+showstr+"</div>";
                    if (updata.x != "" || updata.x != null)
                    {

                        string updatasql = " update " + tablename + " j set x='" + updata.x + "', y='" + updata.y + "',lmatch='f',address='" + updata.addr + "' where dwmc='" + dwname + "'";
                        OleDbCommand update = new OleDbCommand(updatasql, conn);
                        update.CommandTimeout = 20;
                        //  continue;
                        update.ExecuteNonQuery();
                    }
                    else
                    {
                        string updatasql = " update " + tablename + " j set lmatch='f' where dwmc='" + dwname + "'";
                        OleDbCommand update = new OleDbCommand(updatasql, conn);
                        update.CommandTimeout = 20;
                        update.ExecuteNonQuery();
                    }
                    //}
                    //else
                    //{
                    //    continue;
                    //}


                }


                if (!sdr.HasRows)
                {
                    MessageBox.Show("查询无结果");
                }


                sdr.Close();

                conn.Close();
                conn.Dispose();
       //     }
       //     catch (Exception ex)
        //    {
        //        throw new Exception(ex.Message, ex);
        //    }
        }
        private resdata strToObj(string rstring)
        {
            resdata rd = new resdata();
            rd.x = "";
            rd.y = "";
            rd.addr = "";
            if (rstring != "{}")
            {
                JObject o = (JObject)JsonConvert.DeserializeObject(rstring);
                
                    //  JObject o = (JObject)JsonConvert.DeserializeObject(rstring);
                    //  JToken record;
                    rd.name = o["Name"].ToString();
                    string x = o["Envelope"]["MaxX"].ToString();
                    string y = o["Envelope"]["MaxY"].ToString();
                    string posturl = "http://126.10.32.23/Transform/LiquidGIS/HangZhou.gis?x=" + x + "&y=" + y + "&COMMAND=TOWGS84&userkey=a67db68dbfb2752f9b913dff9ece867117c87e95&FORMAT=json";
                    string xy=HttpPost(posturl, "");
                    JObject jo = (JObject)JsonConvert.DeserializeObject(xy);
                rd.x = jo["X"].ToString();
                rd.y = jo["Y"].ToString();
            }
            return rd;

        }

    }
}
