using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.IO;
using Ionic.Zip;
using System.Net.Mail;
using System.Net;
using System.Text.RegularExpressions;
using System.Security.Cryptography;
using MySql.Data.MySqlClient;
using HtmlAgilityPack;


namespace linemedia_WS_1113
{
    public partial class Form1 : Form
    {
        string strQuery;
        string strLogFilePath;
        string strTimeStamp;
        string strLogPath;
        string strLogMessage;
        string strHTMLFilePath;
        string strOutputPath;
        string strOutputPath_csv;
        string strHTMLPath;
        string strHtmlFilePath;
        string strHTMLPath1;
        string FTP_fileSize = "";
        string Pc_FileSize = "";
        string flink1;
        string string2 = "";
        int cat = 0;
        int re = 0;
        mysql obj = new mysql();

        Thread t;
        string Id;
        string[] s = null;
        public Form1()
        {
            InitializeComponent();
        }
        private void createfolder()
        {
            try
            {
                System.Windows.Forms.Control.CheckForIllegalCrossThreadCalls = false;
                strTimeStamp = System.DateTime.Now.Month + "_" + System.DateTime.Now.Day + "_" + System.DateTime.Now.Year + "_" + System.DateTime.Now.Hour + "_" + System.DateTime.Now.Minute + "_" + System.DateTime.Now.Second;
                strLogPath = Application.StartupPath + "\\" + "Log";
                strOutputPath_csv = Application.StartupPath + "\\" + "Output_" + DateTime.Now.ToString("yyyy_MM_dd") + "\\" + "CSV";
                strOutputPath = Application.StartupPath + "\\" + "OutPut_" + DateTime.Now.ToString("yyyy_MM_dd");

                string strhtmlfolder = Application.StartupPath;
                int index = strhtmlfolder.LastIndexOf("\\");
                strhtmlfolder = strhtmlfolder.Remove(index);

                strHTMLPath1 = strhtmlfolder + "\\" + "Html1" + "\\" + DateTime.Now.ToString("dd_MM_yyyy");
                strHTMLPath = strhtmlfolder + "\\" + "Html" + "\\" + DateTime.Now.ToString("dd_MM_yyyy");

                if (Directory.Exists(strLogPath)) { }
                else
                {
                    Directory.CreateDirectory(Application.StartupPath + "\\" + "Log");
                }

                if (Directory.Exists(strOutputPath)) { }
                else
                {
                    Directory.CreateDirectory(Application.StartupPath + "\\" + "OutPut_" + DateTime.Now.ToString("yyyy_MM_dd"));
                }

                if (Directory.Exists(strHTMLPath1)) { }
                else
                {
                    Directory.CreateDirectory(strhtmlfolder + "\\" + "Html1" + "\\" + DateTime.Now.ToString("dd_MM_yyyy"));
                }
                if (Directory.Exists(strHTMLPath)) { }
                else
                {
                    Directory.CreateDirectory(strhtmlfolder + "\\" + "Html" + "\\" + DateTime.Now.ToString("dd_MM_yyyy"));
                }
                if (Directory.Exists(strOutputPath_csv)) { }
                else
                {
                    Directory.CreateDirectory(Application.StartupPath + "\\" + "Output_" + DateTime.Now.ToString("yyyy_MM_dd") + "\\" + "CSV");
                }
            }
            catch (Exception ex1)
            {
                MessageBox.Show(ex1.Message);
            }
        }
        public void createtable(string strquery)
        {
            MySqlConnection con = obj.mysqlconnection();
            int re = 0;
            strquery = strquery.Replace("\\", "\\\\");
            if (con.State == ConnectionState.Closed)
            {
                con = obj.mysqlconnection();
            }
            re = obj.executeDMLSQL(strquery);
            if (!string.IsNullOrEmpty(obj.error))
            {
                writelog(obj.error);
            }
            else
            {
            }
        }
        private string StripTag(string Html)
        {
            Html = WebUtility.HtmlDecode(Html);
            Html = Regex.Replace(Html, "<script[^>]*>([\\w\\W]*?)</script>", "");
            Html = Regex.Replace(Html, "<style[^>]*>([\\w\\W]*?)</style>", "");
            Html = Regex.Replace(Html, "<.*?>", "");
            Html = Regex.Replace(Html, Environment.NewLine, "");
            Html = Html.Replace(System.Environment.NewLine, "");
            Html = System.Text.RegularExpressions.Regex.Replace(Html, "\\s+", " ");
            Html = Html.Replace("\r", " ").Replace("\n", " ").Replace("\r\n", " ");
            Html = Html.Replace("\t", "");
            Html = Html.Replace("\n", "");
            Html = Html.Trim();
            return Html;
        }
        public static string GetMd5Hash(MD5 md5Hash, string input)
        {
            byte[] data = md5Hash.ComputeHash(Encoding.UTF8.GetBytes(input));
            StringBuilder sBuilder = new StringBuilder();
            int i = 0;
            for (i = 0; i <= data.Length - 1; i++)
            {
                sBuilder.Append(data[i].ToString("x2"));
            }
            return sBuilder.ToString();
        }
        public void writelog(string strlogmessage)
        {
            try
            {
                strLogFilePath = strLogPath + "\\Log_" + strTimeStamp + ".txt";
                StreamWriter sw = new StreamWriter(strLogFilePath, false);
                sw.Write(strLogMessage);
                sw.AutoFlush = true;
                sw.Close();
                sw.Dispose();
            }
            catch (Exception ex1)
            {
                MessageBox.Show(ex1.Message);
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            obj.closeconnection();
            MySqlConnection con = obj.mysqlconnection();
            string[] s;
            string str;
            string[] c = { "\\" };

            str = Application.StartupPath;
            s = str.Split(c, StringSplitOptions.None);
            createfolder();
            this.Text = "Autoline - Linemedia" + s[s.Length - 1];
            CheckForIllegalCrossThreadCalls = false;

            string constr1;
            StreamReader objReader = new StreamReader(Application.StartupPath + "\\connectionstring.txt");
            constr1 = objReader.ReadLine();
            obj.constr = constr1 + ";SslMode=none;";
            objReader.Close();

            string process = "";
            StreamReader reader = new StreamReader(Application.StartupPath + "\\" + "process.txt");
            process = reader.ReadToEnd();
            reader.Close();

            string input_path = Application.StartupPath + "\\" + "input.txt";
            StreamReader st = new StreamReader(Application.StartupPath + "\\" + "input.txt");
            string line = st.ReadToEnd();
            string[] str1;
            str1 = Regex.Split(line, "-");
            textBox1.Text = str1[0];
            textBox2.Text = str1[1];
            st.Close();
            DataSet ds = new DataSet();
            int re = 0;
            for (int x = 1; x <= 2; x++)
            {
                if (re == 0)
                {
                    if (process == "Generate_Table")
                    {
                        string checktableexistssqr = "SELECT COUNT(*) FROM information_schema.tables WHERE table_schema = 'linemedia' AND table_name = 'Final_productdata_" + DateTime.Now.ToString("yyyy_MM_dd") + " ' ";
                        checktableexistssqr = checktableexistssqr.Replace("\\", "\\\\");
                        if (con.State == ConnectionState.Closed)
                        {
                            con = obj.mysqlconnection();
                        }
                        ds = obj.executeSQL_dset(checktableexistssqr);

                        if (ds.Tables.Count == 1)
                        {
                            re = Convert.ToInt32(ds.Tables[0].Rows[0][0].ToString());
                        }

                        if (re == 0)
                        {
                            string strquary = "UPDATE `firstlevelcategory` SET status='pending'";
                            if (con.State == ConnectionState.Closed)
                            {
                                con = obj.mysqlconnection();
                            }
                            obj.executeDMLSQL(strquary);
                            string strquary1 = "CREATE TABLE `productdata_" + System.DateTime.Now.ToString("yyyy_MM_dd") + "` (" + "`Id` int(11) NOT NULL AUTO_INCREMENT,"
                                + "`URL` VARCHAR(250) DEFAULT NULL,"
                                + "`Category` VARCHAR(300) DEFAULT NULL,"
                                + "`Name` VARCHAR(300) DEFAULT NULL,"
                                + "`Main_HtmlPath` VARCHAR(300) DEFAULT NULL,"
                                + "`HtmlPath` VARCHAR(300) DEFAULT NULL,"
                                + "`Link` VARCHAR(250) DEFAULT NULL,"
                                + "`Status` VARCHAR(50) DEFAULT 'Pending',"
                                + "unique key(`Link`),"
                                + "PRIMARY KEY (`Id`));";
                            createtable(strquary1);
                            if (con.State == ConnectionState.Closed)
                            {
                                con = obj.mysqlconnection();
                            }
                            re = obj.executeDMLSQL(strquary1);
                            if (obj.srr != "")
                            {
                                writelog(obj.srr + " sql= " + strquary1);
                            }
                            
                            String strqueary1 = "CREATE TABLE final_productdata_" + System.DateTime.Now.ToString("yyyy_MM_dd")
                            + "(" + "  `ID` int(11) NOT NULL AUTO_INCREMENT,"
                            + " `clientIdentifier` varchar(250) DEFAULT NULL,"
                            + "`foreignId` varchar(255) DEFAULT NULL,"
                            + "`url` varchar(250) DEFAULT NULL,"
                            + "`language` varchar(250) DEFAULT NULL,"
                            + "`title` mediumtext,"
                            + "`type` mediumtext,"
                            + "`description` mediumtext,"
                            + "`manufacturer` varchar(250) DEFAULT NULL,"
                            + "`model` varchar(250) DEFAULT NULL,"
                            + "`location` varchar(250) DEFAULT NULL,"
                            + "`yearOfManufacture` varchar(250) DEFAULT NULL,"
                            + "`auctionId` varchar(250) DEFAULT NULL,"
                            + "`auctionUrl` varchar(250) DEFAULT NULL,"
                            + "`auctionTitle` varchar(250) DEFAULT NULL,"
                            + "`endDate` varchar(250) DEFAULT NULL,"
                            + "`currency` varchar(250) DEFAULT NULL,"
                            + "`price` varchar(250) DEFAULT NULL,"
                            + "`startBid` mediumtext,"
                            + "`minimumBid` mediumtext,"
                            + "`currentBid` mediumtext,"
                            + "`category` mediumtext,"
                            + "`attributes` mediumtext,"
                            + "`imageUrls` mediumtext,"
                            + "`JSON_status` varchar(10) DEFAULT 'Pending',"
                            + "`Htmlpath` varchar(500) DEFAULT 'Pending',"
                            + "`Seller` varchar(500) DEFAULT 'Pending',"
                            + "PRIMARY KEY(`ID`),"
                            + "UNIQUE KEY `NewIndex1` (`url`))";
                            createtable(strqueary1);
                            if (con.State == ConnectionState.Closed)
                            {
                                con = obj.mysqlconnection();
                            }
                            re = obj.executeDMLSQL(strqueary1);
                            if (obj.srr != "")
                            {
                                writelog(obj.srr + " sql= " + strqueary1);
                            }
                        }
                    }
                }
            }
            t = new Thread(generatexl1);
            t.IsBackground = true;
            t.Start();

        }
        private void ExtractURL()
        {
            obj.closeconnection();
            MySqlConnection con = obj.mysqlconnection();
            CheckForIllegalCrossThreadCalls = false;
            label1.Text = "ExtractURL...";
            string strhtml = "";
            int StartVal = Convert.ToInt32(textBox1.Text);
            int EndVal = Convert.ToInt32(textBox2.Text);
            string strhtmlpath2 = "";
            string category1 = "";
            for (int x = 0; x <= 100; x++)
            {
                if (con.State == ConnectionState.Closed)
                {
                    con = obj.mysqlconnection();
                }
                DataSet ds = obj.executeSQL_dset("Select * from `firstlevelcategory` where ID between " + StartVal + " and " + EndVal + " and Status='Pending'");
                int count = ds.Tables[0].Rows.Count - 1;
                if (count < 0)
                {
                    goto down2;
                }
                for (int i = 0; i <= count; i++)
                {
                    string Str1 = "";
                    string link = "";
                    //string Prefix = ds.Tables[0].Rows[i]["Name"].ToString();
                    int page = 1;

                    string Str2 = "";
                    string[] str = { "" };
                    Str2 = ds.Tables[0].Rows[i]["URL"].ToString();
                    String mainstr = Str2;
                    String ID = ds.Tables[0].Rows[i]["Id"].ToString();
                    string[] s1 = { "" };
                    string[] s2 = { "" };

                    //string name = "";

                    strhtml = "";
                    //link1 = "";
                    strhtml = "";
                    int cat = 0;
                    int counter = 0;
                    String phone = "";
                    String address = "";
                    String attributes = "";
                    String ms = "";
                    Boolean foundnextpage = true;
                    while (foundnextpage == true)
                    {

                        try
                        {
                            String name = "";
                            HtmlAgilityPack.HtmlWeb web = new HtmlAgilityPack.HtmlWeb();
                            HtmlAgilityPack.HtmlDocument document = web.Load(Str2);
                            String path = strHTMLPath + "//" + ID +"_"+page+ ".html";
                            document.Save(path);
                            String html = document.DocumentNode.SelectSingleNode("/html").InnerHtml;
                            if(html.Contains("<span>next"))
                            {
                                foundnextpage = true;
                                page++;
                            }
                            else
                            {
                                foundnextpage = false;
                            }
                            try
                            {
                                if (html.Contains("<meta name=\"description\""))
                                {
                                    String[] data1 = Regex.Split(html, "<meta name=\"description\"");
                                    data1 = Regex.Split(data1[1], "\">");
                                    String cnt = StripTag(data1[0]);
                                    cnt = new String(cnt.Where(Char.IsDigit).ToArray());
                                    obj.executeDMLSQL("Update `firstlevelcategory` set count = '" + cnt + "' WHERE Id = " + ID);
                                }
                            }
                            catch (Exception ex)
                            {

                            }

                            HtmlNodeCollection nodes = document.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[2]/div[1]/div/div[1]/div[2]/div[2]/div[4]/div");
                            if (nodes != null)
                            {
                                for (int k = 1; k <= nodes.Count; k++)
                                {
                                    try
                                    {
                                        name = document.DocumentNode.SelectSingleNode("//*[@id=\"content\"]/div[2]/div[1]/div/div[1]/div[2]/div[2]/div[4]/div[" + k + "]/div/div[2]/div[1]/a").InnerText;
                                        name = StripTag(name);
                                        try
                                        {
                                            link = document.DocumentNode.SelectSingleNode("//*[@id=\"content\"]/div[2]/div[1]/div/div[1]/div[2]/div[2]/div[4]/div[" + k + "]/div/div[2]/div[1]/a").GetAttributeValue("href", "");
                                        }
                                        catch (Exception ex)
                                        {
                                            link = "";
                                        }
                                        label1.Text = path;
                                        label2.Text = link.ToString();
                                        category1 = "";
                                        try
                                        {
                                            if (html.Contains("<div class=\"cat\">"))
                                            {
                                                String[] cate = Regex.Split(html, "<div class=\"cat\">");
                                                cate = Regex.Split(cate[1], "</");
                                                category1 = StripTag(cate[0]);
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            category1 = "";
                                        }
                                        string strQuery = "INSERT INTO `productdata_" + System.DateTime.Now.ToString("yyyy_MM_dd") + "`(URL,Name,Category,Link,Htmlpath) VALUES ('" + Str2.Replace("'", "''") + "','" + name.Replace("'", "''") + "','" + category1 + "','" + link + "','"+path+"')";



                                        strQuery = strQuery.Replace("\\", "\\\\");
                                        re = 0;
                                        re = obj.executeDMLSQL(strQuery);
                                        if (obj.srr != "")
                                        {

                                            writelog(obj.srr + " sql= " + strQuery);
                                        }
                                    }
                                    catch (Exception)
                                    { }
                                }
                            }
                            else
                            {
                                name = "";
                                category1 = "";
                                link = "";
                                if (html.Contains("sl-item-extension sl-item-extension-photo"))
                                {
                                    String[] data = Regex.Split(html, "sl-item-extension sl-item-extension-photo");
                                    for(int mk=1;mk<data.Length;mk++)
                                    {
                                        if(data[mk].Contains("<a href="))
                                        {
                                            String[] data1 = Regex.Split(data[mk], "<a href=\"");
                                            data1 = Regex.Split(data1[1], "\"");
                                            link = StripTag(data1[0]);

                                            data1 = Regex.Split(data1[3],"</a");
                                            name = data1[0].Replace(">","");
                                            name = StripTag(name);
                                            
                                        }
                                        try
                                        {
                                            if (html.Contains("<div class=\"cat\">"))
                                            {
                                                String[] cate = Regex.Split(html, "<div class=\"cat\">");
                                                cate = Regex.Split(cate[1], "</");
                                                category1 = StripTag(cate[0]);
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            category1 = "";
                                        }
                                        label1.Text = "ELSE "+path;
                                        label2.Text = link.ToString();
                                        string strQuery = "INSERT INTO `productdata_" + System.DateTime.Now.ToString("yyyy_MM_dd") + "`(URL,Name,Category,Link,Htmlpath) VALUES ('" + Str2.Replace("'", "''") + "','" + name.Replace("'", "''") + "','" + category1 + "','" + link + "','" + path + "')";



                                        strQuery = strQuery.Replace("\\", "\\\\");
                                        re = 0;
                                        re = obj.executeDMLSQL(strQuery);
                                        if (obj.srr != "")
                                        {

                                            writelog(obj.srr + " sql= " + strQuery);
                                        }
                                    }
                                }
                            }



                            Str2 = "";
                            Str2 = mainstr+"?page="+page;
                        }

                        catch (Exception ex)
                        {
                            cat++;
                            //label5.Text = cat.ToString();
                            writelog(ex.Message + " " + Str1);
                        }
                        

                    }
                    int id = Convert.ToInt32(ds.Tables[0].Rows[i].ItemArray[0]);
                    if (con.State == ConnectionState.Closed)
                    {
                        con = obj.mysqlconnection();
                    }
                    string sQuery1 = "Update `firstlevelcategory" + "` set " +
                         "Status='Done'" + " WHERE Id =" + id;
                    obj.executeDMLSQL(sQuery1);



                }



                //t = new System.Threading.Thread(ExtractURL1);
                //t.IsBackground = true;
                //t.Start();

            }
            down2:
            label3.Text = "Link Exctration Completed";
            writelog("Completed " + Id + Environment.NewLine);

        }

        private void Extractpage()
        {
            obj.closeconnection();
            MySqlConnection con = obj.mysqlconnection();
            MD5 md5Hash = MD5.Create();
            label1.Text = "Extract Page...";
            string strHtml = "";

            int StartVal = Convert.ToInt32(textBox1.Text);
            int EndVal = Convert.ToInt32(textBox2.Text);

            string import_id = "";
            string sGUID;
            sGUID = System.Guid.NewGuid().ToString();
            import_id = sGUID;
            String erid = "";
            for (int x = 0; x <= 100; x++)
            {
                //DataSet ds = obj.executeSQL_dset("Select * from `productdata_" + System.DateTime.Now.ToString("yyyy_MM_dd") + "` where ID between " + StartVal + " and " + EndVal + " and Status='Pending'");
                DataSet ds = obj.executeSQL_dset("Select * from `productdata_" + DateTime.Now.ToString("yyyy_MM_dd") + "` where ID between " + StartVal + " and " + EndVal + " and Status='Pending'");
                //DataSet ds = obj.executeSQL_dset("Select * from `productdata_2018_10_01` where ID between " + StartVal + " and " + EndVal + " and Status='Pending'");
                if (ds.Tables[0].Rows.Count == 0)
                {
                    goto up;
                }

                int m = ds.Tables[0].Rows.Count;

                for (int i = 0; i < m; i++)
                {
                    label1.Text = ds.Tables[0].Rows[i][0].ToString();
                    string ID = ds.Tables[0].Rows[i].ItemArray[0].ToString();

                    //string htmlpath = ds.Tables[0].Rows[i]["Htmlpath"].ToString();
                    string str5 = ds.Tables[0].Rows[i]["Link"].ToString();

                    string Prefix = ds.Tables[0].Rows[i]["Name"].ToString();
                    string Category = ds.Tables[0].Rows[i]["Category"].ToString();

                    string title = "";
                    //string Main_HtmlPath = ds.Tables[0].Rows[i]["Main_HtmlPath"].ToString();
                    string description = "";
                    string manufacturer = "";
                    string model = "";
                    string Images = "";
                    string currency = "";
                    string attributes = "";
                    string Yearofmanufacture = "";

                    string language = "en";
                    string Type1 = "trader";

                    string price = "";

                    string location = "";
                    try
                    {
                        re = 0;
                        HtmlAgilityPack.HtmlWeb web = new HtmlAgilityPack.HtmlWeb();
                        HtmlAgilityPack.HtmlDocument document = web.Load(str5);



                        //label3.Text = Prefix.ToString();
                        document.Save(strHTMLPath + "\\" + ID + ".html");
                        String htmlpage = strHTMLPath + "\\" + ID + ".html";
                        String html = document.DocumentNode.SelectSingleNode("/html").InnerHtml;
                        String seller = "";
                        if (html.Contains("<span class=\"only-print-inline\">www.autoline.info/</span>"))
                        {
                            String[] data1 = Regex.Split(html, "<span class=\"only-print-inline\">www.autoline.info/</span>");
                            data1 = Regex.Split(data1[1], "</");
                            seller = StripTag(data1[0]);
                        }
                        if (seller == "wantiantrading" || seller == "al-quds" || seller == "gassmann" || seller == "e-farm" || seller== "gamalquiler")
                        {
                            string strQuery2 = "Update `productdata_" + DateTime.Now.ToString("yyyy_MM_dd") + "` set Status='" + seller+"' WHERE Id =" + ds.Tables[0].Rows[i][0];
                            strQuery2 = strQuery2.Replace("\\", "\\\\");
                            if (con.State == ConnectionState.Closed)
                            {
                                obj.mysqlconnection();
                            }
                            int re = 0;
                            re = obj.executeDMLSQL(strQuery2);
                            if (re == 0)
                            {
                                writelog(obj.srr + " sql= " + strQuery2);
                            }
                        }
                        else
                        {
                            try
                            {
                                Prefix = document.DocumentNode.SelectSingleNode("//*[@id=\"content\"]/div[1]/div[1]/h1").InnerText;
                                Prefix = StripTag(Prefix);
                            }
                            catch (Exception)
                            {
                                String[] data1 = Regex.Split(html, "class=\"sf-title\">");
                                data1 = Regex.Split(data1[1], "</h1>");
                                Prefix = StripTag(data1[0]);
                            }
                            label2.Text = Prefix + "  " + str5;
                            string foreignId = "";
                            try
                            {
                                String[] data1 = Regex.Split(str5, "--");
                                foreignId = StripTag(data1[1]);
                                foreignId = foreignId.Replace("-", "");

                            }
                            catch (Exception)
                            {

                            }
                            currency = "";
                            if (html.Contains("<div class=\"value \">"))
                            {
                                String[] data1 = Regex.Split(html, "<div class=\"value \">");
                                data1 = Regex.Split(data1[1], "</");
                                price = StripTag(data1[0]);
                                price = price.Replace("€", "");
                                price = price.Replace(",", "");
                                price = price.Replace(".", ".!");
                                if (price.Contains(".!"))
                                {
                                    data1 = Regex.Split(price, ".!");
                                    price = StripTag(data1[0]);
                                }


                                currency = "EUR";
                                if (price.Contains("$"))
                                {
                                    price = price.Replace("$", "");
                                }
                                if (price == "")
                                {
                                    currency = "";
                                }

                            }
                            if (html.Contains("<span class=\"field\">Brand"))
                            {
                                String[] data1 = Regex.Split(html, "<span class=\"field\">Brand");
                                if (data1[1].Contains("<span class=\"value \">"))
                                {
                                    data1 = Regex.Split(data1[1], "<span class=\"value \">");
                                    data1 = Regex.Split(data1[1], "</");
                                    manufacturer = StripTag(data1[0]);
                                }
                                else
                                {
                                    manufacturer = "";
                                }

                            }

                            if (html.Contains("<span class=\"field\">Model"))
                            {
                                String[] data1 = Regex.Split(html, "<span class=\"field\">Model");
                                if (data1[1].Contains("<span class=\"value \">"))
                                {
                                    data1 = Regex.Split(data1[1], "<span class=\"value \">");
                                    data1 = Regex.Split(data1[1], "</");
                                    model = StripTag(data1[0]);
                                }
                                else
                                {
                                    model = "";
                                }

                            }
                            if (html.Contains("<span class=\"field\">Type"))
                            {
                                String[] data1 = Regex.Split(html, "<span class=\"field\">Type");
                                if (data1[1].Contains("<span class=\"value \">"))
                                {
                                    data1 = Regex.Split(data1[1], "<span class=\"value \">");
                                    data1 = Regex.Split(data1[1], "</");
                                    Category = StripTag(data1[0]);
                                }
                                else
                                {
                                    Category = "";
                                }

                            }
                            location = "";
                            if (html.Contains("class=\"loc-country\""))
                            {
                                String[] data1 = Regex.Split(html, "class=\"loc-country\"");
                                if (data1[1].Contains("<span class=\"loc-city\">"))
                                {
                                    data1 = Regex.Split(data1[1], "<span class=\"loc-city\">");

                                    String[] cntry = Regex.Split(data1[0], ">");
                                    cntry = Regex.Split(cntry[1], "</");
                                    location = StripTag(cntry[0]);
                                    String[] dat2 = Regex.Split(data1[1], "</");
                                    location = location + "," + StripTag(dat2[0]);
                                }


                                else
                                {


                                    try
                                    {

                                        data1 = Regex.Split(data1[1], ">");
                                        data1 = Regex.Split(data1[1], "</");
                                        location = StripTag(data1[0]);
                                        if (location.Contains(">") || location.Contains("<") || location.Contains("="))
                                        {
                                            location = "";
                                        }
                                    }
                                    catch (Exception ex)
                                    {

                                    }
                                }


                            }
                            else if (html.Contains("class=\"loc-city\""))
                            {
                                String[] data = Regex.Split(html, "class=\"loc-city\"");
                                data = Regex.Split(data[1], ">");
                                data = Regex.Split(data[1], "<");
                                location = StripTag(data[0]);
                            }
                            if (html.Contains("<span class=\"field\">Year of manufacture"))
                            {
                                String[] data1 = Regex.Split(html, "<span class=\"field\">Year of manufacture");
                                if (data1[1].Contains("<span class=\"value \">"))
                                {
                                    data1 = Regex.Split(data1[1], "<span class=\"value \">");
                                    data1 = Regex.Split(data1[1], "</");
                                    Yearofmanufacture = StripTag(data1[0]);
                                    if (Yearofmanufacture.Contains("/"))
                                    {
                                        String[] yr = Regex.Split(Yearofmanufacture, "/");
                                        Yearofmanufacture = StripTag(yr[1]);
                                    }
                                    bool isIntString = Yearofmanufacture.All(char.IsDigit);
                                    if (isIntString == false)
                                    {
                                        Yearofmanufacture = "";
                                    }
                                    if (Yearofmanufacture.Length != 4)
                                    {
                                        Yearofmanufacture = "";
                                    }
                                }
                                else
                                {
                                    Yearofmanufacture = "";
                                }

                            }
                            try
                            {
                                if (html.Contains("class=\"thumb thumb"))
                                {
                                    String[] data1 = Regex.Split(html, "class=\"thumb thumb");
                                    for (int img = 1; img < data1.Length; img++)
                                    {
                                        String[] data2 = Regex.Split(data1[img], "data-image-big=\"");
                                        data2 = Regex.Split(data2[1], "\"");
                                        if (Images == "")
                                        {
                                            Images = StripTag(data2[0]);
                                        }
                                        else
                                        {
                                            Images = Images + "|" + StripTag(data2[0]);
                                        }
                                        if (img == 3)
                                        {
                                            break;
                                        }
                                    }

                                }
                            }catch(Exception ex)
                            {

                            }
                            if (Images == "")
                            {
                                try
                                {
                                    String[] img = Regex.Split(html, "data-image-big=\"");
                                    img = Regex.Split(img[1], "\"");
                                    Images = StripTag(img[0]);
                                }
                                catch (Exception)
                                {
                                    Images = "";
                                }
                            }
                            if (html.Contains("<span class=\"field\">"))
                            {
                                String[] data1 = Regex.Split(html, "<span class=\"field\">");
                                for (int at = 1; at < data1.Length; at++)
                                {
                                    String[] data2 = { };
                                    data2 = Regex.Split(data1[at], "</span");
                                    if (data2[0].Contains("Brand") || data2[0].Contains("Model") || data2[0].Contains("Type") || data2[0].Contains("Year of manufacture") || data2[0].Contains("Location"))
                                    {
                                        continue;
                                    }
                                    else
                                    {
                                        if (data1[at].Contains("<span class=\"value \">"))
                                        {
                                            String[] data3 = Regex.Split(data1[at], "<span class=\"value \">");
                                            data3 = Regex.Split(data3[1], "</span>");
                                            if (attributes == "")
                                            {
                                                attributes = StripTag(data2[0] + ":" + data3[0]);
                                            }
                                            else
                                            {
                                                attributes = attributes + "|" + StripTag(data2[0] + ":" + data3[0]);
                                            }
                                        }
                                        else
                                        {
                                            if (data1[at].Contains("class=\"value tick\""))
                                            {
                                                if (attributes == "")
                                                {
                                                    attributes = StripTag(data2[0] + ":Yes");
                                                }
                                                else
                                                {
                                                    attributes = attributes + "|" + StripTag(data2[0] + ":Yes");
                                                }
                                            }
                                        }

                                    }

                                }

                            }

                            string cid = "linemedia";
                            re = 0;
                            strQuery = "INSERT INTO `final_productdata_" + DateTime.Now.ToString("yyyy_MM_dd") + "`(clientIdentifier,htmlpath,foreignId,url,language,title,type,description,manufacturer,model,location,yearOfManufacture,currency,price,category,attributes,Seller,imageUrls) VALUES('" +
                                cid.Replace("'", "''") + "','" +
                                htmlpage.Replace("'", "''") + "','" +
                                       foreignId.Replace("'", "''") + "','" +
                                       str5.Replace("'", "''") + "','" +
                                       language.Replace("'", "''") + "','" +
                                       Prefix.Replace("'", "''") + "','" +
                                       Type1.Replace("'", "''") + "','" +
                                       description.Replace("'", "''") + "','" +
                                       manufacturer.Replace("'", "''") + "','" +
                                       model.Replace("'", "''") + "','" +
                                       location.Replace("'", "''") + "','" +
                                       Yearofmanufacture.Replace("'", "''") + "','" +
                                       currency.Replace("'", "''") + "','" +
                                       price.Replace("'", "''") + "','" +
                                       Category.Replace("'", "''") + "','" +
                                       attributes.Replace("'", "''") + "','" +
                                       seller.Replace("'", "''") + "','" +
                                       Images.Replace("'", "''") + "')";
                            obj.srr = "";
                            strQuery = strQuery.Replace("\\", "\\\\");
                            re = 0;
                            re = obj.executeDMLSQL(strQuery);

                                

                        }
                    }
                    catch (Exception ex)
                    {

                        //label5.Text = (erid = erid + ID + " ").ToString();
                        writelog(ex.Message);
                    }
                
                    if (re == 1)
                    {

                        string strQuery2 = "Update `productdata_" + DateTime.Now.ToString("yyyy_MM_dd") + "` set Status='Done'" + " WHERE Id =" + ds.Tables[0].Rows[i][0];
                        strQuery2 = strQuery2.Replace("\\", "\\\\");
                        if (con.State == ConnectionState.Closed)
                        {
                            obj.mysqlconnection();
                        }
                        int re = 0;
                        re = obj.executeDMLSQL(strQuery2);
                        if (re == 0)
                        {
                            writelog(obj.srr + " sql= " + strQuery2);
                        }

                    }





                }
            }
            up:
            label3.Text = "Extract Page Completed";


            //t = new System.Threading.Thread(generatexl1);
            //t.IsBackground = true;
            //t.Start();


        }
        public void generatexl1()
        {
            label1.Text = "Generate Excel...";
            upper:
            try
            {
                //Boolean foundpending = true;
                string strtimestamp1 = "";
                strtimestamp1 = DateTime.Now.ToString("yyyy_MM_dd");
                //clientIdentifier,foreignId,url,language,title,type,description,manufacturer,model,location,yearOfManufacture,category,attributes,imageUrls,Status1
                //DataSet dset = obj.executeSQL_dset("SELECT * FROM final_productdata_" + DateTime.Now.ToString("yyyy_MM_dd"));
                //DataSet dset = obj.executeSQL_dset("SELECT clientIdentifier,foreignId,url,language,title,type,description,manufacturer,model,location,yearOfManufacture,auctionId,auctionUrl,auctionTitle,endDate,currency,price,startBid,minimumBid,currentBid,category,attributes,imageUrls FROM final_productdata_2018_07_23");
                DataSet dset = obj.executeSQL_dset("SELECT clientIdentifier,foreignId,url,language,title,type,description,manufacturer,model,location,yearOfManufacture,auctionId,auctionUrl,auctionTitle,endDate,currency,price,startBid,minimumBid,currentBid,category,attributes,imageUrls,Seller FROM final_productdata_2018_12_12 where foreignId <>'' and Seller != 'wantiantrading' AND Seller != 'al-quds' AND Seller != 'gassmann' AND Seller != 'gamalquiler' AND Seller != 'e-farm' AND Seller != 'putzmeister'");
                //DataSet dset = obj.executeSQL_dset("SELECT clientIdentifier,foreignId,url,language,title,type,description,manufacturer,model,location,yearOfManufacture,auctionId,auctionUrl,auctionTitle,endDate,currency,price,startBid,minimumBid,currentBid,category,attributes,imageUrls FROM final_productdata where foreignId <>''");
                // DataSet dset = obj.executeSQL_dset("SELECT clientIdentifier,foreignId,url,language,title,type,description,manufacturer,model,location,yearOfManufacture,auctionId,auctionUrl,auctionTitle,endDate,currency,price,startBid,minimumBid,currentBid,category,attributes,imageUrls FROM final_productdata" + strtimestamp1 + " where foreignId <>''");
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook wBook = excel.Application.Workbooks.Add(true);
                Microsoft.Office.Interop.Excel.Worksheet wSheet = new Microsoft.Office.Interop.Excel.Worksheet();
                //wBook = excel.Workbooks.Add();
                int cnt = dset.Tables[0].Rows.Count;
                System.Data.DataTable dt1 = dset.Tables[0];
                int colIndex = 0;
                int rowIndex = 0;
                foreach (System.Data.DataColumn dc in dt1.Columns)
                {
                    colIndex = colIndex + 1;
                    excel.Cells[1, colIndex] = dc.ColumnName;
                }
                foreach (System.Data.DataRow dr in dt1.Rows)
                {
                    rowIndex = rowIndex + 1;
                    colIndex = 0;
                    foreach (System.Data.DataColumn dc in dt1.Columns)
                    {
                        colIndex = colIndex + 1;
                        excel.Cells[rowIndex + 1, colIndex] = dr[dc.ColumnName];
                        label2.Text = "Row Index = " + rowIndex + " Column Index = " + colIndex.ToString();
                    }
                }
                string strtimestamp3 = "";

                strtimestamp3 = DateTime.Now.Year + "_" + DateTime.Now.Month + "_" + DateTime.Now.Day;
                string strFileName = strOutputPath + "\\" + "linemedia" +strtimestamp3+ ".xls";
                Boolean blnFileOpen = false;
                try
                {
                    System.IO.FileStream fileTemp = System.IO.File.OpenWrite(strFileName);
                    fileTemp.Close();
                }
                catch (Exception ex)
                {
                    blnFileOpen = false;
                }
                if (System.IO.File.Exists(strFileName))
                {
                    System.IO.File.Delete(strFileName);
                }
                wBook.SaveAs(strFileName);
                wBook.Close();
                wBook = null;
                excel.Quit();
                excel = null;


                label1.Text = "Generate Excel Done.";

                ExportCSV1(dset.Tables[0]);
            }
            catch (Exception ex)
            {
                Thread.Sleep(3000);
                goto upper;
            }
            label2.Text = "CSV Extract.";

            //t = new Thread(UploadFileToFTP);
            //t.IsBackground = true;
            //t.Start();

        }

        public string FixCSV(string sValue)
        {
            return "\"" + sValue.Replace("\"", "\"\"") + "\"";
        }

        public void ExportCSV1(DataTable dtSource)
        {
            string strtimestamp3 = "";
            createfolder();
            strtimestamp3 = DateTime.Now.ToString("yyyy-MM-dd");
            string strFileName = strOutputPath_csv + "\\" + "linemedia_1" + ".csv";
            StreamWriter sw = new StreamWriter(strFileName, false, System.Text.Encoding.UTF8);

            try
            {
                int nCount = dtSource.Columns.Count;
                StringBuilder sbLine = new StringBuilder();
                for (int iCol = 0; iCol < dtSource.Columns.Count; iCol++)
                {
                    sbLine.Append(",");
                    string strheader = dtSource.Columns[iCol].ColumnName;
                    sbLine.Append(strheader);
                }

                sw.WriteLine(sbLine.ToString().Substring(1));
                int id = 1;
                foreach (DataRow dr in dtSource.Rows)
                {
                    sbLine = new StringBuilder();
                    for (int iCol = 0; iCol < dtSource.Columns.Count; iCol++)
                    {
                        sbLine.Append(",");
                        if (!Convert.IsDBNull(dr[iCol]))
                        {
                            if (dr[iCol].GetType().ToString() == "System.String")
                            {
                                sbLine.Append(FixCSV(dr[iCol].ToString()));
                            }
                            else
                            {
                                sbLine.Append(dr[iCol]);
                            }
                        }
                    }
                    sw.WriteLine(sbLine.ToString().Substring(1));
                }
                //thread = new Thread(EmailSend);
                //thread.IsBackground = true;
                //thread.Start();
            }
            finally
            {
                sw.Close();
            }
        }
        private void update()
        {
            obj.closeconnection();
            MySqlConnection con = obj.mysqlconnection();
            MD5 md5Hash = MD5.Create();
            label1.Text = "Extract Page...";
            string strHtml = "";

            int StartVal = Convert.ToInt32(textBox1.Text);
            int EndVal = Convert.ToInt32(textBox2.Text);

            string import_id = "";
            string sGUID;
            sGUID = System.Guid.NewGuid().ToString();
            import_id = sGUID;
            String erid = "";
            for (int x = 0; x <= 100; x++)
            {
              DataSet ds = obj.executeSQL_dset("SELECT * FROM final_productdata WHERE location = ''");
                //DataSet ds = obj.executeSQL_dset("Select * from `productdata_" + DateTime.Now.ToString("yyyy_MM_dd") + "` where ID between " + StartVal + " and " + EndVal + " and Status='Pending'");
                //DataSet ds = obj.executeSQL_dset("Select * from `final_productdata_2018_07_21` where ID between " + StartVal + " and " + EndVal + " and Status1='Pending'");
                if (ds.Tables[0].Rows.Count == 0)
                {
                    goto up;
                }

                int m = ds.Tables[0].Rows.Count;

                for (int i = 0; i < m; i++)
                {
                    label1.Text = ds.Tables[0].Rows[i][0].ToString();
                    string ID = ds.Tables[0].Rows[i].ItemArray[0].ToString();
                    String st2 = ds.Tables[0].Rows[i]["URL"].ToString();
                    string html1 = ds.Tables[0].Rows[i]["Htmlpath"].ToString();
                    

                    
                    
                    
                    try
                    {
                        re = 0;
                        HtmlAgilityPack.HtmlWeb web = new HtmlAgilityPack.HtmlWeb();
                        HtmlAgilityPack.HtmlDocument document = web.Load(html1);
                        String html = document.DocumentNode.SelectSingleNode("/html").InnerHtml;


                        ////label3.Text = Prefix.ToString();
                        ////document.Save(strHTMLPath + "\\" + ID + ".html");
                        //String htmlpage = strHTMLPath + "\\" + ID + ".html";


                        //if (html.Contains("<div class=\"value \">"))
                        //{
                        //    String[] data1 = Regex.Split(html, "<div class=\"value \">");
                        //    data1 = Regex.Split(data1[1], "</");
                        //    price = StripTag(data1[0]);
                        //    price = price.Replace("€", "");
                        //    price = price.Replace(",", "");

                        //    price = price.Replace(".", ".!");
                        //    if (price.Contains(".!"))
                        //    {
                        //        data1 = Regex.Split(price, ".!");
                        //        price = StripTag(data1[0]);
                        //    }


                        //    currency = "EUR";
                        //    if (price.Contains("$"))
                        //    {
                        //        price = price.Replace("$", "");
                        //    }
                        //    if (price == "")
                        //    {
                        //        currency = "";
                        //    }

                        //}
                        String location = "";
                        if (html.Contains("class=\"loc-country\""))
                        {
                            String[] data1 = Regex.Split(html, "class=\"loc-country\"");
                            if (data1[1].Contains("<span class=\"loc-city\">"))
                            {
                                data1 = Regex.Split(data1[1], "<span class=\"loc-city\">");

                                String[] cntry = Regex.Split(data1[0], ">");
                                cntry = Regex.Split(cntry[1], "</");
                                location = StripTag(cntry[0]);
                                String[] dat2 = Regex.Split(data1[1], "</");
                                location = location + "," + StripTag(dat2[0]);
                            }
                            

                            else
                            {
                                

                                try
                                {

                                    data1 = Regex.Split(data1[1], ">");
                                    data1 = Regex.Split(data1[1], "</");
                                    location = StripTag(data1[0]);
                                    if(location.Contains(">")|| location.Contains("<")|| location.Contains("="))
                                    {
                                        location = "";
                                    }
                                }
                                catch (Exception ex)
                                {

                                }
                            }


                        }
                        else if(html.Contains("class=\"loc-city\""))
                        {
                            String[] data = Regex.Split(html, "class=\"loc-city\"");
                            data = Regex.Split(data[1], ">");
                            data = Regex.Split(data[1], "<");
                            location = StripTag(data[0]);
                        }

                        re = 0;
                        strQuery = "update final_productdata set location = '" + location + "' where id ='"+ID+"'";
                        strQuery = strQuery.Replace("\\", "\\\\");
                        re = 0;
                        re = obj.executeDMLSQL(strQuery);



                    }
                    catch (Exception ex)
                    {

                        //label5.Text = (erid = erid + ID + " ").ToString();
                        writelog(ex.Message);
                    }

                    





                }
            }
            up:
            label3.Text = "Update Completed";


            

        }

        private void button1_Click(object sender, EventArgs e)
        {
            t = new Thread(ExtractURL);
            t.IsBackground = true;
            t.Start();
        }
    }
}
