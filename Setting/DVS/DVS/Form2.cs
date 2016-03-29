using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using System.Net;
using System.Text.RegularExpressions;

namespace DVS
{
    public partial class Form2 : Form
    {

        string xNum = null;
        string connetionString = null, sql = null;
        string reHostName = null, reLocalName = null, xStr = null;
        //string DS = "192.168.0.253";
        string DS = "211.75.132.163"; 
        string localDS = System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString();
        string reHostNameNum = System.Net.Dns.GetHostEntry("LocalHost").HostName.Substring(4, 2);
        string reLocalNameNum = System.Net.Dns.GetHostEntry("LocalHost").HostName.Substring(7, 2);
        bool xUpdateKey = false;
        public Form2()
        {
            InitializeComponent();
            xMain();
        }

        void xMain()
        {
            showDesk();
            timer1.Interval = 1000;
            timer1.Start();
            showLocalName(reLocalNameNum);
            showComboBox();
            showAutoTimer();
        }

        void showDesk()
        {
            deskPannel.Location = new System.Drawing.Point(0, 22);
            deskPannel.Size = new System.Drawing.Size(472, 397);
        }

        void showLocalName(string xKey)
        {
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + DS + ";Initial Catalog=STMS;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "select * from UserBranchID ";
            sql += "where number = '" + xKey + "' ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            if (rd.Read())
            {
                reLocalName = rd["Name"].ToString();
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }

        void showAutoTimer()
        {
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + DS + ";Initial Catalog=STMS;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "SELECT * FROM ClassIn_AutoTimer ";
            sql += "WHERE xHost = '" + reLocalNameNum + "' ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            if (rd.Read())
            {
                label4.Text = rd["xTimer"].ToString();
                comboBox1.Text = label4.Text.Substring(0, 2);
                comboBox2.Text = label4.Text.Substring(3, 2);
                comboBox3.Text = label4.Text.Substring(6, 2);
                this.Text = "三元及第視訊補課伺服器 V1.2 " + reLocalName;
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }

        void showComboBox()
        {
            for (int i = 0; i <= 24; i++)
            {
                switch(i.ToString().Length.ToString())
                {
                    case "1": comboBox1.Items.Add("0" + (i.ToString())); break;
                    case "2": comboBox1.Items.Add(i.ToString()); break;
                }
                
            }
            for (int i = 0; i <= 60; i++)
            {
                switch (i.ToString().Length.ToString())
                {
                    case "1": comboBox2.Items.Add("0" + (i.ToString())); break;
                    case "2": comboBox2.Items.Add(i.ToString()); break;
                }
            }
            for (int i = 0; i <= 60; i++)
            {
                switch (i.ToString().Length.ToString())
                {
                    case "1": comboBox3.Items.Add("0" + (i.ToString())); break;
                    case "2": comboBox3.Items.Add(i.ToString()); break;
                }
            }             
        }

        void autoEditData()
        {
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + DS + ";Initial Catalog=STMS;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "update ClassIn_AutoTimer SET ";
            sql += "xTimer = '" + comboBox1.Text.ToString() + ":" + comboBox2.Text.ToString() + ":" + comboBox3.Text.ToString() + "' ";
            sql += "where xHost = '" + reLocalNameNum + "' ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            rd.Close();
            cmd.Dispose();
            conn.Close();
            showAutoTimer();
        }

        void autoSelectUpdateData()
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "select * from Local_Update ";
            sql += "where xDate = '" + System.DateTime.Now.ToString("yyyy/MM/dd") + "' ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            if (rd.Read())
            {
                xUpdateKey = true;
            }
            else
            {
                xUpdateKey = false;
                autoInsertUpdateData();
                autoSelectUpdateData();
            }
            rd.Close(); cmd.Dispose(); conn.Close();
        }

        void autoInsertUpdateData()
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            //connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "Insert into Local_Update(xDate) VALUES (";
            sql += "'" + System.DateTime.Now.ToString("yyyy/MM/dd") + "') ";    //xCourseDate
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read(); rd.Close(); cmd.Dispose(); conn.Close();
        }

        void autoUpdateUpdateData(string xKey)
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "Update Local_Update SET  ";
            sql += xKey + "= '1' ";
            sql += "where xDate = '" + System.DateTime.Now.ToString("yyyy/MM/dd") + "' ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Close(); cmd.Dispose(); conn.Close();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            label2.Text = System.DateTime.Now.ToString("HH:mm:ss");
            if (label2.Text.ToString() == label4.Text.ToString())
            {
                autoSelectUpdateData();
                if (xUpdateKey)
                {
                    xActionLogistics();
                    xActionAuto();
                    xActionVideo();
                }
            }
            else
            {
                //this.Text = "Not OK";
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            manualSetPannel.Visible = false;
            autoSetPannel.Visible = true;
            autoSetPannel.Location = new Point(136, 66);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            manualSetPannel.Visible = true;
            autoSetPannel.Visible = false;
            selpanel1.Visible = true;
            selpanel2.Visible = true;
            selpanel4.Visible = true;
            showCBa();
            showCBb();
            showCBc();


            //manualSetLabelinfo01.Text = System.DateTime.Now.AddDays(-15).ToString("yyyy/MM/dd") + " ~ " + System.DateTime.Now.ToString("yyyy/MM/dd");
        }

        private void autoCancerBt_Click(object sender, EventArgs e)
        {
            autoSetPannel.Visible = false;
        }

        private void manualClosoe_Click(object sender, EventArgs e)
        {
            manualSetPannel.Visible = false;
        }

        private void autoBt_Click(object sender, EventArgs e)
        {
            Button bt = (Button)sender;
            switch(bt.Name)
            {
                case "autoCancerBt":
                    autoSetPannel.Visible = false;
                    break;
                case "autoSaveBt":
                    autoEditData();
                    break;
            }
        }

        // ~~~~~~~~~~ Logistic Action ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        // ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        void xActionLogistics()
        {
            xActiondvd_logistics();
            xActiondvd_logistics2();
        }

        //dvd_logistics 500 set -----------------------------
        void xActiondvd_logistics()
        {
            xSetNumdvd_logistics();
            xDeldvd_logistics();
            xSelectdvd_logistics();
            autoUpdateUpdateData("xActiondvd_logistics");
        }
        void xSelectdvd_logistics()
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + DS + ";Initial Catalog=STMS;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "SELECT serial,DVD編號, 上課日期,當科堂次 FROM DVD_logistics ";
            sql += "WHERE SUBSTRING(DVD編號, 3, 2) = '" + reHostNameNum + "' ";
            sql += "and serial >= '" + xNum + "' ";
            sql += "ORDER BY  serial DESC ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                xInsertdvd_logistics(Convert.ToInt32(rd["serial"].ToString()), rd["DVD編號"].ToString(), rd["當科堂次"].ToString(), Convert.ToDateTime(rd["上課日期"]));
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void xInsertdvd_logistics(int eKey, string xKey, string xKey2, DateTime dKey)
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            //connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "Insert into Local_dvd_logistics(serial, xDvNum, xCourseNum, xCourseDate) VALUES (";
            sql += "'" + eKey + "', ";    //serial
            sql += "'" + xKey + "', ";    //xDvNum
            sql += "'" + xKey2 + "', ";    //xCourseDate
            sql += "'" + dKey.ToShortDateString() + "') ";    //xCourseDate
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void xSetNumdvd_logistics()
        {
            xNum = "";
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            //connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "select top 500 * from Local_dvd_logistics order by serial desc";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                xNum = rd["serial"].ToString();
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void xDeldvd_logistics()
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            //connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "delete from Local_dvd_logistics WHERE serial >= '" + xNum + "' ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }

        //dvd_logistics2 500 set ---------------------------
        void xActiondvd_logistics2()
        {
            xSetNumdvd_logistics2();
            xDeldvd_logistics2();
            xSelectdvd_logistics2();
            autoUpdateUpdateData("xActiondvd_logistics2");
        }
        void xSelectdvd_logistics2()
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            //connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
            connetionString = "Data Source=" + DS + ";Initial Catalog=STMS;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "select serial, DVD編號, 上課講義名稱, 發放講義名稱, 發放講義編號, 章節進度, 頁數進度, 補充講義頁數 from DVD_logistics2 ";
            sql += "where SUBSTRING(DVD編號, 3, 2) = '" + reHostNameNum + "' ";
            sql += "and serial >= '" + xNum + "' ";
            sql += "ORDER BY  serial DESC ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                xInsertdvd_logistics2(Convert.ToInt32(rd["serial"].ToString()), rd["DVD編號"].ToString(), rd["上課講義名稱"].ToString(), rd["發放講義名稱"].ToString(), rd["發放講義編號"].ToString(), rd["章節進度"].ToString(), rd["頁數進度"].ToString(), rd["補充講義頁數"].ToString());
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void xInsertdvd_logistics2(int eKey, string xKey, string xKey2, string xKey3, string xKey4, string xKey5, string xKey6, string xKey7)
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            //connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "Insert into Local_dvd_logistics2(serial, xDvNum, xAcLectureName, xGrantLectureName, xGrantLectureNum, xChapterRate, xPagesRate, xReplenishPagesRate) VALUES (";
            sql += "'" + eKey + "', ";    //serial
            sql += "'" + xKey + "', ";    //serial
            sql += "'" + xKey2 + "', ";    //xCourseDatesql += "'" + eKey + "', ";    //serial
            sql += "'" + xKey3 + "', ";    //xDvNum
            sql += "'" + xKey4 + "', ";    //xCourseDate
            sql += "'" + xKey5 + "', ";    //xCourseDate
            sql += "'" + xKey6 + "', ";    //xCourseDate
            sql += "'" + xKey7 + "') ";    //xCourseDate
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void xSetNumdvd_logistics2()
        {
            xNum = "";
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            //connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "select top 500 * from Local_dvd_logistics2 order by serial desc";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                xNum = rd["serial"].ToString();
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void xDeldvd_logistics2()
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            //connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "delete from Local_dvd_logistics2 WHERE serial >= '" + xNum + "' ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }

        // ~~~~~~~~~~ action Video ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        // ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        void xActionVideo()
        {
            xActionDVD_video2();
            xActionDVD_video3();
            xActionSeat();
            xActionLocal_SwapSeat("auto");
        }

        //Local_dvd_video2 15day -----------------------------
        void xActionDVD_video2()
        {
            selectDVD_video2();
            autoUpdateUpdateData("xActionDVD_video2");
        }
        void selectDVD_video2()
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + DS + ";Initial Catalog=DVD;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "select serial, id, 地點, 日期, 時段, 已看, 座位, xInTime, xOutTime from DVD_video2 ";
            sql += "where 地點 = '" + reLocalNameNum + "' ";
            sql += "AND 日期 = '" + System.DateTime.Now.ToShortDateString() + "' ";
            sql += "order by 座位 ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                insertDVD_video2(Convert.ToInt32(rd["serial"].ToString()), rd["id"].ToString(), rd["地點"].ToString(), rd["日期"].ToString(), rd["時段"].ToString(), rd["已看"].ToString(), rd["座位"].ToString(), rd["xInTime"].ToString(), rd["xOutTime"].ToString());

            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void insertDVD_video2(int eKey, string xKey, string xKey2, string xKey3, string xKey4, string xKey5, string xKey6, string xKey7, string xKey8)
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            //connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "INSERT INTO Local_dvd_video2 (serial, xid, xHost, xDate, xPeriod, xStatus, xSeat, xInTime, xOutTime) VALUES (";
            sql += "'" + eKey + "', ";    //serial
            sql += "'" + xKey + "', ";    //xid
            sql += "'" + xKey2 + "', ";    //xHost
            sql += "'" + Convert.ToDateTime(xKey3).ToShortDateString() + "', ";    //xDate
            sql += "'" + xKey4 + "', ";    //xPeriod
            sql += "'" + xKey5 + "', ";    //xStatus
            sql += "'" + xKey6 + "', ";    //xSeat
            sql += "'" + xKey7 + "', ";    //xInTime
            sql += "'" + xKey8 + "') ";    //xOutTime
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void deleteDVD_video2(string xKey)
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            //connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "delete from Local_dvd_video2 ";
            sql += "where xDate = '" + System.DateTime.Now.ToShortDateString() + "' ";
            sql += "and xPeriod = '" + xKey + "' ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }

        //Local_dvd_video3 15day -----------------------------
        void xActionDVD_video3()
        {
            deleteDVD_video3();
            selectDVD_video3();
            autoUpdateUpdateData("xActionDVD_video3");
        }
        void selectDVD_video3()
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + DS + ";Initial Catalog=DVD;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "select * from DVD_video3 ";
            sql += "where 地點 = '" + reLocalNameNum + "' ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                insertDVD_video3(rd["地點"].ToString(), rd["星期"].ToString(), rd["時段"].ToString());
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void insertDVD_video3(string xKey, string xKey2, string xKey3)
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "INSERT INTO Local_dvd_video3 (xHost, xWeek, xPeriod) VALUES (";
            sql += "'" + xKey + "', ";    //xHost
            sql += "'" + xKey2 + "', ";    //xWeek
            sql += "'" + xKey3 + "') ";    //xPeriod
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void deleteDVD_video3()
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            //connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "delete from Local_dvd_video3 ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }

        //Local_Seat all -----------------------------
        void xActionSeat()
        {
            deleteSeat();
            selectSeat();
            autoUpdateUpdateData("xActionSeat");
        }
        void selectSeat()
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + DS + ";Initial Catalog=DVD;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "select * from DVD_Seat ";
            sql += "where xHost = '" + reLocalNameNum + "' ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                insertSeat(rd["xHost"].ToString(), rd["xTime"].ToString(), rd["xSeat"].ToString(), rd["xSeatPro"].ToString(), rd["xSeatErr"].ToString());
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void insertSeat(string xKey, string xKey2, string xKey3, string xKey4, string xKey5)
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "INSERT INTO Local_Seat (xHost, xTime, xSeat, xSeatPro, xSeatErr) VALUES (";
            sql += "'" + xKey + "', ";    //xHost
            sql += "'" + xKey2 + "', ";    //xTime
            sql += "'" + xKey3 + "', ";    //xSeat
            sql += "'" + xKey4 + "', ";    //xSeatPro
            sql += "'" + xKey5 + "') ";    //xSeatErr
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void deleteSeat()
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            //connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "delete from Local_Seat ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }

        //Local_SwapSeat -----------------------------
        void xActionLocal_SwapSeat(string xKey)
        {
            switch (xKey)
            {
                case "auto":
                    xActionPeriod();
                    autoUpdateUpdateData("xActionLocal_SwapSeat");
                    break;
                case "manual":
                    deleteLocal_SwapSeat();
                    xActionPeriod();
                    break;
            }
            
        }
        void xActionPeriod()
        {
            string WeekStr = null;
            switch (System.DateTime.Now.DayOfWeek.ToString())
            {
                case "Monday": WeekStr = "1"; break;
                case "Tuesday": WeekStr = "2"; break;
                case "Wednesday": WeekStr = "3"; break;
                case "Thursday": WeekStr = "4"; break;
                case "Friday": WeekStr = "5"; break;
                case "Saturday": WeekStr = "6"; break;
                case "Sunday": WeekStr = "7"; break;
            }

            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "SELECT * FROM Local_dvd_video3  ";
            sql += "WHERE xWeek  = '" + WeekStr + "' ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            int i = 0;
            while (rd.Read())
            {
                if (rd["xPeriod"].ToString().Substring(2, 1).ToString() != "不")
                {
                    if (rd["xPeriod"].ToString().Substring(2, 1).ToString() != "請")
                    {
                        i++;
                    }
                }
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
            selectDataLocal_SwapSeat(System.DateTime.Now.ToShortDateString(), i.ToString());
        }
        void selectDataLocal_SwapSeat(string xKey, string xKey2)
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "SELECT * FROM Local_Seat  ";
            sql += "order by xSeat ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                for (int i = 1; i <= Convert.ToInt32(xKey2); i++)
                {
                    insertLocal_SwapSeat(xKey, rd["xSeat"].ToString(), rd["xSeatPro"].ToString(), rd["xSeatErr"].ToString(), i.ToString());
                    selectVideo2(xKey, i.ToString(), rd["xSeat"].ToString());
                }
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void insertLocal_SwapSeat(string xKey, string xKey2, string xKey3, string xKey4, string xKey5)
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "Insert into Local_SwapSeat(xHost ,xDate ,xSeat ,xSeatP ,xSeatB ,xPeriod) VALUES (   ";
            sql += "'" + reLocalNameNum + "', "; //1014
            sql += "'" + xKey + "', ";
            sql += "'" + xKey2 + "', ";
            sql += "'" + xKey3 + "', ";
            sql += "'" + xKey4 + "', ";
            sql += "'" + xKey5 + "' ";
            sql += ") ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void selectVideo2(string xKey, string xKey2, string xKey3)
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "select * from Local_dvd_video2  ";
            sql += "where xHost = '" + reLocalNameNum + "' ";
            sql += "and xDate = '" + xKey + "' ";
            sql += "and left(xPeriod,1) = '" + xKey2 + "' ";
            sql += "and xSeat = '" + xKey3 + "' ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            if (rd.Read())
            {
                updateLocal_SwapSeat(xKey, xKey2, xKey3, rd["xid"].ToString(), rd["xStatus"].ToString());
                //updateLocal_SwapSeat("2016/3/24", "1,AM8:00~PM12:00", "027", "1552290025", "請假");
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void updateLocal_SwapSeat(string xKey, string xKey2, string xKey3, string xKey4, string xKey5)
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "Update Local_SwapSeat SET  ";
            switch (xKey5)
            {
                case "請假":
                    sql += "xNot='4', ";
                    break;
                case "否":
                    sql += "xNot='1', ";
                    break;
            }
            sql += "xID = '" + xKey4 + "' ";
            sql += "where xHost = '" + reLocalNameNum + "' ";
            sql += "and xDate = '" + xKey + "' ";
            sql += "and xPeriod = '" + xKey2 + "' ";
            sql += "and xSeat = '" + xKey3 + "' ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Close();
            cmd.Dispose();
            conn.Close();
            copyDataPanel.Visible = false;
        }
        void deleteLocal_SwapSeat()
        { 
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "delete from Local_SwapSeat ";
            sql += "where xDate = '" + System.DateTime.Now.ToShortDateString() + "' ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }

        // ~~~~~~~~~~ action auto ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        // ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        void xActionAuto()
        {
            //xActionClassIn_DV();
            //xActionClassIn_DV_OutSide();
            xActionClassCategory();
            xActionClass();
            xActionSubclass();
            xActionSubject();
            xActionParticipation();
            xActionStu();
            xActionUserProfile();
        }

        //ClassInU UserProfile >> Local_UserProfile
        void xActionUserProfile()
        {
            deleteLocal_UserProfile();
            selectLocal_UserProfile();
            autoUpdateUpdateData("xActionUserProfile");
        }
        void selectLocal_UserProfile()
        {
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + DS + ";Initial Catalog=STMS;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "SELECT * FROM UserProfile ";
            sql += "WHERE BranchID = '" + reLocalNameNum + "' ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                string xexHost = rd["BranchID"].ToString();
                string xeID = rd["ID"].ToString();
                string xeUserDesc = rd["UserDesc"].ToString();
                insertLocal_UserProfile(xexHost, xeID, xeUserDesc);
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void insertLocal_UserProfile(string xKey, string xKey2, string xKey3)
        {
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "Insert into Local_UserProfile(xHost, ID, UserDesc) VALUES (";
            sql += "'" + xKey + "', ";
            sql += "'" + xKey2 + "', ";
            sql += "'" + xKey3 + "') ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void deleteLocal_UserProfile()
        {
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            //connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "delete from Local_UserProfile ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }

        //ClassIn_DV  15day -----------------------------
        void xActionClassIn_DV()
        {
            deleteClassIn_DV();
            selectClassIn_DV();
            autoUpdateUpdateData("xActionClassIn_DV");
        }
        void selectClassIn_DV()
        {
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + DS + ";Initial Catalog=STMS;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "SELECT * FROM ClassIn_DV ";
            sql += "WHERE xHost = '" + reLocalNameNum + "' ";
            sql += "AND xCreateDate >= '" + System.DateTime.Now.AddDays(-15).ToShortDateString() + "' ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                string xeSerial = rd["serial"].ToString();
                string xeHost = rd["xHost"].ToString();
                string xeYear = rd["xYear"].ToString();
                string xeClassID = rd["xClassID"].ToString();
                string xeSeqNo = rd["xSeqNo"].ToString();
                string xeSubjectID = rd["xSubjectID"].ToString();
                string xedYear = rd["xdYear"].ToString();
                string xedCategoryName = rd["xdCategoryName"].ToString();
                string xedClassName = rd["xdClassName"].ToString();
                string xedSubClassName = rd["xdSubClassName"].ToString();
                string xedSubClassID = rd["xdSubClassID"].ToString();
                string xedSubjectName = rd["xdSubjectName"].ToString();
                string xedNumber = rd["xdNumber"].ToString();
                string xeCreateDate = rd["xCreateDate"].ToString();
                insertClassIn_DV(xeSerial, xeHost, xeYear, xeClassID, xeSeqNo, xeSubjectID, xedYear, xedCategoryName, xedClassName, xedSubClassName, xedSubClassID, xedSubjectName, xedNumber, xeCreateDate);
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void insertClassIn_DV(string xKey, string xKey2, string xKey3, string xKey4, string xKey5, string xKey6, string xKey7, string xKey8, string xKey9, string xKey10, string xKey11, string xKey12, string xKey13, string xKey14)
        {
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            //connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "Insert into Local_ClassIn_DV(serial, xHost, xYear, xClassID, xSeqNo, xSubjectID, xdYear, xdCategoryName, xdClassName, xdSubClassName, xdSubClassID, xdSubjectName, xdNumber, xCreateDate) VALUES (";
            sql += "'" + xKey + "', ";
            sql += "'" + xKey2 + "', ";
            sql += "'" + xKey3 + "', ";
            sql += "'" + xKey4 + "', ";
            sql += "'" + xKey5 + "', ";
            sql += "'" + xKey6 + "', ";
            sql += "'" + xKey7 + "', ";
            sql += "'" + xKey8 + "', ";
            sql += "'" + xKey9 + "', ";
            sql += "'" + xKey10 + "', ";
            sql += "'" + xKey11 + "', ";
            sql += "'" + xKey12 + "', ";
            sql += "'" + xKey13 + "', ";
            sql += "'" + xKey14 + "') ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void deleteClassIn_DV()
        {
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            //connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "delete from Local_ClassIn_DV ";
            sql += "where xCreateDate >= '" + System.DateTime.Now.AddDays(-15).ToShortDateString() + "' ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }

        //ClassIn_DV_OutSide  15day -----------------------------
        void xActionClassIn_DV_OutSide()
        {
            deleteFirstClassIn_DV_OutSide();
            selectClassIn_DV_OutSide();
            autoUpdateUpdateData("xActionClassIn_DV_OutSide");
        }
        void selectClassIn_DV_OutSide()
        {
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + DS + ";Initial Catalog=STMS;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "SELECT * FROM ClassIn_DV_OutSide ";
            sql += "WHERE xHost = '" + reLocalNameNum + "' ";
            sql += "AND xCreateDate >= '" + System.DateTime.Now.AddDays(-15).ToShortDateString() + "' ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                string xeSerial = rd["serial"].ToString();
                string xeID = rd["xID"].ToString();
                string xeHost = rd["xHost"].ToString();
                string xedYear = rd["xdYear"].ToString();
                string xedCategoryName = rd["xdCategoryName"].ToString();
                string xedClassName = rd["xdClassName"].ToString();
                string xedSubClassName = rd["xdSubClassName"].ToString();
                string xedSubClassID = rd["xdSubClassID"].ToString();
                string xedSubjectName = rd["xdSubjectName"].ToString();
                string xedNumber = rd["xdNumber"].ToString();
                string xeCreateDate = rd["xCreateDate"].ToString();

                insertClassIn_DV_OutSide(xeSerial, xeID, xeHost, xedYear, xedCategoryName, xedClassName, xedSubClassName, xedSubClassID, xedSubjectName, xedNumber, xeCreateDate);
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void insertClassIn_DV_OutSide(string xKey, string xKey2, string xKey3, string xKey4, string xKey5, string xKey6, string xKey7, string xKey8, string xKey9, string xKey10, string xKey11)
        {
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            //connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "Insert into Local_ClassIn_DV_OutSide(serial, xID, xHost, xdYear, xdCategoryName, xdClassName, xdSubClassName,xdSubClassID, xdSubjectName, xdNumber, xCreateDate) VALUES (";
            sql += "'" + xKey + "', ";
            sql += "'" + xKey2 + "', ";
            sql += "'" + xKey3 + "', ";
            sql += "'" + xKey4 + "', ";
            sql += "'" + xKey5 + "', ";
            sql += "'" + xKey6 + "', ";
            sql += "'" + xKey7 + "', ";
            sql += "'" + xKey8 + "', ";
            sql += "'" + xKey9 + "', ";
            sql += "'" + xKey10 + "', ";
            sql += "'" + xKey11 + "') ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void deleteFirstClassIn_DV_OutSide()
        {
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            //connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "delete from Local_ClassIn_DV_OutSide ";
            sql += "where xCreateDate >= '" + System.DateTime.Now.AddDays(-15).ToShortDateString() + "' ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }

        //Local_ClassCategory all -----------------------------
        void xActionClassCategory()
        {
            deleteClassCategory();
            selectClassCategory();
            autoUpdateUpdateData("xActionClassCategory");
        }
        void selectClassCategory()
        {
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + DS + ";Initial Catalog=STMS;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "SELECT * FROM ClassCategory ";
            sql += "WHERE BranchID = '" + reLocalNameNum + "' ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                string xeOID = rd["OID"].ToString();
                string xeBranchID = rd["BranchID"].ToString();
                string xeYear = rd["Year"].ToString();
                string xeName = rd["Name"].ToString();
                insertClassCategory(xeOID, xeBranchID, xeYear, xeName);
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void insertClassCategory(string xKey, string xKey2, string xKey3, string xKey4)
        {
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "Insert into Local_ClassCategory(serial, xHost, xYear, xName) VALUES (";
            sql += "'" + xKey + "', ";
            sql += "'" + xKey2 + "', ";
            sql += "'" + xKey3 + "', ";
            sql += "'" + xKey4 + "') ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void deleteClassCategory()
        {
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "delete from Local_ClassCategory ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }

        //Local_Class 500 -----------------------------
        void xActionClass()
        {
            xSetNumdvd_Class();
            deleteClass();
            selectClass();
            autoUpdateUpdateData("xActionClass");
        }
        void xSetNumdvd_Class()
        {
            xNum = "";
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "select top 1000 * from Local_Class ORDER BY serial desc";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                xNum = rd["serial"].ToString();
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void selectClass()
        {
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + DS + ";Initial Catalog=STMS;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "SELECT * FROM Class ";
            sql += "WHERE oid >= '" + xNum + "' ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                string xeOID = rd["OID"].ToString();
                string xeBranchID = rd["BranchID"].ToString();
                string xeYear = rd["Year"].ToString();
                string xeID = rd["ID"].ToString();
                string xeName = rd["Name"].ToString();
                string xeDeadline = rd["Deadline"].ToString();
                string xeCategoryOID = rd["CategoryOID"].ToString();
                insertClass(xeOID, xeBranchID, xeYear, xeID, xeName, xeDeadline, xeCategoryOID);
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void insertClass(string xKey, string xKey2, string xKey3, string xKey4, string xKey5, string xKey6, string xKey7)
        {
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "Insert into Local_Class(serial, xHost, xYear,xID ,xName ,xDeadline ,xCategoryOID) VALUES (";
            sql += "'" + xKey + "', ";
            sql += "'" + xKey2 + "', ";
            sql += "'" + xKey3 + "', ";
            sql += "'" + xKey4 + "', ";
            sql += "'" + xKey5 + "', ";
            if (xKey6 != "")
            {
                sql += "'" + Convert.ToDateTime(xKey6).ToShortDateString() + "', ";
            }
            else
            {
                sql += "null, ";
            }
            sql += "'" + xKey7 + "') ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void deleteClass()
        {
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "delete from Local_Class WHERE serial >= '" + xNum + "' ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }

        //Local_Subclass 500 -----------------------------
        void xActionSubclass()
        {
            xSetNumSubclass();
            deleteSubclass();
            selectSubclass();
            autoUpdateUpdateData("xActionSubclass");
        }
        void xSetNumSubclass()
        {
            xNum = "";
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "select top 1500 * from Local_SubClass ORDER BY serial desc";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                xNum = rd["serial"].ToString();
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void selectSubclass()
        {
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + DS + ";Initial Catalog=STMS;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "SELECT * FROM Subclass ";
            sql += "WHERE oid >= '" + xNum + "' ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                string xeOID = rd["OID"].ToString();
                string xeBranchID = rd["BranchID"].ToString();
                string xeYear = rd["Year"].ToString();
                string xeName = rd["Name"].ToString();
                string xeClassID = rd["ClassID"].ToString();
                string xeDeadline = rd["Deadline"].ToString();
                string xeSeqNo = rd["SeqNo"].ToString();
                string xeClassOID = rd["ClassOID"].ToString();
                insertSubclass(xeOID, xeBranchID, xeYear, xeName, xeClassID, xeDeadline, xeSeqNo, xeClassOID);
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void insertSubclass(string xKey, string xKey2, string xKey3, string xKey4, string xKey5, string xKey6, string xKey7, string xKey8)
        {
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "Insert into Local_SubClass(serial, xHost, xYear,xName ,xClassID ,xDeadline , xSeqNo, xClassOID) VALUES (";
            sql += "'" + xKey + "', ";
            sql += "'" + xKey2 + "', ";
            sql += "'" + xKey3 + "', ";
            sql += "'" + xKey4 + "', ";
            sql += "'" + xKey5 + "', ";
            if (xKey6 != "")
            {
                sql += "'" + Convert.ToDateTime(xKey6).ToShortDateString() + "', ";
            }
            else
            {
                sql += "null, ";
            }
            sql += "'" + xKey7 + "', ";
            sql += "'" + xKey8 + "') ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void deleteSubclass()
        {
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "delete from Local_SubClass WHERE serial >= '" + xNum + "' ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }

        //Local_Subject 500 -----------------------------
        void xActionSubject()
        {
            xSetNumSubject();
            deleteSubject();
            selectSubject();
            autoUpdateUpdateData("xActionSubject");
        }
        void xSetNumSubject()
        {
            xNum = "";
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "select top 1000 * from Local_Subject ORDER BY serial desc";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                xNum = rd["serial"].ToString();
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void selectSubject()
        {
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + DS + ";Initial Catalog=STMS;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "SELECT * FROM Subject ";
            sql += "WHERE oid >= '" + xNum + "' ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                string xeOID = rd["OID"].ToString();
                string xeBranchID = rd["BranchID"].ToString();
                string xeYear = rd["Year"].ToString();
                string xeID = rd["ID"].ToString();
                string xeName = rd["Name"].ToString();
                insertSubject(xeOID, xeBranchID, xeYear, xeID, xeName);
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void insertSubject(string xKey, string xKey2, string xKey3, string xKey4, string xKey5)
        {
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "Insert into Local_Subject(serial, xHost, xYear, xID, xName) VALUES (";
            sql += "'" + xKey + "', ";
            sql += "'" + xKey2 + "', ";
            sql += "'" + xKey3 + "', ";
            sql += "'" + xKey4 + "', ";
            sql += "'" + xKey5 + "') ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void deleteSubject()
        {
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "delete from Local_Subject WHERE serial >= '" + xNum + "' ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }

        //Local_Participation 500 -----------------------------
        void xActionParticipation()
        {
            xSetNumParticipation();
            deleteParticipation();
            selectParticipation();
            autoUpdateUpdateData("xActionParticipation");
        }
        void xSetNumParticipation()
        {
            xNum = "";
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "select top 500 * from Local_Participation ORDER BY serial desc";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                xNum = rd["serial"].ToString();
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void selectParticipation()
        {
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + DS + ";Initial Catalog=STMS;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "SELECT * FROM Participation ";
            sql += "WHERE oid >= '" + xNum + "' ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                string xeOID = rd["OID"].ToString();
                string xeBranchID = rd["BranchID"].ToString();
                string xeYear = rd["Year"].ToString();
                string xeStudentID = rd["StudentID"].ToString();
                string xeSubjectID = rd["SubjectID"].ToString();
                insertParticipation(xeOID, xeBranchID, xeYear, xeStudentID, xeSubjectID);
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void insertParticipation(string xKey, string xKey2, string xKey3, string xKey4, string xKey5)
        {
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "Insert into Local_Participation(serial, xHost, xYear, xID, xSubjectID) VALUES (";
            sql += "'" + xKey + "', ";
            sql += "'" + xKey2 + "', ";
            sql += "'" + xKey3 + "', ";
            sql += "'" + xKey4 + "', ";
            sql += "'" + xKey5 + "') ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void deleteParticipation()
        {
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "delete from Local_Participation WHERE serial >= '" + xNum + "' ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }

        //Local_Stu_Data 500 -----------------------------
        void xActionStu()
        {
            xSetNumStu();
            deleteStu();
            selectStu();
            autoUpdateUpdateData("xActionStu");
        }
        void xSetNumStu()
        {
            xNum = "";
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "select top 1200 * from Local_Stu_Data ORDER BY serial desc";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                xNum = rd["serial"].ToString();
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void selectStu()
        {
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + DS + ";Initial Catalog=STMS;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "SELECT * FROM Student ";
            sql += "WHERE oid >= '" + xNum + "' ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                string xeOID = rd["OID"].ToString();
                string xeBranchID = rd["BranchID"].ToString();
                string xeYear = rd["Year"].ToString();
                string xeID = rd["ID"].ToString();
                string xeName = rd["Name"].ToString();
                string xeClassID = rd["ClassID"].ToString();
                string xeSeqNo = rd["SeqNo"].ToString();
                string xePID = rd["PID"].ToString();
                string xeBirthDate = rd["BirthDate"].ToString();
                string xeTel = rd["Tel1"].ToString();
                
                string rxeName = Regex.Replace(xeName, @"[\W-]+", " ");

                insertStu(xeOID, xeBranchID, xeYear, xeID, rxeName, xeClassID, xeSeqNo, xePID, xeBirthDate, xeTel);
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void insertStu(string xKey, string xKey2, string xKey3, string xKey4, string xKey5, string xKey6, string xKey7, string xKey8, string xKey9, string xKey10)
        {
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "Insert into Local_Stu_Data(serial, xBranchID, xYear, xID, xName, xClassID, xSeqNo, xPID, xBirthDate, xTel) VALUES (";
            sql += "'" + xKey + "', ";
            sql += "'" + xKey2 + "', ";
            sql += "'" + xKey3 + "', ";
            sql += "'" + xKey4 + "', ";
            sql += "'" + xKey5 + "', ";
            sql += "'" + xKey6 + "', ";
            sql += "'" + xKey7 + "', ";
            sql += "'" + xKey8 + "', ";
            sql += "'" + xKey9 + "', ";            
            sql += "'" + xKey10 + "') ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void deleteStu()
        {
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "delete from Local_Stu_Data WHERE serial >= '" + xNum + "' ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            manualRun01.Text = "查詢";
        }
        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            manualRun02.Text = "查詢";
        }
        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e)
        {
            manualRun03.Text = "查詢";
        }

        private void manualRun_Click(object sender, EventArgs e)
        {
            Button bt = (Button)sender;
            switch(bt.Name)
            {
                case "manualRun01":
                    Button bt01 = (Button)sender;
                    switch (bt01.Text)
                    {
                        case "查詢":
                            if (comboBox5.SelectedIndex.ToString() != "0")
                            {
                                manualRun01.Enabled = false;
                                xcselectStu(comboBox5.SelectedItem.ToString());
                                bt01.Text = "匯入";
                                manualRun01.Enabled = true;
                            }
                            break;
                        case "匯入":
                            if (comboBox5.SelectedIndex.ToString() != "0")
                            {
                                manualRun01.Enabled = false;
                                deleteStu_Data(comboBox5.SelectedItem.ToString());
                                xselectStu_Data(comboBox5.SelectedItem.ToString());
                                MessageBox.Show("匯入結果：" + label9.Text + @"/" + label5.Text);
                                bt01.Text = "查詢";
                                manualRun01.Enabled = true;
                                label9.Text = "";
                                label5.Text = "";
                            }
                            break;
                    }
                    break;
                case "manualRun02":
                    Button bt02 = (Button)sender;
                    switch (bt02.Text)
                    {
                        case "查詢":
                            if (comboBox6.SelectedIndex.ToString() != "0")
                            {
                                manualRun02.Enabled = false;
                                xcselectClassCategory(comboBox6.SelectedItem.ToString());
                                bt02.Text = "匯入";
                                manualRun02.Enabled = true;

                            }
                            break;
                        case "匯入":
                            if (comboBox6.SelectedIndex.ToString() != "0")
                            {
                                manualRun02.Enabled = false;
                                xdeleteClassCategory(comboBox6.SelectedItem.ToString());
                                xselectClassCategory(comboBox6.SelectedItem.ToString());
                                MessageBox.Show("匯入結果：" + label14.Text + @"/" + label12.Text);
                                bt02.Text = "查詢";
                                manualRun02.Enabled = true;
                                label14.Text = "";
                                label12.Text = "";
                            }
                            break;
                    }
                    break;
                case "manualRun03":
                    Button bt03 = (Button)sender;
                    switch (bt03.Text)
                    {
                        case "查詢":
                            if (comboBox7.SelectedIndex.ToString() != "0")
                            {
                                manualRun03.Enabled = false;
                                xcselectDVD_video2(comboBox7.SelectedItem.ToString());
                                bt03.Text = "匯入";
                                manualRun03.Enabled = true;
                            }
                            break;
                        case "匯入":
                            if (comboBox7.SelectedIndex.ToString() != "0")
                            {
                                manualRun03.Enabled = false;
                                deleteDVD_video2(comboBox7.SelectedItem.ToString());
                                xselectDVD_video2(comboBox7.SelectedItem.ToString());                                
                                MessageBox.Show("匯入結果：" + label24.Text + @"/" + label22.Text);
                                bt03.Text = "查詢";
                                manualRun03.Enabled = true;
                                label24.Text = "";
                                label22.Text = "";

                            }
                            break;
                    }
                    break;
            }
        }

        void showCBa()
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            comboBox5.Items.Clear();
            comboBox5.Items.Add("請選擇");
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "SELECT DISTINCT RIGHT(xYear,2) AS E1 ";
            sql += "FROM Local_Stu_Data ";
            sql += "WHERE (xYear >= '" + System.DateTime.Now.AddYears(-7).ToString() + "')  order by E1 desc ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                comboBox5.Items.Add(rd["E1"].ToString());
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
            comboBox5.SelectedIndex = 0;
        }
        void showCBb()
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            comboBox6.Items.Clear();
            comboBox6.Items.Add("請選擇");
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "SELECT DISTINCT RIGHT(xYear,2) AS E1 ";
            sql += "FROM Local_Stu_Data ";
            sql += "WHERE (xYear >= '" + System.DateTime.Now.AddYears(-7).ToString() + "')  order by E1 desc ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                comboBox6.Items.Add(rd["E1"].ToString());
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
            comboBox6.SelectedIndex = 0;
        }
        void showCBc()
        {
            string connetionString = null, sql = null;
            string WeekStr = null;
            switch (System.DateTime.Now.DayOfWeek.ToString())
            {
                case "Monday": WeekStr = "1"; break;
                case "Tuesday": WeekStr = "2"; break;
                case "Wednesday": WeekStr = "3"; break;
                case "Thursday": WeekStr = "4"; break;
                case "Friday": WeekStr = "5"; break;
                case "Saturday": WeekStr = "6"; break;
                case "Sunday": WeekStr = "7"; break;
            }
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            comboBox7.Items.Clear();
            comboBox7.Items.Add("請選擇");
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "select xPeriod AS E1 from Local_dvd_video3  ";
            sql += "where xHost = '" + reLocalNameNum + "' ";
            sql += "and xWeek = '" + WeekStr + "' ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                if (rd["E1"].ToString().Substring(2, 1).ToString() != "不")
                {
                    if (rd["E1"].ToString().Substring(2, 1).ToString() != "請")
                    {
                        comboBox7.Items.Add(rd["E1"].ToString());
                    }
                }                
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
            comboBox7.SelectedIndex = 0;
        }

        private void autoSaveBt_Click(object sender, EventArgs e)
        {

        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            if (copyDataLabel.Visible == true)
            {
                timer2.Stop();
                copyDataPanel.Visible = false;
                manualSetPannel.Visible = false;
            }
        }

        //Local_Stu_Data -----------------------------
        void xcselectStu(string xxYear)
        {
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + DS + ";Initial Catalog=STMS;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "SELECT COUNT(*) AS E1 FROM Student ";
            sql += "WHERE (SUBSTRING(Year, 3, 2) = '" + xxYear + "' ) ";
            sql += "AND (Status = '正常') ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            if (rd.Read())
            {
                label9.Text = rd["E1"].ToString();
            }
            rd.Close();
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void deleteStu_Data(string xKey)
        {
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "delete from Local_Stu_Data ";
            sql += "WHERE (SUBSTRING(xID, 1, 2) = '" + xKey + "')";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void xselectStu_Data(string xKey)
        {
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + DS + ";Initial Catalog=STMS;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "SELECT * FROM Student ";
            sql += "WHERE (SUBSTRING(ID, 1, 2) = '" + xKey + "') ";
            sql += "AND (Status = '正常')";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
	        int i = 0;
            while (rd.Read())
            {
                i++;
                string xeOID = rd["OID"].ToString();
                string xeBranchID = rd["BranchID"].ToString();
                string xeYear = rd["Year"].ToString();
                string xeID = rd["ID"].ToString();
                string xeName = rd["Name"].ToString();
                string xeClassID = rd["ClassID"].ToString();
                string xeSeqNo = rd["SeqNo"].ToString();
                string xePID = rd["PID"].ToString();
                string xeBirthDate = rd["BirthDate"].ToString();
                string xeTel = rd["Tel1"].ToString();
                string rxeName = Regex.Replace(xeName, @"[\W-]+", " ");
                insertStu(xeOID, xeBranchID, xeYear, xeID, rxeName, xeClassID, xeSeqNo, xePID, xeBirthDate, xeTel);
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
            label5.Text = i.ToString();
        }

        //Local_Class -----------------------------
        void xcselectClassCategory(string xKey)
        {
            int xCountClassCategory = 0;
            int xCountClass = 0;
            int xCountSubClass = 0;
            int xCountSubject = 0;
            int xCountParticipation = 0;

            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + DS + ";Initial Catalog=STMS;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "SELECT COUNT(*) AS E1 FROM ClassCategory ";
            sql += "WHERE (SUBSTRING(Year, 3, 2) = '" + xKey + "') ";
            sql += "AND (BranchID = '" + reLocalNameNum + "') ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            if (rd.Read())
            {
                label16.Text = rd["E1"].ToString();
                xCountClassCategory = Convert.ToInt16(label16.Text);
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();

            connetionString = "Data Source=" + DS + ";Initial Catalog=STMS;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "SELECT COUNT(*) AS E1 FROM Class ";
            sql += "WHERE (SUBSTRING(Year, 3, 2) = '" + xKey + "') ";
            sql += "AND (BranchID = '" + reLocalNameNum + "') ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            if (rd.Read())
            {
                label17.Text = rd["E1"].ToString();
                xCountClass = Convert.ToInt32(label17.Text);
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();

            connetionString = "Data Source=" + DS + ";Initial Catalog=STMS;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "SELECT COUNT(*) AS E1 FROM Subclass ";
            sql += "WHERE (SUBSTRING(Year, 3, 2) = '" + xKey + "') ";
            sql += "AND (BranchID = '" + reLocalNameNum + "') ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            if (rd.Read())
            {
                label18.Text = rd["E1"].ToString();
                xCountSubClass = Convert.ToInt32(label18.Text);
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();

            connetionString = "Data Source=" + DS + ";Initial Catalog=STMS;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "SELECT COUNT(*) AS E1 FROM Subject ";
            sql += "WHERE (SUBSTRING(Year, 3, 2) = '" + xKey + "') ";
            sql += "AND (BranchID = '" + reLocalNameNum + "') ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            if (rd.Read())
            {
                label19.Text = rd["E1"].ToString();
                xCountSubject = Convert.ToInt32(label19.Text);
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();

            connetionString = "Data Source=" + DS + ";Initial Catalog=STMS;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "SELECT COUNT(*) AS E1 FROM Participation ";
            sql += "WHERE (SUBSTRING(Year, 3, 2) = '" + xKey + "') ";
            sql += "AND (BranchID = '" + reLocalNameNum + "') ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            if (rd.Read())
            {
                label20.Text = rd["E1"].ToString();
                xCountParticipation = Convert.ToInt32(label20.Text);
                label14.Text = Convert.ToString(xCountClassCategory + xCountClass + xCountSubClass + xCountSubject + xCountParticipation);
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }

        void xdeleteClassCategory(string xKey)
        {
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "delete from Local_ClassCategory ";
            sql += "where (SUBSTRING(xYear, 3, 2) = '" + xKey + "') ";
            sql += "and (xHost = '" + reLocalNameNum + "') ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            rd.Close();
            cmd.Dispose();
            conn.Close();

            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "delete from Local_Class ";
            sql += "where (SUBSTRING(xYear, 3, 2) = '" + xKey + "') ";
            sql += "and (xHost = '" + reLocalNameNum + "') ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            rd.Close();
            cmd.Dispose();
            conn.Close();

            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "delete from Local_Subclass ";
            sql += "where (SUBSTRING(xYear, 3, 2) = '" + xKey + "') ";
            sql += "and (xHost = '" + reLocalNameNum + "') ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            rd.Close();
            cmd.Dispose();
            conn.Close();

            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "delete from Local_Subject ";
            sql += "where (SUBSTRING(xYear, 3, 2) = '" + xKey + "') ";
            sql += "and (xHost = '" + reLocalNameNum + "') ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            rd.Close();
            cmd.Dispose();
            conn.Close();

            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "delete from Local_Participation ";
            sql += "where (SUBSTRING(xYear, 3, 2) = '" + xKey + "') ";
            sql += "and (xHost = '" + reLocalNameNum + "') ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }

        void xselectClassCategory(string xKey)
        {

            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + DS + ";Initial Catalog=STMS;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "SELECT * FROM ClassCategory ";
            sql += "WHERE (SUBSTRING(Year, 3, 2) = '" + xKey + "') ";
            sql += "AND (BranchID = '" + reLocalNameNum + "') ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            int i = 0;
            while (rd.Read())
            {
                i++;
                string xeOID = rd["OID"].ToString();
                string xeBranchID = rd["BranchID"].ToString();
                string xeYear = rd["Year"].ToString();
                string xeName = rd["Name"].ToString();
                insertClassCategory(xeOID, xeBranchID, xeYear, xeName);
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();

            connetionString = "Data Source=" + DS + ";Initial Catalog=STMS;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "SELECT * FROM Class ";
            sql += "WHERE (SUBSTRING(Year, 3, 2) = '" + xKey + "') ";
            sql += "AND (BranchID = '" + reLocalNameNum + "') ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            int j = 0;
            while (rd.Read())
            {
                j++;
                string xeOID = rd["OID"].ToString();
                string xeBranchID = rd["BranchID"].ToString();
                string xeYear = rd["Year"].ToString();
                string xeID = rd["ID"].ToString();
                string xeName = rd["Name"].ToString();
                string xeDeadline = rd["Deadline"].ToString();
                string xeCategoryOID = rd["CategoryOID"].ToString();
                insertClass(xeOID, xeBranchID, xeYear, xeID, xeName, xeDeadline, xeCategoryOID);
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();

            connetionString = "Data Source=" + DS + ";Initial Catalog=STMS;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "SELECT * FROM Subclass ";
            sql += "WHERE (SUBSTRING(Year, 3, 2) = '" + xKey + "') ";
            sql += "AND (BranchID = '" + reLocalNameNum + "') ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            int k = 0;
            while (rd.Read())
            {
                k++;
                string xeOID = rd["OID"].ToString();
                string xeBranchID = rd["BranchID"].ToString();
                string xeYear = rd["Year"].ToString();
                string xeName = rd["Name"].ToString();
                string xeClassID = rd["ClassID"].ToString();
                string xeDeadline = rd["Deadline"].ToString();
                string xeSeqNo = rd["SeqNo"].ToString();
                string xeClassOID = rd["ClassOID"].ToString();
                insertSubclass(xeOID, xeBranchID, xeYear, xeName, xeClassID, xeDeadline, xeSeqNo, xeClassOID);
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();

            connetionString = "Data Source=" + DS + ";Initial Catalog=STMS;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "SELECT * FROM Subject ";
            sql += "WHERE (SUBSTRING(Year, 3, 2) = '" + xKey + "') ";
            sql += "AND (BranchID = '" + reLocalNameNum + "') ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            int l = 0;
            while (rd.Read())
            {
                l++;
                string xeOID = rd["OID"].ToString();
                string xeBranchID = rd["BranchID"].ToString();
                string xeYear = rd["Year"].ToString();
                string xeID = rd["ID"].ToString();
                string xeName = rd["Name"].ToString();
                insertSubject(xeOID, xeBranchID, xeYear, xeID, xeName);
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();

            connetionString = "Data Source=" + DS + ";Initial Catalog=STMS;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "SELECT * FROM Participation ";
            sql += "WHERE (SUBSTRING(Year, 3, 2) = '" + xKey + "') ";
            sql += "AND (BranchID = '" + reLocalNameNum + "') ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            int m = 0;
            while (rd.Read())
            {
                m++;
                string xeOID = rd["OID"].ToString();
                string xeBranchID = rd["BranchID"].ToString();
                string xeYear = rd["Year"].ToString();
                string xeStudentID = rd["StudentID"].ToString();
                string xeSubjectID = rd["SubjectID"].ToString();
                insertParticipation(xeOID, xeBranchID, xeYear, xeStudentID, xeSubjectID);
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
            label31.Text = m.ToString();
            label12.Text = Convert.ToString(i + j + k + l + m);
        }

        //Local_DVD_video2 -----------------------------
        void xcselectDVD_video2(string xKey)
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + DS + ";Initial Catalog=DVD;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "select count(*) AS E1 from DVD_video2 ";
            sql += "where 地點 = '" + reLocalNameNum + "' ";
            sql += "AND 日期 = '" + System.DateTime.Now.ToShortDateString() + "' ";
            sql += "AND 時段 = '" + xKey + "' ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            if (rd.Read())
            {
                label24.Text = rd["E1"].ToString();
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void xselectDVD_video2(string xKey)
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + DS + ";Initial Catalog=DVD;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "select serial, id, 地點, 日期, 時段, 已看, 座位, xInTime, xOutTime from DVD_video2 ";
            sql += "WHERE (地點 = '" + reLocalNameNum + "') ";
            sql += "AND (日期 = '" + System.DateTime.Now.ToShortDateString() + "') ";
            sql += "AND (時段 = '" + xKey + "') ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            int i = 0;
            while (rd.Read())
            {
                i++;
                insertDVD_video2(Convert.ToInt32(rd["serial"].ToString()), rd["id"].ToString(), rd["地點"].ToString(), rd["日期"].ToString(), rd["時段"].ToString(), rd["已看"].ToString(), rd["座位"].ToString(), rd["xInTime"].ToString(), rd["xOutTime"].ToString());
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
            label22.Text = i.ToString();
        }

    }
}
