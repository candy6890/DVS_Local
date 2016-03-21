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

namespace DVS
{
    public partial class Main : Form
    {
        string xNum = null;
        string DS = "211.75.132.163";
        string localDS = System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString();
        public Main()
        {
            InitializeComponent();
            xMain();
        }

        void xMain()
        {
            xObject();
            //xActionClassIn_DV();
         
        }

        void xObject()
        {
            nowTime.Text = "";
            nowTime2.Text = "";
            timer1.Interval = 1000;
            timer1.Start();
            timer1.Interval = 1000;
            timer2.Start();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            nowTime.Text = System.DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            nowTime2.Text = System.DateTime.Now.ToString("yyyyMMddHHmmss");
            if (nowTime2.Text == "20150605025700")
            {
                xActionDVD_video2();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {

            xActionClassIn_DV();
            xActionClassIn_DV_OutSide();

            xActiondvd_logistics();
            xActiondvd_logistics2();


            xActionDVD_video2();
            xActionDVD_Seat();
            

            //mForm1 mf = new mForm1();
            //mf.MdiParent = this;
            //mf.Show();


            //xActionDVD_video2();
        }

        //dvd_logistics 15day -----------------------------
        void xActiondvd_logistics()
        {
            xDeldvd_logistics();
            xSelectdvd_logistics();
        }
        void xSelectdvd_logistics()
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + DS + ";Initial Catalog=STMS;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "SELECT serial,DVD編號, 上課日期,當科堂次 FROM DVD_logistics ";
            sql += "WHERE SUBSTRING(DVD編號, 3, 2) = '33' ";
            sql += "AND 上課日期 >= '" + System.DateTime.Now.AddDays(-15).ToShortDateString() + "' ";
            sql += "ORDER BY  上課日期 DESC";
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
            connetionString = "Data Source=DVS-33-00;Initial Catalog=sourceData;User ID=sa;Password=";
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
        void xDeldvd_logistics()
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            //connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
            connetionString = "Data Source=DVS-33-00;Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "delete from Local_dvd_logistics ";
            sql += "where xCourseDate >= '" + System.DateTime.Now.AddDays(-15).ToShortDateString() + "' ";
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
            sql += "where SUBSTRING(DVD編號, 3, 2) = '33' ";
            sql += "and serial >= '" + xNum + "' ";
            sql += "order by DVD編號 ";
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
            connetionString = "Data Source=DVS-33-00;Initial Catalog=sourceData;User ID=sa;Password=";
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
            connetionString = "Data Source=DVS-33-00;Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "select top 500 * from Local_dvd_logistics2 order by serial desc";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            while(rd.Read())
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
            connetionString = "Data Source=DVS-33-00;Initial Catalog=sourceData;User ID=sa;Password=";
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

        //ClassIn_DV  15day -----------------------------
        void xActionClassIn_DV()
        {
            deleteClassIn_DV();
            selectClassIn_DV();
        }
        void selectClassIn_DV()
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + DS + ";Initial Catalog=STMS;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "SELECT * FROM ClassIn_DV ";
            sql += "WHERE xHost = '33' ";
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
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            //connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
            connetionString = "Data Source=DVS-33-00;Initial Catalog=sourceData;User ID=sa;Password=";
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
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            //connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
            connetionString = "Data Source=DVS-33-00;Initial Catalog=sourceData;User ID=sa;Password=";
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
        }
        void selectClassIn_DV_OutSide()
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + DS + ";Initial Catalog=STMS;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "SELECT * FROM ClassIn_DV_OutSide ";
            sql += "WHERE xHost = '33' ";
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
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            //connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
            connetionString = "Data Source=DVS-33-00;Initial Catalog=sourceData;User ID=sa;Password=";
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
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            //connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
            connetionString = "Data Source=DVS-33-00;Initial Catalog=sourceData;User ID=sa;Password=";
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

        //Local_dvd_video2 15day -----------------------------
        void xActionDVD_video2()
        {
            selectDVD_video2();
        }
        void selectDVD_video2()
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + DS + ";Initial Catalog=DVD;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "select serial, id, 地點, 日期, 時段, 已看, 座位, xInTime, xOutTime from DVD_video2 ";
            sql += "where 地點 = '33' ";
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
            connetionString = "Data Source=DVS-33-00;Initial Catalog=sourceData;User ID=sa;Password=";
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
        void deleteDVD_video2()
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            //connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
            connetionString = "Data Source=DVS-33-00;Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "delete from Local_dvd_video2 ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }

        //Local_ClassCategory  -----------------------------
        void xActionClassCategory()
        {

        }
        void selectClassCategory()
        {

        }
        void insertClassCategory()
        {

        }
        void deleteClassCategory()
        {

        }

        //Local_Class  -----------------------------
        void xActionClass()
        {

        }
        void selectClass()
        {

        }
        void insertClass()
        {

        }
        void deleteClass()
        {

        }

        //Local_Subclass  -----------------------------
        void xActionSubclass()
        {

        }
        void selectSubclass()
        {

        }
        void insertSubclass()
        {

        }
        void deleteSubclass()
        {

        }

        //Local_Subject  -----------------------------
        void xActionSubject()
        {

        }
        void selectSubject()
        {

        }
        void insertSubject()
        {

        }
        void deleteSubject()
        {

        }

        //Local_Association  -----------------------------
        void xActionAssociation()
        {

        }
        void selectAssociation()
        {

        }
        void insertAssociation()
        {

        }
        void deleteAssociation()
        {

        }

        //Local_dvd_video5  未知數
        void xActionDVD_video5()
        { 
            
        }
        void selectDVD_video5()
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + DS + ";Initial Catalog=DVD;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "select * from DVD_video5 ";
            sql += "where SUBSTRING(id, 3, 2) = '33' ";
            sql += "order by serial ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {

                //insertFirstDVD_video5(Convert.ToInt32(rd["serial"].ToString()), rd["開通時間"].ToString(), rd["開通"].ToString(), rd["id"].ToString(), Convert.ToInt32(rd["免費次數"].ToString()));
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void insertDVD_video5(int eKey, string xKey, string xKey2, string xKey3, int eKey2)
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            //connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
            connetionString = "Data Source=DVS-33-00;Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "INSERT INTO Local_dvd_video5 (serial, xOpenDate, xOpen, xid, xPoint) VALUES (";
            sql += "'" + eKey + "', ";    //serial
            sql += "'" + xKey + "', ";    //xOpenDate
            sql += "'" + xKey2 + "', ";    //xOpen
            sql += "'" + xKey3 + "', ";    //xid
            sql += "'" + eKey2 + "') ";    //xPoint
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            rd.Close();
            cmd.Dispose();
            //connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
            connetionString = "Data Source=DVS-33-00;Initial Catalog=sourceData;User ID=sa;Password=";
            conn.Close();
        }
        void deleteDVD_video5()
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            //connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
            connetionString = "Data Source=DVS-33-00;Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "delete from Local_dvd_video5 ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }

        //DVD_video3
        void xActionDVD_video3()
        { 
        
        }
        void selectDVD_video3()
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + DS + ";Initial Catalog=DVD;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "select * from DVD_video3 ";
            sql += "where 地點 = '33' ";
            sql += "order by serial ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                //listBox3.Items.Add(rd["星期"].ToString() + "-" + rd["時段"].ToString());
                //label7.Text = System.DateTime.Now.ToString("HH:mm:ss");
                //insertFirstDVD_video3(rd["地點"].ToString(), rd["星期"].ToString(), rd["時段"].ToString());
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void insertDVD_video3(string xKey, string xKey2, string xKey3)
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            //connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
            connetionString = "Data Source=DVS-33-00;Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "INSERT INTO Local_dvd_video3 (xHost, xWeek, xPeriod) VALUES (";
            sql += "'" + xKey + "', ";    //serial
            sql += "'" + xKey2 + "', ";    //xOpenDate
            sql += "'" + xKey3 + "') ";    //xPoint
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
            connetionString = "Data Source=DVS-33-00;Initial Catalog=sourceData;User ID=sa;Password=";
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


        //DVD_Seat
        void xActionDVD_Seat()
        {
            //selectDVD_Seat();
            selectPeriod();
        }
        void selectPeriod()
        {
            string connetionString = null;
            SqlConnection conn;
            SqlCommand cmd;
            string sql = null;
            SqlDataReader rd;
            //connetionString = "Data Source=211.75.132.163;Initial Catalog=DVD;User ID=sa;Password=";
            connetionString = "Data Source=" + DS + ";Initial Catalog=DVD;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "SELECT * FROM DVD_video3  ";
            sql += "WHERE 地點  = '33' ";
            switch (System.DateTime.Now.DayOfWeek.ToString())
            {
                case "Monday": sql += "AND 星期  = '1' "; break;
                case "Tuesday": sql += "AND 星期  = '2' "; break;
                case "Wednesday": sql += "AND 星期  = '3' "; break;
                case "Thursday": sql += "AND 星期  = '4' "; break;
                case "Friday": sql += "AND 星期  = '5' "; break;
                case "Saturday": sql += "AND 星期  = '6' "; break;
                case "Sunday": sql += "AND 星期  = '7' "; break;
            }
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            int i = 0;
            while (rd.Read())
            {
                if (rd["時段"].ToString().Substring(2, 1).ToString() != "不" && rd["時段"].ToString().Substring(2, 1).ToString() != "請")
                {
                    i++;
                }
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
            selectDVD_Seat(i);
        }
        void selectDVD_Seat(int eKey)
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + DS + ";Initial Catalog=DVD;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "SELECT * FROM DVD_Seat ";
            sql += "WHERE xHost = '33' ";
            sql += "order by xSeat ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                for (int i = 1; i <= eKey; i++)
                {
                    insertLocal_SwapSeat(rd["xHost"].ToString(), System.DateTime.Now.ToShortDateString(), rd["xSeat"].ToString(), rd["xSeatPro"].ToString(), rd["xSeatErr"].ToString(), i.ToString());
                    selectStuData(rd["xHost"].ToString(), System.DateTime.Now.ToShortDateString(), i.ToString(), rd["xSeat"].ToString());
                }
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void insertLocal_SwapSeat(string xKey, string xKey2, string xKey3, string xKey4, string xKey5, string xKey6)
        {
            string connetionString = null;
            SqlConnection conn;
            SqlCommand cmd;
            string sql = null;
            SqlDataReader rd;
            //connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
            connetionString = "Data Source=DVS-33-00;Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "Insert into Local_SwapSeat(xHost ,xDate ,xSeat ,xSeatP ,xSeatB ,xPeriod) VALUES (   ";
            sql += "'" + xKey + "', ";
            sql += "'" + xKey2 + "', ";
            sql += "'" + xKey3 + "', ";
            sql += "'" + xKey4 + "', ";
            sql += "'" + xKey5 + "', ";
            sql += "'" + xKey6 + "' ";
            sql += ") ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void selectStuData(string xKey, string xKey2, string xKey3, string xKey4)
        {
            string connetionString = null;
            SqlConnection conn;
            SqlCommand cmd;
            string sql = null;
            SqlDataReader rd;
            //connetionString = "Data Source=211.75.132.163;Initial Catalog=DVD;User ID=sa;Password=";
            //connetionString = "Data Source=211.75.132.163;Initial Catalog=DVD;User ID=sa;Password=";
            //connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
            connetionString = "Data Source=DVS-33-00;Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "select * from Local_dvd_video2  ";
            sql += "where xHost = '" + xKey + "' ";
            sql += "and xDate = '" + xKey2 + "' ";
            sql += "and left(xPeriod,1) = '" + xKey3 + "' ";
            sql += "and xSeat = '" + xKey4 + "' ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            if (rd.Read())
            {
                xUpdateData(xKey, xKey2, xKey3, xKey4, rd["xid"].ToString(), rd["xStatus"].ToString());
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void xUpdateData(string xKey, string xKey2, string xKey3, string xKey4, string xKey5, string xKey6)
        {
            string connetionString = null;
            SqlConnection conn;
            SqlCommand cmd;
            string sql = null;
            SqlDataReader rd;
            //connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
            connetionString = "Data Source=DVS-33-00;Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "Update Local_SwapSeat SET  ";
            //switch (xKey6)
            //{
            //    case "請假":
            //        sql += "xNot='4', ";
            //        break;
            //    case "否":
            //        sql += "xNot='1', ";
            //        break;
            //}
            if (xKey6 == "請假")
            {
                sql += "xNot='4', ";
            }
            else
            {
                sql += "xNot='1', ";
            }
            sql += "xID = '" + xKey5 + "' ";
            sql += "where xHost = '" + xKey + "' ";
            sql += "and xDate = '" + xKey2 + "' ";
            sql += "and xPeriod = '" + xKey3 + "' ";
            sql += "and xSeat = '" + xKey4 + "' ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }

        private void Main_Load(object sender, EventArgs e)
        {

        }

    }
}
