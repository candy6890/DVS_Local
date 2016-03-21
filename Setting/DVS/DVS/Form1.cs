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
    public partial class Form1 : Form
    {
        string reHostNameNum = System.Net.Dns.GetHostEntry("LocalHost").HostName.Substring(4, 2);
        string reLocalNameNum = System.Net.Dns.GetHostEntry("LocalHost").HostName.Substring(7, 2);
        string reHostName = null, reLocalName = null, xStr = null;
        int chkDataNum = 0,insertDataNum = 0;
        //string DS = "data";
        string DS = "211.75.132.163";
        string localDS = System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString();
        public Form1()
        {
            InitializeComponent();
            xMain();
        }

        void xMain()
        {
            this.Size = new System.Drawing.Size(299, 117);
            chkDvData(System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString());
            timer1.Interval = 1000;
            timer1.Start();
            
            
            showLocalName(reLocalNameNum);
            this.Text = "三元及第視訊補課伺服器 v1.0 " + reLocalName;
            if (chkDataNum == 0)
            {
                button1.Enabled = true;
            }
        }

        void showLocalName(string xKey)
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + DS + ";Initial Catalog=STMS;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "select * from UserBranchID ";
            sql += "where number = '" + xKey +"' ";
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

        void chkDvData(string xKey)
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + xKey + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "select COUNT(*) AS E1 from Local_dvd_logistics ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            chkDataNum = Convert.ToInt32(rd["E1"].ToString());
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }

        //----------------------------------------------------------------------------------------
        //-* A-1 *- 查詢學生資料 Stms-Student >> Local_Stu_Data
        void selectFirstStu_Data()  
        {
            label6.Text = System.DateTime.Now.ToString("HH:mm:ss");
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            
            
            if (reHostNameNum == reLocalNameNum) { xStr = reHostNameNum; } else { xStr = reHostNameNum; }
            connetionString = "Data Source=" + DS + ";Initial Catalog=STMS;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            int i = 0;
            sql = "SELECT OID, BranchID, Year, ID, Name, ClassID, SeqNo, PID, BirthDate, Tel1 ";
            sql += "FROM Student ";
            sql += "WHERE year >= '" + System.DateTime.Now.AddYears(-7).ToString() + "' ";
            sql += "and status ='正常' ";
            sql += "order by oid desc ";

            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                i++;
                string xOID = rd["OID"].ToString(); string xBranchID = rd["BranchID"].ToString();
                string xYear = rd["Year"].ToString(); string xID = rd["ID"].ToString();
                string xName = rd["Name"].ToString(); string xClassID = rd["ClassID"].ToString();
                string xSeqNo  = rd["SeqNo"].ToString(); string xPID = rd["PID"].ToString();
                string xBirthDate = rd["BirthDate"].ToString(); string xTel1 = rd["Tel1"].ToString();
                listBox1.Items.Add(rd["id"].ToString());
                label7.Text = System.DateTime.Now.ToString("HH:mm:ss");
                insertFirstStu_Data(xOID, xBranchID, xYear, xID, xName, xClassID, xSeqNo, xPID, xBirthDate, xTel1);
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
            label5.Text = i.ToString();
        }
        void insertFirstStu_Data(string xKey, string xKey2, string xKey3, string xKey4, string xKey5, string xKey6, string xKey7, string xKey8, string xKey9, string xKey10)
        { 
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            //connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "Insert into Local_Stu_Data(serial, xBranchID, xYear, xID, xName, xClassID, xSeqNo, xPID, xBirthDate,xTel) VALUES (";
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
        void deleteFirstStu_Data()
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            //connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "delete from Local_Stu_Data ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }

        //-* A-2 *- 查詢學生資料 Stms-ClassIn_DV >> Local_ClassIn_DV
        void selectFirstClassIn_DV()
        {
            label6.Text = System.DateTime.Now.ToString("HH:mm:ss");
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            
            
            if (reHostNameNum == reLocalNameNum) { xStr = reHostNameNum; } else { xStr = reHostNameNum; }
            connetionString = "Data Source=" + DS + ";Initial Catalog=STMS;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            int i = 0;
            sql = "SELECT * ";
            sql += "FROM ClassIn_DV ";
            sql += "where xYear >= '" + System.DateTime.Now.AddYears(-7).ToString() + "' ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                i++;
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

                listBox2.Items.Add(rd["xdNumber"].ToString());
                label7.Text = System.DateTime.Now.ToString("HH:mm:ss");
                insertFirstClassIn_DV(xeSerial, xeHost, xeYear, xeClassID, xeSeqNo , xeSubjectID, xedYear, xedCategoryName, xedClassName, xedSubClassName, xedSubClassID, xedSubjectName, xedNumber, xeCreateDate);
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
            label5.Text = i.ToString();
        }
        void insertFirstClassIn_DV(string xKey, string xKey2, string xKey3, string xKey4, string xKey5, string xKey6, string xKey7, string xKey8, string xKey9, string xKey10, string xKey11, string xKey12, string xKey13, string xKey14)
        {
            string connetionString = null, sql = null;
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
        void deleteFirstClassIn_DV()
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            //connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "delete from Local_ClassIn_DV ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }

        //-* A-3 *- 查詢學生資料 Stms-ClassIn_DV >> ClassIn_DV_OutSide
        void selectFirstClassIn_DV_OutSide()
        {
            label6.Text = System.DateTime.Now.ToString("HH:mm:ss");
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            
            
            if (reHostNameNum == reLocalNameNum) { xStr = reHostNameNum; } else { xStr = reHostNameNum; }
            connetionString = "Data Source=" + DS + ";Initial Catalog=STMS;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            int i = 0;
            sql = "SELECT * ";
            sql += "FROM ClassIn_DV_OutSide ";
            sql += "where xdyear >= '" + System.DateTime.Now.AddYears(-7).ToString() + "' ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                i++;
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

                listBox3.Items.Add(rd["xID"].ToString());
                label7.Text = System.DateTime.Now.ToString("HH:mm:ss");
                insertFirstClassIn_DV_OutSide(xeSerial, xeID, xeHost, xedYear, xedCategoryName, xedClassName, xedSubClassName, xedSubClassID, xedSubjectName, xedNumber, xeCreateDate);
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
            label5.Text = i.ToString();
        }
        void insertFirstClassIn_DV_OutSide(string xKey, string xKey2, string xKey3, string xKey4, string xKey5, string xKey6, string xKey7, string xKey8, string xKey9, string xKey10, string xKey11)
        {
            string connetionString = null, sql = null;
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
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            //connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "delete from Local_ClassIn_DV_OutSide ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }

        //--------------------------------------------------------------------------------------------------------------------------------------
        private void firstBtA_Click(object sender, EventArgs e)
        {
            selectFirstStu_Data();
            selectFirstClassIn_DV();
            selectFirstClassIn_DV_OutSide();
            selectFirstLocal_ClassCategory();
            selectFirstLocal_Class();
            selectFirstLocal_SubClass();
            selectFirstLocal_Subject();
            selectFirstLocal_Participation();

            firstBtA.Enabled = false;
        }
        //-* A-4 *- 類別 Stms-ClassCategory >> Local_ClassCategory
        void selectFirstLocal_ClassCategory()
        {
            label6.Text = System.DateTime.Now.ToString("HH:mm:ss");
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            
            
            if (reHostNameNum == reLocalNameNum) { xStr = reHostNameNum; } else { xStr = reHostNameNum; }
            connetionString = "Data Source=" + DS + ";Initial Catalog=STMS;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "SELECT * ";
            sql += "FROM ClassCategory ";
            sql += "WHERE year >= '" + System.DateTime.Now.AddYears(-7).ToString() + "' ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                insertFirstLocal_ClassCategory(rd["OID"].ToString(), rd["BranchID"].ToString(), rd["Year"].ToString(), rd["Name"].ToString());
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void insertFirstLocal_ClassCategory(string xKey, string xKey2, string xKey3, string xKey4)
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            //connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
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
        void deleteFirstLocal_ClassCategory()
        {

            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            //connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
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

        //-* A-5 *- 班級 Stms-Class >> Local_Class
        void selectFirstLocal_Class()
        {
            label6.Text = System.DateTime.Now.ToString("HH:mm:ss");
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            
            
            if (reHostNameNum == reLocalNameNum) { xStr = reHostNameNum; } else { xStr = reHostNameNum; }
            connetionString = "Data Source=" + DS + ";Initial Catalog=STMS;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "SELECT * ";
            sql += "FROM Class ";
            sql += "WHERE year >= '" + System.DateTime.Now.AddYears(-7).ToString() + "' ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                insertFirstLocal_Class(rd["OID"].ToString(), rd["BranchID"].ToString(), rd["Year"].ToString(), rd["ID"].ToString(), rd["Name"].ToString(), rd["Deadline"].ToString(), rd["CategoryOID"].ToString());
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void insertFirstLocal_Class(string xKey, string xKey2, string xKey3, string xKey4, string xKey5, string xKey6, string xKey7)
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            //connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "Insert into Local_Class(serial, xHost, xYear, xID, xName,xDeadline ,xCategoryOID) VALUES (";
            sql += "'" + xKey + "', ";
            sql += "'" + xKey2 + "', ";
            sql += "'" + xKey3 + "', ";
            sql += "'" + xKey4 + "', ";
            sql += "'" + xKey5 + "', ";
            if (xKey6 != "") { sql += "'" + Convert.ToDateTime(xKey6).ToShortDateString() + "', "; } else { sql += "NULL, "; }
            sql += "'" + xKey7 + "') ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void deleteFirstLocal_Class()
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            //connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "delete from Local_Class ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }

        //-* A-6 *- 子班級 Stms-SubClass >> Local_SubClass 
        void selectFirstLocal_SubClass()
        {
            label6.Text = System.DateTime.Now.ToString("HH:mm:ss");
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            
            
            if (reHostNameNum == reLocalNameNum) { xStr = reHostNameNum; } else { xStr = reHostNameNum; }
            connetionString = "Data Source=" + DS + ";Initial Catalog=STMS;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "SELECT * ";
            sql += "FROM Subclass ";
            sql += "WHERE year >= '" + System.DateTime.Now.AddYears(-7).ToString() + "' ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                insertFirstLocal_SubClass(rd["OID"].ToString(), rd["BranchID"].ToString(), rd["Year"].ToString(), rd["Name"].ToString(), rd["ClassID"].ToString(), rd["Deadline"].ToString(), rd["SeqNo"].ToString(), rd["ClassOID"].ToString());
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void insertFirstLocal_SubClass(string xKey, string xKey2, string xKey3, string xKey4, string xKey5, string xKey6, string xKey7, string xKey8)
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            //connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "Insert into Local_SubClass(serial, xHost, xYear, xName, xClassID, xDeadline, xSeqNo, xClassOID) VALUES (";
            sql += "'" + xKey + "', ";
            sql += "'" + xKey2 + "', ";
            sql += "'" + xKey3 + "', ";
            sql += "'" + xKey4 + "', ";
            sql += "'" + xKey5 + "', ";
            if (xKey6 != "") { sql += "'" + Convert.ToDateTime(xKey6).ToShortDateString() + "', "; } else { sql += "NULL, "; }
            sql += "'" + xKey7 + "', ";
            sql += "'" + xKey8 + "') ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void deleteFirstLocal_SubClass()
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            //connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "delete from Local_SubClass ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }

        //-* A-7 *- 科目 Stms-Subject >> Local_Subject
        void selectFirstLocal_Subject()
        {
            label6.Text = System.DateTime.Now.ToString("HH:mm:ss");
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            
            
            if (reHostNameNum == reLocalNameNum) { xStr = reHostNameNum; } else { xStr = reHostNameNum; }
            connetionString = "Data Source=" + DS + ";Initial Catalog=STMS;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "SELECT * ";
            sql += "FROM Subject ";
            sql += "WHERE year >= '" + System.DateTime.Now.AddYears(-7).ToString() + "' ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                insertFirstLocal_Subject(rd["OID"].ToString(), rd["BranchID"].ToString(), rd["Year"].ToString(), rd["ID"].ToString(), rd["Name"].ToString());
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void insertFirstLocal_Subject(string xKey, string xKey2, string xKey3, string xKey4, string xKey5)
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            //connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
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
        void deleteFirstLocal_Subject()
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            //connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "delete from Local_Subject ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }

        //-* A-8 *- 學生資料 Stms-Participation >> Local_Association
        void selectFirstLocal_Participation()
        {
            label6.Text = System.DateTime.Now.ToString("HH:mm:ss");
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            
            
            if (reHostNameNum == reLocalNameNum) { xStr = reHostNameNum; } else { xStr = reHostNameNum; }
            connetionString = "Data Source=" + DS + ";Initial Catalog=STMS;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "SELECT * ";
            sql += "FROM Participation ";
            sql += "WHERE year >= '" + System.DateTime.Now.AddYears(-7).ToString() + "' ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                insertFirstLocal_Participation(rd["OID"].ToString(), rd["BranchID"].ToString(), rd["Year"].ToString(), rd["StudentID"].ToString(), rd["SubjectID"].ToString());
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void insertFirstLocal_Participation(string xKey, string xKey2, string xKey3, string xKey4, string xKey5)
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            //connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
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
        void deleteFirstLocal_Participation()
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            //connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "delete from Local_Participation ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }

        //----------------------------------------------------------------------------------------
        //-* B-1 *- DVD_logistics
        void selectFirstDVD_Logistics()
        {
            label6.Text = System.DateTime.Now.ToString("HH:mm:ss");
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            
            
            if (reHostNameNum == reLocalNameNum) { xStr = reHostNameNum; } else { xStr = reHostNameNum; }
            connetionString = "Data Source=" + DS + ";Initial Catalog=STMS;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            int i = 0;
            sql = "select serial,DVD編號, 上課日期,當科堂次 from DVD_logistics ";
            sql += "where SUBSTRING(DVD編號, 3, 2) = '" + xStr + "' ";
            sql += "order by DVD編號 desc";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                i++;
                listBox1.Items.Add(rd["DVD編號"].ToString());
                insertFirstDVD_Logistics(Convert.ToInt32(rd["serial"].ToString()), rd["DVD編號"].ToString(), rd["當科堂次"].ToString(), Convert.ToDateTime(rd["上課日期"]));
                label7.Text = System.DateTime.Now.ToString("HH:mm:ss");
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
            label5.Text = i.ToString();
        }
        void insertFirstDVD_Logistics(int eKey, string xKey, string xKey2, DateTime dKey)
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
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
        void deleteFirstDVD_Logistics()
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "delete from Local_dvd_logistics ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }

        //-* B-2 *- DVD_logistics2
        void selectFirstDVD_Logistics2()
        {
            int i = 0;
            label6.Text = System.DateTime.Now.ToString("HH:mm:ss");
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            
            
            if (reHostNameNum == reLocalNameNum) { xStr = reHostNameNum; } else { xStr = reHostNameNum; }
            connetionString = "Data Source=" + DS + ";Initial Catalog=STMS;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "select serial, DVD編號, 上課講義名稱, 發放講義名稱, 發放講義編號, 章節進度, 頁數進度, 補充講義頁數 from DVD_logistics2 ";
            sql += "where SUBSTRING(DVD編號, 3, 2) = '" + xStr + "' ";
            sql += "order by DVD編號 desc";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                i++;
                listBox2.Items.Add(rd["DVD編號"].ToString());
                label7.Text = System.DateTime.Now.ToString("HH:mm:ss");
                insertFirstDVD_Logistics2(Convert.ToInt32(rd["serial"].ToString()), rd["DVD編號"].ToString(), rd["上課講義名稱"].ToString(), rd["發放講義名稱"].ToString(), rd["發放講義編號"].ToString(), rd["章節進度"].ToString(), rd["頁數進度"].ToString(), rd["補充講義頁數"].ToString());
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
            label5.Text = i.ToString();
        }
        void insertFirstDVD_Logistics2(int eKey, string xKey, string xKey2, string xKey3, string xKey4, string xKey5, string xKey6, string xKey7)
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
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
        void deleteFirstDVD_Logistics2()
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "delete from Local_dvd_logistics2 ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }

        //----------------------------------------------------------------------------------------
        //-* C-1 *- DVD_video2
        void selectFirstDVD_video2()
        {
            label6.Text = System.DateTime.Now.ToString("HH:mm:ss");
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;


            if (reHostNameNum == reLocalNameNum) { xStr = reHostNameNum; } else { xStr = reLocalNameNum; }
            connetionString = "Data Source=" + DS + ";Initial Catalog=DVD;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            int i = 0;
            sql = "select serial, id, 地點, 日期, 時段, 已看, 座位, xInTime, xOutTime from DVD_video2 ";
            sql += "where 地點 = '" + xStr + "' ";
            sql += "and LEFT(id, 2) >= '" + System.DateTime.Now.AddYears(-7).ToString().Substring(2,2) + "' ";
            sql += "order by serial desc";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                i++;
                listBox1.Items.Add(rd["日期"].ToString() + "-" + rd["時段"].ToString());
                label7.Text = System.DateTime.Now.ToString("HH:mm:ss");
                insertFirstDVD_video2(Convert.ToInt32(rd["serial"].ToString()), rd["id"].ToString(), rd["地點"].ToString(), rd["日期"].ToString(), rd["時段"].ToString(), rd["已看"].ToString(), rd["座位"].ToString(), rd["xInTime"].ToString(), rd["xOutTime"].ToString());
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
            label5.Text = i.ToString();
        }
        void insertFirstDVD_video2(int eKey, string xKey, string xKey2, string xKey3, string xKey4, string xKey5, string xKey6, string xKey7, string xKey8)
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "INSERT INTO Local_dvd_video2 (serial, xid, xHost, xDate, xPeriod, xStatus, xSeat, xInTime, xOutTime) VALUES (";
            sql += "'" + eKey + "', ";    //serial
            sql += "'" + xKey + "', ";    //xid
            sql += "'" + xKey2 + "', ";    //xHost
            sql += "'" + Convert.ToDateTime(xKey3).ToShortDateString() + "', ";    //xDate
            sql += "'" + xKey4 + "', ";    //xPeriod
            sql += "'" + xKey5 + "', ";    //xCourseDate
            sql += "'" + xKey6 + "', ";    //xCourseDate
            sql += "'" + xKey7 + "', ";    //xCourseDate
            sql += "'" + xKey8 + "') ";    //xCourseDate
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void deleteFirstDVD_video2()
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
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

        //-* C-3 *- DVD_video3
        void selectFirstDVD_video3()
        {
            int i = 0;
            label6.Text = System.DateTime.Now.ToString("HH:mm:ss");
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;


            if (reHostNameNum == reLocalNameNum) { xStr = reHostNameNum; } else { xStr = reLocalNameNum; }
            connetionString = "Data Source=" + DS + ";Initial Catalog=DVD;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "select * from DVD_video3 ";
            sql += "where 地點 = '" + xStr + "' ";
            sql += "order by serial ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                i++;
                listBox3.Items.Add(rd["星期"].ToString() + "-" + rd["時段"].ToString());
                label7.Text = System.DateTime.Now.ToString("HH:mm:ss");
                insertFirstDVD_video3(rd["地點"].ToString(), rd["星期"].ToString(), rd["時段"].ToString());
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
            label5.Text = i.ToString();
        }
        void insertFirstDVD_video3(string xKey, string xKey2, string xKey3)
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
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
        void deleteFirstDVD_video3()
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
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

        //-* C-2 *- DVD_video5
        void selectFirstDVD_video5()
        {
            int i = 0;
            label6.Text = System.DateTime.Now.ToString("HH:mm:ss");
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            
            
            if (reHostNameNum == reLocalNameNum) { xStr = reHostNameNum; } else { xStr = reHostNameNum; }
            connetionString = "Data Source=" + DS + ";Initial Catalog=DVD;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "select * from DVD_video5 ";
            sql += "where SUBSTRING(id, 3, 2) = '" + xStr + "' ";
            sql += "and LEFT(id, 2) >= '" + System.DateTime.Now.AddYears(-7).ToString().Substring(2, 2) + "' ";
            sql += "order by serial ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                i++;
                listBox2.Items.Add(rd["開通時間"].ToString() + "-" + rd["id"].ToString());
                label7.Text = System.DateTime.Now.ToString("HH:mm:ss");
                insertFirstDVD_video5(Convert.ToInt32(rd["serial"].ToString()), rd["開通時間"].ToString(), rd["開通"].ToString(), rd["id"].ToString(), Convert.ToInt32(rd["免費次數"].ToString()));
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
            label5.Text = i.ToString();
        }
        void insertFirstDVD_video5(int eKey, string xKey, string xKey2, string xKey3, int eKey2)
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
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
            conn.Close();
        }
        void deleteFirstDVD_video5()
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
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

        void selectFirstDvNum()
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            
            
            if (reHostNameNum == reLocalNameNum) { xStr = reHostNameNum; } else { xStr = reHostNameNum; }
            connetionString = "Data Source=" + DS + ";Initial Catalog=STMS;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            int i = 0;
            sql = "select serial,DVD編號, 上課日期,當科堂次 from DVD_logistics ";
            sql += "where SUBSTRING(DVD編號, 3, 2) = '" + xStr + "' ";
            sql += "order by DVD編號 ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                i++;
                insertFirstDvNum(Convert.ToInt32(rd["serial"].ToString()), rd["DVD編號"].ToString(), rd["當科堂次"].ToString(), Convert.ToDateTime(rd["上課日期"]));
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
            insertDataNum = i;
        }

        //DVD_logistics
        void insertFirstDvNum(int eKey, string xKey, string xKey2, DateTime dKey)
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
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
        
        void selectFirstDvs()
        {
            int i = 0;
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            
            
            if (reHostNameNum == reLocalNameNum) { xStr = reHostNameNum; } else { xStr = reHostNameNum; }
            connetionString = "Data Source=" + DS + ";Initial Catalog=STMS;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "select serial, DVD編號, 上課講義名稱, 發放講義名稱, 發放講義編號, 章節進度, 頁數進度, 補充講義頁數 from DVD_logistics2 ";
            sql += "where SUBSTRING(DVD編號, 3, 2) = '" + xStr + "' ";
            sql += "order by DVD編號 ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                i++;
                listBox2.Items.Add(rd["DVD編號"].ToString());
                //insertFirstDvs(Convert.ToInt32(rd["serial"].ToString()), rd["DVD編號"].ToString(), rd["上課講義名稱"].ToString(), rd["發放講義名稱"].ToString(), rd["發放講義編號"].ToString(), rd["章節進度"].ToString(), rd["頁數進度"].ToString(), rd["補充講義頁數"].ToString());
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }

        void insertFirstDvs(int eKey, string xKey, string xKey2,string xKey3, string xKey4,string xKey5, string xKey6, string xKey7)
        { 
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
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

        private void button1_Click(object sender, EventArgs e)
        {
            firstPL.Visible = true;
            secPL.Visible = false;
            this.Size = new System.Drawing.Size(299, 248);
            firstPL.Location = new System.Drawing.Point(12, 88);
            button1.Enabled = false;
            button3.Enabled = false;
            switch (button1.Text.ToString())
            {
                case "第一次匯入資料[開啟視窗]":
                    button1.Text = "第一次匯入資料[關閉視窗]";
                    firstPL.Visible = true;
                    button3.Enabled = true;
                    break;
                case "第一次匯入資料[關閉視窗]":
                    button1.Text = "第一次匯入資料[開啟視窗]";
                    firstPL.Visible = false;
                    button3.Enabled = true;
                    break;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Size = new System.Drawing.Size(299, 248);
            firstPL.Visible = false;
            secPL.Visible = true;
            secPL.Location = new System.Drawing.Point(12, 88);
            button1.Enabled = false;
            button3.Enabled = false;
            button1.Enabled = true;
            showCBa();
            showCBb();

        }

        void showCBa()
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            
            
            comboBox2.Items.Clear();
            comboBox2.Items.Add("請選擇");
            connetionString = "Data Source=" + DS + ";Initial Catalog=STMS;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "SELECT DISTINCT LEFT(DVD編號, 2) AS E1 FROM DVD_logistics ";
            sql += "WHERE SUBSTRING(DVD編號, 3, 2 )= '" + reHostNameNum + "' order by E1 desc ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                comboBox2.Items.Add(rd["E1"].ToString());
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
            comboBox2.SelectedIndex = 0;
        }

        void showCBb()
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            
            
            comboBox1.Items.Clear();
            comboBox1.Items.Add("請選擇");
            connetionString = "Data Source=" + DS + ";Initial Catalog=STMS;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "SELECT DISTINCT LEFT(DVD編號, 2) AS E1 FROM DVD_logistics2 ";
            sql += "WHERE SUBSTRING(DVD編號, 3, 2 )= '" + reHostNameNum + "' order by E1 desc ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                comboBox1.Items.Add(rd["E1"].ToString());
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
            comboBox1.SelectedIndex = 0;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            label1.Text = System.DateTime.Now.ToString("HH:mm:ss");
            if (label1.Text == label2.Text)
            {

            }
            else
            {
                //┤┴├┬┼ ○● ∣―
                if (label3.Text == "○")
                {
                    label3.Text = "●";
                }
                else
                {
                    label3.Text = "○";
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            
            
            if (reHostNameNum == reLocalNameNum) { xStr = reHostNameNum; } else { xStr = reHostNameNum; }
            connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "select * from Local_dvd_logistics ";
            sql += "order by serial ";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            if (rd.Read())
            {

            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }

        private void firstBtB_Click(object sender, EventArgs e)
        {
            selectFirstDVD_Logistics();
            selectFirstDVD_Logistics2();
            firstBtB.Enabled = false;
        }

        private void firstBtC_Click(object sender, EventArgs e)
        {
            selectFirstDVD_video2();
            //selectFirstDVD_video5();
            selectFirstDVD_video3();
            firstBtC.Enabled = false;
        }

        private void deleteBtA_Click(object sender, EventArgs e)
        {
            DialogResult myResult = MessageBox.Show("刪除資料(是＆否)?", "你確定"+ deleteBtA.Text , MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (myResult == DialogResult.Yes)
            {
                firstPL.Enabled = false;
                deleteFirstStu_Data();
                deleteFirstClassIn_DV();
                deleteFirstClassIn_DV_OutSide();
                deleteFirstLocal_ClassCategory();
                deleteFirstLocal_Class();
                deleteFirstLocal_SubClass();
                deleteFirstLocal_Subject();
                deleteFirstLocal_Participation();
                firstPL.Enabled = true;
            }
        }

        private void deleteBtB_Click(object sender, EventArgs e)
        {
            DialogResult myResult = MessageBox.Show("刪除資料(是＆否)?", "你確定" + deleteBtB.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (myResult == DialogResult.Yes)
            {
                firstPL.Enabled = false;
                deleteFirstDVD_Logistics();
                deleteFirstDVD_Logistics2();
                firstPL.Enabled = true;
            }
        }

        private void deleteBtC_Click(object sender, EventArgs e)
        {
            DialogResult myResult = MessageBox.Show("刪除資料(是＆否)?", "你確定" + deleteBtC.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (myResult == DialogResult.Yes)
            {
                firstPL.Enabled = false;
                deleteFirstDVD_video2();
                deleteFirstDVD_video3();
                deleteFirstDVD_video5();
                firstPL.Enabled = true;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Button bt = (Button)sender;
            switch (bt.Text)
            {
                case "查詢":
                    if (comboBox1.SelectedIndex.ToString() != "0")
                    {
                        viewDVD_Logistics2(comboBox1.SelectedItem.ToString());
                        bt.Text = "匯入";
                    }
                    break;
                case "匯入":
                    if (comboBox1.SelectedIndex.ToString() != "0")
                    {
                        deleteDVD_Logistics2(comboBox1.SelectedItem.ToString());
                        selectDVD_Logistics2(comboBox1.SelectedItem.ToString());
                        bt.Text = "查詢";
                    }
                    break;
            }
        }

        void viewDVD_Logistics2(string xKey)
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            
            
            connetionString = "Data Source=" + DS + ";Initial Catalog=STMS;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "SELECT COUNT(*) AS E1 FROM DVD_logistics2 ";
            sql += "WHERE SUBSTRING(DVD編號, 3, 2) = '" + reHostNameNum + "' ";
            sql += "AND (SUBSTRING(DVD編號, 1, 2) = '" + xKey + "')";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            if (rd.Read())
            {
                label13.Text = rd["E1"].ToString();
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
            
        }
        void selectDVD_Logistics2(string xKey)
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            
            
            connetionString = "Data Source=" + DS + ";Initial Catalog=STMS;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "SELECT * FROM DVD_logistics2 ";
            sql += "WHERE SUBSTRING(DVD編號, 3, 2) = '" + reHostNameNum + "' ";
            sql += "AND (SUBSTRING(DVD編號, 1, 2) = '" + xKey + "')";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            int i = 0;
            while (rd.Read())
            {
                i++;
                insertDVD_Logistics2(Convert.ToInt32(rd["serial"].ToString()), rd["DVD編號"].ToString(), rd["上課講義名稱"].ToString(), rd["發放講義名稱"].ToString(), rd["發放講義編號"].ToString(), rd["章節進度"].ToString(), rd["頁數進度"].ToString(), rd["補充講義頁數"].ToString());

            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
            label14.Text = i.ToString();
        }
        void insertDVD_Logistics2(int eKey, string xKey, string xKey2, string xKey3, string xKey4, string xKey5, string xKey6, string xKey7)
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
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
        void deleteDVD_Logistics2(string xKey)
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            
            
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "delete FROM Local_dvd_logistics2 ";
            sql += "WHERE SUBSTRING(xDvNum, 3, 2) = '" + reHostNameNum + "' ";
            sql += "AND (SUBSTRING(xDvNum, 1, 2) = '" + xKey + "')";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            button4.Text = "查詢";
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            button5.Text = "查詢";
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Button bt = (Button)sender;
            switch (bt.Text)
            {
                case "查詢":
                    if (comboBox2.SelectedIndex.ToString() != "0")
                    {
                        viewDVD_Logistics(comboBox2.SelectedItem.ToString());
                        bt.Text = "匯入";
                    }
                    break;
                case "匯入":
                    if (comboBox2.SelectedIndex.ToString() != "0")
                    {
                        deleteDVD_Logistics(comboBox2.SelectedItem.ToString());
                        selectDVD_Logistics(comboBox2.SelectedItem.ToString());
                        bt.Text = "查詢";
                    }
                    break;
            }
        }

        void viewDVD_Logistics(string xKey)
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            
            
            connetionString = "Data Source=" + DS + ";Initial Catalog=STMS;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "SELECT COUNT(*) AS E1 FROM DVD_logistics ";
            sql += "WHERE SUBSTRING(DVD編號, 3, 2) = '" + reHostNameNum + "' ";
            sql += "AND (SUBSTRING(DVD編號, 1, 2) = '" + xKey + "')";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            if (rd.Read())
            {
                label19.Text = rd["E1"].ToString();
            }
            rd.Close();
            cmd.Dispose();
            conn.Close();

        }
        void selectDVD_Logistics(string xKey)
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            
            
            connetionString = "Data Source=" + DS + ";Initial Catalog=STMS;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "SELECT * FROM DVD_logistics ";
            sql += "WHERE SUBSTRING(DVD編號, 3, 2) = '" + reHostNameNum + "' ";
            sql += "AND (SUBSTRING(DVD編號, 1, 2) = '" + xKey + "')";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            int i = 0;
            while (rd.Read())
            {
                i++;
                insertDVD_Logistics(Convert.ToInt32(rd["serial"].ToString()), rd["DVD編號"].ToString(), rd["當科堂次"].ToString(), rd["上課日期"].ToString());

            }
            rd.Close();
            cmd.Dispose();
            conn.Close();
            label17.Text = i.ToString();
        }
        void insertDVD_Logistics(int eKey, string xKey, string xKey2, string xKey3)
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            connetionString = "Data Source=" + System.Net.Dns.GetHostEntry("LocalHost").HostName.ToString() + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "Insert into Local_dvd_logistics(serial, xDvNum, xCourseNum, xCourseDate) VALUES (";
            sql += "'" + eKey + "', ";    //serial
            sql += "'" + xKey + "', ";    //xDvNum
            sql += "'" + xKey2 + "', ";    //xCourseNum
            sql += "'" + Convert.ToDateTime(xKey3).ToShortDateString() + "') ";    //xCourseDate
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
        void deleteDVD_Logistics(string xKey)
        {
            string connetionString = null, sql = null;
            SqlConnection conn; SqlCommand cmd; SqlDataReader rd;
            
            
            connetionString = "Data Source=" + localDS + ";Initial Catalog=sourceData;User ID=sa;Password=";
            conn = new SqlConnection(connetionString);
            conn.Open();
            sql = "delete FROM Local_dvd_logistics ";
            sql += "WHERE SUBSTRING(xDvNum, 3, 2) = '" + reHostNameNum + "' ";
            sql += "AND (SUBSTRING(xDvNum, 1, 2) = '" + xKey + "')";
            cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            rd.Close();
            cmd.Dispose();
            conn.Close();
        }
    }
}
