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
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Xml;
using System.Data.OleDb;


using System.Management;
using System.Threading.Tasks;
using DAL;



using System.Drawing.Drawing2D;
using System.Drawing.Imaging;



using System.Windows.Forms.DataVisualization.Charting;


//using System.Reflection; // 引用这个才能使用Missing字段  using Excel;



namespace Array_Eyes
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            
        }
         //SqlConnection conn;
         public SqlConnection conn = new SqlConnection(SqlHelper.conStr);
       // string str = "Data Source=10.120.20.47;Initial Catalog=db_15;Integrated Security=True";
        string str = "server=DESKTOP-3UCF97E\\SQLEXPRESS;database=db_15;uid=sa;pwd=a123456.";
         //string str = "server=10.120.20.47;database=db_15;uid=sa;pwd=a123456.";
         DataTable dt = new DataTable();
         DataTable dt1 = new DataTable();
         DataTable dt2 = new DataTable();

         DataTable dt3 = new DataTable();
         public static int BMDT_Flag=0;
         public static int count1 = 0, count2 = 0, count3 = 0, count4 = 0;
         bool Button1_Flag = false, Button2_Flag = false, Button3_Flag = false, Button4_Flag = false, Button5_Flag = false, Button10_Flag = false;
         public static int Flag = 0;

         DataBase database = new DataBase();

         string[] x = new string[1000];
         double[] y = new double[1000];

         public static string Product_Type = string.Empty;
        

         private void Form1_Load(object sender, EventArgs e)
         {
             try
             {
                 string FilePath = string.Empty;
                 int Row = 0;
                 FilePath = @"D:\BMDT\Q_SingleCount.txt";
                 conn = new SqlConnection(str);
                 conn.Open();

                 radioButton1.Checked = true;
                 radioButton2.Checked = false;

                 if (conn.State == ConnectionState.Open)
                     label3.Text = "数据库连接\n状态：成功";
                 else
                     label3.Text = "数据库连接\n状态：失败";


                 DataTable dt4 = new DataTable();
                 dt4 = database.getDs("select ParamterID from  Parameter_Panel ").Tables[0];

                // MessageBox.Show(dt4.Rows.Count.ToString());
                  for (int i = 0; i < dt4.Rows.Count; i++)
                  {
                      listBox1.Items.Add(dt4.Rows[i][0].ToString());
                     
                  }
                  if (listBox1.Items.Count > 0)
                  
                      listBox1.SelectedIndex = 0;
                    
                  

                 if (!Directory.Exists("D:\\BMDT"))
                     Directory.CreateDirectory("D:\\BMDT");
                 if (File.Exists("D:\\BMDT\\Q_SingleCount.txt"))
                 {
                
                     dt.Columns.Add("Q_Name");
                     dt.Columns.Add("SingleCount");

                     dataGridView4.DataSource = dt;
                     foreach (string str1 in File.ReadAllLines(FilePath, Encoding.Default))
                     {
                         dt.Rows.Add("");
                         dataGridView4.Rows[Row].Cells[0].Value = StringSplit(str1, 1);
                         dataGridView4.Rows[Row].Cells[1].Value = StringSplit(str1, 2);
                         Row++;
                     }

                 }
                

                 ReadFromSQL("select * from Q_SingleCount ", "Q_SingleCount", 3);

                 SendToSQL("delete from Table_1");
                 SendToSQL("delete from Table_2");
                 SendToSQL("delete from EQP_Info");

                 FilePath = "D:\\BMDT\\code.txt";
                 Row = 0;
                 if (File.Exists("D:\\BMDT\\code.txt"))
                 {

                     dt3.Columns.Add("F1");
                     dt3.Columns.Add("F2");

                     dataGridView9.DataSource = dt3;
                     foreach (string str1 in File.ReadAllLines(FilePath, Encoding.Default))
                     {
                         dt3.Rows.Add("");
                         dataGridView9.Rows[Row].Cells[0].Value = StringSplit(str1, 1);
                         dataGridView9.Rows[Row].Cells[1].Value = StringSplit(str1, 2);
                         Row++;
                     }

                 }


                 ReadFromSQL("select * from code ", "code", 8);

                 dataGridView7.Visible = false;

                 chart1.Visible = false;

             }
             catch
             {
                 MessageBox.Show("窗体加载失败");
             }

                   
         }


         public string  StringSplit(string ImageName,int e)
         {
             string[] aa = new string[100];
             char[] separator = { ',' };
             aa = ImageName.Split(separator);
             switch (e)
             {
                 case 1: return aa[0];
                 case 2: return aa[1];
                 default: return aa[0]; break;

             }

         }

         public void ReadFromSQL(String SQL,String Table,int e)
         {

             conn = new SqlConnection(str);
             conn.Open();
             SqlCommand cm = new SqlCommand();
             cm.CommandTimeout = 0;
             cm.Connection = conn;
             cm.CommandText = SQL;


             SqlDataAdapter da = new SqlDataAdapter();
             da.SelectCommand = cm;
             DataSet ds = new DataSet();
             da.Fill(ds, Table);
             switch (e)
             {
                 case 1: dataGridView1.DataSource = ds.Tables[Table]; MessageBox.Show("共有"+dataGridView1.Rows.Count.ToString()+"条数据！"); break;
                 case 2: dataGridView2.DataSource = ds.Tables[Table];break;
                 case 3: dataGridView3.DataSource = ds.Tables[Table]; break;
                 case 4: dataGridView7.DataSource = ds.Tables[Table]; break;
                 case 5: dataGridView5.DataSource = ds.Tables[Table]; break;
                // case 6: dataGridView6.DataSource = ds.Tables[Table]; break;
                 case 8: dataGridView8.DataSource = ds.Tables[Table]; break;

                 default:   break;
             }
            
             conn.Close();
         
         
         }
         public void SendToSQL(String SQL)
         {
             conn = new SqlConnection(str);
             conn.Open();
             SqlCommand cmd = new SqlCommand(SQL, conn);
             //cmd.CommandTimeout = 1800;
             cmd.CommandTimeout = 0;
             cmd.ExecuteNonQuery();
             conn.Close();
         }





         #region /* 函数getData(int t) */
        
         public DataSet getData(int t=0)
         {
             //打开文件
             string sql_select=string.Empty;


             OpenFileDialog file = new OpenFileDialog();
             file.Filter = "Excel(*.xlsx)|*.xlsx|Excel(*.xls)|*.xls|(*.csv)|*.csv";
             file.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
             file.Multiselect = false;
             if (file.ShowDialog() == DialogResult.Cancel)
                 return null;
             //判断文件后缀
             var path = file.FileName;
             string fileSuffix = System.IO.Path.GetExtension(path);
             if (string.IsNullOrEmpty(fileSuffix))
                 return null;
             using (DataSet ds = new DataSet())
             {
                 //判断Excel文件是2003版本还是2007版本
                 string connString = "";
                 if (fileSuffix == ".xls")
                     connString = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + path + ";" + ";Extended Properties=\"Excel 8.0;HDR=YES;IMEX=1\"";
                 else
                     connString = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + path + ";" + ";Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1\"";
                 //读取文件

                 if (t == 1)
                     sql_select = " SELECT ID,CREATETIME,REASONCODE1,PROCESSID FROM [Sheet1$] ";
                 else if (t == 2)
                     sql_select = " SELECT PANELID,EVENTTIME,DESCRIPTION,PROCESSID FROM [Grid$] ";
                 else if (t == 3)
                     sql_select = " SELECT Operation,Lot_ID,EQP_ID,Event_Time FROM [报表1$] ";

               
                 using (OleDbConnection conn = new OleDbConnection(connString))
                 using (OleDbDataAdapter cmd = new OleDbDataAdapter(sql_select, conn))
                 {
                     conn.Open();
                     cmd.Fill(ds);
                 }
                 if (ds == null || ds.Tables.Count <= 0) return null;
                 return ds;
             }
         }
         #endregion

         #region /* 数据导出到CSV */
         public void ExportCSV()
          {
            if (dataGridView4.Rows.Count == 0)
            {
                MessageBox.Show("没有数据可导出!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "CSV files (*.csv)|*.csv";
            saveFileDialog.FilterIndex = 0;
            saveFileDialog.RestoreDirectory = true;
            saveFileDialog.CreatePrompt = true;
            saveFileDialog.FileName = null;
            saveFileDialog.Title = "保存";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                Stream stream = saveFileDialog.OpenFile();
                StreamWriter sw = new StreamWriter(stream, System.Text.Encoding.GetEncoding(-0));
                string strLine = "";
                try
                {
                    for (int i = 0; i < dataGridView2.ColumnCount; i++)
                    {
                        if (i > 0)
                            strLine += ",";

                        strLine += dataGridView2.Columns[i].HeaderText;

                    }

                    strLine.Remove(strLine.Length - 1);

                    sw.WriteLine(strLine);

                    strLine = "";

                    //表的内容

                    for (int j = 0; j < dataGridView2.Rows.Count; j++)
                    {

                        strLine = "";

                        int colCount = dataGridView2.Columns.Count;

                        for (int k = 0; k < colCount; k++)
                        {

                            if (k > 0 && k < colCount)

                                strLine += ",";

                            if (dataGridView2.Rows[j].Cells[k].Value == null)

                                strLine += "";

                            else
                            {

                                string cell = dataGridView2.Rows[j].Cells[k].Value.ToString().Trim();

                                //防止里面含有特殊符号

                                cell = cell.Replace("\"", "\"\"");

                                cell = "\"" + cell + "\"";

                                strLine += cell;

                            }

                        }

                        sw.WriteLine(strLine);

                    }

                    sw.Close();

                    stream.Close();

                    MessageBox.Show("数据被导出到：" + saveFileDialog.FileName.ToString(), "导出完毕", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }

                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "导出错误", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
              }
           }
         #endregion



         #region /* 数据导出到excel */
         public void ExportExcel()
         {
             try
             {
                 this.Cursor = Cursors.WaitCursor;

                 if (!Directory.Exists(@"D:\BMDT"))
                     Directory.CreateDirectory(@"D:\BMDT");


                 string fileName = "";
                 string saveFileName = "";
                 SaveFileDialog saveDialog = new SaveFileDialog();
                 saveDialog.DefaultExt = "xlsx";
                 saveDialog.InitialDirectory = @"D:\BMDT";
                 saveDialog.Filter = "Excel文件|*.xlsx";
                 // saveDialog.FileName = fileName;
                 saveDialog.FileName = "BMDT_Data_" + DateTime.Now.ToLongDateString().ToString();
                 saveDialog.ShowDialog();
                 saveFileName = saveDialog.FileName;



                if (saveFileName.IndexOf(":") < 0)
                 {
                     this.Cursor = Cursors.Default;
                     return; //被点了取消
                 }
                 Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                 if (xlApp == null)
                 {
                     MessageBox.Show("无法创建Excel对象，您的电脑可能未安装Excel");
                     return;
                 }
                 Microsoft.Office.Interop.Excel.Workbooks workbooks = xlApp.Workbooks;
                 Microsoft.Office.Interop.Excel.Workbook workbook = workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);
                 Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[1];//取得sheet1 
                 Microsoft.Office.Interop.Excel.Range range = worksheet.Range[worksheet.Cells[4, 1], worksheet.Cells[8, 1]];

                 //写入标题             
                 for (int i = 0; i < dataGridView2.ColumnCount; i++)
                 { worksheet.Cells[1, i + 1] = dataGridView2.Columns[i].HeaderText; }
                 //写入数值
                 for (int r = 0; r < dataGridView2.Rows.Count; r++)
                 {
                     for (int i = 0; i < dataGridView2.ColumnCount; i++)
                     {
                         worksheet.Cells[r + 2, i + 1] = dataGridView2.Rows[r].Cells[i].Value;

                         if (this.dataGridView2.Rows[r].Cells[i].Style.BackColor == Color.Red)
                         {

                             range = worksheet.Range[worksheet.Cells[r + 2, i + 1], worksheet.Cells[r + 2, i + 1]];
                             range.Interior.ColorIndex = 10;

                         }



                     }
                     System.Windows.Forms.Application.DoEvents();
                 }
                 worksheet.Columns.EntireColumn.AutoFit();//列宽自适应

                 MessageBox.Show(fileName + "资料保存成功", "提示", MessageBoxButtons.OK);
                 if (saveFileName != "")
                 {
                     try
                     {
                         workbook.Saved = true;
                         workbook.SaveCopyAs(saveFileName);  //fileSaved = true;  

                     }
                     catch (Exception ex)
                     {//fileSaved = false;                      
                         MessageBox.Show("导出文件时出错,文件可能正被打开！\n" + ex.Message);
                     }
                 }
                 xlApp.Quit();
                 GC.Collect();//强行销毁           

                 this.Cursor = Cursors.Default;
             }
             catch
             {
                 this.Cursor = Cursors.Default;
                 MessageBox.Show("处理异常1");

             }




         }
         #endregion


         #region /* 按钮触发事件*/
         public void ExcelToCsv(string FilePath1, string FilePath2)
         {
             Microsoft.Office.Interop.Excel.Application xApp = new Microsoft.Office.Interop.Excel.Application();
             // xApp.Visible = true;
             Excel.Workbook xBook = xApp.Workbooks._Open(FilePath1, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                 Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);

             Object format = Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV;
             xBook.SaveAs(FilePath2, format, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
             xApp.Quit();
         
         }
       
         public static void DataTableToSQLServer(DataTable dt, string tableName)
        {
            using (SqlConnection destinationConnection = new SqlConnection(SqlHelper.conStr))
            {
                destinationConnection.Open();
                using (SqlBulkCopy bulkCopy = new SqlBulkCopy(destinationConnection))
                {
                    try
                    {
                        bulkCopy.DestinationTableName = tableName;//要插入的表的表名
                        bulkCopy.BatchSize = dt.Rows.Count;
                       // bulkCopy.BatchSize = dt.Rows.Count;
                      //  bulkCopy.BatchSize = 100;
                     //   bulkCopy.NotifyAfter = 100;
                        
                        /*数据太多或者 excel被打开等会报错*/
                        if (Flag == 1)
                        {
                            bulkCopy.ColumnMappings.Add("ID", "ID");//映射字段名 DataTable列名 ,数据库 对应的列名  
                            bulkCopy.ColumnMappings.Add("DATE", "CREATETIME");
                            bulkCopy.ColumnMappings.Add("CODE", "REASONCODE1");
                            bulkCopy.ColumnMappings.Add("STEPID", "PROCESSID");
                            bulkCopy.ColumnMappings.Add("FAB", "FAB");
                        }
                            /*表的必须从第一个单元格开始*/
                        else if (Flag == 2)
                        {
                                          
                                bulkCopy.ColumnMappings.Add("Operation", "Operation");//映射字段名 DataTable列名 ,数据库 对应的列名  
                                bulkCopy.ColumnMappings.Add("EQP ID", "EQP_ID");
                                bulkCopy.ColumnMappings.Add("Lot ID", "Lot_ID");
                                bulkCopy.ColumnMappings.Add("Event Time", "Event_Time");
                        
                        }
                        /*gls yield*/
                        else if (Flag == 3)
                        {

                            bulkCopy.ColumnMappings.Add("Operation", "Operation");//映射字段名 DataTable列名 ,数据库 对应的列名  
                            bulkCopy.ColumnMappings.Add("EQP ID", "EQP_ID");
                            bulkCopy.ColumnMappings.Add("Lot ID", "Lot_ID");
                            bulkCopy.ColumnMappings.Add("Event Time", "Event_Time");

                        }

                        bulkCopy.WriteToServer(dt);
                    }
                    catch (Exception ex)
                    {
                        System.Windows.Forms.MessageBox.Show(ex.Message);
                    }
                    finally
                    {
                        if (destinationConnection.State == ConnectionState.Open)
                        {
                            destinationConnection.Close();
                        }
                    }
                }
            }
        }
         
       // public SqlConnection conn = new SqlConnection(SqlHelper.conStr);
        private void button2_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
          
          
             try
            {
  
                    Flag = 1;


                //将excel数据读取到dataTable中,然后将dataTable数据写入到Sql中
                //读取到Excel数据
                //1.选择Excel路径(为了灵活性,暂时手动进行,后续应该考虑自动获取近期数据)
                string localFilePath = string.Empty;
                OpenFileDialog openFile = new OpenFileDialog();
                openFile.Title = "请选择文件";
                //设置文件类型  
                openFile.Filter = "XLSX(*.xlsx)|*.xlsx|xls(*.xls)|*.xls";
                //设置默认文件类型显示顺序  
                openFile.FilterIndex = 1;
                //保存对话框是否记忆上次打开的目录  
                openFile.RestoreDirectory = true;
                //设置是否允许多选  
                //save.Multiselect = false;
                //按下确定选择的按钮  
                if (openFile.ShowDialog() == DialogResult.OK)
                {
                    //获得文件路径  
                    localFilePath = openFile.FileName.ToString();
                }
                //读取数据到DataTable中
                DataTable rawData = new DataTable();
                

                //nopi会自动建列
                //rawData.Columns.Add("PANELID");
                //rawData.Columns.Add("EVENTTIME");
                //rawData.Columns.Add("DESCRIPTION");

                rawData = NOPI.ExcelToDataTable(localFilePath, true);
               // MessageBox.Show("111");
                //写入到数据库
                DataTableToSQLServer(rawData, "Table_2");
      
                MessageBox.Show("成功插入了" + rawData.Rows.Count + "条数据!");
              
            
             }
             catch
             {
                
                 this.Cursor = Cursors.Default;
                 MessageBox.Show("处理异常2");
             }

            
            this.Cursor = Cursors.Default;

        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
         
            try
            {
                if (radioButton1.Checked == true)
                    Flag = 2;
                else
                    Flag = 3;


                string localFilePath = string.Empty;
                OpenFileDialog openFile = new OpenFileDialog();
                openFile.Title = "请选择文件";
                openFile.Filter = "XLSX(*.xlsx)|*.xlsx|xls(*.xls)|*.xls";
                openFile.FilterIndex = 1; 
                openFile.RestoreDirectory = true;
                if (openFile.ShowDialog() == DialogResult.OK)
                {
                    localFilePath = openFile.FileName.ToString();
                }
                DataTable rawData = new DataTable();
                rawData = NOPI.ExcelToDataTable(localFilePath, true);
                DataTableToSQLServer(rawData, "EQP_Info");
                MessageBox.Show("成功插入了" + rawData.Rows.Count + "条数据!");
            }
            catch
            {
               
                this.Cursor = Cursors.Default;
                MessageBox.Show("处理异常3");
            }
         
            this.Cursor = Cursors.Default;
        }

        private void button12_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;

            try
            {

                ReadFromSQL("select * from (select distinct REASONCODE1,code.f2 as NewCode from Table_2 left join code on Table_2.REASONCODE1 = code.f1) as a1 ", "code", 1);
               
            }
            catch
            {

                this.Cursor = Cursors.Default;
                MessageBox.Show("处理异常8");
            }
            this.Cursor = Cursors.Default;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
         
            try
            {
               // if (Button1_Flag == true || Button2_Flag == true && Button3_Flag == true)
                if(1==1)
                {
                    Button4_Flag = true;
                    Button1_Flag = false;
                    Button2_Flag = false;
                    Button3_Flag = false;


                    string InsertSql = "";
                   /* conn = new SqlConnection(str);
                    conn.Open();
                    SqlCommand cmd4 = new SqlCommand("delete from Defect_Data_temp", conn);
                    cmd4.CommandTimeout = 0;
                    cmd4.ExecuteNonQuery();
                    conn.Close();*/

                    SendToSQL("delete from Defect_Data_temp");
                    SendToSQL("delete from Raw_Data");
                    SendToSQL("delete from Defect_Data");
                  //  SendToSQL("delete from Table_1");
                  //  SendToSQL("delete from Table_2");
                   // SendToSQL("delete from EQP_Info");


                    SendToSQL("insert Table_1 select * from (select distinct Table_2.ID,Table_2.CREATETIME,"
                               +" isnull(code.f2,'后端不良') as REASONCODEID,Table_2.PROCESSID,Table_2.FAB "
                                + " from Table_2 left join code on Table_2.REASONCODE1 = code.f1) as a1 ");

                  


                  /*  InsertSql = "declare @rowNum bigint;"
                                + "set @rowNum=(select COUNT(*) from Defect_Data);"
                                + "SELECT @rowNum;"
                                + "insert into Defect_Data_temp "
                                + "select (ROW_NUMBER() over(order by s10.LotID))+@rowNum As rownum, s10.LotID,s10.Model,CONVERT(date,s10.BMDTDate,101) AS BMDTDate,"
                                + "CONVERT(varchar(100),DATENAME(Month,s10.BMDTDate),120)+'M' AS BMDTMonth,"
                                + "'CW'+CONVERT(varchar(100), datepart(wk,s10.BMDTDate), 120) AS BMDTWeek,"
                                + "s10.batch,s10.DefectCode,"
                                + "s10.Count,s10.SingleCount,s10.DefectRate,s10.Yield,s11.EQP_ID,s11.Operation,s11.Event_Time,"
                  
                                + "cast(CONVERT(varchar(20),s11.Event_Time,112) as bigint) AS Date," 

                                + "CONVERT(varchar(100),DATENAME(Month,s11.Event_Time),120) AS Month,"
                                + "CONVERT(varchar(100), datepart(wk,s11.Event_Time), 120) AS Week,s10.PROCESSID,s10.FAB,SUBSTRING(s10.LotID,7,1) as Floor from "
                                + "(select s8.*,s9.Yield,"
                                + "LEFT(s8.LotID,4) as Model,SUBSTRING(s8.LotID,5,2) as batch,s9.PROCESSID,s9.FAB AS FAB from "
                                + "(select s3.*,SingleCount,COALESCE(round(CAST(Count as FLOAT)/SingleCount,4),0) as DefectRate from "
                                + "(select s1.*,s2.BMDTDate from "
                                + "(select distinct LEFT(ID,10) as LotID,REASONCODE1 as DefectCode,COUNT(REASONCODE1) as Count "
                                + "from Table_1 GROUP BY LEFT(ID,10),REASONCODE1) as s1 left join "
                                + "(select distinct LEFT(ID,10) as LotID,min(CREATETIME) as BMDTDate from Table_1 "
                                + "GROUP BY LEFT(ID,10)) as s2 on s1.LotID= s2.LotID)  as s3 left join "
                                + "(select LEFT(Q_ID,10) as LotID,SUM(SingleCount) AS SingleCount FROM "
                                + "(select distinct LEFT(ID,13) as Q_ID,SingleCount from (select * from Table_1 left join "
                                + "Q_SingleCount on LEFT(Table_1.ID,4)+ substring(Table_1.ID,13,1)"
                                + "= Q_SingleCount.Q_Name) as t ) as t1 GROUP BY LEFT(Q_ID,10)) as s4 on s3.LotID=s4.LotID ) as s8 left join "

                                + "(select z1.*,z3.PROCESSID,z3.FAB from  "
                                + "(select a,1-sum(DefectRate) as Yield from "
                                + "(select a,b,c,SingleCount,COALESCE(round(CAST(c as FLOAT)/SingleCount,4),0)  as DefectRate "
                                + "from (select distinct LEFT(ID,10) as a,REASONCODE1 as b,COUNT(REASONCODE1) as c "
                                + "from Table_1 GROUP BY LEFT(ID,10),REASONCODE1 ) as s5 left join "
                                + "(select LEFT(Q_ID,10) as LotID,SUM(SingleCount) AS SingleCount FROM "
                                + "(select distinct LEFT(ID,13) as Q_ID,SingleCount from (select * from Table_1 left join "
                                + "Q_SingleCount on LEFT(Table_1.ID,4)+ substring(Table_1.ID,13,1)"
                                + "= Q_SingleCount.Q_Name) as t ) as t1 GROUP BY LEFT(Q_ID,10)) as s6 "

                                + "on s5.a=s6.LotID ) as s7 group by a )  as z1 left join "
                                + "(select distinct LEFT(z2.ID,10) as LotID,z2.PROCESSID,z2.FAB from Table_1 as Z2) as z3 on z1.a=z3.LotID) "
                                
                                +" as s9 on s8.LotID=s9.a) as s10 left join "
                                + "(select distinct  Lot_ID,EQP_ID,Operation,Event_Time from EQP_Info)"
                                + " as s11 on s10.LotID=s11.Lot_ID order by Date";*/
                    if(radioButton1.Checked == true)
                    InsertSql = "declare @rowNum bigint;set @rowNum=(select COUNT(*) from Defect_Data);"
                               + "SELECT @rowNum;"
                               + "select * into ##tempz1 from(select distinct  SUBSTRING(a.ID,1,10) as LotID from Table_1 as a) as z1;"
                               + "select * into ##tempz2 from(select distinct b.F2 as DefectCode  from code as b) as z2;"
                               + "select * into ##tempz3 from(select *  from ##tempz1,##tempz2) as z3;"
                               +"select * into ##tempz4 from(select distinct LEFT(ID,10) as LotID,REASONCODE1 as DefectCode,"
                               +"COUNT(REASONCODE1) as Count from Table_1 GROUP BY LEFT(ID,10),REASONCODE1) as z4;"
                               +"select * into ##temp1 from(select a.*,isnull(b.Count,0) as Count from ##tempz3 as a left join "
                               +" ##tempz4 as b on (a.LotID=b.LotID and a.DefectCode=b.DefectCode)) as a1;"
                               + "select * into ##temp2 from(select distinct LEFT(ID,10) as LotID,min(CREATETIME) as BMDTDate from Table_1 GROUP BY LEFT(ID,10)) as a2;"
                               +"select * into ##temp3 from(select * from Table_1 left join Q_SingleCount on LEFT(Table_1.ID,4)+ substring(Table_1.ID,13,1)= Q_SingleCount.Q_Name) as a3;"
                               + "select * into ##temp4 from(select distinct LEFT(ID,13) as Q_ID,SingleCount from ##temp3) as a4;"
                               + "select * into ##temp5 from(select LEFT(Q_ID,10) as LotID,SUM(SingleCount) AS SingleCount FROM ##temp4 GROUP BY LEFT(Q_ID,10)) as a5;"
                               +"select * into ##temp6 from(select ##temp1.*,##temp2.BMDTDate from ##temp1 left join ##temp2 on ##temp1.LotID= ##temp2.LotID) as a6;"
                               + "select * into ##temp7 from(select a.*,b.SingleCount from ##temp2 as a ,##temp5 as b where a.LotID=b.LotID ) as a7;"
                               +"select * into ##temp8 from(select a.*,b.BMDTDate,b.SingleCount,COALESCE(round(CAST(Count as FLOAT)/SingleCount,4),0) as DefectRate "
                              +"  from ##temp1 as a left join ##temp7 as b on a.LotID =b.LotID ) as a8;"
                              + "select * into ##temp9 from(select distinct LotID,1-sum(DefectRate) as Yield from ##temp8 group by LotID ) as a9 ;"
                              + "select * into ##temp10 from "
                              +"(select a.*, b.Yield,LEFT(a.LotID,4) as Model,SUBSTRING(a.LotID,7,1) as Floor,(ROW_NUMBER() over(order by a.LotID))+@rowNum As rownum,"
                              +" CONVERT(date,a.BMDTDate,101) AS BMDTDate1,CONVERT(varchar(100),DATENAME(Month,a.BMDTDate),120)+'M' AS BMDTMonth,"
                              +" 'CW'+CONVERT(varchar(100), datepart(wk,a.BMDTDate), 120) AS BMDTWeek,SUBSTRING(a.LotID,5,2) as batch"
                              +" from ##temp8 as a left join ##temp9 as b on a.LotID =b.LotID) as a10 ;"
                              +"select * into ##tempb1 from(select distinct LEFT(z1.ID,10) as LotID,z1.PROCESSID,z1.FAB from Table_1 as z1) as b1;"
                              + "select * into ##tempc1 from(select distinct  Lot_ID,EQP_ID,Operation,Event_Time,"
                              +"cast(CONVERT(varchar(20),a.Event_Time,112) as bigint) AS Date,CONVERT(varchar(100),"
                              +"DATENAME(Month,a.Event_Time),120) AS Month,CONVERT(varchar(100), "
                              +"datepart(wk,a.Event_Time), 120) AS Week from EQP_Info as a) as c1;"
                              +"select * into ##tempc2 from(select a.*,b.EQP_ID,b.Operation,b.Event_Time,b.Date,"
                              +"b.Month,b.Week from ##temp10 as a left join ##tempc1 as b on  a.LotID = b.Lot_ID ) as c2;"
                              +"select * into ##tempc3 from(select ##tempc2.*,##tempb1.PROCESSID,##tempb1.FAB "
                              +"from  ##tempc2 left join ##tempb1 on ##tempc2.LotID = ##tempb1.LotID ) as c3;"
                              +"insert into Defect_Data_temp select a.rownum,a.LotID,a.Model,a.BMDTDate,a.Month,a.Week,a.batch,"
                              +"a.DefectCode,a.Count,a.SingleCount,a.DefectRate,a.Yield,a.EQP_ID,"
                              +"a.Operation,a.Event_Time,a.Date,a.Month,a.Week,a.PROCESSID,a.FAB,a.Floor from ##tempc3 as a;"
                              +"drop table ##tempz1;drop table ##tempz2;drop table ##tempz3;drop table ##tempz4;drop table ##temp1;drop table ##temp2;"
                              +"drop table ##temp3;drop table ##temp4;drop table ##temp5;drop table ##temp6;drop table ##temp7;drop table ##temp8;"
                              +"drop table ##temp9;drop table ##temp10;drop table ##tempb1;drop table ##tempc1;drop table ##tempc2;drop table ##tempc3;"
                               ;
                    else

                          InsertSql = "declare @rowNum bigint;set @rowNum=(select COUNT(*) from Defect_Data);"
                               + "SELECT @rowNum;"
                               + "select * into ##tempz1 from(select distinct  SUBSTRING(a.ID,1,12) as LotID from Table_1 as a) as z1;"
                               + "select * into ##tempz2 from(select distinct b.F2 as DefectCode  from code as b) as z2;"
                               + "select * into ##tempz3 from(select *  from ##tempz1,##tempz2) as z3;"
                               +"select * into ##tempz4 from(select distinct LEFT(ID,12) as LotID,REASONCODE1 as DefectCode,"
                               +"COUNT(REASONCODE1) as Count from Table_1 GROUP BY LEFT(ID,12),REASONCODE1) as z4;"
                               +"select * into ##temp1 from(select a.*,isnull(b.Count,0) as Count from ##tempz3 as a left join "
                               +" ##tempz4 as b on (a.LotID=b.LotID and a.DefectCode=b.DefectCode)) as a1;"
                               + "select * into ##temp2 from(select distinct LEFT(ID,12) as LotID,min(CREATETIME) as BMDTDate from Table_1 GROUP BY LEFT(ID,12)) as a2;"
                               +"select * into ##temp3 from(select * from Table_1 left join Q_SingleCount on LEFT(Table_1.ID,4)+ substring(Table_1.ID,13,1)= Q_SingleCount.Q_Name) as a3;"
                               + "select * into ##temp4 from(select distinct LEFT(ID,13) as Q_ID,SingleCount from ##temp3) as a4;"
                               + "select * into ##temp5 from(select LEFT(Q_ID,12) as LotID,SUM(SingleCount) AS SingleCount FROM ##temp4 GROUP BY LEFT(Q_ID,12)) as a5;"
                               +"select * into ##temp6 from(select ##temp1.*,##temp2.BMDTDate from ##temp1 left join ##temp2 on ##temp1.LotID= ##temp2.LotID) as a6;"
                               + "select * into ##temp7 from(select a.*,b.SingleCount from ##temp2 as a ,##temp5 as b where a.LotID=b.LotID ) as a7;"
                               +"select * into ##temp8 from(select a.*,b.BMDTDate,b.SingleCount,COALESCE(round(CAST(Count as FLOAT)/SingleCount,4),0) as DefectRate "
                              +"  from ##temp1 as a left join ##temp7 as b on a.LotID =b.LotID ) as a8;"
                              + "select * into ##temp9 from(select distinct LotID,1-sum(DefectRate) as Yield from ##temp8 group by LotID ) as a9 ;"
                              + "select * into ##temp10 from "
                              +"(select a.*, b.Yield,LEFT(a.LotID,4) as Model,SUBSTRING(a.LotID,7,1) as Floor,(ROW_NUMBER() over(order by a.LotID))+@rowNum As rownum,"
                              +" CONVERT(date,a.BMDTDate,121) AS BMDTDate1,CONVERT(varchar(100),DATENAME(Month,a.BMDTDate),120)+'M' AS BMDTMonth,"
                              +" 'CW'+CONVERT(varchar(100), datepart(wk,a.BMDTDate), 120) AS BMDTWeek,SUBSTRING(a.LotID,5,2) as batch"
                              +" from ##temp8 as a left join ##temp9 as b on a.LotID =b.LotID) as a10 ;"
                              +"select * into ##tempb1 from(select distinct LEFT(z1.ID,12) as LotID,z1.PROCESSID,z1.FAB from Table_1 as z1) as b1;"
                              + "select * into ##tempc1 from(select distinct  Lot_ID,EQP_ID,Operation,Event_Time,"
                              +"cast(CONVERT(varchar(20),a.Event_Time,112) as bigint) AS Date,CONVERT(varchar(100),"
                              +"DATENAME(Month,a.Event_Time),120) AS Month,CONVERT(varchar(100), "
                              +"datepart(wk,a.Event_Time), 120) AS Week from EQP_Info as a) as c1;"
                              +"select * into ##tempc2 from(select a.*,b.EQP_ID,b.Operation,b.Event_Time,b.Date,"
                              +"b.Month,b.Week from ##temp10 as a left join ##tempc1 as b on  a.LotID = b.Lot_ID ) as c2;"
                              +"select * into ##tempc3 from(select ##tempc2.*,##tempb1.PROCESSID,##tempb1.FAB "
                              +"from  ##tempc2 left join ##tempb1 on ##tempc2.LotID = ##tempb1.LotID ) as c3;"
                              +"insert into Defect_Data_temp select a.rownum,a.LotID,a.Model,a.BMDTDate,a.Month,a.Week,a.batch,"
                              +"a.DefectCode,a.Count,a.SingleCount,a.DefectRate,a.Yield,a.EQP_ID,"
                              +"a.Operation,a.Event_Time,a.Date,a.Month,a.Week,a.PROCESSID,a.FAB,a.Floor from ##tempc3 as a;"
                              +"drop table ##tempz1;drop table ##tempz2;drop table ##tempz3;drop table ##tempz4;drop table ##temp1;drop table ##temp2;"
                              +"drop table ##temp3;drop table ##temp4;drop table ##temp5;drop table ##temp6;drop table ##temp7;drop table ##temp8;"
                              +"drop table ##temp9;drop table ##temp10;drop table ##tempb1;drop table ##tempc1;drop table ##tempc2;drop table ##tempc3;"
                               ;



                    conn = new SqlConnection(str);
                    conn.Open();
                    SqlCommand cmd1 = new SqlCommand(InsertSql, conn);
                    cmd1.CommandTimeout = 0;        
                    cmd1.ExecuteNonQuery();

                    //SqlParameter pms = new SqlParameter()
                    //{
                    //};
                   // SqlHelper.ExecuteNonQuery(conn,CommandType.StoredProcedure,"dbo.Yield",0)


                    SendToSQL("delete from Defect_Data_temp where LotID is NULL");// 删除空白行，不然生成图表会报错
                    ReadFromSQL("select * from Defect_Data_temp order by rownum", "Defect_Data_temp", 1);
                }
                else
                {
                    MessageBox.Show("你还未导入数据，请检查！");

                }
            }
            catch
            {
               
                this.Cursor = Cursors.Default;
                MessageBox.Show("处理异常4");
            }

        
            this.Cursor = Cursors.Default;

        }


        private void button5_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
         
            try
            {
                DialogResult dr = MessageBox.Show("确定将Yield数据导入数据库吗？", "提示", MessageBoxButtons.OKCancel);
                if (dr == DialogResult.OK)
                {
                   // if (Button4_Flag == true)
                   // {
                        Button5_Flag = true;
                        Button4_Flag = false;
                        //button12.Enabled = true;

                        conn = new SqlConnection(str);
                        conn.Open();
                        SqlCommand cmd = new SqlCommand("select count(*) cnt from Defect_Data", conn);
                        cmd.CommandTimeout = 0;
                        // cmd.CommandText = "select count(*) cnt from Defect_Data";
                        SqlDataReader dr1 = cmd.ExecuteReader();
                        dr1.Read();
                        String count1 = dr1["cnt"].ToString();
                        conn.Close();

                        //MessageBox.Show("1");
                        conn = new SqlConnection(str);
                        conn.Open();
                        SqlCommand cmd3 = new SqlCommand("insert into Defect_Data select * from Defect_Data_temp ", conn);
                        cmd3.CommandTimeout = 0;
                        cmd3.ExecuteNonQuery();
                        conn.Close();
                       // MessageBox.Show("2");

                        conn = new SqlConnection(str);
                        conn.Open();
                        SqlCommand cmd2 = new SqlCommand("select count(*) cnt from Defect_Data", conn);
                        cmd2.CommandTimeout = 0;
                        SqlDataReader dr2 = cmd2.ExecuteReader();
                        dr2.Read();
                        String count2 = dr2["cnt"].ToString();
                        conn.Close();

                        if (!Directory.Exists("D:\\BMDT"))
                            Directory.CreateDirectory("D:\\BMDT");
                        if (!File.Exists("D:\\BMDT\\Defect_Data_log.txt"))
                        {
                            FileStream NewText = File.Create("D:\\BMDT\\Defect_Data_log.txt");
                            NewText.Close();
                        }

                        StreamWriter sw = new StreamWriter("D:\\BMDT\\Defect_Data_log.txt", true);
                        sw.WriteLine(DateTime.Now.ToString() + " 原数据条数：" + count1 + " 追加后数据条数：" + count2);
                        sw.Close();
                        MessageBox.Show("恭喜你，数据成功插入数据库");
                  //  }
                  //  else
                  //  {
                   //     MessageBox.Show("您还没加工导入的数据！");
                    //}
                }
                else if (dr == DialogResult.Cancel)
                {
                    //用户选择取消的操作
                    MessageBox.Show("您【取消】了数据导入");
                }
            }
            catch
            {
               
                this.Cursor = Cursors.Default;
                MessageBox.Show("处理异常5");
            }
          
            this.Cursor = Cursors.Default;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            
            try
            {
                String StartTime = string.Empty, EndTime = string.Empty;
                string Sql = string.Empty, Sql1 = string.Empty;
                int j = 0,i=0 ;
                string output = string.Empty;

                chart1.Visible = false;
                dataGridView2.Visible = true;

                StartTime = dateTimePicker1.Value.Date.ToShortDateString();
                EndTime = dateTimePicker2.Value.Date.ToShortDateString();
                StartTime = dateTimePicker1.Value.Date.ToString("yyyyMMdd");
                EndTime = dateTimePicker2.Value.Date.ToString("yyyyMMdd");

                //Sql = "select * from Raw_Data where   Date BETWEEN " + StartTime + " and " + EndTime;
                // ReadFromSQL(Sql, "Defect_Data", 2);
                Sql = "select * from Defect_Data where  1=1 ";
                Sql1 = "select distinct LotID,Yield from Defect_Data where  1=1 ";

                if (checkBox8.Checked == true)
                {
                    Sql = Sql + " and Date BETWEEN " + StartTime + " and " + EndTime;
                    Sql1 = Sql1 + " and Date BETWEEN " + StartTime + " and " + EndTime;
                }
                if (checkBox9.Checked == true)
                {
                    output = string.Empty;
                    j = 0;
                    for (i = 0; i < checkedListBox1.Items.Count; i++)
                    {
                      if (checkedListBox1.GetItemChecked(i))
                        {
                         if (j == 0)
                            output += "'" + checkedListBox1.GetItemText(checkedListBox1.Items[i]) + "'";
                         if (j > 0)
                            output += ",'" + checkedListBox1.GetItemText(checkedListBox1.Items[i]) + "'";
                             j++;
                         }
                    }
                    if (j > 0)
                    {
                        Sql = Sql + " and  Model in(" + output + ")";
                        Sql1 = Sql1 + " and  Model in(" + output + ")"; ;
                    }
                    else
                        MessageBox.Show("请至少选择一个Model值");
                }

 
                    if (checkBox10.Checked == true)
                    {
                        output = string.Empty;
                        j = 0;
                        for (i = 0; i < checkedListBox2.Items.Count; i++)
                        {
                            if (checkedListBox2.GetItemChecked(i))
                            {
                                if (j == 0)
                                    output += "'" + checkedListBox2.GetItemText(checkedListBox2.Items[i]) + "'";
                                if (j > 0)
                                    output += ",'" + checkedListBox2.GetItemText(checkedListBox2.Items[i]) + "'";
                                j++;
                            }
                        }
                        if (j > 0)
                        {
                            Sql = Sql + " and  batch in(" + output + ")";
                            Sql1 = Sql1 + " and  batch in(" + output + ")";
                        }
                        else
                            MessageBox.Show("请至少选择一个batch值");
                    }

               // if (checkBox11.Checked == true)
                //    Sql = Sql + " and DefectCode=" + "'" + comboBox9.Text + "'";
                if (checkBox11.Checked == true)
                {
                    output = string.Empty;
                    j = 0;
                    for (i = 0; i < checkedListBox3.Items.Count; i++)
                    {
                        if (checkedListBox3.GetItemChecked(i))
                        {
                            if (j == 0)
                                output += "'" + checkedListBox3.GetItemText(checkedListBox3.Items[i]) + "'";
                            if (j > 0)
                                output += ",'" + checkedListBox3.GetItemText(checkedListBox3.Items[i]) + "'";
                            j++;
                        }
                    }
                    if (j > 0)
                    {
                        Sql = Sql + " and  DefectCode in(" + output + ")";
                        Sql1 = Sql1 + " and  DefectCode in(" + output + ")";
                    }
                    else
                        MessageBox.Show("请至少选择一个DefectCode值");
                }

               // if (checkBox12.Checked == true)
               //     Sql = Sql + " and Operation=" + "'" + comboBox10.Text + "'";
                if (checkBox12.Checked == true)
                {
                    output = string.Empty;
                    j = 0;
                    for (i = 0; i < checkedListBox4.Items.Count; i++)
                    {
                        if (checkedListBox4.GetItemChecked(i))
                        {
                            if (j == 0)
                                output += "'" + checkedListBox4.GetItemText(checkedListBox4.Items[i]) + "'";
                            if (j > 0)
                                output += ",'" + checkedListBox4.GetItemText(checkedListBox4.Items[i]) + "'";
                            j++;
                        }
                    }
                    if (j > 0)
                    {
                        Sql = Sql + " and  Operation in(" + output + ")";
                        Sql1 = Sql1 + " and  Operation in(" + output + ")";
                    }
                    else
                        MessageBox.Show("请至少选择一个Operation值");
                }

              //  if (checkBox13.Checked == true)
             //       Sql = Sql + " and  LotID=" + "'" + comboBox11.Text + "'";
                if (checkBox13.Checked == true)
                {
                    output = string.Empty;
                    j = 0;
                    for (i = 0; i < checkedListBox5.Items.Count; i++)
                    {
                        if (checkedListBox5.GetItemChecked(i))
                        {
                            if (j == 0)
                                output += "'" + checkedListBox5.GetItemText(checkedListBox5.Items[i]) + "'";
                            if (j > 0)
                                output += ",'" + checkedListBox5.GetItemText(checkedListBox5.Items[i]) + "'";
                            j++;
                        }
                    }
                    if (j > 0)
                    {
                        Sql = Sql + " and  LotID in(" + output + ")";
                        Sql1 = Sql1 + " and  LotID in(" + output + ")";
                    }
                    else
                        MessageBox.Show("请至少选择一个LotID值");
                }

               // if (checkBox14.Checked == true)
                //    Sql = Sql + " and  EQP_ID=" + "'" + comboBox12.Text + "'";

                if (checkBox14.Checked == true)
                {
                    output = string.Empty;
                    j = 0;
                    for (i = 0; i < checkedListBox6.Items.Count; i++)
                    {
                        if (checkedListBox6.GetItemChecked(i))
                        {
                            if (j == 0)
                                output += "'" + checkedListBox6.GetItemText(checkedListBox6.Items[i]) + "'";
                            if (j > 0)
                                output += ",'" + checkedListBox6.GetItemText(checkedListBox6.Items[i]) + "'";
                            j++;
                        }
                    }
                    if (j > 0)
                    {
                        Sql = Sql + " and  EQP_ID in(" + output + ")";
                        Sql1 = Sql1 + " and  EQP_ID in(" + output + ")";
                    }
                    else
                        MessageBox.Show("请至少选择一个EQP_ID值");
                }

                if (checkBox17.Checked == true)
                {
                    output = string.Empty;
                    j = 0;
                    for (i = 0; i < checkedListBox16.Items.Count; i++)
                    {
                        if (checkedListBox16.GetItemChecked(i))
                        {
                            if (j == 0)
                                output += "'" + checkedListBox16.GetItemText(checkedListBox16.Items[i]) + "'";
                            if (j > 0)
                                output += ",'" + checkedListBox16.GetItemText(checkedListBox16.Items[i]) + "'";
                            j++;
                        }
                    }
                    if (j > 0)
                    {
                        Sql = Sql + " and  PROCESSID in(" + output + ")";
                        Sql1 = Sql1 + " and  PROCESSID in(" + output + ")";
                    }
                    else
                        MessageBox.Show("请至少选择一个PROCESSID值");
                }
                if (checkBox18.Checked == true)
                {
                    output = string.Empty;
                    j = 0;
                    for (i = 0; i < checkedListBox17.Items.Count; i++)
                    {
                        if (checkedListBox17.GetItemChecked(i))
                        {
                            if (j == 0)
                                output += "'" + checkedListBox17.GetItemText(checkedListBox17.Items[i]) + "'";
                            if (j > 0)
                                output += ",'" + checkedListBox17.GetItemText(checkedListBox17.Items[i]) + "'";
                            j++;
                        }
                    }
                    if (j > 0)
                    {
                        Sql = Sql + " and  FAB in(" + output + ")";
                        Sql1 = Sql1 + " and  FAB in(" + output + ")";
                    }
                    else
                        MessageBox.Show("请至少选择一个FAB值");
                }

                ReadFromSQL(Sql, "Raw_Data", 2);
                ReadFromSQL(Sql1, "Raw_Data", 4);


                label2.Text = "共搜索到" + dataGridView2.RowCount.ToString() + "条数据";
            }
            catch
            {
                
                this.Cursor = Cursors.Default;
                MessageBox.Show("处理异常7");
            }
           
            this.Cursor = Cursors.Default;

        }

        private void button7_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
        
            try
            {
                BMDT_Flag = 4;
                conn = new SqlConnection(str);
                conn.Open();
                SqlCommand cmd = new SqlCommand("delete from Q_SingleCount", conn);
                cmd.CommandTimeout = 0;
                cmd.ExecuteNonQuery();


                for (int Row = 0; Row < dataGridView4.Rows.Count - 1; Row++)
                    insertToSql1(dataGridView4.Rows[Row].Cells[0].Value.ToString(), dataGridView4.Rows[Row].Cells[1].Value.ToString());

                ReadFromSQL("select * from Q_SingleCount ", "Q_SingleCount", 3);
                MessageBox.Show("注册完成");
            }
            catch
            {
               
                this.Cursor = Cursors.Default;
                MessageBox.Show("处理异常8");
            }
            this.Cursor = Cursors.Default;
        }


        private void button13_Click(object sender, EventArgs e)
        {

            this.Cursor = Cursors.WaitCursor;

            try
            {
                BMDT_Flag = 4;
                conn = new SqlConnection(str);
                conn.Open();
                SqlCommand cmd = new SqlCommand("delete from code", conn);
                cmd.CommandTimeout = 0;
                cmd.ExecuteNonQuery();

               
                for (int Row = 0; Row < dataGridView9.Rows.Count - 1; Row++)
                    insertToSql2(dataGridView9.Rows[Row].Cells[0].Value.ToString(), dataGridView9.Rows[Row].Cells[1].Value.ToString());

                ReadFromSQL("select * from code ", "code", 8);
                MessageBox.Show("注册完成");
            }
            catch
            {

                this.Cursor = Cursors.Default;
                MessageBox.Show("处理异常9");
            }
            this.Cursor = Cursors.Default;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
        
            try
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "CSV files (*.csv)|*.csv";
                saveFileDialog.FilterIndex = 0;
                saveFileDialog.RestoreDirectory = true;
                saveFileDialog.CreatePrompt = true;
                saveFileDialog.FileName = "BMDT_Data_" + DateTime.Now.ToLongDateString().ToString();
                saveFileDialog.Title = "保存";
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {


                    // string fileName = "D:\\BMDT\\1.csv";
                    string fileName = saveFileDialog.FileName;
                    StreamWriter sw = new StreamWriter(fileName, false, Encoding.Default);
                    for (int row = 0; row < dataGridView2.Rows.Count - 1; row++)
                    {
                        if (row == 0)
                        {
                            sw.WriteLine(dataGridView2.Columns[0].HeaderText + "," +
                                         dataGridView2.Columns[1].HeaderText + "," +
                                         dataGridView2.Columns[2].HeaderText + "," +
                                         dataGridView2.Columns[3].HeaderText + "," +
                                         dataGridView2.Columns[4].HeaderText + "," +
                                         dataGridView2.Columns[5].HeaderText + "," +
                                         dataGridView2.Columns[6].HeaderText + "," +
                                         dataGridView2.Columns[7].HeaderText + "," +
                                         dataGridView2.Columns[8].HeaderText + "," +
                                         dataGridView2.Columns[9].HeaderText + "," +
                                         dataGridView2.Columns[10].HeaderText + "," +
                                         dataGridView2.Columns[11].HeaderText + "," +
                                         dataGridView2.Columns[12].HeaderText + "," +
                                         dataGridView2.Columns[13].HeaderText + "," +
                                         dataGridView2.Columns[14].HeaderText + "," +
                                         dataGridView2.Columns[15].HeaderText + "," +
                                         dataGridView2.Columns[16].HeaderText + "," +
                                         dataGridView2.Columns[17].HeaderText + "," +
                                         dataGridView2.Columns[18].HeaderText + "," +
                                         dataGridView2.Columns[19].HeaderText);
                        }
                        else
                        {
                            sw.WriteLine(dataGridView2.Rows[row].Cells[0].Value.ToString() + "," +
                                         dataGridView2.Rows[row].Cells[1].Value.ToString() + "," +
                                         dataGridView2.Rows[row].Cells[2].Value.ToString() + "," +
                                         dataGridView2.Rows[row].Cells[3].Value.ToString() + "," +
                                         dataGridView2.Rows[row].Cells[4].Value.ToString() + "," +
                                         dataGridView2.Rows[row].Cells[5].Value.ToString() + "," +
                                         dataGridView2.Rows[row].Cells[6].Value.ToString() + "," +
                                         dataGridView2.Rows[row].Cells[7].Value.ToString() + "," +
                                         dataGridView2.Rows[row].Cells[8].Value.ToString() + "," +
                                         dataGridView2.Rows[row].Cells[9].Value.ToString() + "," +
                                         dataGridView2.Rows[row].Cells[10].Value.ToString() + "," +
                                         dataGridView2.Rows[row].Cells[11].Value.ToString() + "," +
                                         dataGridView2.Rows[row].Cells[12].Value.ToString() + "," +
                                         dataGridView2.Rows[row].Cells[13].Value.ToString() + "," +
                                         dataGridView2.Rows[row].Cells[14].Value.ToString() + "," +
                                         dataGridView2.Rows[row].Cells[15].Value.ToString() + "," +
                                         dataGridView2.Rows[row].Cells[16].Value.ToString() + "," +
                                         dataGridView2.Rows[row].Cells[17].Value.ToString() + "," +
                                         dataGridView2.Rows[row].Cells[18].Value.ToString() + "," +
                                         dataGridView2.Rows[row].Cells[19].Value.ToString());
                        }
                    }
                    sw.Close();

                    MessageBox.Show("数据保存完成");
                }

            }
            catch
            {
               
                this.Cursor = Cursors.Default;
                MessageBox.Show("处理异常9");
            }
           
            this.Cursor = Cursors.Default;
           
            //ExportCSV();
        }

    

        private void button10_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
          
            try
            {
                DialogResult dr = MessageBox.Show("确定执行Map库数据汇总吗？", "提示", MessageBoxButtons.OKCancel);
                if (dr == DialogResult.OK)
                {
                    if (Button5_Flag == true)
                    {
                        Button10_Flag = true;
                        Button5_Flag = false;


                        string InsertSql = string.Empty;

                   

                        InsertSql = "declare @rowNum bigint;"
                                    + "set @rowNum=(select COUNT(*) from Map_SingleID);"
                                    + "SELECT @rowNum;"
                                    + "insert into Map_SingleID "
                                    + "select (ROW_NUMBER() over(order by k2.LotID))+@rowNum As rownum,"
                                    + "k2.*,k3.*,cast(CONVERT(varchar(20),k3.Event_Time,112) as bigint) AS Date from "
                                    + "(select k1.*,Q_SingleCount.SingleCount from "
                                    + "(select LEFT(ID,10) as LotID,LEFT(ID,4) as Model ,SUBSTRING(ID,5,2) as "
                                    + "batch,ID,LEFT(ID,13) AS Q_ID,SUBSTRING(ID,13,3) AS SingleName,CREATETIME,REASONCODE1,PROCESSID,FAB  "
                                    + "from Table_1 ) as k1 left join	Q_SingleCount on "
                                    + "k1.Model+ substring(k1.ID,13,1)=Q_SingleCount.Q_Name) as k2 left join "
                                    + "EQP_Info as k3 On k2.LotID=k3.Lot_ID order by Date ";



                        conn = new SqlConnection(str);
                        conn.Open();
                        SqlCommand cmd2 = new SqlCommand("select count(*) cnt from Map_SingleID", conn);
                        cmd2.CommandTimeout = 0;
                        SqlDataReader dr2 = cmd2.ExecuteReader();
                        dr2.Read();
                        String count3 = dr2["cnt"].ToString();
                        conn.Close();


                        SendToSQL(InsertSql);


                        conn = new SqlConnection(str);
                        conn.Open();
                        SqlCommand cmd3 = new SqlCommand("select count(*) cnt from Map_SingleID", conn);
                        cmd3.CommandTimeout = 0;
                        SqlDataReader dr3 = cmd3.ExecuteReader();
                        dr3.Read();
                        String count4 = dr3["cnt"].ToString();
                        conn.Close();


                        if (!Directory.Exists("D:\\BMDT"))
                            Directory.CreateDirectory("D:\\BMDT");
                        if (!File.Exists("D:\\BMDT\\log.txt"))
                        {
                            FileStream NewText = File.Create("D:\\BMDT\\log_Map_SingleID.txt");
                            NewText.Close();
                        }

                        StreamWriter sw = new StreamWriter("D:\\BMDT\\log_Map_SingleID.txt", true);
                        sw.WriteLine(DateTime.Now.ToString() + " 原数据条数：" + count3 + " 追加后数据条数：" + count4);
                        sw.Close();



                        MessageBox.Show("库数据汇总完成");
                    }
                    else
                    {
                        MessageBox.Show("请先点击 数据入库!");
                    }

                }
                else if (dr == DialogResult.Cancel)
                {
                    //用户选择取消的操作
                    MessageBox.Show("您【取消】了数据导入");
                }
            }
            catch
            {
              
                this.Cursor = Cursors.Default;
                MessageBox.Show("处理异常11");
            }
          
            this.Cursor = Cursors.Default;

        }

  


        private void button11_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
           
            try
            {
                if (checkBox5.Checked == true && checkBox2.Checked == true)
                {
                    String StartTime = string.Empty, EndTime = string.Empty;
   
                    string Sql = string.Empty;
                    int j = 0, i = 0;
                    string output = string.Empty;

                    StartTime = dateTimePicker3.Value.Date.ToShortDateString();
                    EndTime = dateTimePicker4.Value.Date.ToShortDateString();
                    StartTime = dateTimePicker3.Value.Date.ToString("yyyyMMdd");
                    EndTime = dateTimePicker4.Value.Date.ToString("yyyyMMdd");

                    Sql = "delete from Map_Data;";
                    SendToSQL(Sql);

                    Sql =  "select * into ##tempa1 from( ";
                    Sql = Sql+ "select SingleName,count(SingleName) as Count from Map_SingleID  where 1=1 ";

                    if (checkBox1.Checked == true)
                        Sql = Sql + " and Date BETWEEN " + StartTime + " and " + EndTime;
 
                    if (checkBox2.Checked == true)
                    {
                        output = string.Empty;
                        j = 0;
                        for (i = 0; i < checkedListBox7.Items.Count; i++)
                        {
                            if (checkedListBox7.GetItemChecked(i))
                            {
                                if (j == 0)
                                    output += "'" + checkedListBox7.GetItemText(checkedListBox7.Items[i]) + "'";
                                if (j > 0)
                                    output += ",'" + checkedListBox7.GetItemText(checkedListBox7.Items[i]) + "'";
                                j++;
                            }
                        }
                        if (j > 0)
                            Sql = Sql + " and  Model in(" + output + ")";
                        else
                            MessageBox.Show("请至少选择一个Model值");
                    }

                    if (checkBox3.Checked == true)
                    {
                        output = string.Empty;
                        j = 0;
                        for (i = 0; i < checkedListBox8.Items.Count; i++)
                        {
                            if (checkedListBox8.GetItemChecked(i))
                            {
                                if (j == 0)
                                    output += "'" + checkedListBox8.GetItemText(checkedListBox8.Items[i]) + "'";
                                if (j > 0)
                                    output += ",'" + checkedListBox8.GetItemText(checkedListBox8.Items[i]) + "'";
                                j++;
                            }
                        }
                        if (j > 0)
                            Sql = Sql + " and  batch in(" + output + ")";
                        else
                            MessageBox.Show("请至少选择一个batch值");
                    }

                    if (checkBox4.Checked == true)
                    {
                        output = string.Empty;
                        j = 0;
                        for (i = 0; i < checkedListBox9.Items.Count; i++)
                        {
                            if (checkedListBox9.GetItemChecked(i))
                            {
                                if (j == 0)
                                    output += "'" + checkedListBox9.GetItemText(checkedListBox9.Items[i]) + "'";
                                if (j > 0)
                                    output += ",'" + checkedListBox9.GetItemText(checkedListBox9.Items[i]) + "'";
                                j++;
                            }
                        }
                        if (j > 0)
                            Sql = Sql + " and  REASONCODE1 in(" + output + ")";
                        else
                            MessageBox.Show("请至少选择一个REASONCODE1值");
                    }


                    if (checkBox5.Checked == true)
                    {
                        output = string.Empty;
                        j = 0;
                        for (i = 0; i < checkedListBox10.Items.Count; i++)
                        {
                            if (checkedListBox10.GetItemChecked(i))
                            {
                                if (j == 0)
                                    output += "'" + checkedListBox10.GetItemText(checkedListBox10.Items[i]) + "'";
                                if (j > 0)
                                    output += ",'" + checkedListBox10.GetItemText(checkedListBox10.Items[i]) + "'";
                                j++;
                            }
                        }
                        if (j > 0)
                            Sql = Sql + " and  Operation in(" + output + ")";
                        else
                            MessageBox.Show("请至少选择一个Operation值");
                    }

                    if (checkBox6.Checked == true)
                    {
                        output = string.Empty;
                        j = 0;
                        for (i = 0; i < checkedListBox11.Items.Count; i++)
                        {
                            if (checkedListBox11.GetItemChecked(i))
                            {
                                if (j == 0)
                                    output += "'" + checkedListBox11.GetItemText(checkedListBox11.Items[i]) + "'";
                                if (j > 0)
                                    output += ",'" + checkedListBox11.GetItemText(checkedListBox11.Items[i]) + "'";
                                j++;
                            }
                        }
                        if (j > 0)
                            Sql = Sql + " and  Lot_ID in(" + output + ")";
                        else
                            MessageBox.Show("请至少选择一个Lot_ID值");
                    }

                    if (checkBox7.Checked == true)
                    {
                        output = string.Empty;
                        j = 0;
                        for (i = 0; i < checkedListBox12.Items.Count; i++)
                        {
                            if (checkedListBox12.GetItemChecked(i))
                            {
                                if (j == 0)
                                    output += "'" + checkedListBox12.GetItemText(checkedListBox12.Items[i]) + "'";
                                if (j > 0)
                                    output += ",'" + checkedListBox12.GetItemText(checkedListBox12.Items[i]) + "'";
                                j++;
                            }
                        }
                        if (j > 0)
                            Sql = Sql + " and  EQP_ID in(" + output + ")";
                        else
                            MessageBox.Show("请至少选择一个EQP_ID值");
                    }

                    if (checkBox15.Checked == true)
                    {
                        output = string.Empty;
                        j = 0;
                        for (i = 0; i < checkedListBox14.Items.Count; i++)
                        {
                            if (checkedListBox14.GetItemChecked(i))
                            {
                                if (j == 0)
                                    output += "'" + checkedListBox14.GetItemText(checkedListBox14.Items[i]) + "'";
                                if (j > 0)
                                    output += ",'" + checkedListBox14.GetItemText(checkedListBox14.Items[i]) + "'";
                                j++;
                            }
                        }
                        if (j > 0)
                            Sql = Sql + " and  PROCESSID in(" + output + ")";
                        else
                            MessageBox.Show("请至少选择一个PROCESSID值");
                    }
                    if (checkBox16.Checked == true)
                    {
                        output = string.Empty;
                        j = 0;
                        for (i = 0; i < checkedListBox15.Items.Count; i++)
                        {
                            if (checkedListBox15.GetItemChecked(i))
                            {
                                if (j == 0)
                                    output += "'" + checkedListBox15.GetItemText(checkedListBox15.Items[i]) + "'";
                                if (j > 0)
                                    output += ",'" + checkedListBox15.GetItemText(checkedListBox15.Items[i]) + "'";
                                j++;
                            }
                        }
                        if (j > 0)
                            Sql = Sql + " and  FAB in(" + output + ")";
                        else
                            MessageBox.Show("请至少选择一个FAB值");
                    }


                    Sql = Sql + " GROUP BY SingleName  ) as a1 ;";


                    //ReadFromSQL(Sql, "Map_SingleID", 5);
                    //ReadFromSQL("select * from ##tempa1 ", "Map_SingleID", 5);
                   // SendToSQL(Sql);
                   // MessageBox.Show("11");

                    Sql = Sql+ "select * into ##tempa2 from( ";
                    Sql = Sql + "select Model,LEFT(Q_ID,4)+SUBSTRING(Q_ID,13,1) AS Q_Name,sum(cast(SingleCount as bigint)) as Count from "
                          + "(select distinct  LotID,Model,Q_ID,SingleCount from Map_SingleID where 1=1 ";
                    if (checkBox1.Checked == true)
                        Sql = Sql + " and Date BETWEEN " + StartTime + " and " + EndTime;
                    if (checkBox2.Checked == true )
                    {
                        output = string.Empty;
                        j = 0;
                        for (i = 0; i < checkedListBox7.Items.Count; i++)
                        {
                            if (checkedListBox7.GetItemChecked(i))
                            {
                                if (j == 0)
                                    output += "'" + checkedListBox7.GetItemText(checkedListBox7.Items[i]) + "'";
                                if (j > 0)
                                    output += ",'" + checkedListBox7.GetItemText(checkedListBox7.Items[i]) + "'";
                                j++;
                            }
                        }
                        if (j > 0)
                            Sql = Sql + " and  Model in(" + output + ")";
                        else
                            MessageBox.Show("请至少选择一个Model值");
                    }


                    if (checkBox3.Checked == true)
                    {
                        output = string.Empty;
                        j = 0;
                        for (i = 0; i < checkedListBox8.Items.Count; i++)
                        {
                            if (checkedListBox8.GetItemChecked(i))
                            {
                                if (j == 0)
                                    output += "'" + checkedListBox8.GetItemText(checkedListBox8.Items[i]) + "'";
                                if (j > 0)
                                    output += ",'" + checkedListBox8.GetItemText(checkedListBox8.Items[i]) + "'";
                                j++;
                            }
                        }
                        if (j > 0)
                            Sql = Sql + " and  batch in(" + output + ")";
                        else
                            MessageBox.Show("请至少选择一个batch值");
                    }



                    if (checkBox5.Checked == true)
                    {
                        output = string.Empty;
                        j = 0;
                        for (i = 0; i < checkedListBox10.Items.Count; i++)
                        {
                            if (checkedListBox10.GetItemChecked(i))
                            {
                                if (j == 0)
                                    output += "'" + checkedListBox10.GetItemText(checkedListBox10.Items[i]) + "'";
                                if (j > 0)
                                    output += ",'" + checkedListBox10.GetItemText(checkedListBox10.Items[i]) + "'";
                                j++;
                            }
                        }
                        if (j > 0)
                            Sql = Sql + " and  Operation in(" + output + ")";
                        else
                            MessageBox.Show("请至少选择一个Operation值");
                    }

                    if (checkBox6.Checked == true)
                    {
                        output = string.Empty;
                        j = 0;
                        for (i = 0; i < checkedListBox11.Items.Count; i++)
                        {
                            if (checkedListBox11.GetItemChecked(i))
                            {
                                if (j == 0)
                                    output += "'" + checkedListBox11.GetItemText(checkedListBox11.Items[i]) + "'";
                                if (j > 0)
                                    output += ",'" + checkedListBox11.GetItemText(checkedListBox11.Items[i]) + "'";
                                j++;
                            }
                        }
                        if (j > 0)
                            Sql = Sql + " and  Lot_ID in(" + output + ")";
                        else
                            MessageBox.Show("请至少选择一个Lot_ID值");
                    }

                    //  if (checkBox7.Checked == true)
                    //     Sql = Sql + " and  EQP_ID=" + "'" + comboBox6.Text + "'";
                    if (checkBox7.Checked == true)
                    {
                        output = string.Empty;
                        j = 0;
                        for (i = 0; i < checkedListBox12.Items.Count; i++)
                        {
                            if (checkedListBox12.GetItemChecked(i))
                            {
                                if (j == 0)
                                    output += "'" + checkedListBox12.GetItemText(checkedListBox12.Items[i]) + "'";
                                if (j > 0)
                                    output += ",'" + checkedListBox12.GetItemText(checkedListBox12.Items[i]) + "'";
                                j++;
                            }
                        }
                        if (j > 0)
                            Sql = Sql + " and  EQP_ID in(" + output + ")";
                        else
                            MessageBox.Show("请至少选择一个EQP_ID值");
                    }

                    if (checkBox15.Checked == true)
                    {
                        output = string.Empty;
                        j = 0;
                        for (i = 0; i < checkedListBox14.Items.Count; i++)
                        {
                            if (checkedListBox14.GetItemChecked(i))
                            {
                                if (j == 0)
                                    output += "'" + checkedListBox14.GetItemText(checkedListBox14.Items[i]) + "'";
                                if (j > 0)
                                    output += ",'" + checkedListBox14.GetItemText(checkedListBox14.Items[i]) + "'";
                                j++;
                            }
                        }
                        if (j > 0)
                            Sql = Sql + " and  PROCESSID in(" + output + ")";
                        else
                            MessageBox.Show("请至少选择一个PROCESSID值");
                    }
                    if (checkBox16.Checked == true)
                    {
                        output = string.Empty;
                        j = 0;
                        for (i = 0; i < checkedListBox15.Items.Count; i++)
                        {
                            if (checkedListBox15.GetItemChecked(i))
                            {
                                if (j == 0)
                                    output += "'" + checkedListBox15.GetItemText(checkedListBox15.Items[i]) + "'";
                                if (j > 0)
                                    output += ",'" + checkedListBox15.GetItemText(checkedListBox15.Items[i]) + "'";
                                j++;
                            }
                        }
                        if (j > 0)
                            Sql = Sql + " and  FAB in(" + output + ")";
                        else
                            MessageBox.Show("请至少选择一个FAB值");
                    }


                    Sql = Sql + ") as m1 group by Model,LEFT(Q_ID,4)+SUBSTRING(Q_ID,13,1) ) as a2 ;";
                   // SendToSQL(Sql);

                   Sql = Sql + "Select * into ##tempa3 from( ";
                   Sql = Sql+ "Select a.*,b.Q_Name ,b.Count as Count1 ,cast(COALESCE(round(CAST(a.Count as FLOAT)/b.Count,4),0)*100 as varchar(10))+'%' as DefectRate "
                                   + " from  ##tempa1 as a left join ##tempa2 as b on left(a.SingleName,1) = SubString(b.Q_Name,5,1) ) as a3 ;";


                   /* Sql = Sql + "Select a.*,b.Q_Name ,b.Count ,cast(COALESCE(round(CAST(a.Count as FLOAT)/b.Count,4),0)*100 as varchar(10))+'%' as DefectRate "
                                    + " from  ##tempa1 as a left join ##tempa2 as b on left(a.SingleName,1) = SubString(b.Q_Name,5,1)";*/
                  // SendToSQL(Sql);
                  // Sql = Sql + " Select * from ##tempa3 ";
                   //MessageBox.Show(Sql);
                   //SendToSQL(Sql);
                 //  ReadFromSQL(Sql, "##tempa3", 5);
                 //  MessageBox.Show("111");
                    

                   DataTable dt5 = new DataTable();
                   dt5 = database.getDs("select " +listBox1.SelectedItem.ToString()+","+ listBox1.SelectedItem.ToString() +"_X," +listBox1.SelectedItem.ToString()+ "_Y "+" from  Pattern_coord ").Tables[0];
                  // dataGridView5.DataSource = dt5;
                   Product_Type = listBox1.SelectedItem.ToString();
                   //MessageBox.Show(Product_Type);
                   DataTableToSQLServer1(dt5, "Map_Data");

                   Sql = Sql + "select Map_Data.*,##tempa3.* from  Map_Data left join  ##tempa3 on Map_Data.PanelID = ##tempa3.SingleName where Map_Data.PanelID is not NULL ";
                 

                   ReadFromSQL(Sql, "Map_SingleID", 5);
             
                 /*  DataTable dt4 = new DataTable();
                   dt4 = database.getDs("select ParamterID from  Parameter_Panel ").Tables[0];

                   // MessageBox.Show(dt4.Rows.Count.ToString());
                   for (int i = 0; i < dt4.Rows.Count; i++)
                   {
                       listBox1.Items.Add(dt4.Rows[i][0].ToString());

                   }*/

                  //  ReadFromSQL(Sql, "Map_SingleID", 5);


                    Map("1");

                   // MessageBox.Show("3");
                    //ReadFromSQL(Sql, "Map_SingleID", 6);

                    MessageBox.Show("数据查询完成");
                }
                else
                {
                    MessageBox.Show("请选择Operation & Model");
                }
            }
            catch
            {
               
                this.Cursor = Cursors.Default;
                MessageBox.Show("处理异常13");
            }
          
            this.Cursor = Cursors.Default;


        }
       

        public static void DataTableToSQLServer1(DataTable dt, string tableName)
        {
           
            using (SqlConnection destinationConnection = new SqlConnection(SqlHelper.conStr))
            {
                destinationConnection.Open();
                using (SqlBulkCopy bulkCopy = new SqlBulkCopy(destinationConnection))
                {
                    try
                    {
                            bulkCopy.DestinationTableName = tableName;//要插入的表的表名
                     
                            bulkCopy.ColumnMappings.Add(Product_Type, "PanelID");//映射字段名 DataTable列名 ,数据库 对应的列名  
                            //MessageBox.Show("Product_Type");
                            bulkCopy.ColumnMappings.Add( Product_Type+"_X","X");
                            bulkCopy.ColumnMappings.Add(Product_Type+"_Y","Y");


                        bulkCopy.WriteToServer(dt);
                    }
                    catch (Exception ex)
                    {
                        System.Windows.Forms.MessageBox.Show(ex.Message);
                    }
                    finally
                    {
                        if (destinationConnection.State == ConnectionState.Open)
                        {
                            destinationConnection.Close();
                        }
                    }
                }
            }
        }

      

         #endregion

      
      /*  public void WriteToSQL()
        {
   
            conn = new SqlConnection(str);
            conn.Open();
            if (BMDT_Flag == 1 || BMDT_Flag == 2)
            {
                SqlCommand cmd = new SqlCommand("delete from db_15.dbo.Table_1", conn);
                cmd.CommandTimeout = 0;
                cmd.ExecuteNonQuery();
            }
         
            if (BMDT_Flag == 3)
            {
                SqlCommand cmd = new SqlCommand("delete from db_15.dbo.EQP_Info", conn);
                cmd.CommandTimeout = 0;
                cmd.ExecuteNonQuery();
            }
         
     

            if (dataGridView1.Rows.Count > 0)
            {
                DataRow dr = null;
               
                for (int i = 0; i < dt1.Rows.Count; i++)
                {
                    dr = dt1.Rows[i];
                    insertToSql(dr);
                  
                }
                conn.Close();
                MessageBox.Show("成功导入" + dt1.Rows.Count.ToString()+"条数据！");
            }
            else
            {
                MessageBox.Show("没有数据！");
            }
        }
        private void insertToSql(DataRow dr)
        {
            if (BMDT_Flag == 1 )
            {
                //excel表中的列名和数据库中的列名一定要对应  
                string ID = dr["ID"].ToString();
                string CREATETIME = dr["CREATETIME"].ToString();
                string REASONCODE1 = dr["REASONCODE1"].ToString();
                string PROCESSID = dr["PROCESSID"].ToString();
               // string FAB = checkedListBox13.SelectedItem.ToString();
                string FAB = "BMDT";
         

                string sql = "insert into db_15.dbo.Table_1 values('" + ID + "','" + CREATETIME + "','" + REASONCODE1 + "','" + PROCESSID + "','" + FAB + "')";
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.CommandTimeout = 0;
                cmd.ExecuteNonQuery();
            }
            else if (BMDT_Flag == 2)
            {
              
                string PANELID = dr["PANELID"].ToString();
                string EVENTTIME = dr["EVENTTIME"].ToString();
                string DESCRIPTION = dr["DESCRIPTION"].ToString();
                string PROCESSID = dr["PROCESSID"].ToString();
                //string FAB = checkedListBox13.SelectedItem.ToString();
                string FAB = "BMDT";

                string sql = "insert into db_15.dbo.Table_1 values('" + PANELID + "','" + EVENTTIME + "','" + DESCRIPTION + "','" + PROCESSID + "','" + FAB + "')";
                
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.CommandTimeout = 0;
                cmd.ExecuteNonQuery();
            }
            else if (BMDT_Flag == 3)
            {
                //excel表中的列名和数据库中的列名一定要对应  
                string Operation = dr["Operation"].ToString();
                string Lot_ID = dr["Lot_ID"].ToString();
                string EQP_ID = dr["EQP_ID"].ToString();
                string Event_Time = dr["Event_Time"].ToString();

                string sql =  "insert into db_15.dbo.EQP_Info values('" + Operation + "','" + Lot_ID + "','" + EQP_ID + "','" + Event_Time + "')";
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.CommandTimeout = 0;
                cmd.ExecuteNonQuery();
            }
           
        }*/

        private void insertToSql1(String str1,String str2)
        {
            string sql = "insert into Q_SingleCount values('" + str1 + "','" + str2 + "')";
            SqlCommand cmd = new SqlCommand(sql, conn);
            cmd.CommandTimeout = 0;
            cmd.ExecuteNonQuery();          
        }
        private void insertToSql2(String str1, String str2)
        {
            string sql = "insert into code values('" + str1 + "','" + str2 + "')";
            SqlCommand cmd = new SqlCommand(sql, conn);
            cmd.CommandTimeout = 0;
            cmd.ExecuteNonQuery();
        }
      

       
        #region /* 添加过滤值 */
        private void tabControl1_Selected(object sender, TabControlEventArgs e)
        {
            if (tabControl1.SelectedIndex == 2 || tabControl1.SelectedIndex == 1)
            {
                conn = new SqlConnection(str);
                conn.Open();
                string Sql = "select distinct Model from Defect_Data where Date> cast(getdate()-180 as bigint) order by Model  ";
                DataSet Ds = new DataSet();
                SqlDataAdapter Da1 = new SqlDataAdapter(Sql, conn);
                Da1.Fill(Ds, "Model");

                checkedListBox7.DataSource = Ds.Tables["Model"];
                checkedListBox7.DisplayMember = "Model";

          
                checkedListBox1.DataSource = Ds.Tables["Model"];
                checkedListBox1.DisplayMember = "Model";





                Sql = "select distinct batch from Defect_Data where Date> cast(getdate()-180 as bigint) order by batch";
                SqlDataAdapter Da2 = new SqlDataAdapter(Sql, conn);
                Da2.Fill(Ds, "batch");
                checkedListBox8.DataSource = Ds.Tables["batch"];
                checkedListBox8.DisplayMember = "batch";
                checkedListBox2.DataSource = Ds.Tables["batch"];
                checkedListBox2.DisplayMember = "batch";

                Sql = "select distinct DefectCode from Defect_Data where Date> cast(getdate()-180 as bigint) order by DefectCode";
                SqlDataAdapter Da3 = new SqlDataAdapter(Sql, conn);
                Da3.Fill(Ds, "DefectCode");
                checkedListBox9.DataSource = Ds.Tables["DefectCode"];
                checkedListBox9.DisplayMember = "DefectCode";
                checkedListBox3.DataSource = Ds.Tables["DefectCode"];
                checkedListBox3.DisplayMember = "DefectCode";


                Sql = "select distinct Operation from Defect_Data where Operation!='NULL' and Date> cast(getdate()-180 as bigint) order by Operation";
                SqlDataAdapter Da4 = new SqlDataAdapter(Sql, conn);
                Da4.Fill(Ds, "Operation");
                checkedListBox10.DataSource = Ds.Tables["Operation"];
                checkedListBox10.DisplayMember = "Operation";
                checkedListBox4.DataSource = Ds.Tables["Operation"];
                checkedListBox4.DisplayMember = "Operation";



                Sql = "select distinct LotID from Defect_Data  where left(LotID,3)!='1A5' and  Date> cast(getdate()-180 as bigint) order by LotID";
                SqlDataAdapter Da5 = new SqlDataAdapter(Sql, conn);
                Da5.Fill(Ds, "LotID");
                checkedListBox11.DataSource = Ds.Tables["LotID"];
                checkedListBox11.DisplayMember = "LotID";
                checkedListBox5.DataSource = Ds.Tables["LotID"];
                checkedListBox5.DisplayMember = "LotID";

                Sql = "select distinct EQP_ID from Defect_Data where EQP_ID!='NULL' and Date> cast(getdate()-180 as bigint) order by EQP_ID";
                SqlDataAdapter Da6 = new SqlDataAdapter(Sql, conn);
                Da6.Fill(Ds, "EQP_ID");
                checkedListBox12.DataSource = Ds.Tables["EQP_ID"];
                checkedListBox12.DisplayMember = "EQP_ID";
                checkedListBox6.DataSource = Ds.Tables["EQP_ID"];
                checkedListBox6.DisplayMember = "EQP_ID";

                Sql = "select distinct PROCESSID from Defect_Data where PROCESSID!='NULL' and Date> cast(getdate()-180 as bigint) order by PROCESSID";
                SqlDataAdapter Da7 = new SqlDataAdapter(Sql, conn);
                Da7.Fill(Ds, "PROCESSID");
                checkedListBox14.DataSource = Ds.Tables["PROCESSID"];
                checkedListBox14.DisplayMember = "PROCESSID";
                checkedListBox16.DataSource = Ds.Tables["PROCESSID"];
                checkedListBox16.DisplayMember = "PROCESSID";


                Sql = "select distinct FAB from Defect_Data where FAB!='NULL' and Date> cast(getdate()-180 as bigint) order by FAB";
                SqlDataAdapter Da8 = new SqlDataAdapter(Sql, conn);
                Da8.Fill(Ds, "FAB");
                checkedListBox15.DataSource = Ds.Tables["FAB"];
                checkedListBox15.DisplayMember = "FAB";
                checkedListBox17.DataSource = Ds.Tables["FAB"];
                checkedListBox17.DisplayMember = "FAB";

                 
            }

        }
        #endregion

     

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                this.Cursor = Cursors.WaitCursor;
                conn = new SqlConnection(str);
                conn.Open();
                SqlCommand cmd4 = new SqlCommand("delete from Defect_Data_temp", conn);
                cmd4.CommandTimeout = 0;
                cmd4.ExecuteNonQuery();
                conn.Close();

                conn = new SqlConnection(str);
                conn.Open();
                SqlCommand cmd5 = new SqlCommand("delete from Defect_Data", conn);
                cmd5.CommandTimeout = 0;
                cmd5.ExecuteNonQuery();
                conn.Close();

                conn = new SqlConnection(str);
                conn.Open();
                SqlCommand cmd6 = new SqlCommand("delete from Map_SingleID", conn);
                cmd6.CommandTimeout = 0;
                cmd6.ExecuteNonQuery();
                conn.Close();

           
                this.Cursor = Cursors.Default;
                MessageBox.Show("库信息删除完成");
            }
            catch
            {

                this.Cursor = Cursors.Default;
                MessageBox.Show("处理异常14");
            }

        }
        private void button14_Click(object sender, EventArgs e)
        {
            try
            {
                this.Cursor = Cursors.WaitCursor;
                // string c1 = "", c2 = "";
                conn = new SqlConnection(str);
                conn.Open();
                SqlCommand cmd1 = new SqlCommand("select count(*) cnt from Map_SingleID", conn);
                cmd1.CommandTimeout = 0;
                SqlDataReader dr1 = cmd1.ExecuteReader();
                dr1.Read();
                String c1 = dr1["cnt"].ToString();
                conn.Close();

                conn = new SqlConnection(str);
                conn.Open();
                SqlCommand cmd2 = new SqlCommand("select count(*) cnt from Defect_Data", conn);
                cmd2.CommandTimeout = 0;
                SqlDataReader dr2 = cmd2.ExecuteReader();
                dr2.Read();
                String c2 = dr2["cnt"].ToString();
                conn.Close();

                conn = new SqlConnection(str);
                conn.Open();
                SqlCommand cmd3 = new SqlCommand("select count(*) cnt from Raw_Data", conn);
                cmd3.CommandTimeout = 0;
                SqlDataReader dr3 = cmd3.ExecuteReader();
                dr3.Read();
                String c3 = dr3["cnt"].ToString();
                conn.Close();

                label5.Text = "Panel信息库Map_SingleID共有数据" + c1 + "条\n"
                                + "Lot良率信息库Defect_Data共有数据" + c2 + "条\n"
                                + "Lot良率信息库(汇总后)Raw_Data共有数据" + c3 + "条\n";
                this.Cursor = Cursors.Default;
            }
            catch
            {

                this.Cursor = Cursors.Default;
                MessageBox.Show("处理异常14");
            }

        }

        private void button9_Click(object sender, EventArgs e)
        {
            string TempS = string.Empty;
            string[] aa = new string[100];

            chart1.Visible = true;
            dataGridView2.Visible = false;
            chart1.Series[0].Points.Clear();


            chart1.Series[0].ChartType = SeriesChartType.Spline;
           // chart1.Series[1].ChartType = SeriesChartType.Spline; 

            /* https://www.cnblogs.com/topmount/p/8430689.html */
            chart1.ChartAreas[0].AxisX.Interval = 1;   //设置X轴坐标的间隔为1
            chart1.ChartAreas[0].AxisX.IntervalOffset = 1;  //设置X轴坐标偏移为1
            //chart1.ChartAreas[0].AxisX.LabelStyle.IsStaggered = true;   //设置是否交错显示,比如数据多的时间分成两行来显示 
            chart1.ChartAreas[0].AxisX.LabelStyle.Angle = -45;
            //chart1.Series.is

            chart1.ChartAreas[0].AxisY.LabelStyle.Format = "0%";
            chart1.ChartAreas[0].AxisY.Minimum = 0.7;
            chart1.ChartAreas[0].AxisY.Maximum = 1.0;
            chart1.ChartAreas[0].AxisY.Interval = 0.05;
            //chart1.ChartAreas[0].AxisX.MajorGrid.Enabled = false;
           // chart1.ChartAreas[0].AxisY.MajorGrid.Enabled = false;//不显示网格线
            chart1.Legends[0].Enabled = false;//不显示图例
            chart1.ChartAreas[0].BackColor = Color.White;//设置背景为白色
           // chart1.ChartAreas[0].Area3DStyle.Enable3D = true;//设置3D效果
          //  chart1.ChartAreas[0].Area3DStyle.PointDepth = 50;
           // chart1.ChartAreas[0].Area3DStyle.PointGapDepth = 50;//设置一下深度，看起来舒服点……
           // chart1.ChartAreas[0].Area3DStyle.WallWidth = 0;//设置墙的宽度为0；
            chart1.Series[0].Label = "#VAL{P}";//设置标签文本 (在设计期通过属性窗口编辑更直观) 标签变成百分数
            chart1.Series[0].IsValueShownAsLabel = true;//显示标签
           // chart1.Series[0].CustomProperties = "DrawingStyle=Cylinder, PointWidth=1";//设置为圆柱形 (在设计期通过属性窗口编辑更直观)
            chart1.Series[0].Palette = System.Windows.Forms.DataVisualization.Charting.ChartColorPalette.Pastel;//设置调色板


            for (int j = 0; j < dataGridView7.RowCount-1; j++)
            {
              TempS = dataGridView7.Rows[j].Cells[1].Value.ToString();
             /* aa=TempS.Split('%');
              chart1.Series[0].Points.AddXY(dataGridView7.Rows[j].Cells[0].Value.ToString(),
                  Convert.ToDouble(aa[0])/100);*/
              chart1.Series[0].Points.AddXY(dataGridView7.Rows[j].Cells[0].Value.ToString(),
                              Convert.ToDouble(TempS));

           
            }

        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void tabPage4_Click(object sender, EventArgs e)
        {

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                string FilePath = Environment.CurrentDirectory+"\\map\\"+listBox1.SelectedItem.ToString() + ".jpg";
                if (listBox1.SelectedIndex.ToString() != null)
                {
                    // MessageBox.Show(Environment.CurrentDirectory + "\\DesktopFile\\map\\" + listBox5.SelectedItem.ToString() + ".jpg");
                    /*在c盘就报警！！！*/
                    if (File.Exists(FilePath))
                    {
                        Stream s = File.Open(FilePath, FileMode.Open);
                        pictureBox1.Image = Image.FromStream(s);
                        s.Close();
                    }
                    else
                    {

                        MessageBox.Show("你选择的map不存在");
                    }

                }
            }
            catch
            {

                MessageBox.Show("处理异常4");

            }

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

      

        public void Map(string e)
        {

            try
            {
                float x = 0, y = 0,temp1=0,temp2=0,temp3=0;
                int Tem1 = 0, Tem2 = 0;
                float CellSizeX = 0, CellSizeY = 0,DefectRate =0 ;
                string str=string.Empty;

                Image myimage = new Bitmap(pictureBox1.Width, pictureBox1.Height);

               // Graphics dg = this.CreateGraphics();
                Graphics g = Graphics.FromImage(myimage);

                g.Clear(Color.White);
                Pen mypen1 = new Pen(Color.Red, 2);
                Pen mypen2 = new Pen(Color.Blue, 1);
                Pen mypen3 = new Pen(Color.Black, 1);

                g.DrawRectangle(mypen1, 0, 0, 750, 650);
                g.TranslateTransform(375, 325);


                DataTable dt6 = new DataTable();
               // dt1 = database.getDs("select * from " + listBox1.SelectedItem.ToString()).Tables[0];
                dt6 = database.getDs("select * from  Parameter_Panel where ParamterID =" +"'"+ listBox1.SelectedItem.ToString()+"'").Tables[0];


           
                    CellSizeX = float.Parse(dt6.Rows[0][2].ToString()) / 2;
                    CellSizeY = float.Parse(dt6.Rows[0][3].ToString()) / 2;


                    
                    for (int i = 0; i < dataGridView5.RowCount - 1; i++)
                    {
                        if (dataGridView5.Rows[i].Cells[7].Value.ToString() != "")
                        {
                            if (temp1 > temp2)
                                temp2 = temp1;
                            if (temp1 < temp3)
                                temp3 = temp1;
                            temp1 = float.Parse(dataGridView5.Rows[i].Cells[7].Value.ToString().Substring(0, (dataGridView5.Rows[i].Cells[7].Value.ToString().Length - 1)));
                        }

                    
                    }
                    temp1 = temp2 - temp3;

                    //http://www.114la.com/other/rgb.htm


                  for (int i = 0; i < dataGridView5.RowCount-1; i++)
                  {
                      //计算panel的左上角
                     // x = float.Parse(dataGridView5.Rows[i].Cells[1].Value.ToString()) / 2 - float.Parse((CellSizeX / 2).ToString());
                     // y = -float.Parse(dataGridView5.Rows[i].Cells[2].Value.ToString()) / 2 - float.Parse((CellSizeY / 2).ToString());
                      x = float.Parse(dataGridView5.Rows[i].Cells[1].Value.ToString())/2 ;
                      y = float.Parse(dataGridView5.Rows[i].Cells[2].Value.ToString())/2 ;
                      
                    // g.DrawRectangle(mypen2, x, y, CellSizeX, CellSizeY);

                    //  if()
                      if (dataGridView5.Rows[i].Cells[7].Value.ToString() != "")
                      {

                      DefectRate = float.Parse(dataGridView5.Rows[i].Cells[7].Value.ToString().Substring(0, (dataGridView5.Rows[i].Cells[7].Value.ToString().Length - 1)));
                      if (DefectRate <= temp1 / 3)
                      {
                          g.FillRectangle(new System.Drawing.SolidBrush(System.Drawing.Color.FromArgb(255, 245, 245)), x + 2, y + 2, CellSizeX - 2, CellSizeY - 2);//画实心椭圆
                          g.DrawRectangle(mypen3, x + 1, y + 1, CellSizeX - 1, CellSizeY - 1);
                      }
                      if (DefectRate > temp1 / 3)
                      {
                          g.FillRectangle(new System.Drawing.SolidBrush(System.Drawing.Color.FromArgb(255, 175, 175)), x + 2, y + 2, CellSizeX - 2, CellSizeY - 2);//画实心椭圆
                          g.DrawRectangle(mypen3, x + 1, y + 1, CellSizeX - 1, CellSizeY - 1);
                      }
                      if (DefectRate > temp1 * 2 / 3)
                      {
                          g.FillRectangle(new System.Drawing.SolidBrush(System.Drawing.Color.FromArgb(255, 105, 105)), x + 2, y + 2, CellSizeX - 2, CellSizeY - 2);//画实心椭圆
                          g.DrawRectangle(mypen3, x + 1, y + 1, CellSizeX - 1, CellSizeY - 1);
                      }
                           
                      }
                      else
                          //g.FillRectangle(new System.Drawing.SolidBrush(System.Drawing.Color.FromArgb(255, 255, 255)), x + 2, y + 2, CellSizeX - 2, CellSizeY - 2);//画实心椭圆
                          g.DrawRectangle(mypen3, x + 2, y + 2, CellSizeX - 2, CellSizeY - 2);
                          


                      Brush brush = System.Drawing.Brushes.Black;
                      //MessageBox.Show(listBox1.SelectedItem.ToString().Substring(3, 1));

                      if (listBox1.SelectedItem.ToString().Substring(3,1) == "V")
                          g.DrawString(dataGridView5.Rows[i].Cells[7].Value.ToString(), new Font("微软雅黑", 10, FontStyle.Regular), brush, new PointF(x + CellSizeX / 4, y + CellSizeY / 6), new StringFormat(StringFormatFlags.DirectionVertical));  
                      else
                          g.DrawString(dataGridView5.Rows[i].Cells[7].Value.ToString(), new Font("微软雅黑", 10), brush, new PointF(x + CellSizeX / 4, y + CellSizeY / 6));  

                     myimage.Save(@"d:\image.png");


                  }



                 pictureBox1.Image = myimage;
              

             
            }
            catch
            {
                MessageBox.Show("处理异常8：Map生成异常");
            }
        }

        private void Form1_Paint(object sender, PaintEventArgs e)
        {
           

        }


      

    }
}
