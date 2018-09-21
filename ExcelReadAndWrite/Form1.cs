using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.Data.Odbc;
using System.IO;

/**
 * 作者：XieGuoPing_ Wutiezhu 
 * 参考：https://www.cnblogs.com/mora1988/p/5715097.html
 *       http://www.cnblogs.com/restran/p/3889479.html
 *       http://www.bubuko.com/infodetail-292769.html
 *       https://blog.okazuki.jp/entry/20091128/1259405232
 *       初次接触NPOI算是学习上面的（Excel读写部分）内容一个小结，Excel格式的设定还没有试
 *       我用的数据库是 Odbc 连的 MariaDB 10
 *       
 * **/
namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
		//选择你所安装的ODBC驱动 服务器
        string connectionString = "Driver={MariaDB ODBC 3.0 Driver};DBQ=localhost;password = yourPassword;uid = root;DATABASE = yourDataBase";
        //参见帮助文档 OdbcConnection
        OdbcConnection odbcConn;
        OdbcDataAdapter oDA;
        DataSet oDs = new DataSet();
        string queryText = "select * from acc_userinfo";
        //文件格式过滤规则
        string filterReg = "Excel 2003兼容格式|*.xls|Excel 2007文件格式|*.xlsx";    
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "就绪        ";
        }

        private void saveToFile_Click(object sender, EventArgs e)
        {
            #region  最开始只是输出一个格子
            //IWorkbook  wbook = new XSSFWorkbook();
            //ISheet  sheet = wbook.CreateSheet("sheet1") ;
            //IRow row = sheet.CreateRow(0);
            //ICell cell = row.CreateCell(0);
            //cell.SetCellValue("this is test cell text ");
            //SaveFileDialog save = new SaveFileDialog();
            //save.Filter = "xlsx|*.xlsx";
            //DialogResult isSave = new DialogResult();
            //isSave = save.ShowDialog();
            //if (isSave == DialogResult.OK)
            //{
            //    FileStream stream = new FileStream(save.FileName,FileMode.Create);
            //    wbook.Write(stream);
            // }
            #endregion

            try   //有时文件读写出问题
            {
                //创建一个工作薄
                IWorkbook wbook = null;
                //创建一个工作表
                ISheet sheet = null;

                //如果dataGridView 有数据
                if (this.dataGridView1.Rows.Count > 0)
                {
                    //创建保存对话框
                    SaveFileDialog saveFile = new SaveFileDialog();
                    //设置文件过滤器
                    saveFile.Filter = filterReg;
                    //保存用户的操作
                    DialogResult isSave = new DialogResult();
                    isSave = saveFile.ShowDialog();
                    if (DialogResult.OK == isSave && saveFile.FileName != null)
                    {
                        wbook = CreateIWorkBook(@saveFile.FileName.ToString(),null);
                        //this.toolStripProgressBar1.Visible = true;
                        toolStripStatusLabel1.Text = "正在保存……";
                        //创建文件流
                        FileStream fs = new FileStream(saveFile.FileName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                        //给sheet赋值
                        sheet = wbook.CreateSheet("Sheet1");
                        IRow row = null;
                        ICell cell = null;
                        int hasTitle = 0;    //标记是否包含表头
						//表头选项
                        if (includeHead.Checked)
                        {
                            row = sheet.CreateRow(0);
                            for (int i = 0; i < this.dataGridView1.ColumnCount; ++i)
                            {
                                //将 dataGridView1的表头写入到文件中
                                cell = row.CreateCell(i);
                                cell.SetCellValue(this.dataGridView1.Columns[i].HeaderText);
                            }
                            hasTitle = 1;
                        }
                        //如果有表头则行从1开始，多创建一行，否则从0开始
                        for (int rowIndex = hasTitle; rowIndex < this.dataGridView1.Rows.Count + hasTitle; ++rowIndex)
                        {
                            //创建数据行
                            row = sheet.CreateRow(rowIndex);
                            //创建行内单元格
                            for (int cellIndex = 0; cellIndex < this.dataGridView1.Columns.Count; ++cellIndex)
                            {
                                cell = row.CreateCell(cellIndex);
                                cell.SetCellValue(this.dataGridView1[cellIndex, rowIndex - hasTitle].Value.ToString());
                                this.progressBar1.Value += 100 / (this.dataGridView1.Rows.Count * this.dataGridView1.Columns.Count);    //只是用于试验用的，无实际用途
                            }
                        }
                        //写入文件流
                        wbook.Write(fs);
                        this.toolStripProgressBar1.Value = 100;                  
                        MessageBox.Show("文件保存成功！");
                        toolStripStatusLabel1.Text = "就绪        ";
                        this.toolStripProgressBar1.Value = 0;
                        //this.progressBar1.Visible = false;
                    }
                    else
                    {
                        //MessageBox.Show("用户取消文件保存！");
                        return;
                    }
                }
                else    //如果datagridview 没有数据 
                {
                    MessageBox.Show("没有可用数据可以导出！");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }



        private void readDataBase_Click(object sender, EventArgs e)
        {
            this.dataGridView1.DataSource = null;
            DataTable dataTable = new DataTable();
            oDs.Clear();
            //DataRow dr = dataTable.NewRow();
            odbcConn = new OdbcConnection(connectionString);
            try
            {
                odbcConn.Open();
                oDA = new OdbcDataAdapter(queryText, odbcConn);
                oDA.Fill(oDs);
                dataTable = oDs.Tables[0];
                this.dataGridView1.DataSource = dataTable;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void readFromFile_Click(object sender, EventArgs e)
        {
            //this.progressBar1.Visible = true;
            DataTable excelTable = new DataTable();

            IWorkbook wbook = null;
            ISheet sheet = null;

            //1 先打开文件
            OpenFileDialog openFile = new OpenFileDialog();
            //设置文件过滤器
            openFile.Filter = filterReg;
            DialogResult opened = new DialogResult();
            opened = openFile.ShowDialog();
            //用户点了 OK 或 Cancel
            if (DialogResult.OK == opened && openFile.FileName != null)
            {
                //2判断当前dataGridView 是否有数据,有则清空，无则不管
                if (dataGridView1.Rows.Count > 0)
                    dataGridView1.DataSource = null;
                //3打开文件流,只读形式打开
                FileStream fs = new FileStream(openFile.FileName, FileMode.Open, FileAccess.Read);
                //4判断文件格式
                wbook = CreateIWorkBook(@openFile.FileName, fs);
                //关闭文件流
                fs.Close();
                sheet = wbook.GetSheet("Sheet1");
                if (sheet == null)
                    sheet = wbook.GetSheetAt(0);
                IRow row = null;
                //ICell cell = null;
                int hasTitle = 0;
                //如果Excel包含表头
                if (includeHead.Checked)
                {
                    row = sheet.GetRow(0);
                    for (int i = 0; i < row.LastCellNum; i++)
                    {
                        excelTable.Columns.Add(row.GetCell(i).ToString());
                        //因为本身上没有数据列的，所以下面的方式只能用于事先创建列了的dataGridView,但是一般人都不是先知除非有格式规定
                        //this.dataGridView1.Columns [i].HeaderText = row.GetCell(i).ToString ();              
                    }
                    hasTitle = 1;
                }
                else    //如果Excel不包含表头，创建空列
                {
                    row = sheet.GetRow(0);
                    for (int j = 0; j < row.LastCellNum; j++)
                    {
                        excelTable.Columns.Add();
                    }
                }
                //这里ISheet.LastRowNum 返回的值为（RowCount - 1）,所以要用 <=,
                for (int rowIndex = hasTitle; rowIndex <= sheet.LastRowNum; rowIndex++)
                {
                    //加载RowIndex行的数据
                    row = sheet.GetRow(rowIndex);
                    //excelTable添加一个新数据行
                    DataRow dRow = excelTable.NewRow();

                    for (int columnIndex = 0; columnIndex < row.LastCellNum; columnIndex++)
                    {
                        // if(row.GetCell(columnIndex).ToString () != null)           //这里本想是如果为null值就给它一个默认值，但是结果全为空
                        dRow[columnIndex] = row.GetCell(columnIndex).ToString();
                        //dRow[columnIndex] = "";                                    //同上
                    }
                    excelTable.Rows.Add(dRow);
                }
                this.dataGridView1.DataSource = excelTable;
            }
            else
            {
                //MessageBox.Show("用户取消打开文件！");
                return;
            }

        }
		#region  由下面的函数代替了
/*
        /// <summary>
        /// 根据文件格式返回表文件
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <returns>IWorkBook 类型</returns>
        public IWorkbook CreateIWorkBook(string filePath)
        {
            IWorkbook wbook = null;
            //读取用户选择的保存格式
            string extesion = (Path.GetExtension(filePath)).ToLower();         //当然不能排除有大写后缀的可能性，还有极少数的大小写混合的后缀
            if (extesion.Equals(".xlsx"))
            //if( saveFile.FileName.IndexOf(".xlsx") > 0);
            {
                wbook = new XSSFWorkbook();
            }
            else if (extesion.Equals(".xls"))
            //else if (saveFile.FileName.IndexOf(".xls") > 0) ;            //这种方法也可以实现但.xlsx必须在前面
            {
                wbook = new HSSFWorkbook();
            }
            else    //如果不是上述两种格式,因为设置了过滤器几乎不会出现下面的情况
            {
                MessageBox.Show("对不起，您所输入的文件格式不受支持！");
                return null;
            }
            return wbook;
        }
*/
#endregion
        /// <summary>
        /// 根据文件格式返回表文件
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <param name="fs">文件流</param>
        /// <returns>IWorkBook 表</returns>
        public IWorkbook CreateIWorkBook(string filePath, FileStream fs)
        {
            IWorkbook wbook = null;
            //读取用户选择的保存格式
            string extesion = (Path.GetExtension(filePath)).ToLower();
            if (extesion.Equals(".xlsx"))
            //if( saveFile.FileName.IndexOf(".xlsx") > 0);
            {
                wbook = new XSSFWorkbook(fs);
            }
            else if (extesion.Equals(".xls"))
            //else if (saveFile.FileName.IndexOf(".xls") > 0) ;    //这种方法也可以实现但.xlsx必须在前面
            {
                wbook = new HSSFWorkbook(fs);
            }
            else    //如果不是上述两种格式,因为设置了过滤器几乎不会出现下面的情况
            {
                MessageBox.Show("对不起，您所输入的文件格式不受支持！");
                return null;
            }
            return wbook;
        }
    }
}
