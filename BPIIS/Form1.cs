using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Aspose.Words;
using Aspose.Words.Tables;
using OfficeOpenXml;

using System.Windows.Forms;

//新代码
using Ninject;
using BPIIS.IRepository;

namespace BPIIS
{
    public partial class Form1 : Form
    {
        IKernel kernel;
        IContractRepository contractRepository;
        IProjectRepository projectRepository;

        BindingList<BridgeInspection> myGridView = new BindingList<BridgeInspection>();
        BindingSource mBbindingSource = new BindingSource();

        private void dataGridView1_Load()
        {
            //myGridView.Add(new BridgeInspection("栏杆推力", 30000,750));
            myGridView.Add(new BridgeInspection("", 0, 0));
            
            mBbindingSource.DataSource = myGridView;
            //dataGridView1.Dock = DockStyle.Fill;    //挤满
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView1.AutoGenerateColumns = false;
            AddColumns();
            //AddColumns();

            dataGridView1.DataSource = mBbindingSource;
            dataGridView1.CellClick +=
             new DataGridViewCellEventHandler(dataGridView1_CellClick);
        }

        //添加新的自定义检测类型
        private void button8_Click(object sender, EventArgs e)
        {
            myGridView.Add(new BridgeInspection("", 0, 0));
        }

        private void AddColumns()
        {

            //“检测类型”在数据库中为“备注”字段
            DataGridViewTextBoxColumn commentColumn = new DataGridViewTextBoxColumn
            {
                Name = "检测类型",
                DataPropertyName = "Comment"
            };

            DataGridViewTextBoxColumn stdValueColumn = new DataGridViewTextBoxColumn
            {
                Name = "标准产值",
                DataPropertyName = "stdValue"
            };

            DataGridViewTextBoxColumn calcValueColumn = new DataGridViewTextBoxColumn
            {
                Name = "计算产值",
                DataPropertyName = "calcValue"
            };

            DataGridViewButtonColumn insertColumn =
             new DataGridViewButtonColumn
             {
                 HeaderText = "",
                 Name = "insertColumn",
                 Text = "插入",
                 UseColumnTextForButtonValue = true
             };


            DataGridViewButtonColumn deleteColumn =
            new DataGridViewButtonColumn
            {
                HeaderText = "",
                Name = "deleteColumn",
                Text = "删除",
                UseColumnTextForButtonValue = true
            };

            dataGridView1.Columns.Add(commentColumn);
            dataGridView1.Columns.Add(stdValueColumn);
            dataGridView1.Columns.Add(calcValueColumn);
            dataGridView1.Columns.Add(insertColumn);
            dataGridView1.Columns.Add(deleteColumn);

        }


        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {


            //增加空行
            if (e.ColumnIndex == 3)
            {
                myGridView.Insert(e.RowIndex, new BridgeInspection("", 0,0));
            }

            //删除当前行
            if (e.ColumnIndex == 4)
            {
               myGridView.RemoveAt(e.RowIndex);
            }
        }


        private void dataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            MessageBox.Show(e.Exception.Message.ToString());

        }


        //对列1进行进一步验证
        private void dataGridView1_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            if (e.ColumnIndex == 0)
            {
                if (e.FormattedValue.ToString() == "123")
                {
                    MessageBox.Show("不能为123");
                    e.Cancel = true;
                    dataGridView1.CancelEdit();
                }
                else
                {
                    e.Cancel = false;
                }
            }
        }



        public class BridgeInspection
        {
            public string Comment { get; set; }

            public decimal StdValue { get; set; }

            public decimal CalcValue { get; set; }

            public BridgeInspection(string comment,decimal stdValue,decimal calcValue)
            {
                Comment = comment;
                StdValue = stdValue;
                CalcValue = calcValue;
            }





        }



        public Form1()
        {
            InitializeComponent();
            kernel = new StandardKernel(new Infrastructure.NinjectDependencyResolver());
            contractRepository = kernel.Get<IContractRepository>();
            projectRepository = kernel.Get<IProjectRepository>();
            dataGridView1_Load();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            // The path to the documents directory.
            //string dataDir = GetDataDir_QuickStart();

            // Create a blank document.
            //Document doc = new Document();

            //// DocumentBuilder provides members to easily add content to a document.
            //DocumentBuilder builder = new DocumentBuilder(doc);

            //// Write a new paragraph in the document with the text "Hello World!"
            //builder.Writeln("Hello World!");

            //// Save the document in DOCX format. The format to save as is inferred from the extension of the file name.
            //// Aspose.Words supports saving any document in many more formats.
            ////dataDir = dataDir + "HelloWorld_out.docx";
            //doc.Save("HelloWorld_out.docx");

            //MessageBox.Show("想松李博头！");

            //打开word文档，fileName是路径地址，需要扩展名
            string fileName = "南平市莲花大桥-改.doc";
            Aspose.Words.Document doc = new Document(fileName);

            //获取word文档中的第一个表格
            var table0 = doc.GetChildNodes(NodeType.Table, true)[1] as Aspose.Words.Tables.Table;

            Aspose.Words.Tables.Cell cell = table0.Rows[0].Cells[2];
            //用GetText()的方法来获取cell中的值
            string cbfbm = cell.GetText();
            cbfbm = cbfbm.Replace("\a", "");
            cbfbm = cbfbm.Replace("\r", "");

            textBox1.Text = cbfbm;
            //MessageBox.Show(cbfbm);

        }

        //智能读取合同
        private void button3_Click(object sender, EventArgs e)
        {
            //IKernel kernel = new StandardKernel(new Infrastructure.NinjectDependencyResolver());
            //IRepository.IContractRepository contractRepository = kernel.Get<IRepository.IContractRepository>();

            //TODO：增加匹配不到时的异常处理
            //打开word文档，fileName是路径地址，需要扩展名
            //string fileName = "合同--后屿路桥检测-褚工改.doc";

            string fileName = listBox1.SelectedItems[0].ToString();    //多选只算第一个

            string rootPath = Directory.GetCurrentDirectory();

            Aspose.Words.Document doc = new Document($"{rootPath}\\合同\\{fileName}");

            string originalWholeText = doc.Range.Text;    //原始全文

            //半角括号替换为全角括号
            string wholeText = originalWholeText.Replace("(", "（");
            wholeText = wholeText.Replace(")", "）");

            //合同编号
            textBox12.Text = contractRepository.GetNo(wholeText);
            //合同名称
            textBox1.Text = contractRepository.GetName(wholeText);

            //合同金额
            textBox8.Text = contractRepository.GetAmount(wholeText);

            //合同签订日期
            textBox2.Text = contractRepository.GetSignedDate(wholeText);

            //合同地点
            textBox3.Text = contractRepository.GetProjectLocation(wholeText);

            //合同约定工作内容
            textBox6.Text = contractRepository.GetJobContent(wholeText);

            //委托单位
            textBox4.Text = contractRepository.GetClient(wholeText);

            //委托单位联系人

            textBox5.Text = contractRepository.GetClientContactPerson(wholeText);

            //委托单位联系人电话

            textBox7.Text = contractRepository.GetClientContactPersonPhone(wholeText);

            textBox10.Text = contractRepository.GetDeadline(wholeText);

            //MessageBox.Show(cbfbm);
        }


        private void button2_Click_1(object sender, EventArgs e)
        {
            string source = @"桥隧项目管理系统导入模板-空白.xlsx";
            string destination = @"桥隧项目管理系统导入模板-导出.xlsx";

            try
            {
                FileInfo sourceFile = new FileInfo(source);
                FileInfo destinationFile = sourceFile.CopyTo(destination, true);

                using (ExcelPackage package = new ExcelPackage(destinationFile))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["合同"];
                    //worksheet.Cells["A1"].Value = "名称";//直接指定单元格进行赋值
                    worksheet.Cells[2, 3].Value = textBox1.Text;//直接指定行列数进行赋值
                    package.Save();
                }
            }
            catch (Exception)
            {

                label19.Text = "写入异常，请联系管理员";
            }



        }

        //读取所有项目word文件
        private void button5_Click(object sender, EventArgs e)
        {
            listBox2.Items.Clear();

            string rootPath = Directory.GetCurrentDirectory();

            DirectoryInfo folder = new DirectoryInfo($"{rootPath}\\项目");

            foreach (FileInfo file in folder.GetFiles("*.doc"))
            {
                listBox2.Items.Add(file.Name);
            }
        }

        //读取所有合同word文件
        private void button4_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();

            string rootPath = Directory.GetCurrentDirectory();

            DirectoryInfo folder = new DirectoryInfo($"{rootPath}\\合同");

            foreach (FileInfo file in folder.GetFiles("*.doc"))
            {
                listBox1.Items.Add(file.Name);
            }

            //集成测试代码
            //StreamReader sr = new StreamReader("项目\\a.txt", Encoding.Default);
            //String line;
            //while ((line = sr.ReadLine()) != null)
            //{
            //    MessageBox.Show(line.ToString());
            //}

        }

        //同步合同文件
        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            label12.Text = listBox1.SelectedItems[0].ToString();
        }
        //同步项目文件
        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            label29.Text = listBox2.SelectedItems[0].ToString();
        }

        private void textBox28_TextChanged(object sender, EventArgs e)
        {

        }

        //识别项目信息
        private void button6_Click(object sender, EventArgs e)
        {
            string rootPath = Directory.GetCurrentDirectory();
            string fileName = listBox2.SelectedItems[0].ToString();    //多选只算第一个
            
            Document doc = new Document($"{rootPath}\\项目\\{fileName}");
                  
            textBox13.Text = projectRepository.GetName(doc);

            textBox14.Text = projectRepository.GetContractNo(doc);

            textBox15.Text = projectRepository.GetBridgeName(doc);

            checkBox1.Checked = projectRepository.IsExistRegularPeriod(doc);

            checkBox2.Checked = projectRepository.IsExistStructurePeriod(doc);

            checkBox3.Checked = projectRepository.IsExistStaticLoad(doc);

            checkBox4.Checked = projectRepository.IsExistDynamicLoad(doc);

        }


    }



}
