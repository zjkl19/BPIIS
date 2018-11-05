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

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        IKernel kernel;
        IRepository.IContractRepository contractRepository;

        BindingList<Employee> mEmployees = new BindingList<Employee>();
        BindingSource mBbindingSource = new BindingSource();
        private void dataGridView1_Load()
        {
            mEmployees.Add(new Employee("Tom", 23));
            mEmployees.Add(new Employee("Harry", 24));
            mEmployees.Add(new Employee("John", 26));
            mBbindingSource.DataSource = mEmployees;
            //dataGridView1.Dock = DockStyle.Fill;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView1.AutoGenerateColumns = false;
            AddColumns();
            dataGridView1.DataSource = mBbindingSource;
            dataGridView1.CellClick +=
             new DataGridViewCellEventHandler(dataGridView1_CellClick);
        }


        private void AddColumns()
        {



            DataGridViewTextBoxColumn nameColumn = new DataGridViewTextBoxColumn();
            nameColumn.Name = "Name";
            nameColumn.DataPropertyName = "Name";



            DataGridViewTextBoxColumn ageColumn = new DataGridViewTextBoxColumn();
            ageColumn.Name = "Age";
            ageColumn.DataPropertyName = "Age";


            DataGridViewButtonColumn insertColumn =
             new DataGridViewButtonColumn();
            insertColumn.HeaderText = "";
            insertColumn.Name = "insertColumn";
            insertColumn.Text = "插入";
            insertColumn.UseColumnTextForButtonValue = true;


            DataGridViewButtonColumn deleteColumn =
            new DataGridViewButtonColumn();
            deleteColumn.HeaderText = "";
            deleteColumn.Name = "deleteColumn";
            deleteColumn.Text = "删除";
            deleteColumn.UseColumnTextForButtonValue = true;


            dataGridView1.Columns.Add(nameColumn);
            dataGridView1.Columns.Add(ageColumn);
            dataGridView1.Columns.Add(insertColumn);
            dataGridView1.Columns.Add(deleteColumn);




        }


        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {


            //增加空行
            if (e.ColumnIndex == 2)
            {
                mEmployees.Insert(e.RowIndex, new Employee("", 0));
            }

            //删除当前行
            if (e.ColumnIndex == 3)
            {
                mEmployees.RemoveAt(e.RowIndex);
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



        public class Employee
        {
            public Employee(String name)
            {
                nameValue = name;
            }
            public Employee(String name, int age)
            {
                nameValue = name;
                ageValue = age;
            }


            private String nameValue;
            public String Name
            {
                get { return nameValue; }
                set { nameValue = value; }
            }


            public Employee Self
            {
                get { return this; }
            }


            private int? ageValue;
            public int? Age
            {
                get { return ageValue; }
                set { ageValue = value; }
            }




        }



        public Form1()
        {
            InitializeComponent();
            kernel = new StandardKernel(new Infrastructure.NinjectDependencyResolver());
            contractRepository = kernel.Get<IRepository.IContractRepository>();
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

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            label12.Text = listBox1.SelectedItems[0].ToString();
        }

        private void textBox28_TextChanged(object sender, EventArgs e)
        {

        }


    }



}
