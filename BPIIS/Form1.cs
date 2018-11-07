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

            //myGridView.Add(new BridgeInspection("", 0, 0));

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
                myGridView.Insert(e.RowIndex, new BridgeInspection("", 0, 0));
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

            public BridgeInspection(string comment, decimal stdValue, decimal calcValue)
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

        //将项目信息写入excel
        private void button2_Click_1(object sender, EventArgs e)
        {
            string source = @"桥隧项目管理系统导入模板-空白.xlsx";
            string destination = @"桥隧项目管理系统导入模板-导出.xlsx";

            try
            {
                FileInfo sourceFile = new FileInfo(source);
                FileInfo destinationFile = null;
                if (!File.Exists(destination))    //不存在则复制
                {
                    destinationFile = sourceFile.CopyTo(destination, true);
                }
                else    //存在则直接打开
                {
                    destinationFile = new FileInfo(destination);
                }
                

                using (ExcelPackage package = new ExcelPackage(destinationFile))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["合同"];
                    //worksheet.Cells["A1"].Value = "名称";//直接指定单元格进行赋值

                    //以追加方式写入

                    int rowIndex = 2;   //写入行

                    //首行：表头不导入
                    bool rowCur = true;    //行游标指示器
                                           //rowCur=false表示到达行尾
                    while (rowCur)
                    {
                        try
                        {
                            //跳过表头
                            if (string.IsNullOrEmpty(worksheet.Cells[rowIndex, 1].Value.ToString()))
                            {
                                rowCur = false;
                            }
                        }
                        catch (Exception)   //读取异常则终止
                        {
                            rowCur = false;
                        }

                        if (rowCur)
                        {
                            rowIndex++;
                        }
                    }

                    //写入excel
                    worksheet.Cells[rowIndex, 1].Value = (rowIndex-1).ToString();    //序号
                    worksheet.Cells[rowIndex, 2].Value = textBox12.Text;    //合同编号
                    worksheet.Cells[rowIndex, 3].Value = textBox1.Text;    //合同名称
                    worksheet.Cells[rowIndex, 4].Value = textBox8.Text;    //合同金额
                    worksheet.Cells[rowIndex, 5].Value = textBox2.Text;    //合同签订日期
                    worksheet.Cells[rowIndex, 6].Value = textBox10.Text;    //合同期限
                    worksheet.Cells[rowIndex, 7].Value = textBox6.Text;    //合同约定工作内容
                    worksheet.Cells[rowIndex, 8].Value = textBox3.Text;    //项目地点
                    worksheet.Cells[rowIndex, 9].Value = textBox4.Text;    //委托单位
                    worksheet.Cells[rowIndex, 10].Value = textBox5.Text;    //委托单位联系人
                    worksheet.Cells[rowIndex, 11].Value = textBox7.Text;    //委托单位联系人电话
                    //承接方式
                    if(radioButton1.Checked)
                    {
                        worksheet.Cells[rowIndex, 12].Value = 1;
                    }
                    else if(radioButton2.Checked)
                    {
                        worksheet.Cells[rowIndex, 12].Value = 2;
                    }
                    else if (radioButton3.Checked)
                    {
                        worksheet.Cells[rowIndex, 12].Value = 3;
                    }
                    else if (radioButton4.Checked)
                    {
                        worksheet.Cells[rowIndex, 12].Value = 4;
                    }
                    //合同签订状态
                    if (radioButton5.Checked)
                    {
                        worksheet.Cells[rowIndex, 13].Value = 1;
                    }
                    else if (radioButton6.Checked)
                    {
                        worksheet.Cells[rowIndex, 13].Value = 2;
                    }
                    else if (radioButton7.Checked)
                    {
                        worksheet.Cells[rowIndex, 13].Value = 3;
                    }
                    else if (radioButton8.Checked)
                    {
                        worksheet.Cells[rowIndex, 13].Value = 4;
                    }
                    //合同承接人工号
                    worksheet.Cells[rowIndex, 14].Value = textBox11.Text;    //委托单位联系人电话

                    package.Save();
                    label19.Text = $"成功写入！当前已写入{rowIndex - 1}行";
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

        //同步合同文件名
        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                label12.Text = listBox1.SelectedItems[0].ToString();
            }
            catch (Exception)
            {

                label12.Text = "无";
            }
            
        }
        //同步项目文件名
        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                label29.Text = listBox2.SelectedItems[0].ToString();
            }
            catch (Exception)
            {

                label29.Text = "无";
            }
            
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

            checkBox5.Checked = projectRepository.IsExistBearingCapacity(doc);

            //栏杆水平推力
            if(projectRepository.IsExistRailThrusting(doc))
            {
                checkBox6.Checked = true;
                myGridView.Add(new BridgeInspection("栏杆推力", 0, 0));
            }
            
        }

        //设置检测类型字符串
        private string SetInspectionString()
        {
            List<string> stringList = new List<string>();
            string inspString = "";    //最终结果
            string tempString = "";

            //常规定期检测
            if(checkBox1.Checked)
            {
                tempString = "常规定期检测";
                if(!checkBox7.Checked)
                {
                    tempString = $"{tempString},{textBox23.Text},{textBox24.Text}";
                }
                stringList.Add(tempString);
            }

            //结构定期检测
            if (checkBox2.Checked)
            {
                tempString = "结构定期检测";
                if (!checkBox8.Checked)
                {
                    tempString = $"{tempString},{textBox25.Text},{textBox26.Text}";
                }
                stringList.Add(tempString);
            }

            //静力荷载试验
            if (checkBox3.Checked)
            {
                tempString = "静力荷载试验";
                if (Convert.ToInt32(textBox21.Text)>1)
                {
                    tempString = $"{tempString},{textBox21.Text}";
                }
                stringList.Add(tempString);
            }

            //动力荷载试验
            if (checkBox4.Checked)
            {
                tempString = "动力荷载试验";
                if (Convert.ToInt32(textBox22.Text) > 1)
                {
                    tempString = $"{tempString},{textBox22.Text}";
                }
                stringList.Add(tempString);
            }

            //承载能力检算(不含基础)
            if (checkBox5.Checked)
            {
                tempString = "承载能力检算(不含基础)";
                stringList.Add(tempString);
            }

            //其它
            if(checkBox6.Checked)
            {
                tempString = "";
                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    tempString = $"{tempString}{dataGridView1.Rows[i].Cells[0].Value.ToString()},{dataGridView1.Rows[i].Cells[1].Value.ToString()},{dataGridView1.Rows[i].Cells[2].Value.ToString()}";              
                    //不是最后一行
                    if (i!= dataGridView1.RowCount-1)
                    {
                        tempString = $"{tempString};";
                    }    
                }
                stringList.Add(tempString);
            }

            inspString = "";
            for (int i=0;i<stringList.Count;i++)
            {
                inspString = $"{inspString}{stringList[i]}";
                if(i!=stringList.Count-1)
                {
                    inspString = $"{inspString};";
                }
            }

            return inspString;

        }

        //生成检测字符串
        private void button9_Click(object sender, EventArgs e)
        {
            textBox18.Text = SetInspectionString();
        }

        //常规定检
        private void checkBox7_CheckedChanged(object sender, EventArgs e)
        {
            //常规定检检测全桥则检测参数无效
            if(checkBox7.Checked)
            {
                textBox23.Enabled = false;
                textBox24.Enabled = false;
                textBox23.Text = "";
                textBox24.Text = "";
            }
            else
            {
                textBox23.Enabled = true;
                textBox24.Enabled = true;

            }
        }

        //结构定检
        private void checkBox8_CheckedChanged(object sender, EventArgs e)
        {
            //结构定检检测全桥则检测参数无效
            if (checkBox8.Checked)
            {
                textBox25.Enabled = false;
                textBox26.Enabled = false;
                textBox25.Text = "";
                textBox26.Text = "";
            }
            else
            {
                textBox25.Enabled = true;
                textBox26.Enabled = true;
            }
        }


        //桥梁写入excel
        private void button10_Click(object sender, EventArgs e)
        {
            string source = @"桥隧项目管理系统导入模板-空白.xlsx";
            string destination = @"桥隧项目管理系统导入模板-导出.xlsx";

            try
            {
                FileInfo sourceFile = new FileInfo(source);
                FileInfo destinationFile = null;
                if (!File.Exists(destination))    //不存在则复制
                {
                    destinationFile = sourceFile.CopyTo(destination, true);
                }
                else    //存在则直接打开
                {
                    destinationFile = new FileInfo(destination);
                }


                using (ExcelPackage package = new ExcelPackage(destinationFile))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["桥梁"];

                    //以追加方式写入
                    int rowIndex = 2;   //写入行

                    //首行：表头不导入
                    bool rowCur = true;    //行游标指示器
                                           //rowCur=false表示到达行尾
                    while (rowCur)
                    {
                        try
                        {
                            //跳过表头
                            if (string.IsNullOrEmpty(worksheet.Cells[rowIndex, 1].Value.ToString()))
                            {
                                rowCur = false;
                            }
                        }
                        catch (Exception)   //读取异常则终止
                        {
                            rowCur = false;
                        }

                        if (rowCur)
                        {
                            rowIndex++;
                        }
                    }

                    //写入excel
                    //桥梁名称	桥梁地理位置	桥梁长度	桥梁宽度	桥梁跨数	桥梁结构形式	备注
                    worksheet.Cells[rowIndex, 1].Value = (rowIndex - 1).ToString();    //序号
                    worksheet.Cells[rowIndex, 2].Value = textBox15.Text;    //桥梁名称
                    worksheet.Cells[rowIndex, 3].Value = textBox27.Text;    //地理位置
                    worksheet.Cells[rowIndex, 4].Value = textBox16.Text;    //桥长
                    worksheet.Cells[rowIndex, 5].Value = textBox17.Text;    //桥宽
                    worksheet.Cells[rowIndex, 6].Value = textBox20.Text;    //跨/联数
                    worksheet.Cells[rowIndex, 7].Value = (comboBox1.SelectedIndex + 1).ToString();    //结构形式
                    worksheet.Cells[rowIndex, 8].Value = textBox19.Text+(String.IsNullOrEmpty(textBox9.Text)?"":" 最大跨径："+textBox9.Text);     //备注

                    package.Save();
                    label41.Text = $"成功写入！当前已写入{rowIndex - 1}行";
                }
                
            }
            catch (Exception)
            {
                label41.Text = "写入异常，请联系管理员";
            }
        }
        //项目写入excel
        private void button7_Click(object sender, EventArgs e)
        {
            string source = @"桥隧项目管理系统导入模板-空白.xlsx";
            string destination = @"桥隧项目管理系统导入模板-导出.xlsx";

            try
            {
                FileInfo sourceFile = new FileInfo(source);
                FileInfo destinationFile = null;
                if (!File.Exists(destination))    //不存在则复制
                {
                    destinationFile = sourceFile.CopyTo(destination, true);
                }
                else    //存在则直接打开
                {
                    destinationFile = new FileInfo(destination);
                }


                using (ExcelPackage package = new ExcelPackage(destinationFile))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["项目"];

                    //以追加方式写入
                    int rowIndex = 2;   //写入行

                    //首行：表头不导入
                    bool rowCur = true;    //行游标指示器
                                           //rowCur=false表示到达行尾
                    while (rowCur)
                    {
                        try
                        {
                            //跳过表头
                            if (string.IsNullOrEmpty(worksheet.Cells[rowIndex, 1].Value.ToString()))
                            {
                                rowCur = false;
                            }
                        }
                        catch (Exception)   //读取异常则终止
                        {
                            rowCur = false;
                        }

                        if (rowCur)
                        {
                            rowIndex++;
                        }
                    }

                    //写入excel
                    worksheet.Cells[rowIndex, 1].Value = (rowIndex - 1).ToString();    //序号
                    worksheet.Cells[rowIndex, 2].Value = textBox14.Text;    //关联合同编号
                    worksheet.Cells[rowIndex, 3].Value = textBox13.Text;    //项目名称
                    worksheet.Cells[rowIndex, 4].Value = textBox15.Text;    //关联桥梁
                    worksheet.Cells[rowIndex, 5].Value = SetInspectionString();//textBox18.Text;    //检测类型

                    package.Save();
                    label40.Text = $"成功写入！当前已写入{rowIndex - 1}行";
                }
                
            }
            catch (Exception)
            {
                label40.Text = "写入异常，请联系管理员";
            }
        }

        //"跨/联"跟收费标准同步（福建省城市桥梁检测评估费用定额）
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(comboBox1.SelectedItem.ToString().Contains("连续梁桥")
                || comboBox1.SelectedItem.ToString().Contains("连续刚构桥"))
            {
                label30.Text = "联数";
                label25.Text = "联";
                label26.Text = "联";
            }
            else
            {
                label30.Text = "跨数";
                label25.Text = "跨";
                label26.Text = "跨";
            }
        }


    }



}
