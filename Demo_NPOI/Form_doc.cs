using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Demo_NPOI
{
    public partial class Form_doc : Form
    {
        public Form_doc()
        {
            InitializeComponent();
        }

        private void btn_Read_Click(object sender, EventArgs e)
        {
            ReadDocx();
        }

        void ReadDoc()
        {

            try
            {

                #region 打开文档
                string fileName = @"test.doc";
                if (!File.Exists(fileName))
                {
                    MessageBox.Show(string.Format("{0}不存在！", fileName));
                    return;
                }
                XWPFDocument document = null;
                using (FileStream file = new FileStream(fileName, FileMode.Open, FileAccess.Read))
                {
                    document = new XWPFDocument(file);
                }
                MessageBox.Show(string.Format("在文档中找到[{0}]个表格", document.Tables.Count));

                #endregion

                List<string> rowList = new List<string>();
                int iRow = 0;//表中行的循环索引
                int iCell = 0;//表中列的循环索引

                //表格里嵌套表格
                var cell = document.Tables[0].Rows[11].GetCell(1);
                var strValues = cell.Tables[0].Text;
                Console.WriteLine(strValues);
                MessageBox.Show(string.Format("{0}", strValues));

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.GetBaseException().Message);
            }
        }
        void ReadDocx() {

            try
            {

                #region 打开文档
                string fileName = @"test.docx";
                if (!File.Exists(fileName))
                {
                    MessageBox.Show(string.Format("{0}不存在！", fileName));
                    return;
                }
                XWPFDocument document = null;
                using (FileStream file = new FileStream(fileName, FileMode.Open, FileAccess.Read))
                {
                    document = new XWPFDocument(file);
                }
                MessageBox.Show(string.Format("在文档中找到[{0}]个表格", document.Tables.Count));

                #endregion

                List<string> rowList = new List<string>();
                int iRow = 0;//表中行的循环索引
                int iCell = 0;//表中列的循环索引

                //表格里嵌套表格
                var cell = document.Tables[0].Rows[11].GetCell(1);
                var strValues = cell.Tables[0].Text;
                Console.WriteLine(strValues);
                MessageBox.Show(string.Format("{0}", strValues));

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.GetBaseException().Message);
            }
        }
    }
}
