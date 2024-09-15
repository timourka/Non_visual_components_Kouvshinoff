using Non_visual_components_Kouvshinoff.Enums;
using Non_visual_components_Kouvshinoff.InfoModels;
using Range = Non_visual_components_Kouvshinoff.InfoModels.Range;

namespace test
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string[] strings = { "aboba1", "abobfgsdgsdfgdsfgdsgsdgsdfgdfsgdfsgsdfga2" };
            using var dialog = new SaveFileDialog { Filter = "xlsx|*.xlsx" };
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    customComponentExcelBigText.createExcel(dialog.FileName, "aboba", strings);
                    MessageBox.Show("Выполнено", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private class Student
        {
            public int number;
            public string name;
            public int doneExam, notDoneExam;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            List<Student> students = new List<Student>
            {
                new Student { number = 1, doneExam = 10, name = "Timur", notDoneExam = 0 },
                new Student { number = 2, doneExam = 0, name = "Andrei", notDoneExam = 10 },
                new Student { number = 3, doneExam = 5, name = "Avos", notDoneExam = 5 }
            };

            List<ColumnInfo> header = new List<ColumnInfo>();
            ColumnInfo col1 = new ColumnInfo("№", 5, "number");
            ColumnInfo col2 = new ColumnInfo("имя", 20, "name");
            ColumnInfo col3 = new ColumnInfo("сдано", 6, "doneExam");
            ColumnInfo col4 = new ColumnInfo("несдано", 6.5, "notDoneExam");
            ColumnInfo colExam = new ColumnInfo("экзамены", new List<ColumnInfo>{ col3, col4 });
            ColumnInfo colStu = new ColumnInfo("студент", new List<ColumnInfo> { col2, colExam });
            header.Add(col1);
            header.Add(colStu);

            using var dialog = new SaveFileDialog { Filter = "xlsx|*.xlsx" };
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    customComponentExcelTableWithHeader.createExcel(dialog.FileName, "aboba", header, students);
                    MessageBox.Show("Выполнено", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            List<string> header = new List<string> { "a", "b", "c", "d" };
            List<Range> ranges = new List<Range>{
                new("aboba", new Dictionary<string, int> { 
                    { "a", 11 },
                    { "b", 22 },
                    { "c", 33 },
                    { "d", 44 }
                    }),
                new("babab", new Dictionary<string, int> {
                    { "a", 44 },
                    { "b", 33 },
                    { "c", 1 },
                    { "d", 11 }
                    }),
            };

            using var dialog = new SaveFileDialog { Filter = "xlsx|*.xlsx" };
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    customComponentExcelLineDiagram.createExcel(dialog.FileName, "sadf", "adsfa", DiagramLegendLocation.Left, header, ranges);
                    MessageBox.Show("Выполнено", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
    }
}
