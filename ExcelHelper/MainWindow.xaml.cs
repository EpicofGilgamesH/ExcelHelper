using Model;
using Operation;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ExcelHelper
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            _filePathPort = string.Empty;
            _filePath = string.Empty;
            InitializeComponent();
            this.WindowStartupLocation = WindowStartupLocation.CenterScreen;
        }

        private string _filePath;
        private string _filePathPort;
        private void Choice_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.OpenFileDialog myopenFileDialog = new System.Windows.Forms.OpenFileDialog();
            myopenFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            myopenFileDialog.Filter = "*.xls|*.xlsx";
            myopenFileDialog.FilterIndex = 2;
            myopenFileDialog.RestoreDirectory = true;
            if (myopenFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                _filePath = myopenFileDialog.FileName;
                msg.Content = _filePath;
            }
        }

        private void Add_Click(object sender, RoutedEventArgs e)
        {
            Stopwatch sw = new Stopwatch();
            sw.Start();
            //picture.Source = new BitmapImage(new Uri("/images/tuanzi1.gif", UriKind.Relative));
            //string sheetName = "透视结果一";
            ReadExcelBase.ReadExcelBase reb = new ReadExcelBase.ReadExcelBase(_filePath);
            try
            {
                reb.ReadExcelToList(new DefalutModel());
            }
            catch (Exception ex)
            {
                msg.Content = "文件格式不对或者被操作的Excel文件未关闭，请联系管理员~ ^^" + ex;
            }
            //IList<DefalutModel> list = reb.List as IList<DefalutModel>;
            //string[] tittle = new string[13] { "城市", "区域", "片区", "店铺", "经纪人", "经纪人号码", "成交类型", "总佣金", "成交时间", "交易编号", "进线精准匹配", "客户来源", "文本" };
            string[] tittle = new string[12] { "城市", "区域", "片区", "店铺", "经纪人", "经纪人号码", "成交类型", "总佣金", "成交时间", "交易编号", "客户来源", "文本" };
            string[] colTittle = new string[7] { "城市", "售单", "售佣金", "租单", "租佣金", "总单数", "总佣金" };
            //导入
            try
            {
                WriteExcelBase.WriteExcelBase web = new WriteExcelBase.WriteExcelBase(filePathPort(_filePath), tittle, colTittle);
                IList<ModelBase> listmb = reb.List.ToList();
                //对集合进行一些列操作
                OperatDefault od = new OperatDefault();
                try
                {
                    //筛选
                    IEnumerable<ModelBase> mb = CheckBox.IsChecked == true ? od.FilterByMonth(listmb.ToList(), false) : od.FilterByMonth(listmb.ToList());
                    //排除
                    mb = od.FilterByColumn(mb.ToList());
                    //修复
                    mb = od.RepairByColumn(mb.ToList());
                    mb = od.GetExactCity(mb.ToList());
                    //排序
                    mb = od.OrederOperation(mb.ToList());
                    //导出
                    int i = web.ListToExcel(mb.ToList());
                    sw.Stop();
                    picture.Source = new BitmapImage(new Uri(GetRandomImagPath(), UriKind.Relative));
                    msg.Content = i == -1 ? string.Format("导出失败，请检查Excel格式") : string.Format("导出成功，共数据：{0}条,共耗时{1}ms", i, sw.ElapsedMilliseconds);
                    msg.Foreground = new SolidColorBrush(Color.FromArgb(255, 255, 0, 0));
                }
                catch
                {
                    msg.Content = "不包含所选月份的数据 或 Excel正在工作~";
                }
            }
            catch (Exception ex)
            {
                msg.Content = "文件数据为空，请检查月份是否对应~ ^^" + ex;
            }
        }

        protected static string filePathPort(string str)
        {
            FileInfo fi = new FileInfo(str);
            string fullName = fi.Name;
            string ext = fi.Extension;
            string name = fullName.Replace(ext, "");
            return str.Replace(name, name + "_Lizzy");
        }


        private void CheckBox_Checked(object sender, RoutedEventArgs e)
        {

        }

        private string GetRandomImagPath()
        {
            Random r = new Random();
            int i = r.Next(1, 10);
            return string.Format("/images/tuanzi{0}.jpg", i);
        }

    }
}
