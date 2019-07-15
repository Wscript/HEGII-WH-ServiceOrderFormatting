using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;

namespace HEGII_WH_ServiceOrderFormatting
{
    public partial class Ribbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void ButtonCallerFormatting_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet ActiveSheet = Globals.ThisAddIn.Application.ActiveSheet;
            string[] TitleList_Survey = {"ID","业务单据号","当前业务单据流程","服务类型","客户编码","客户名称","客户手机","部门名称",
                                         "所属路线片区","所属片区","小区名称","客户地址","另约日期","安排服务技师","安排上门日期",
                                         "单据主表备注","回访备注","下一指定流程","历史服务记录","历史服务人员","服务产品名称汇总",
                                         "送货车牌号","送货司机","送货成本运费","送货服务收费","联系人","固定电话" };

            string[] TitleList_Delete = {"ID","业务单据号","当前业务单据流程","客户编码","客户名称","所属路线片区","小区名称","另约日期",
                                         "下一指定流程","历史服务记录","历史服务人员","送货车牌号","送货司机","送货成本运费","送货服务收费","固定电话" };

            string[] TitleList_Sequence = {"回访备注", "完成（是 / 否）", "单据主表备注", "服务产品名称汇总", "客户手机", "客户地址", "销售员",
                                           "所属片区", "联系人", "安排上门日期", "报装日期", "部门名称", "服务类型", "安排服务技师"};

            int[] Columns_Width = { 10, 10, 15, 8, 8, 8, 12, 8, 20, 10, 35, 25, 10, 25 };

            if (CheckFileFormat(TitleList_Survey, 6))
            {
                ActiveSheet.Range["1:5"].Delete();                                                                  //删除表头（去掉合并单元格）
                ColumnsDeltet(TitleList_Delete, 1);                                                                 //删除多余的列
                ColumnsSort(TitleList_Sequence, 1);                                                                 //按设定的列标题排序
                ColumnsWidthSet(Columns_Width);                                                                     //按列宽表设置列宽
                ActiveSheet.Range["A:N"].WrapText = true;                                                           //设置单元格自动换行
                ActiveSheet.Range["1:" + ActiveSheet.UsedRange.Rows.Count.ToString()].EntireRow.AutoFit();          //设置自动调整行高
                ActiveSheet.Application.ActiveWindow.FreezePanes = false;
            }
            else
            {
                MessageBox.Show("当前工作表(" + ActiveSheet.Name.ToString() + ")的格式错误，请确认需要整理的单据在当前页面！");
            }
        }
        private void ButtonForemanFormatting_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet ActiveSheet = Globals.ThisAddIn.Application.ActiveSheet;
            string[] TitleList_Survey = {"ID","业务单据号","当前业务单据流程","服务类型","客户编码","客户名称","客户手机","部门名称",
                                         "所属路线片区","所属片区","小区名称","客户地址","另约日期","安排服务技师","安排上门日期",
                                         "单据主表备注","回访备注" ,"下一指定流程","历史服务记录","历史服务人员","服务产品名称汇总",
                                         "送货车牌号","送货司机","送货成本运费","送货服务收费","联系人","固定电话" };
            string[] TitleList_Delete = {"ID","当前业务单据流程","客户编码","客户名称","所属路线片区","小区名称","另约日期","回访备注",
                                         "下一指定流程","历史服务记录","历史服务人员","送货车牌号","送货司机","送货成本运费","送货服务收费","固定电话" };

            string[] TitleList_Sequence = {"单据主表备注", "服务产品名称汇总", "客户地址", "所属片区", "客户手机", "联系人",
                                           "部门名称","安排上门日期", "服务类型", "安排服务技师","业务单据号"};

            int[] Columns_Width = { 10, 10, 10, 8, 15, 10, 10, 15, 25, 40, 30 };

            if (CheckFileFormat(TitleList_Survey, 6))
            {
                ActiveSheet.Range["1:5"].Delete();                                                                  //删除表头（去掉合并单元格）
                ColumnsDeltet(TitleList_Delete, 1);                                                                 //删除多余的列
                ColumnsSort(TitleList_Sequence, 1);                                                                 //按设定的列标题排序
                ColumnsWidthSet(Columns_Width);                                                                     //按列宽表设置列宽
                ActiveSheet.Range["A:N"].WrapText = true;                                                           //设置单元格自动换行
                ActiveSheet.Range["1:" + ActiveSheet.UsedRange.Rows.Count.ToString()].EntireRow.AutoFit();          //设置自动调整行高
            }
            else
            {
                MessageBox.Show("当前工作表(" + ActiveSheet.Name.ToString() + ")的格式错误，请确认需要整理的单据在当前页面！");
            }
        }
        private void ColumnsWidthSet(int[] Columns_Width)
        {
            Worksheet ActiveSheet = Globals.ThisAddIn.Application.ActiveSheet;
            for (int i = 0; i < Columns_Width.Length; i++)
            {
                ActiveSheet.Columns[i + 1].ColumnWidth = Columns_Width[i];
            }
        }
        private void ColumnsSort(string[] TitleList_Sequence, int TitleRow)
        {
            Worksheet ActiveSheet = Globals.ThisAddIn.Application.ActiveSheet;
            foreach (string TitleItem in TitleList_Sequence)
            {
                Range rangeFindItem = ActiveSheet.Range[TitleRow.ToString() + ":" + TitleRow.ToString()].Find(TitleItem);
                if (rangeFindItem != null)
                {
                    ActiveSheet.Columns[rangeFindItem.Column].Cut();
                    ActiveSheet.Columns[1].Insert();
                }
                else
                {
                    ActiveSheet.Columns[1].Insert();
                    ActiveSheet.Cells[TitleRow, 1].Value = TitleItem;
                }
            }
        }
        private void ColumnsDeltet(string[] TitleList_Delete, int TitleRow)
        {
            Worksheet ActiveSheet = Globals.ThisAddIn.Application.ActiveSheet;
            foreach (string TitleItem in TitleList_Delete)
            {
                Range rangeFindItem = ActiveSheet.Range[TitleRow.ToString() + ":" + TitleRow.ToString()].Find(TitleItem);
                if (rangeFindItem != null)
                {
                    ActiveSheet.Columns[rangeFindItem.Column].Delete();
                }
            }
        }
        private bool CheckFileFormat(string[] TitleList, int TitleRow)
        {
            Worksheet ActiveSheet = Globals.ThisAddIn.Application.ActiveSheet;
            bool boolNotFound = true;
            foreach (string TitleItem in TitleList)
            {
                Range rangeFindItem = ActiveSheet.Range[TitleRow.ToString() + ":" + TitleRow.ToString()].Find(TitleItem);
                if (rangeFindItem == null)
                {
                    boolNotFound = false;
                }
            }
            return (boolNotFound);
        }
    }
}
