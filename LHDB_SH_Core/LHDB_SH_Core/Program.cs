using FISCA.Permission;
using FISCA.Presentation;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace LHDB_SH_Core
{
    public class Program
    {
        static BackgroundWorker _bgLLoadUDT = new BackgroundWorker();

        [FISCA.MainMethod()]
        public static void Main()
        {

            _bgLLoadUDT.DoWork += _bgLLoadUDT_DoWork;
            _bgLLoadUDT.RunWorkerCompleted += _bgLLoadUDT_RunWorkerCompleted;
            _bgLLoadUDT.RunWorkerAsync();

        }

        static void _bgLLoadUDT_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            // 產生學習歷程
            Catalog catalog01 = RoleAclSource.Instance["學生"]["學習歷程資料"];
            catalog01.Add(new RibbonFeature("LHDB_SH_Core.Report.AttendanceReport", "缺勤紀錄名冊"));

            RibbonBarItem item01 = K12.Presentation.NLDPanels.Student.RibbonBarItems["學習歷程資料"];
            item01["報表"].Image = Properties.Resources.Report;
            item01["報表"].Size = RibbonBarButton.MenuButtonSize.Large;
            item01["報表"]["缺勤紀錄名冊"].Enable = UserAcl.Current["LHDB_SH_Core.Report.AttendanceReport"].Executable;
            item01["報表"]["缺勤紀錄名冊"].Click += delegate
            {
                if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0)
                {
                    Report.AttendanceReport ar = new Report.AttendanceReport(K12.Presentation.NLDPanels.Student.SelectedSource);
                    ar.ShowDialog();
                }
                else
                {
                    FISCA.Presentation.Controls.MsgBox.Show("請選擇學生!");
                }
            };

            catalog01.Add(new RibbonFeature("LHDB_SH_Core.Report.StudentDataNReport", "學生資料名冊(國教署主管學校)"));

            RibbonBarItem item02 = K12.Presentation.NLDPanels.Student.RibbonBarItems["學習歷程資料"];
            item02["報表"].Image = Properties.Resources.Report;
            item02["報表"].Size = RibbonBarButton.MenuButtonSize.Large;
            item02["報表"]["學生資料名冊(國教署主管學校)"].Enable = UserAcl.Current["LHDB_SH_Core.Report.StudentDataNReport"].Executable;
            item02["報表"]["學生資料名冊(國教署主管學校)"].Click += delegate
            {
                if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0)
                {
                    Report.StudentDataNReport ar = new Report.StudentDataNReport(K12.Presentation.NLDPanels.Student.SelectedSource);
                    ar.ShowDialog();
                }
                else
                {
                    FISCA.Presentation.Controls.MsgBox.Show("請選擇學生!");
                }
            };

            catalog01.Add(new RibbonFeature("LHDB_SH_Core.Report.StudentSCETakeScoreReport", "定期考查成績名冊"));

            RibbonBarItem item03 = K12.Presentation.NLDPanels.Student.RibbonBarItems["學習歷程資料"];
            item03["報表"].Image = Properties.Resources.Report;
            item03["報表"].Size = RibbonBarButton.MenuButtonSize.Large;
            item03["報表"]["定期考查成績名冊"].Enable = UserAcl.Current["LHDB_SH_Core.Report.StudentSCETakeScoreReport"].Executable;
            item03["報表"]["定期考查成績名冊"].Click += delegate
            {
                if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0)
                {
                    Report.StudentSCETakeScoreReport ar = new Report.StudentSCETakeScoreReport(K12.Presentation.NLDPanels.Student.SelectedSource);
                    ar.ShowDialog();
                }
                else
                {
                    FISCA.Presentation.Controls.MsgBox.Show("請選擇學生!");
                }
            };


        }

        static void _bgLLoadUDT_DoWork(object sender, DoWorkEventArgs e)
        {
            // 建立使用到 UDT
            Utility.CreateUDTTable();
        }
    }
}
