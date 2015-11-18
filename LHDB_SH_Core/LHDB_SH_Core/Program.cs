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

            catalog01.Add(new RibbonFeature("LHDB_SH_Core.Report.SubjectReport", "科目名冊"));

            RibbonBarItem item04 = K12.Presentation.NLDPanels.Student.RibbonBarItems["學習歷程資料"];
            item04["報表"].Image = Properties.Resources.Report;
            item04["報表"].Size = RibbonBarButton.MenuButtonSize.Large;
            item04["報表"]["科目名冊"].Enable = UserAcl.Current["LHDB_SH_Core.Report.SubjectReport"].Executable;
            item04["報表"]["科目名冊"].Click += delegate
            {
                if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0)
                {
                    Report.SubjectReport ar = new Report.SubjectReport(K12.Presentation.NLDPanels.Student.SelectedSource);
                    ar.ShowDialog();
                }
                else
                {
                    FISCA.Presentation.Controls.MsgBox.Show("請選擇學生!");
                }
            };


            catalog01.Add(new RibbonFeature("LHDB_SH_Core.Report.SubjectScoreReport", "成績名冊"));

            RibbonBarItem item05 = K12.Presentation.NLDPanels.Student.RibbonBarItems["學習歷程資料"];
            item05["報表"].Image = Properties.Resources.Report;
            item05["報表"].Size = RibbonBarButton.MenuButtonSize.Large;
            item05["報表"]["成績名冊"].Enable = UserAcl.Current["LHDB_SH_Core.Report.SubjectScoreReport"].Executable;
            item05["報表"]["成績名冊"].Click += delegate
            {
                if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0)
                {
                    Report.SubjectScoreReport ar = new Report.SubjectScoreReport(K12.Presentation.NLDPanels.Student.SelectedSource);
                    ar.ShowDialog();
                }
                else
                {
                    FISCA.Presentation.Controls.MsgBox.Show("請選擇學生!");
                }
            };

            catalog01.Add(new RibbonFeature("LHDB_SH_Core.Config.DepConfigForm", "部別班別代碼"));

            RibbonBarItem item06 = K12.Presentation.NLDPanels.Student.RibbonBarItems["學習歷程資料"];
            item06["設定"].Image = Properties.Resources.設定;
            item06["設定"].Size = RibbonBarButton.MenuButtonSize.Large;
            item06["設定"]["部別班別代碼"].Enable = UserAcl.Current["LHDB_SH_Core.Config.DepConfigForm"].Executable;
            item06["設定"]["部別班別代碼"].Click += delegate
            {
                Config.DepConfigForm dcg = new Config.DepConfigForm();
                dcg.ShowDialog();
            };

            catalog01.Add(new RibbonFeature("LHDB_SH_Core.Config.ClassCodeConfigForm", "班級代碼"));

            RibbonBarItem item07 = K12.Presentation.NLDPanels.Student.RibbonBarItems["學習歷程資料"];
            item07["設定"].Image = Properties.Resources.設定;
            item07["設定"].Size = RibbonBarButton.MenuButtonSize.Large;
            item07["設定"]["班級代碼"].Enable = UserAcl.Current["LHDB_SH_Core.Config.ClassCodeConfigForm"].Executable;
            item07["設定"]["班級代碼"].Click += delegate
            {
                Config.ClassCodeConfigForm ccf = new Config.ClassCodeConfigForm();
                ccf.ShowDialog();
            };


        }

        static void _bgLLoadUDT_DoWork(object sender, DoWorkEventArgs e)
        {
            // 建立使用到 UDT
            Utility.CreateUDTTable();
        }
    }
}
