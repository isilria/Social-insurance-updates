using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.VBA;

[assembly: AssemblyTitle("사회보험 재원별 대사 보조 도우미")]
[assembly: AssemblyProduct("사회보험 재원별 대사 보조 도우미")]
[assembly: AssemblyDescription("급여대장과 사회보험 부과자료의 재원별 대사 및 제출서 생성을 돕는 프로그램")]
[assembly: AssemblyVersion("2.0.2.0")]
[assembly: AssemblyFileVersion("2.0.2.0")]

namespace InsurancePayrollValidator
{
    static class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            AppDomain.CurrentDomain.AssemblyResolve += delegate(object sender, ResolveEventArgs e)
            {
                if (new AssemblyName(e.Name).Name == "EPPlus")
                {
                    using (Stream s = Assembly.GetExecutingAssembly().GetManifestResourceStream("InsurancePayrollValidator.EPPlus.dll"))
                    {
                        byte[] b = new byte[s.Length]; s.Read(b, 0, b.Length); return Assembly.Load(b);
                    }
                }
                return null;
            };
            if(args.Length>=7&&args[0]=="--convert-package")
            {
                InputSet input=new InputSet{PayrollPackage=args[2],HealthGov=args[3],Pension=args[4],Employment=args[5],Industrial=args[6],ShortTermPayroll=args.Length>=8?args[7]:""};
                try{Processor.Run(input,args[1]);}
                catch(Exception ex){File.WriteAllText(args[1]+".error.txt",ex.ToString(),Encoding.UTF8);Environment.ExitCode=1;}return;
            }
            if (args.Length >= 10 && args[0] == "--convert")
            {
                InputSet input = new InputSet { GovPayroll=args[2], WorkerPayrollSpecial=args[3], WorkerPayrollSchool=args[4], HealthGov=args[5], HealthOther=args[6], Pension=args[7], Employment=args[8], Industrial=args[9], ShortTermPayroll=args.Length>=11?args[10]:"" };
                try { Processor.Run(input, args[1]); }
                catch (Exception ex) { File.WriteAllText(args[1]+".error.txt",ex.ToString(),Encoding.UTF8); Environment.ExitCode=1; }
                return;
            }
            if(args.Length>=3&&(args[0]=="--submit-worker"||args[0]=="--submit-teacher"))
            {
                SubmissionInfo info=args.Length>=9?new SubmissionInfo{RecipientCode=args[3],InstitutionName=args[4],ManagerName=args[5],Phone=args[6],BankName=args[7],AccountNumber=args[8],Round=args.Length>=10?args[9]:"",IndustrialRate=args.Length>=11?args[10]:"0.008",Site=args.Length>=12?args[11]:""}:new SubmissionInfo();
                try{SubmissionGenerator.Create(args[1],args[2],args[0]=="--submit-teacher",info);}
                catch(Exception ex){File.WriteAllText(Path.Combine(args[2],"제출생성_오류.txt"),ex.ToString(),Encoding.UTF8);Environment.ExitCode=1;}return;
            }
            Application.EnableVisualStyles(); Application.SetCompatibleTextRenderingDefault(false);
            if(args.Length>=3&&args[0]=="--submission-flow-test"){try{using(MainForm test=new MainForm())test.SubmissionFlowTest202(args[1],args[2]);}catch(Exception ex){File.WriteAllText(Path.Combine(args[2],"error.txt"),ex.ToString(),Encoding.UTF8);Environment.ExitCode=1;}return;}
            if(args.Length>=3&&args[0]=="--manual-test"){try{using(MainForm test=new MainForm())test.ManualTest202(args[1],args[2]);}catch(Exception ex){File.WriteAllText(args[2]+".error.txt",ex.ToString(),Encoding.UTF8);Environment.ExitCode=1;}return;}
            if(args.Length>=3&&args[0]=="--review-fix-test"){try{using(MainForm test=new MainForm())test.ReviewFixTest202(args[1],args[2]);}catch(Exception ex){File.WriteAllText(args[2]+".error.txt",ex.ToString(),Encoding.UTF8);Environment.ExitCode=1;}return;}
            if(args.Length>=3&&args[0]=="--selftest202"){try{using(MainForm test=new MainForm())test.SelfTest202(args[1],args[2]);}catch(Exception ex){File.WriteAllText(args[2]+".error.txt",ex.ToString(),Encoding.UTF8);Environment.ExitCode=1;}return;}
            Application.ThreadException+=(s,e)=>MessageBox.Show("작업 중 오류가 발생했습니다.\r\n\r\n"+e.Exception.Message,"사회보험 대사 보조 도우미",MessageBoxButtons.OK,MessageBoxIcon.Error);
            if(args.Length>=2&&args[0]=="--screenshot-ui"){using(MainForm preview=new MainForm())preview.SavePreview(args[1]);return;}
            if(args.Length>=2&&args[0]=="--screenshot-ui-collapsed"){using(MainForm preview=new MainForm())preview.SaveCollapsedPreview(args[1]);return;}
            if(args.Length>=3&&args[0]=="--screenshot-ui-page"){using(MainForm preview=new MainForm())preview.SavePagePreview(args[1],args[2]);return;}
            if(args.Length>=4&&args[0]=="--screenshot-ui-theme"){using(MainForm preview=new MainForm())preview.SaveThemePreview(args[1],args[2],args[3]);return;}
            if(args.Length>=5&&args[0]=="--screenshot-ui-individual-mode"){using(MainForm preview=new MainForm())preview.SaveIndividualModePreview(args[1],args[2],args[3],args[4]);return;}
            if(args.Length>=5&&args[0]=="--screenshot-ui-result-theme"){using(MainForm preview=new MainForm())preview.SaveResultThemePreview(args[1],args[2],args[3],args[4]);return;}
            if(args.Length>=4&&args[0]=="--screenshot-ui-nav-cycle"){string theme=args.Length>=4?args[3]:"블루";using(MainForm preview=new MainForm())preview.SaveNavigationCyclePreview(args[1],args[2],theme);return;}
            if(args.Length>=4&&args[0]=="--screenshot-ui-result"){int siteIndex=0,fundIndex=-1,scrollOffset=0;if(args.Length>=5)Int32.TryParse(args[4],out siteIndex);if(args.Length>=6)Int32.TryParse(args[5],out fundIndex);string search=args.Length>=7?args[6]:"";if(args.Length>=8)Int32.TryParse(args[7],out scrollOffset);using(MainForm preview=new MainForm())preview.SaveResultPagePreview(args[1],args[2],args[3],siteIndex,fundIndex,search,scrollOffset);return;}
            if(args.Length>=3&&args[0]=="--screenshot-ui-review-detail"){using(MainForm preview=new MainForm())preview.SaveReviewDetailPreview(args[1],args[2]);return;}
            if(args.Length>=3&&args[0]=="--screenshot-ui-review-draft"){string fund=args.Length>=4?args[3]:"계약제교원";using(MainForm preview=new MainForm())preview.SaveReviewDraftPreview(args[1],args[2],fund);return;}
            if(args.Length>=4&&args[0]=="--export-adjustment"){int siteIndex=0,fundIndex=0;if(args.Length>=5)Int32.TryParse(args[4],out siteIndex);if(args.Length>=6)Int32.TryParse(args[5],out fundIndex);string mode=args.Length>=7?args[6]:"전체";try{using(MainForm export=new MainForm())export.ExportAdjustmentForTest(args[1],args[2],args[3],siteIndex,fundIndex,mode);}catch(Exception ex){File.WriteAllText(args[3]+".error.txt",ex.ToString(),Encoding.UTF8);Environment.ExitCode=1;}return;}
            if(args.Length>=2&&args[0]=="--save-adjustment-test"){try{using(MainForm save=new MainForm())save.SaveAdjustmentForTest(args[1]);}catch(Exception ex){File.WriteAllText(args[1]+".save-error.txt",ex.ToString(),Encoding.UTF8);Environment.ExitCode=1;}return;}
            if(args.Length>=2&&args[0]=="--apply-fund-test"){string fund=args.Length>=3?args[2]:"계약제교원";try{using(MainForm apply=new MainForm())apply.ApplyFundForTest(args[1],fund);}catch(Exception ex){File.WriteAllText(args[1]+".apply-error.txt",ex.ToString(),Encoding.UTF8);Environment.ExitCode=1;}return;}
            if(args.Length>=2&&args[0]=="--save-review-fund-test"){string fund=args.Length>=3?args[2]:"계약제교원";try{using(MainForm apply=new MainForm())apply.SaveReviewFundForTest(args[1],fund);}catch(Exception ex){File.WriteAllText(args[1]+".review-save-error.txt",ex.ToString(),Encoding.UTF8);Environment.ExitCode=1;}return;}
            if(args.Length>=2&&args[0]=="--save-review-complete-test"){try{using(MainForm apply=new MainForm())apply.SaveReviewCompletionForTest(args[1]);}catch(Exception ex){File.WriteAllText(args[1]+".review-complete-error.txt",ex.ToString(),Encoding.UTF8);Environment.ExitCode=1;}return;}
            if(args.Length>=2&&args[0]=="--save-discount-test"){try{using(MainForm apply=new MainForm())apply.SaveDiscountForTest(args[1]);}catch(Exception ex){File.WriteAllText(args[1]+".discount-save-error.txt",ex.ToString(),Encoding.UTF8);Environment.ExitCode=1;}return;}
            if(args.Length>=2&&args[0]=="--save-school-discount-test"){try{using(MainForm apply=new MainForm())apply.SaveSchoolDiscountForTest(args[1]);}catch(Exception ex){File.WriteAllText(args[1]+".school-discount-error.txt",ex.ToString(),Encoding.UTF8);Environment.ExitCode=1;}return;}
            if(args.Length>=3&&args[0]=="--audit-dashboard-stats"){try{using(MainForm audit=new MainForm())audit.AuditDashboardStatsForTest(args[1],args[2]);}catch(Exception ex){File.WriteAllText(args[2]+".error.txt",ex.ToString(),Encoding.UTF8);Environment.ExitCode=1;}return;}
            if(args.Length>=4&&args[0]=="--export-approval-test"){int siteIndex=0;if(args.Length>=5)Int32.TryParse(args[4],out siteIndex);try{using(MainForm export=new MainForm())export.ExportApprovalForTest(args[1],args[2],args[3],siteIndex);}catch(Exception ex){File.WriteAllText(args[2]+".approval-error.txt",ex.ToString(),Encoding.UTF8);Environment.ExitCode=1;}return;}
            if(args.Length>=4&&args[0]=="--screenshot-ui-files-theme"){using(MainForm preview=new MainForm())preview.SavePreviewWithFilesTheme(args[1],args[2],args.Skip(3));return;}
            if(args.Length>=3&&args[0]=="--screenshot-ui-files"){using(MainForm preview=new MainForm())preview.SavePreviewWithFiles(args[1],args.Skip(2));return;}
            if(args.Length>=3&&args[0]=="--ui-test-run"){try{using(MainForm test=new MainForm())test.RunRegisteredForTest(args[1],args.Skip(2));}catch(Exception ex){File.WriteAllText(args[1]+".error.txt",ex.ToString(),Encoding.UTF8);Environment.ExitCode=1;}return;}
            Application.Run(new MainForm());
        }
    }

    class InputSet
    {
        public string PayrollPackage, GovPayroll, WorkerPayrollSpecial, WorkerPayrollSchool, ShortTermPayroll, HealthGov, HealthOther, Pension, Employment, Industrial;
        public IEnumerable<Tuple<string,string>> All()
        {
            yield return Tuple.Create("급여대장 통합", PayrollPackage);
            yield return Tuple.Create("공무원 급여대장", GovPayroll); yield return Tuple.Create("교육공무직 급여대장(교특)", WorkerPayrollSpecial); yield return Tuple.Create("교육공무직 급여대장(학회)", WorkerPayrollSchool);
            yield return Tuple.Create("단기기간제 근로자 인건비 신청", ShortTermPayroll);
            yield return Tuple.Create("건강보험(공무원)", HealthGov); yield return Tuple.Create("건강보험(비공무원)", HealthOther);
            yield return Tuple.Create("국민연금", Pension); yield return Tuple.Create("고용보험", Employment); yield return Tuple.Create("산재보험", Industrial);
        }
    }

    static class UiTheme
    {
        public static string Name="파랑";public static bool Dark;public static Color Accent,Secondary,Page,Sidebar,Card,Surface,Input,Border,Text,Muted,Header,Grid;
        static UiTheme(){Set("파랑");}
        public static void Set(string name)
        {
            if(name=="블루")name="파랑";if(name=="그린")name="초록";if(name=="다크")name="검정";
            Name=new[]{"파랑","초록","빨강","살구","회색","검정"}.Contains(name)?name:"파랑";Dark=Name=="검정"||Name=="회색";
            if(Name=="초록"){Accent=Color.FromArgb(27,145,82);Secondary=Color.FromArgb(47,170,104);}
            else if(Name=="빨강"){Accent=Color.FromArgb(205,62,62);Secondary=Color.FromArgb(226,83,83);}
            else if(Name=="살구"){Accent=Color.FromArgb(221,113,61);Secondary=Color.FromArgb(241,145,92);}
            else if(Name=="회색"){Accent=Color.FromArgb(205,214,229);Secondary=Color.FromArgb(160,174,196);}
            else if(Dark){Accent=Color.FromArgb(117,145,255);Secondary=Color.FromArgb(171,126,255);}
            else{Accent=Color.FromArgb(48,63,220);Secondary=Color.FromArgb(91,55,238);}
            if(Dark){Page=Color.FromArgb(18,22,31);Sidebar=Color.FromArgb(24,29,40);Card=Color.FromArgb(29,35,48);Surface=Color.FromArgb(35,42,57);Input=Color.FromArgb(38,46,62);Border=Color.FromArgb(65,76,98);Text=Color.FromArgb(239,243,253);Muted=Color.FromArgb(174,185,208);Header=Color.FromArgb(42,51,69);Grid=Color.FromArgb(67,78,101);}
            else{Page=Color.White;Sidebar=Color.FromArgb(248,250,255);Card=Color.White;Surface=Color.FromArgb(248,250,255);Input=Color.White;Border=Color.FromArgb(226,230,242);Text=Color.FromArgb(30,43,91);Muted=Color.FromArgb(102,111,142);Header=Color.FromArgb(249,250,254);Grid=Color.FromArgb(225,230,242);}
            if(Name=="회색"){Page=Color.FromArgb(65,69,76);Sidebar=Color.FromArgb(57,61,68);Card=Color.FromArgb(78,83,91);Surface=Color.FromArgb(85,90,99);Input=Color.FromArgb(91,96,105);Border=Color.FromArgb(112,118,128);Text=Color.FromArgb(244,246,249);Muted=Color.FromArgb(204,209,217);Header=Color.FromArgb(87,92,101);Grid=Color.FromArgb(116,122,132);}
        }
    }

    static class AppUpdater
    {
        public const string ManifestUrl="https://raw.githubusercontent.com/isilria/Social-insurance-updates/main/latest.ini";
        public static string CheckAndInstall(IWin32Window owner,Version current,bool interactive)
        {
            try
            {
                ServicePointManager.SecurityProtocol|=SecurityProtocolType.Tls12;
                string manifest;
                using(WebClient client=new WebClient()){client.Headers[HttpRequestHeader.UserAgent]="SocialInsuranceUpdater/"+current;client.Encoding=Encoding.UTF8;manifest=client.DownloadString(ManifestUrl);}
                Dictionary<string,string> values=manifest.Replace("\r","").Split('\n').Select(x=>x.Trim()).Where(x=>x.Length>0&&!x.StartsWith("#")&&x.Contains("=")).Select(x=>x.Split(new[]{'='},2)).ToDictionary(x=>x[0].Trim(),x=>x[1].Trim(),StringComparer.OrdinalIgnoreCase);
                Version latest;string versionText,url,sha;
                if(!values.TryGetValue("version",out versionText)||!Version.TryParse(versionText,out latest)||!values.TryGetValue("url",out url)||!values.TryGetValue("sha256",out sha))throw new InvalidDataException("업데이트 정보 형식이 올바르지 않습니다.");
                if(latest<=current)return "최신 여부  최신 버전입니다.";
                string available="최신 여부  Ver. "+latest+" 업데이트 가능";
                if(!interactive)return available;
                string notes;values.TryGetValue("notes",out notes);
                if(MessageBox.Show("새 버전 Ver. "+latest+"이 있습니다."+(String.IsNullOrWhiteSpace(notes)?"":"\r\n\r\n"+notes)+"\r\n\r\n지금 내려받아 실행할까요?","업데이트 확인",MessageBoxButtons.YesNo,MessageBoxIcon.Information)!=DialogResult.Yes)return available;
                string temp=Path.Combine(Path.GetTempPath(),"SocialInsurance_"+latest+"_"+Guid.NewGuid().ToString("N")+".exe");
                using(WebClient client=new WebClient()){client.Headers[HttpRequestHeader.UserAgent]="SocialInsuranceUpdater/"+current;client.DownloadFile(url,temp);}
                string actual;using(SHA256 hash=SHA256.Create())using(FileStream stream=File.OpenRead(temp))actual=BitConverter.ToString(hash.ComputeHash(stream)).Replace("-","");
                if(!String.Equals(actual,Regex.Replace(sha,"[^0-9A-Fa-f]","").ToUpperInvariant(),StringComparison.OrdinalIgnoreCase)){File.Delete(temp);throw new InvalidDataException("내려받은 업데이트 파일의 무결성 확인에 실패했습니다.");}
                string fileName=Path.GetFileName(new Uri(url).LocalPath),folder=Path.GetDirectoryName(Application.ExecutablePath),target=Path.Combine(folder,fileName);
                try{File.Copy(temp,target,true);File.Delete(temp);}catch{folder=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),"Downloads");Directory.CreateDirectory(folder);target=Path.Combine(folder,fileName);File.Copy(temp,target,true);File.Delete(temp);}
                Process.Start(new ProcessStartInfo(target){UseShellExecute=true});Application.Exit();return "업데이트 실행 중";
            }
            catch(Exception ex)
            {
                if(interactive)MessageBox.Show("업데이트를 확인하지 못했습니다.\r\n\r\n"+ex.Message,"업데이트 확인",MessageBoxButtons.OK,MessageBoxIcon.Warning);
                return "최신 여부  확인 실패";
            }
        }
    }

    partial class MainForm : Form
    {
        readonly Dictionary<string, TextBox> boxes = new Dictionary<string, TextBox>();
        readonly Dictionary<string,List<string>> registeredFiles = new Dictionary<string,List<string>>();
        readonly Dictionary<string,HashSet<string>> registeredFileSites = new Dictionary<string,HashSet<string>>(StringComparer.OrdinalIgnoreCase);
        readonly Dictionary<string,HashSet<string>> registeredFilePeople = new Dictionary<string,HashSet<string>>(StringComparer.OrdinalIgnoreCase);
        readonly Dictionary<string,HashSet<string>> insurancePeopleBySite = new Dictionary<string,HashSet<string>>();
        readonly Dictionary<string,HashSet<string>> regularPeopleBySite = new Dictionary<string,HashSet<string>>();
        readonly Dictionary<string,HashSet<string>> employmentPeopleBySite = new Dictionary<string,HashSet<string>>();
        readonly List<Panel> pendingSiteCards = new List<Panel>();
        readonly Dictionary<string, Control> pages = new Dictionary<string, Control>();
        readonly Dictionary<string, Button> navigationButtons = new Dictionary<string, Button>();
        readonly List<Control> reconciliationNavigationItems = new List<Control>();
        SidebarNavButton reconciliationSectionButton;
        bool reconciliationExpanded=true;
        Timer reconciliationAnimationTimer;
        Panel contentHost,sidebar;
        FlowLayoutPanel navigationPanel;
        FlowLayoutPanel siteCardsHost;
        Panel dropZone;
        Label fileAnalysisStatus,siteCountLabel,readinessLabel,readinessDetail,sidebarVersionLabel,sidebarEmailLabel;
        Button runButton;bool runInProgress;
        DashboardStatCard[] adjustmentStatCards;
        Label adjustmentPeriodLabel,adjustmentRangeLabel,adjustmentSelectionLabel,adjustmentAmountLabel;
        ModernSiteSelector adjustmentSiteSelector,adjustmentFundSelector;
        RoundedPanel adjustmentFilterPanel;
        AdjustmentTableControl adjustmentTable;
        AdjustmentTabButton[] adjustmentTabs;
        Button adjustmentExcelButton,adjustmentPdfButton;
        readonly List<string> adjustmentSiteKeys=new List<string>();
        readonly HashSet<string> adjustmentSelections=new HashSet<string>(StringComparer.Ordinal);
        bool adjustmentFilterLoading;
        string adjustmentMode="전체";
        DashboardStatCard[] reviewStatCards;
        Label reviewPeriodLabel,reviewRangeLabel,reviewSelectionLabel,reviewAmountLabel;
        ModernSiteSelector reviewSiteSelector,reviewFundSelector,reviewApplyFundSelector;
        ReviewTableControl reviewTable;
        ReviewDetailBubble reviewBubble;
        readonly List<string> reviewSiteKeys=new List<string>();
        readonly HashSet<string> reviewSelections=new HashSet<string>(StringComparer.Ordinal);
        readonly HashSet<string> reviewCheckedKeys=new HashSet<string>(StringComparer.Ordinal);
        readonly Dictionary<string,string> reviewFundDrafts=new Dictionary<string,string>(StringComparer.Ordinal);
        bool reviewFilterLoading;
        Label discountPeriodLabel,discountRangeLabel;
        ModernSiteSelector discountSiteSelector,discountFundSelector;
        DiscountTotalsControl discountBilledTotals,discountAppliedTotals,discountAfterTotals;
        DiscountTableControl discountTable;
        readonly List<string> discountSiteKeys=new List<string>();
        readonly Dictionary<string,DiscountEntry> discountSaved=new Dictionary<string,DiscountEntry>(StringComparer.Ordinal);
        readonly Dictionary<string,DiscountEntry> discountDrafts=new Dictionary<string,DiscountEntry>(StringComparer.Ordinal);
        bool discountFilterLoading;
        bool discountStateUsesEmployer=true,discountStateUsesSubtraction=true;
        DashboardStatCard[] summaryStatCards;
        Label summaryPeriodLabel;
        ModernSiteSelector summarySiteSelector;
        PremiumTotalsControl summaryPremiumTotals;
        SummaryTableControl summaryTable;
        SummaryDashboardModel summaryDashboard;
        readonly List<string> summarySiteKeys=new List<string>();
        bool summaryComboLoading;
        DashboardStatCard[] individualStatCards;
        Label individualPeriodLabel,individualRangeLabel;
        ModernSiteSelector individualSiteSelector,individualFundSelector;
        TextBox individualSearchBox;
        IndividualTableControl individualTable;
        IndividualModeTabButton[] individualModeTabs;
        string individualAmountMode="개인부담금";
        IndividualDashboardModel individualDashboard;
        readonly ReconciliationUiState reconciliationState=new ReconciliationUiState();
        readonly List<string> individualSiteKeys=new List<string>();
        bool individualFilterLoading;
        const int IndividualPageSize=6;
        Label submissionSourceLabel;
        Label submissionPeriodLabel;
        ModernSiteSelector submissionSiteSelector,submissionRoundSelector;
        SubmissionSummaryControl submissionWorkerSummary,submissionTeacherSummary;
        readonly List<string> submissionSiteKeys=new List<string>();
        bool submissionFilterLoading;
        Label approvalPeriodLabel,approvalDescriptionLabel;
        ModernSiteSelector approvalSiteSelector;
        ApprovalPreviewControl approvalExcelPreview,approvalPdfPreview;
        readonly List<string> approvalSiteKeys=new List<string>();
        bool approvalFilterLoading;
        Timer cardAnimationTimer;
        int cardAnimationIndex;
        TextBox output, submitOutput, validationResult, recipientCode, institutionName, managerName, phone, bankName, accountNumber, submissionRound, industrialRate; Label status, submitStatus,themeStatusLabel,updateStatusLabel;Button[] themeChoiceButtons;string temporaryResultPath;bool openResultAfterSave,automaticUpdateCheck=true;
        public MainForm()
        {
            Dictionary<string,string> startupSettings=AppSettings.Load();UiTheme.Set(GetSetting(startupSettings,"Theme"));
            Text="사회보험 재원별 대사 보조 도우미 Ver. 2.0.2"; Icon=LoadAppIcon(); ClientSize=new Size(1280,650); MinimumSize=new Size(1180,650); StartPosition=FormStartPosition.CenterScreen; Font=new Font("맑은 고딕",9F); BackColor=UiTheme.Page;DoubleBuffered=true;
            sidebar=new Panel{Dock=DockStyle.Left,Width=205,BackColor=UiTheme.Sidebar,Padding=new Padding(12,15,12,12)};Controls.Add(sidebar);
            var brandImage=LoadReferenceIcon();var brandPanel=new Panel{Dock=DockStyle.Top,Height=90,BackColor=Color.Transparent};brandPanel.Controls.Add(new PictureBox{Image=brandImage==null?null:new Bitmap(brandImage,44,44),Location=new Point(4,8),Size=new Size(44,44),SizeMode=PictureBoxSizeMode.Zoom,BackColor=Color.Transparent});brandPanel.Controls.Add(new Label{Text="사회보험 대사\r\n보조 도우미",Location=new Point(55,8),Size=new Size(112,44),Font=new Font("맑은 고딕",9.5F,FontStyle.Bold),ForeColor=UiTheme.Accent,TextAlign=ContentAlignment.MiddleLeft,Tag="SidebarTitle"});sidebar.Controls.Add(brandPanel);
            navigationPanel=new BufferedFlowLayoutPanel{Dock=DockStyle.Fill,FlowDirection=FlowDirection.TopDown,WrapContents=false,AutoScroll=true,Padding=new Padding(0,8,0,52),BackColor=UiTheme.Sidebar};sidebar.Controls.Add(navigationPanel);navigationPanel.BringToFront();
            contentHost=new Panel{Location=new Point(205,0),Size=new Size(ClientSize.Width-205,ClientSize.Height),Anchor=AnchorStyles.Top|AnchorStyles.Bottom|AnchorStyles.Left|AnchorStyles.Right,BackColor=UiTheme.Page,Padding=new Padding(18,8,18,0)};Controls.Add(contentHost);
            sidebar.BringToFront();
            AddPage("파일 등록",BuildPage(BuildHomeScreen));
            AddPage("총괄표",BuildPage(BuildSummaryScreen));
            AddPage("개인별 내역",BuildPage(BuildIndividualScreen));
            AddPage("반환 / 추징",BuildPage(BuildAdjustmentScreen));
            AddPage("확인 필요",BuildPage(BuildReviewScreen));
            AddPage("감면 적용",BuildPage(BuildDiscountScreen));
            AddPage("제출서 생성",BuildPage(BuildSubmissionScreen));
            AddPage("내부결재자료 생성",BuildPage(BuildApprovalScreen));
            AddPage("설정",BuildPage(BuildSettingsScreen));
            // 작업 관리 화면은 이번 수정에서 제거했습니다.
            AddNav(navigationPanel,"파일 등록",false);AddNav(navigationPanel,"대사 결과",true);AddNav(navigationPanel,"총괄표",false);AddNav(navigationPanel,"개인별 내역",false);AddNav(navigationPanel,"반환 / 추징",false);AddNav(navigationPanel,"확인 필요",false);AddNav(navigationPanel,"감면 적용",false);AddNav(navigationPanel,"제출서 생성",false);AddNav(navigationPanel,"내부결재자료 생성",false);AddNav(navigationPanel,"설정",false);var sidebarFooter=new Panel{Width=165,Height=40,BackColor=Color.Transparent};sidebarVersionLabel=new Label{Text="Ver 2.0.2",Location=new Point(8,2),Size=new Size(157,14),Font=new Font("맑은 고딕",6.7F,FontStyle.Regular),ForeColor=UiTheme.Muted,BackColor=Color.Transparent,Tag="SidebarVersion"};sidebarEmailLabel=new Label{Text="e-mail : isilria@ice.go.kr",Location=new Point(8,19),Size=new Size(157,14),Font=new Font("맑은 고딕",6.7F,FontStyle.Regular),ForeColor=UiTheme.Muted,BackColor=Color.Transparent,Tag="SidebarVersion"};sidebarFooter.Controls.Add(sidebarVersionLabel);sidebarFooter.Controls.Add(sidebarEmailLabel);navigationPanel.Controls.Add(sidebarFooter);Action placeSidebarFooter=()=>{int used=navigationPanel.Controls.Cast<Control>().Where(x=>x!=sidebarFooter&&x.Visible).Sum(x=>x.Height+x.Margin.Vertical);int top=Math.Max(10,navigationPanel.ClientSize.Height-navigationPanel.Padding.Vertical-used-sidebarFooter.Height);if(sidebarFooter.Margin.Top!=top)sidebarFooter.Margin=new Padding(0,top,0,0);};navigationPanel.SizeChanged+=(s,e)=>placeSidebarFooter();navigationPanel.Layout+=(s,e)=>placeSidebarFooter();placeSidebarFooter();
            LoadSavedSubmissionInfo();Initialize202();ApplyTheme(UiTheme.Name,false);FormClosing+=(s,e)=>{if(e.Cancel)return;SaveSubmissionInfo();CleanupTemporaryResult();};ShowPage("파일 등록");sidebar.BringToFront();
        }
        delegate void PageBuilder(Control page);
        Control BuildPage(PageBuilder builder){var page=new Panel{Dock=DockStyle.Fill,BackColor=UiTheme.Page,AutoScroll=true,Tag="ThemePage"};builder(page);return page;}
        Control BuildPlaceholderPage(string title,string description)
        {
            var page=new Panel{Dock=DockStyle.Fill,BackColor=Color.White};
            page.Controls.Add(new Label{Text=title,Font=new Font("맑은 고딕",20F,FontStyle.Bold),ForeColor=Color.FromArgb(28,52,138),AutoSize=true,Location=new Point(18,20)});
            page.Controls.Add(new Label{Text=description,Font=new Font("맑은 고딕",10F),ForeColor=Color.DimGray,AutoSize=true,Location=new Point(21,68)});
            var card=new Panel{Location=new Point(20,112),Size=new Size(850,180),BackColor=Color.FromArgb(248,250,255),BorderStyle=BorderStyle.FixedSingle};
            card.Controls.Add(new Label{Text="1단계 UI 골격 연결 완료",Font=new Font("맑은 고딕",13F,FontStyle.Bold),ForeColor=Color.FromArgb(79,70,229),AutoSize=true,Location=new Point(28,38)});
            card.Controls.Add(new Label{Text="대사 실행 후 생성되는 1.8 결과 구조를 이 화면의 표·필터·상태 카드에 연결할 예정입니다.\r\n현재 단계에서는 기존 계산 및 Excel 생성 로직을 변경하지 않았습니다.",Font=new Font("맑은 고딕",10F),ForeColor=Color.FromArgb(70,75,90),AutoSize=true,Location=new Point(28,80)});
            page.Controls.Add(card);return page;
        }
        void BuildSettingsScreen(Control page)
        {
            page.Controls.Add(TitleLabel("설정",8,10,20F));
            var update=(RoundedPanel)Card(8,61,1030,104,UiTheme.Card);update.Controls.Add(new Label{Text="업데이트 확인",Location=new Point(22,13),AutoSize=true,ForeColor=UiText,Font=new Font("맑은 고딕",11F,FontStyle.Bold),Tag="ThemeText"});update.Controls.Add(new Label{Text="현재 버전  Ver. 2.0.2",Location=new Point(22,51),AutoSize=true,ForeColor=UiTheme.Accent,Font=new Font("맑은 고딕",9F,FontStyle.Bold),Tag="ThemeAccent"});updateStatusLabel=new Label{Text="최신 여부  확인 전",Location=new Point(225,51),AutoSize=true,ForeColor=UiMuted,Font=new Font("맑은 고딕",8.5F),Tag="ThemeMuted"};update.Controls.Add(updateStatusLabel);var check=OutputButton("업데이트 확인","refresh",555,36,155,38,UiTheme.Accent,false);check.Tag="ThemeAccentAction";check.Click+=(s,e)=>CheckForUpdates(true);update.Controls.Add(check);var auto=new CheckBox{Text="자동 확인",Checked=automaticUpdateCheck,Location=new Point(758,45),AutoSize=true,ForeColor=UiText,BackColor=Color.Transparent,Tag="ThemeText"};auto.CheckedChanged+=(s,e)=>{automaticUpdateCheck=auto.Checked;SaveSubmissionInfo();};update.Controls.Add(auto);page.Controls.Add(update);
            var themeCard=(RoundedPanel)Card(8,179,1030,150,UiTheme.Card);themeCard.Controls.Add(new Label{Text="색상 테마 선택",Location=new Point(22,14),AutoSize=true,ForeColor=UiText,Font=new Font("맑은 고딕",11F,FontStyle.Bold),Tag="ThemeText"});string[] names={"파랑","초록","빨강","살구","회색","검정"};Color[] accents={Color.FromArgb(48,63,220),Color.FromArgb(27,145,82),Color.FromArgb(205,62,62),Color.FromArgb(221,113,61),Color.FromArgb(91,103,121),Color.FromArgb(117,145,255)};themeChoiceButtons=new Button[6];for(int i=0;i<6;i++){string choice=names[i];var button=new ThemeChoiceButton{ThemeName=choice,Description="",Accent=accents[i],DarkPreview=choice=="검정",Location=new Point(18+i*168,52),Size=new Size(154,70),Active=UiTheme.Name==choice};button.Click+=(s,e)=>ApplyTheme(choice,true);themeChoiceButtons[i]=button;themeCard.Controls.Add(button);}page.Controls.Add(themeCard);
            var program=(RoundedPanel)Card(8,343,1030,92,UiTheme.Card);program.Controls.Add(new Label{Text="프로그램 설정",Location=new Point(22,14),AutoSize=true,ForeColor=UiText,Font=new Font("맑은 고딕",11F,FontStyle.Bold),Tag="ThemeText"});var open=new CheckBox{Text="저장 후 결과 파일 자동 열기",Checked=openResultAfterSave,Location=new Point(24,53),AutoSize=true,ForeColor=UiText,BackColor=Color.Transparent,Tag="ThemeText"};open.CheckedChanged+=(s,e)=>{openResultAfterSave=open.Checked;SaveSubmissionInfo();};program.Controls.Add(open);page.Controls.Add(program);
        }
        void ApplyTheme(string name,bool persist)
        {
            UiTheme.Set(name);BackColor=UiTheme.Page;if(contentHost!=null)contentHost.BackColor=UiTheme.Page;if(sidebar!=null){sidebar.BackColor=UiTheme.Sidebar;ApplySidebarTheme(sidebar);}if(sidebarVersionLabel!=null){sidebarVersionLabel.BackColor=Color.Transparent;sidebarVersionLabel.ForeColor=UiTheme.Muted;}if(sidebarEmailLabel!=null){sidebarEmailLabel.BackColor=Color.Transparent;sidebarEmailLabel.ForeColor=UiTheme.Muted;}if(navigationPanel!=null)navigationPanel.BackColor=UiTheme.Sidebar;foreach(Control page in pages.Values)ApplyThemeToControl(page);foreach(Button b in navigationButtons.Values)b.Invalidate();if(reconciliationSectionButton!=null)reconciliationSectionButton.Invalidate();if(themeChoiceButtons!=null)foreach(Button b in themeChoiceButtons){ThemeChoiceButton choice=b as ThemeChoiceButton;if(choice!=null){choice.Active=choice.ThemeName==UiTheme.Name;choice.Invalidate();}}if(themeStatusLabel!=null){themeStatusLabel.Text="현재 테마  ·  "+UiTheme.Name;themeStatusLabel.ForeColor=UiTheme.Accent;}ApplyPeriodThemeAccent();Invalidate(true);if(persist)SaveSubmissionInfo();
        }
        void CheckForUpdates(bool interactive){string status=AppUpdater.CheckAndInstall(this,new Version(2,0,2),interactive);if(updateStatusLabel!=null)updateStatusLabel.Text=status;}
        void ApplyPeriodThemeAccent(){foreach(Label label in new[]{summaryPeriodLabel,individualPeriodLabel,adjustmentPeriodLabel,reviewPeriodLabel,discountPeriodLabel,submissionPeriodLabel,approvalPeriodLabel})if(label!=null){label.ForeColor=UiTheme.Accent;label.Invalidate();}}
        void ApplySidebarTheme(Control control){Panel panel=control as Panel;if(panel!=null)panel.BackColor=panel.Height<=1?UiTheme.Border:UiTheme.Sidebar;Label label=control as Label;if(label!=null){label.ForeColor=(label.Tag as string)=="SidebarVersion"?UiTheme.Muted:UiTheme.Accent;label.BackColor=Color.Transparent;}foreach(Control child in control.Controls)ApplySidebarTheme(child);control.Invalidate();}
        void ApplyThemeToControl(Control control)
        {
            string role=control.Tag as string;if(role=="ThemePage")control.BackColor=UiTheme.Page;else if(control is RoundedPanel){control.BackColor=UiTheme.Card;((RoundedPanel)control).BorderColor=UiTheme.Border;}else if(control is TextBox){control.BackColor=UiTheme.Input;control.ForeColor=UiTheme.Text;}else if(control is FlowLayoutPanel||control is Panel)control.BackColor=control==sidebar||control==navigationPanel?UiTheme.Sidebar:UiTheme.Page;
            Label label=control as Label;if(label!=null){if(role=="ThemeMuted")label.ForeColor=UiTheme.Muted;else if(role=="ThemeAccent"||role=="ThemeTitle")label.ForeColor=UiTheme.Accent;else if(!IsSemanticColor(label.ForeColor))label.ForeColor=label.Font.Size>=13F?UiTheme.Accent:label.Font.Bold?UiTheme.Text:UiTheme.Muted;label.BackColor=Color.Transparent;}CheckBox checkBox=control as CheckBox;if(checkBox!=null){checkBox.ForeColor=UiTheme.Text;checkBox.BackColor=Color.Transparent;}OutputActionButton accentAction=control as OutputActionButton;if(accentAction!=null&&role=="ThemeAccentAction")accentAction.Accent=UiTheme.Accent;
            foreach(Control child in control.Controls)ApplyThemeToControl(child);control.Invalidate();
        }
        static bool IsSemanticColor(Color color){return color.R>180&&color.G<150||color.G>105&&color.R<100||color.B>170&&color.R<100&&color.G>70||color.R>190&&color.G>80&&color.G<180&&color.B<110;}
        static Color UiBlue{get{return UiTheme.Accent;}}static Color UiPurple{get{return UiTheme.Secondary;}}static Color UiBorder{get{return UiTheme.Border;}}static Color UiText{get{return UiTheme.Text;}}static Color UiMuted{get{return UiTheme.Muted;}}static readonly Color UiGreen=Color.FromArgb(24,164,91),UiRed=Color.FromArgb(239,68,68),UiOrange=Color.FromArgb(245,124,45);
        Label TitleLabel(string text,int x,int y,float size=20F){return new Label{Text=text,AutoSize=true,Location=new Point(x,y),Font=new Font("맑은 고딕",size,FontStyle.Bold),ForeColor=UiText,Tag="ThemeTitle"};}
        Panel Card(int x,int y,int w,int h,Color? back=null){return new RoundedPanel{Location=new Point(x,y),Size=new Size(w,h),BackColor=back??UiTheme.Card,Radius=12,BorderColor=UiBorder,Tag="ThemeCard"};}
        void RoundControl(Control c,int radius){c.Resize+=(s,e)=>{using(GraphicsPath p=new GraphicsPath()){int d=radius*2,w=Math.Max(1,c.Width-1),h=Math.Max(1,c.Height-1);p.AddArc(0,0,d,d,180,90);p.AddArc(w-d,0,d,d,270,90);p.AddArc(w-d,h-d,d,d,0,90);p.AddArc(0,h-d,d,d,90,90);p.CloseFigure();c.Region=new Region(p);}};}
        Button ActionButton(string text,int x,int y,int w,int h,Color color){var b=new Button{Text=text,Location=new Point(x,y),Size=new Size(w,h),BackColor=color,ForeColor=Color.White,FlatStyle=FlatStyle.Flat,Font=new Font("맑은 고딕",9F,FontStyle.Bold),Cursor=Cursors.Hand,UseVisualStyleBackColor=false};b.FlatAppearance.BorderSize=0;b.FlatAppearance.MouseOverBackColor=Color.FromArgb(Math.Max(0,color.R-8),Math.Max(0,color.G-8),Math.Max(0,color.B-8));b.FlatAppearance.MouseDownBackColor=Color.FromArgb(Math.Max(0,color.R-24),Math.Max(0,color.G-24),Math.Max(0,color.B-24));RoundControl(b,8);return b;}
        OutputActionButton OutputButton(string text,string icon,int x,int y,int w,int h,Color color,bool filled=true){return new OutputActionButton{Text=text,IconKind=icon,Location=new Point(x,y),Size=new Size(w,h),Accent=color,Filled=filled};}
        Label Muted(string text,int x,int y){return new Label{Text=text,AutoSize=true,Location=new Point(x,y),ForeColor=UiMuted,Font=new Font("맑은 고딕",9F),Tag="ThemeMuted"};}
        void AddHeader(Control page,string title,string description)
        {
            page.Controls.Add(TitleLabel(title,8,12));page.Controls.Add(Muted(description,12,53));
            var refresh=new Button{Text="⟳  새로고침",Location=new Point(820,10),Size=new Size(92,30),FlatStyle=FlatStyle.Flat,BackColor=Color.White,ForeColor=UiBlue};refresh.FlatAppearance.BorderColor=UiBorder;page.Controls.Add(refresh);
            var excel=new Button{Text="▣  엑셀 내보내기",Location=new Point(918,10),Size=new Size(120,30),FlatStyle=FlatStyle.Flat,BackColor=Color.White,ForeColor=UiGreen};excel.FlatAppearance.BorderColor=UiBorder;page.Controls.Add(excel);
        }
        void AddStat(Control page,int x,string caption,string value,string note,Color color)
        {
            var c=Card(x,82,190,82,Color.FromArgb(252,253,255));c.Controls.Add(new Label{Text=caption,AutoSize=true,Location=new Point(14,11),ForeColor=UiMuted,Font=new Font("맑은 고딕",8.5F,FontStyle.Bold)});c.Controls.Add(new Label{Text=value,AutoSize=true,Location=new Point(14,31),ForeColor=color,Font=new Font("맑은 고딕",17F,FontStyle.Bold)});c.Controls.Add(new Label{Text=note,AutoSize=true,Location=new Point(14,61),ForeColor=UiMuted,Font=new Font("맑은 고딕",7.5F)});page.Controls.Add(c);
        }
        DataGridView MakeGrid(int x,int y,int w,int h,string[] columns,int rows)
        {
            var g=new DataGridView{Location=new Point(x,y),Size=new Size(w,h),BackgroundColor=Color.White,BorderStyle=BorderStyle.FixedSingle,ReadOnly=true,AllowUserToAddRows=false,AllowUserToDeleteRows=false,AllowUserToResizeRows=false,RowHeadersVisible=false,SelectionMode=DataGridViewSelectionMode.FullRowSelect,AutoSizeColumnsMode=DataGridViewAutoSizeColumnsMode.Fill,GridColor=Color.FromArgb(235,238,247),ColumnHeadersHeight=38,RowTemplate=new DataGridViewRow{Height=34},EnableHeadersVisualStyles=false};
            g.ColumnHeadersDefaultCellStyle.BackColor=Color.FromArgb(247,249,254);g.ColumnHeadersDefaultCellStyle.ForeColor=UiText;g.ColumnHeadersDefaultCellStyle.Font=new Font("맑은 고딕",8F,FontStyle.Bold);g.DefaultCellStyle.Font=new Font("맑은 고딕",8F);g.DefaultCellStyle.ForeColor=Color.FromArgb(58,66,96);g.DefaultCellStyle.SelectionBackColor=Color.FromArgb(239,241,255);g.DefaultCellStyle.SelectionForeColor=UiText;
            foreach(string c in columns)g.Columns.Add(c,c);for(int r=0;r<rows;r++){object[] values=new object[columns.Length];for(int i=0;i<columns.Length;i++)values[i]=i==0?(r+1).ToString():"-";g.Rows.Add(values);}return g;
        }
        void AddFilterBar(Control page,int y,bool fund=true)
        {
            var filter=Card(8,y,1030,66,Color.White);filter.Controls.Add(new Label{Text="고지 년월",Location=new Point(14,9),AutoSize=true,ForeColor=UiMuted,Font=new Font("맑은 고딕",8F,FontStyle.Bold)});filter.Controls.Add(new Label{Text="2026년 8월",Location=new Point(14,31),AutoSize=true,ForeColor=UiBlue,Font=new Font("맑은 고딕",13F,FontStyle.Bold)});filter.Controls.Add(new Label{Text="사업장 관리번호 선택",Location=new Point(210,9),AutoSize=true,ForeColor=UiMuted,Font=new Font("맑은 고딕",8F,FontStyle.Bold)});filter.Controls.Add(new ComboBox{Location=new Point(210,30),Size=new Size(300,26),DropDownStyle=ComboBoxStyle.DropDownList,Items={"12345 - 행복학교 급식실 (공무원)"},SelectedIndex=0});if(fund){filter.Controls.Add(new Label{Text="재원 선택",Location=new Point(535,9),AutoSize=true,ForeColor=UiMuted,Font=new Font("맑은 고딕",8F,FontStyle.Bold)});filter.Controls.Add(new ComboBox{Location=new Point(535,30),Size=new Size(150,26),DropDownStyle=ComboBoxStyle.DropDownList,Items={"전체","공무원","계약제교원","교육공무직","학교회계"},SelectedIndex=0});}page.Controls.Add(filter);
        }
        void BuildHomeScreen(Control page)
        {
            page.Controls.Add(TitleLabel("파일 등록",10,12,18F));
            dropZone=Card(10,58,1025,180,Color.FromArgb(253,252,255));dropZone.AllowDrop=true;var roundedDrop=(RoundedPanel)dropZone;roundedDrop.BorderColor=Color.FromArgb(167,155,255);roundedDrop.BorderWidth=1;roundedDrop.BorderDashStyle=System.Drawing.Drawing2D.DashStyle.Dash;dropZone.Controls.Add(new PictureBox{Image=LoadReferenceIcon(.70F),Location=new Point(66,12),Size=new Size(184,150),SizeMode=PictureBoxSizeMode.Zoom,BackColor=Color.Transparent});dropZone.Controls.Add(new Label{Text="대사 파일을 여기로 드래그하세요",Location=new Point(350,34),AutoSize=true,ForeColor=UiText,Font=new Font("맑은 고딕",13F,FontStyle.Bold)});dropZone.Controls.Add(new Label{Text="급여대장, 사회보험 개인별 부과내역 파일,\r\n1개월 미만 대체근로자 인건비 신청 서식",Location=new Point(350,64),AutoSize=true,ForeColor=UiMuted,Font=new Font("맑은 고딕",8.5F)});var select=ActionButton("▱  파일 선택",350,112,142,34,UiTheme.Accent);select.Click+=(s,e)=>ChooseMultipleFiles();dropZone.Controls.Add(select);var reset=ActionButton("초기화",500,112,78,34,Color.FromArgb(105,115,137));reset.Click+=(s,e)=>ResetWorkspace();dropZone.Controls.Add(reset);fileAnalysisStatus=new Label{Text="ⓘ  파일을 올려놓으면 자료 종류와 사업장관리번호를 자동 분석합니다.",Location=new Point(596,121),AutoSize=true,ForeColor=UiMuted,BackColor=Color.FromArgb(246,244,255),Padding=new Padding(8,3,8,3),Font=new Font("맑은 고딕",7.5F)};RoundControl(fileAnalysisStatus,8);dropZone.Controls.Add(fileAnalysisStatus);dropZone.DragEnter+=OnFilesDragEnter;dropZone.DragLeave+=OnFilesDragLeave;dropZone.DragDrop+=OnFilesDropped;foreach(Control child in dropZone.Controls){child.AllowDrop=true;child.DragEnter+=OnFilesDragEnter;child.DragLeave+=OnFilesDragLeave;child.DragDrop+=OnFilesDropped;}page.Controls.Add(dropZone);
            siteCardsHost=new FlowLayoutPanel{Location=new Point(10,252),Size=new Size(1025,202),AutoScroll=true,WrapContents=false,FlowDirection=FlowDirection.LeftToRight,BackColor=UiTheme.Page,Padding=new Padding(8,3,8,3)};page.Controls.Add(siteCardsHost);var empty=Card(170,18,680,145,UiTheme.Card);empty.Name="EmptySiteState";empty.Controls.Add(new Label{Text="아직 분석된 사업장이 없습니다",Location=new Point(210,38),AutoSize=true,ForeColor=UiText,Font=new Font("맑은 고딕",12F,FontStyle.Bold)});empty.Controls.Add(Muted("위 영역에 파일을 드래그하거나 ‘파일 선택’을 눌러 등록해 주세요.",154,75));siteCardsHost.Controls.Add(empty);
            output=new TextBox{Visible=false,Text=Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)};page.Controls.Add(output);
            page.Controls.Add(new Label{Text="●  ○  ○  ○",Location=new Point(485,458),AutoSize=true,ForeColor=Color.FromArgb(205,198,244),Font=new Font("맑은 고딕",7F)});
            var ready=Card(28,486,1005,67,Color.FromArgb(249,252,255));siteCountLabel=new Label{Text="▱  사업장 0개",Location=new Point(18,15),AutoSize=true,ForeColor=UiText,Font=new Font("맑은 고딕",10F,FontStyle.Bold)};ready.Controls.Add(siteCountLabel);readinessLabel=new Label{Text="○  파일 등록 대기",Location=new Point(255,12),AutoSize=true,ForeColor=UiMuted,Font=new Font("맑은 고딕",10F,FontStyle.Bold)};ready.Controls.Add(readinessLabel);readinessDetail=new Label{Text="분석할 파일을 등록해 주세요.",Location=new Point(277,36),AutoSize=true,ForeColor=UiMuted,Font=new Font("맑은 고딕",8F)};ready.Controls.Add(readinessDetail);runButton=ActionButton("▷  대사 시작",745,12,235,42,Color.FromArgb(181,184,201));runButton.Enabled=false;runButton.Click+=(s,e)=>Run();ready.Controls.Add(runButton);page.Controls.Add(ready);status=new Label{Text="",Visible=false};page.Controls.Add(status);
            var homeNotice=(RoundedPanel)Card(28,563,1005,30,Color.FromArgb(255,249,249));homeNotice.BorderColor=Color.FromArgb(244,224,224);var warningText=new Label{Text="⚠  본 프로그램은 업무 효율성 향상을 위해 개인이 제작한 도구입니다. 사용 상황에 따라 결과값이 부정확할 수 있으니 보조용으로 사용하시고 결과 값을 꼭 확인하시기 바랍니다",Location=new Point(14,7),AutoSize=true,ForeColor=Color.FromArgb(126,32,32),Font=new Font("맑은 고딕",7.2F,FontStyle.Bold)};homeNotice.Controls.Add(warningText);int authorX=Math.Min(936,warningText.Right+4);var authorAt=new Label{Text="@",Location=new Point(authorX,7),AutoSize=true,ForeColor=Color.FromArgb(30,30,30),Font=new Font("맑은 고딕",7.2F,FontStyle.Bold)};homeNotice.Controls.Add(authorAt);homeNotice.Controls.Add(new Label{Text="살구아빠",Location=new Point(authorAt.Right,7),AutoSize=true,ForeColor=Color.FromArgb(222,130,84),Font=new Font("맑은 고딕",7.2F,FontStyle.Bold)});page.Controls.Add(homeNotice);
            foreach(string key in new[]{"급여대장 통합","건강보험","국민연금","고용보험","산재보험","단기기간제 근로자"}){boxes[key]=new TextBox{Visible=false};page.Controls.Add(boxes[key]);registeredFiles[key]=new List<string>();}
        }
        void ChooseMultipleFiles(){using(OpenFileDialog d=new OpenFileDialog{Filter="대사 자료 (*.xlsx;*.xlsm;*.xls;*.zip)|*.xlsx;*.xlsm;*.xls;*.zip|모든 파일 (*.*)|*.*",Multiselect=true})if(d.ShowDialog()==DialogResult.OK)AnalyzeRegisteredFiles(d.FileNames);}
        void ResetWorkspace(){importErrors202.Clear();importOrigins202.Clear();foreach(List<string> files in registeredFiles.Values)files.Clear();foreach(TextBox box in boxes.Values)box.Text="";foreach(HashSet<string> values in registeredFileSites.Values)values.Clear();foreach(HashSet<string> values in registeredFilePeople.Values)values.Clear();registeredFileSites.Clear();registeredFilePeople.Clear();insurancePeopleBySite.Clear();regularPeopleBySite.Clear();employmentPeopleBySite.Clear();CleanupTemporaryResult();if(validationResult!=null)validationResult.Text="";RenderSiteCards(new string[0]);UpdateReadiness(0,0);if(fileAnalysisStatus!=null)fileAnalysisStatus.Text="ⓘ  파일을 올려놓으면 자료 종류와 사업장관리번호를 자동 분석합니다.";MessageBox.Show("등록 파일과 현재 작업 상태를 초기화했습니다.","초기화",MessageBoxButtons.OK,MessageBoxIcon.Information);}
        void OnFilesDragEnter(object sender,DragEventArgs e){if(e.Data.GetDataPresent(DataFormats.FileDrop)){e.Effect=DragDropEffects.Copy;dropZone.BackColor=Color.FromArgb(244,242,255);((RoundedPanel)dropZone).BorderColor=UiPurple;}}
        void OnFilesDragLeave(object sender,EventArgs e){dropZone.BackColor=Color.FromArgb(252,251,255);((RoundedPanel)dropZone).BorderColor=Color.FromArgb(179,172,255);}
        void OnFilesDropped(object sender,DragEventArgs e){OnFilesDragLeave(sender,EventArgs.Empty);string[] paths=e.Data.GetData(DataFormats.FileDrop) as string[];if(paths!=null)AnalyzeRegisteredFiles(paths);}
        void AnalyzePreparedFiles(IEnumerable<string> paths)
        {
            string[] valid=paths.Where(File.Exists).Where(p=>new[]{".xlsx",".xlsm",".xls",".zip"}.Contains(Path.GetExtension(p).ToLowerInvariant())).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();if(valid.Length==0){MessageBox.Show("등록할 수 있는 Excel 또는 ZIP 파일이 없습니다.","파일 등록",MessageBoxButtons.OK,MessageBoxIcon.Information);return;}
            fileAnalysisStatus.Text="●  파일 내용을 분석하고 있습니다...";fileAnalysisStatus.ForeColor=UiPurple;UseWaitCursor=true;Application.DoEvents();
            foreach(string path in valid)
            {
                string[] kinds=ClassifyInput(path);foreach(string kind in kinds){if(!registeredFiles[kind].Contains(path,StringComparer.OrdinalIgnoreCase))registeredFiles[kind].Add(path);}
                registeredFileSites[path]=DetectWorkplaceNumbers(path);registeredFilePeople[path]=DetectPersonKeys(path);
            }
            foreach(string key in registeredFiles.Keys)boxes[key].Text=registeredFiles[key].Count==1?registeredFiles[key][0]:"";
            BuildInsuranceSiteMap();ResolveShortTermSites();
            HashSet<string> sites=new HashSet<string>();var insurancePaths=new[]{"건강보험","국민연금","고용보험","산재보험"}.SelectMany(k=>registeredFiles[k]).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();foreach(string path in insurancePaths){HashSet<string> found;if(registeredFileSites.TryGetValue(path,out found))sites.UnionWith(found);}if(sites.Count==0)foreach(HashSet<string> found in registeredFileSites.Values)sites.UnionWith(found);if(sites.Count==0)sites.Add("미확인");
            RenderSiteCards(sites.OrderBy(x=>x).ToArray());UpdateReadiness(sites.Count,valid.Length);UseWaitCursor=false;
        }
        string[] ClassifyInput(string path)
        {
            string fileName=Path.GetFileName(path);if(Regex.IsMatch(fileName,"단기기간제|단기 기간제|대체근로|대체 근로|1개월[ ]*미만|일용근로|일용 근로",RegexOptions.IgnoreCase))return new[]{"단기기간제 근로자"};if(Regex.IsMatch(fileName,"급여대장|급여 대장|급여명세",RegexOptions.IgnoreCase))return new[]{"급여대장 통합"};if(Regex.IsMatch(fileName,"건강보험|장기요양",RegexOptions.IgnoreCase))return new[]{"건강보험"};if(Regex.IsMatch(fileName,"국민연금|연금보험료",RegexOptions.IgnoreCase))return new[]{"국민연금"};if(Regex.IsMatch(fileName,"고용보험",RegexOptions.IgnoreCase))return new[]{"고용보험"};if(Regex.IsMatch(fileName,"산재보험",RegexOptions.IgnoreCase))return new[]{"산재보험"};string probe=ReadClassificationText(path);
            int pension=ScoreProbe(probe,new[]{"기준소득월액","본인기여금","사용자부담금","소급분","취득일","상실일","국민연금"});int health=ScoreProbe(probe,new[]{"장기요양보험료","건강보험료","감면사유","정산보험료","건강보험"});int employment=ScoreProbe(probe,new[]{"고용보험","피보험자격","실업급여","고용안정"});int industrial=ScoreProbe(probe,new[]{"산재보험","산재근로자","업종요율"});int shortTerm=ScoreProbe(probe,new[]{"1개월미만","대체근로자","단기기간제","일용근로","근무일수","인건비신청"});int payroll=ScoreProbe(probe,new[]{"급여대장","지급총액","공제총액","실지급액","급여직종","재원"});int best=Math.Max(pension,Math.Max(health,Math.Max(employment,Math.Max(industrial,Math.Max(shortTerm,payroll)))));if(best==pension&&pension>=3)return new[]{"국민연금"};if(best==health&&health>=2)return new[]{"건강보험"};if(best==employment&&employment>=2)return new[]{"고용보험"};if(best==industrial&&industrial>=2)return new[]{"산재보험"};if(best==shortTerm&&shortTerm>=2)return new[]{"단기기간제 근로자"};if(payroll>=2)return new[]{"급여대장 통합"};return new string[0];
        }
        int ScoreProbe(string text,IEnumerable<string> tokens){int score=0;foreach(string token in tokens)if(text.IndexOf(token,StringComparison.OrdinalIgnoreCase)>=0)score++;return score;}
        string ReadClassificationText(string path)
        {
            var text=new StringBuilder(Path.GetFileName(path));try{string ext=Path.GetExtension(path).ToLowerInvariant();if(ext==".zip"){using(ZipArchive z=ZipFile.OpenRead(path))foreach(ZipArchiveEntry e in z.Entries){text.Append(' ').Append(e.FullName);string ee=Path.GetExtension(e.FullName).ToLowerInvariant();if(ee==".xlsx"||ee==".xlsm"){string temp=Path.Combine(Path.GetTempPath(),"classify_"+Guid.NewGuid().ToString("N")+ee);try{e.ExtractToFile(temp,true);text.Append(' ').Append(ReadClassificationText(temp));}catch{}finally{try{File.Delete(temp);}catch{}}}}}else if(ext==".xlsx"||ext==".xlsm"){using(ExcelPackage p=new ExcelPackage(new FileInfo(path)))foreach(ExcelWorksheet ws in p.Workbook.Worksheets){text.Append(' ').Append(ws.Name);if(ws.Dimension==null)continue;for(int r=1;r<=Math.Min(18,ws.Dimension.End.Row);r++)for(int c=1;c<=Math.Min(24,ws.Dimension.End.Column);c++){string v=ws.Cells[r,c].Text;if(!String.IsNullOrWhiteSpace(v))text.Append(' ').Append(v);}}}}catch{}return text.ToString();
        }
        HashSet<string> DetectWorkplaceNumbers(string path)
        {
            var result=new HashSet<string>();Action<string> scan=text=>{if(String.IsNullOrEmpty(text))return;foreach(Match m in Regex.Matches(text,@"(?<!\d)(?:\d[\s-]?){10,11}(?!\d)")){string digits=Regex.Replace(m.Value,@"\D","");if(digits.Length==10||digits.Length==11)result.Add(digits);}};scan(Path.GetFileName(path));
            try
            {
                if(Path.GetExtension(path).Equals(".zip",StringComparison.OrdinalIgnoreCase)){using(ZipArchive z=ZipFile.OpenRead(path))foreach(ZipArchiveEntry e in z.Entries){scan(e.FullName);string ext=Path.GetExtension(e.FullName).ToLowerInvariant();if(ext==".xlsx"||ext==".xlsm"){string temp=Path.Combine(Path.GetTempPath(),"site_scan_"+Guid.NewGuid().ToString("N")+ext);try{e.ExtractToFile(temp,true);result.UnionWith(DetectWorkplaceNumbers(temp));}catch{}finally{try{File.Delete(temp);}catch{}}}}}
                else if(!Path.GetExtension(path).Equals(".xls",StringComparison.OrdinalIgnoreCase)){using(ExcelPackage p=new ExcelPackage(new FileInfo(path)))foreach(ExcelWorksheet ws in p.Workbook.Worksheets){if(ws.Dimension==null)continue;int maxRow=Math.Min(ws.Dimension.End.Row,45),maxCol=Math.Min(ws.Dimension.End.Column,28);for(int r=1;r<=maxRow;r++)for(int c=1;c<=maxCol;c++){string value=Convert.ToString(ws.Cells[r,c].Value,CultureInfo.InvariantCulture)??"";if(!Regex.IsMatch(value,"사업장.*관리번호|관리번호.*사업장",RegexOptions.IgnoreCase))continue;scan(value);for(int rr=r+1;rr<=Math.Min(maxRow,r+8);rr++)for(int cc=c;cc<=Math.Min(maxCol,c+1);cc++)scan(Convert.ToString(ws.Cells[rr,cc].Value,CultureInfo.InvariantCulture));}}}
            }catch{}
            return result;
        }
        HashSet<string> DetectPersonKeys(string path)
        {
            var result=new HashSet<string>();try{string ext=Path.GetExtension(path).ToLowerInvariant();if(ext==".zip"){using(ZipArchive z=ZipFile.OpenRead(path))foreach(ZipArchiveEntry e in z.Entries){string ee=Path.GetExtension(e.FullName).ToLowerInvariant();if(ee!=".xlsx"&&ee!=".xlsm")continue;string temp=Path.Combine(Path.GetTempPath(),"people_"+Guid.NewGuid().ToString("N")+ee);try{e.ExtractToFile(temp,true);result.UnionWith(DetectPersonKeys(temp));}catch{}finally{try{File.Delete(temp);}catch{}}}}else if(ext==".xlsx"||ext==".xlsm"){using(ExcelPackage p=new ExcelPackage(new FileInfo(path)))foreach(ExcelWorksheet ws in p.Workbook.Worksheets){if(ws.Dimension==null)continue;int header=0,nameCol=0,birthCol=0,jobCol=0;for(int r=1;r<=Math.Min(40,ws.Dimension.End.Row)&&header==0;r++){for(int c=1;c<=Math.Min(40,ws.Dimension.End.Column);c++){string h=Regex.Replace(ws.Cells[r,c].Text??"",@"\s","");if(nameCol==0&&Regex.IsMatch(h,"^(성명|가입자명|근로자명|피보험자명|이름)$"))nameCol=c;if(birthCol==0&&Regex.IsMatch(h,"주민등록번호|생년월일|외국인등록번호|주민번호"))birthCol=c;if(jobCol==0&&Regex.IsMatch(h,"^(직종|직종명|고용형태|직급|공무직급여직종|시도직종)$"))jobCol=c;}if(nameCol>0&&(birthCol>0||jobCol>0))header=r;else{nameCol=0;birthCol=0;jobCol=0;}}if(header==0)continue;int blank=0;for(int r=header+1;r<=Math.Min(ws.Dimension.End.Row,header+5000);r++){string name=Regex.Replace(ws.Cells[r,nameCol].Text??"",@"\s","");if(name.Length==0){if(++blank>200)break;continue;}blank=0;if(name=="성명"||name=="합계"||name=="총계")continue;result.Add("N:"+name);if(birthCol>0){string digits=Regex.Replace(ws.Cells[r,birthCol].Text??"",@"\D","");if(digits.Length>=6)result.Add("K:"+name+"|"+digits.Substring(0,6));}}}}}catch{}return result;
        }
        void CollectInsurancePeopleBySite(string path)
        {
            try{if(Path.GetExtension(path).Equals(".zip",StringComparison.OrdinalIgnoreCase)){using(ZipArchive z=ZipFile.OpenRead(path))foreach(ZipArchiveEntry e in z.Entries){string ext=Path.GetExtension(e.FullName).ToLowerInvariant();if(ext!=".xlsx"&&ext!=".xlsm")continue;string temp=Path.Combine(Path.GetTempPath(),"site_people_"+Guid.NewGuid().ToString("N")+ext);try{e.ExtractToFile(temp,true);CollectInsurancePeopleBySite(temp);}catch{}finally{try{File.Delete(temp);}catch{}}}}else{HashSet<string> sites=DetectWorkplaceNumbers(path),people=DetectPersonKeys(path);foreach(string site in sites){HashSet<string> target;if(!insurancePeopleBySite.TryGetValue(site,out target)){target=new HashSet<string>();insurancePeopleBySite[site]=target;}target.UnionWith(people);}}}catch{}
        }
        void ForEachWorkbook(string path,Action<string> action)
        {
            if(!Path.GetExtension(path).Equals(".zip",StringComparison.OrdinalIgnoreCase)){action(path);return;}try{using(ZipArchive z=ZipFile.OpenRead(path))foreach(ZipArchiveEntry e in z.Entries){string ext=Path.GetExtension(e.FullName).ToLowerInvariant();if(ext!=".xlsx"&&ext!=".xlsm")continue;string temp=Path.Combine(Path.GetTempPath(),"map_"+Guid.NewGuid().ToString("N")+ext);try{e.ExtractToFile(temp,true);action(temp);}catch{}finally{try{File.Delete(temp);}catch{}}}}catch{}
        }
        void AddPeople(Dictionary<string,HashSet<string>> map,string site,IEnumerable<string> people){HashSet<string> target;if(!map.TryGetValue(site,out target)){target=new HashSet<string>();map[site]=target;}target.UnionWith(people);}
        void BuildInsuranceSiteMap()
        {
            regularPeopleBySite.Clear();employmentPeopleBySite.Clear();insurancePeopleBySite.Clear();string[] regularPaths=new[]{"건강보험","국민연금"}.SelectMany(k=>registeredFiles[k]).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();foreach(string outer in regularPaths)ForEachWorkbook(outer,book=>{HashSet<string> sites=DetectWorkplaceNumbers(book),people=DetectPersonKeys(book);foreach(string site in sites)AddPeople(regularPeopleBySite,site,people);});
            string[] workPaths=new[]{"고용보험","산재보험"}.SelectMany(k=>registeredFiles[k]).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();foreach(string outer in workPaths){var mappedForFile=new HashSet<string>();ForEachWorkbook(outer,book=>{HashSet<string> people=DetectPersonKeys(book),detected=DetectWorkplaceNumbers(book);string best=null;int bestCount=0;foreach(var site in regularPeopleBySite){int count=people.Intersect(site.Value).Count();if(count>bestCount){bestCount=count;best=site.Key;}}if(best!=null&&bestCount>0){AddPeople(employmentPeopleBySite,best,people);mappedForFile.Add(best);}else foreach(string site in detected){AddPeople(employmentPeopleBySite,site,people);mappedForFile.Add(site);}});if(mappedForFile.Count>0)registeredFileSites[outer]=mappedForFile;}
            foreach(var site in regularPeopleBySite)AddPeople(insurancePeopleBySite,site.Key,site.Value);foreach(var site in employmentPeopleBySite)AddPeople(insurancePeopleBySite,site.Key,site.Value);
        }
        void ResolveShortTermSites()
        {
            foreach(string shortPath in registeredFiles["단기기간제 근로자"]){var matched=new HashSet<string>();HashSet<string> shortPeople;registeredFilePeople.TryGetValue(shortPath,out shortPeople);if(shortPeople!=null&&shortPeople.Count>0){foreach(var site in employmentPeopleBySite)if(shortPeople.Overlaps(site.Value))matched.Add(site.Key);if(matched.Count==0)foreach(var site in insurancePeopleBySite)if(shortPeople.Overlaps(site.Value))matched.Add(site.Key);}registeredFileSites[shortPath]=matched;}
        }
        string FormatSite(string digits){if(digits=="미확인")return "관리번호 미확인";if(digits.Length==11)return digits.Substring(0,3)+"-"+digits.Substring(3,2)+"-"+digits.Substring(5,6);if(digits.Length==10)return digits.Substring(0,3)+"-"+digits.Substring(3,2)+"-"+digits.Substring(5,5);return digits;}
        bool HasKindForSite(string kind,string site)
        {
            if(kind=="급여대장 통합")return registeredFiles[kind].Count>0;if(kind=="단기기간제 근로자"){foreach(string path in registeredFiles[kind]){HashSet<string> sites;if(registeredFileSites.TryGetValue(path,out sites)&&sites.Contains(site))return true;}return false;}
            foreach(string path in registeredFiles[kind]){HashSet<string> sites;if(!registeredFileSites.TryGetValue(path,out sites)||sites.Count==0||sites.Contains(site))return true;}return false;
        }
        void RenderSiteCards(string[] sites)
        {
            if(cardAnimationTimer!=null)cardAnimationTimer.Stop();siteCardsHost.Controls.Clear();pendingSiteCards.Clear();string[] kinds={"급여대장 통합","건강보험","국민연금","고용보험","산재보험","단기기간제 근로자"};string[] labels={"급여대장","건강보험","국민연금","고용보험","산재보험","대체근로자"};int visualIndex=0;
            foreach(string site in sites)
            {
                bool darkCards=UiTheme.Dark||UiTheme.Name=="회색";Color cardBack=darkCards?UiTheme.Card:(visualIndex++%3==2?Color.FromArgb(252,251,255):Color.FromArgb(251,254,252));var card=new RoundedPanel{Size=new Size(315,172),Margin=new Padding(8,3,8,3),BackColor=cardBack,Radius=13,BorderColor=darkCards?UiTheme.Border:(visualIndex%3==0?Color.FromArgb(226,220,249):Color.FromArgb(218,235,226)),Visible=false,Tag="ThemeCard"};bool required=HasKindForSite("급여대장 통합",site)&&(HasKindForSite("건강보험",site)||HasKindForSite("국민연금",site)||HasKindForSite("고용보험",site)||HasKindForSite("산재보험",site));card.Controls.Add(new Label{Text=FormatSite(site),Location=new Point(14,11),AutoSize=true,ForeColor=UiText,Font=new Font("맑은 고딕",8.5F,FontStyle.Bold)});var badge=new Label{Text=required?"●  완료":"●  확인",Location=new Point(242,9),Size=new Size(58,20),TextAlign=ContentAlignment.MiddleCenter,BackColor=darkCards?UiTheme.Surface:(required?Color.FromArgb(232,249,238):Color.FromArgb(255,243,227)),ForeColor=required?UiGreen:UiOrange,Font=new Font("맑은 고딕",7F,FontStyle.Bold)};RoundControl(badge,9);card.Controls.Add(badge);
                for(int i=0;i<kinds.Length;i++){bool present=HasKindForSite(kinds[i],site),uploaded=registeredFiles[kinds[i]].Count>0;string state=present?"등록":(uploaded?"미해당":"미등록");card.Controls.Add(new Label{Text=(present?"●  ":"○  ")+labels[i],Location=new Point(14,37+i*21),AutoSize=true,ForeColor=present?UiGreen:UiMuted,Font=new Font("맑은 고딕",8F,present?FontStyle.Bold:FontStyle.Regular)});card.Controls.Add(new Label{Text=state,Location=new Point(250,37+i*21),Size=new Size(50,18),TextAlign=ContentAlignment.MiddleRight,ForeColor=present?UiGreen:UiMuted,Font=new Font("맑은 고딕",7.5F,present?FontStyle.Bold:FontStyle.Regular)});}
                siteCardsHost.Controls.Add(card);pendingSiteCards.Add(card);
            }
            cardAnimationIndex=0;cardAnimationTimer=new Timer{Interval=120};cardAnimationTimer.Tick+=(s,e)=>{if(cardAnimationIndex>=pendingSiteCards.Count){cardAnimationTimer.Stop();return;}Panel card=pendingSiteCards[cardAnimationIndex++];card.Visible=true;card.Margin=new Padding(8,18,8,3);var rise=new Timer{Interval=18};rise.Tick+=(rs,re)=>{if(card.Margin.Top<=3){card.Margin=new Padding(8,3,8,3);((Timer)rs).Stop();((Timer)rs).Dispose();}else card.Margin=new Padding(8,Math.Max(3,card.Margin.Top-3),8,3);};rise.Start();};cardAnimationTimer.Start();
        }
        void UpdateReadiness(int siteCount,int addedCount)
        {
            int total=registeredFiles.Sum(x=>x.Value.Count);bool payroll=registeredFiles["급여대장 통합"].Count>0||registeredFiles["단기기간제 근로자"].Count>0;bool insurance=new[]{"건강보험","국민연금","고용보험","산재보험"}.Any(k=>registeredFiles[k].Count>0);bool shortUploaded=registeredFiles["단기기간제 근로자"].Count>0,shortMatched=registeredFiles["단기기간제 근로자"].Any(p=>registeredFileSites.ContainsKey(p)&&registeredFileSites[p].Count>0),shortUnmatched=shortUploaded&&!shortMatched;bool ready=payroll&&insurance;siteCountLabel.Text="▱  사업장 "+siteCount+"개";readinessLabel.Text=ready?"✓  대사 준비 완료":"△  필수 파일 확인 필요";readinessLabel.ForeColor=ready?UiGreen:UiOrange;readinessDetail.Text=ready?(shortUnmatched?"대사는 가능하지만 대체근로자 대상 사업장을 찾지 못했습니다.":"등록된 "+total+"개 파일을 이용해 대사를 시작할 수 있습니다."):"급여대장과 사회보험 부과자료를 각각 하나 이상 등록해 주세요.";runButton.Enabled=ready;runButton.BackColor=ready?Color.FromArgb(89,54,238):Color.FromArgb(181,184,201);fileAnalysisStatus.Text=(shortUnmatched?"△  ":"✓  ")+addedCount+"개 파일 분석 완료 · 사업장 "+siteCount+"개 인식"+(shortUnmatched?" · 대체근로자 매칭 0명":"");fileAnalysisStatus.ForeColor=shortUnmatched?UiOrange:UiGreen;
        }
        void BuildSummaryScreen(Control page)
        {
            page.Controls.Add(TitleLabel("총괄표",8,10,20F));
            string[] captions={"총 사업장 수","처리 완료 파일","총 근로자 수","대체근로자 수","확인 필요 항목"};string[] notes={"전체 등록 사업장","모든 필수 파일 완료","전체 사업장 합계","1개월 미만 근로자","추징/환급/분류필요"};string[] icons={"building","files","worker","short","warning"};Color[] colors={Color.FromArgb(102,78,238),UiGreen,Color.FromArgb(37,133,235),UiOrange,UiRed};summaryStatCards=new DashboardStatCard[5];for(int i=0;i<5;i++){summaryStatCards[i]=new DashboardStatCard{Location=new Point(8+i*206,59),Size=new Size(198,96),Caption=captions[i],Value="-",Note=notes[i],IconKind=icons[i],Accent=colors[i]};page.Controls.Add(summaryStatCards[i]);}
            var overview=(RoundedPanel)Card(8,168,1030,108,Color.White);overview.Radius=13;overview.BorderColor=Color.FromArgb(222,227,241);overview.Controls.Add(new Label{Text="고지 년월",Location=new Point(18,17),AutoSize=true,ForeColor=UiText,Font=new Font("맑은 고딕",8.5F,FontStyle.Bold)});summaryPeriodLabel=new Label{Text="-",Location=new Point(18,49),AutoSize=true,ForeColor=Color.FromArgb(33,91,235),Font=new Font("맑은 고딕",18F,FontStyle.Bold)};overview.Controls.Add(summaryPeriodLabel);overview.Controls.Add(new Panel{Location=new Point(205,16),Size=new Size(1,76),BackColor=UiBorder});summaryPremiumTotals=new PremiumTotalsControl{Location=new Point(224,12),Size=new Size(558,84)};overview.Controls.Add(summaryPremiumTotals);overview.Controls.Add(new Panel{Location=new Point(794,16),Size=new Size(1,76),BackColor=UiBorder});overview.Controls.Add(new Label{Text="사업장 관리번호 선택",Location=new Point(814,16),AutoSize=true,ForeColor=UiText,Font=new Font("맑은 고딕",8.5F,FontStyle.Bold)});summarySiteSelector=new ModernSiteSelector{Location=new Point(812,45),Size=new Size(202,44)};summarySiteSelector.SelectedIndexChanged+=(s,e)=>{if(!summaryComboLoading)UpdateSummaryForSelectedSite();};overview.Controls.Add(summarySiteSelector);page.Controls.Add(overview);
            summaryTable=new SummaryTableControl{Location=new Point(8,288),Size=new Size(1030,306)};page.Controls.Add(summaryTable);
            var legend=(RoundedPanel)Card(8,602,1030,38,Color.FromArgb(253,254,255));legend.Radius=10;legend.Controls.Add(new Label{Text="대조결과 안내",Location=new Point(16,10),AutoSize=true,ForeColor=UiText,Font=new Font("맑은 고딕",8F,FontStyle.Bold)});legend.Controls.Add(new Label{Text="●",Location=new Point(125,9),AutoSize=true,ForeColor=UiGreen});legend.Controls.Add(Muted("정상 (고지금액 = 급여대장 대조 결과)",145,10));legend.Controls.Add(new Label{Text="●",Location=new Point(455,9),AutoSize=true,ForeColor=UiRed});legend.Controls.Add(Muted("추징 필요",475,10));legend.Controls.Add(new Label{Text="●",Location=new Point(640,9),AutoSize=true,ForeColor=Color.FromArgb(45,119,231)});legend.Controls.Add(Muted("환급 필요",660,10));page.Controls.Add(legend);
        }
        void RefreshSummaryDashboard(){if(validationResult!=null&&!String.IsNullOrWhiteSpace(validationResult.Text)&&File.Exists(validationResult.Text))LoadResultIntoUi(validationResult.Text);else MessageBox.Show("먼저 파일 등록 화면에서 대사 작업을 실행해 주세요.","총괄표",MessageBoxButtons.OK,MessageBoxIcon.Information);}
        void ExportSummaryWorkbook(){if(validationResult==null||String.IsNullOrWhiteSpace(validationResult.Text)||!File.Exists(validationResult.Text)){MessageBox.Show("내보낼 대사 결과가 없습니다.","엑셀 내보내기",MessageBoxButtons.OK,MessageBoxIcon.Information);return;}using(SaveFileDialog d=new SaveFileDialog{Filter="Excel 매크로 통합 문서 (*.xlsm)|*.xlsm",FileName="4대보험_급여검증결과_"+DateTime.Now.ToString("yyyyMMdd_HHmmss")+".xlsm"})if(d.ShowDialog()==DialogResult.OK){File.Copy(validationResult.Text,d.FileName,true);MessageBox.Show("엑셀 파일을 내보냈습니다.","엑셀 내보내기",MessageBoxButtons.OK,MessageBoxIcon.Information);OpenGeneratedFileIfEnabled(d.FileName);}}
        void BuildIndividualScreen(Control page)
        {
            page.Controls.Add(TitleLabel("개인별 내역",8,10,20F));
            string[] captions={"전체 인원","정상","추징 필요","환급 필요","확인 필요 항목"};string[] notes={"모든 재원 합계","이상 없음","고지금액 > 급여대장","고지금액 < 급여대장","추징/환급/분류필요"};string[] icons={"people","normal","collection","refund","warning"};Color[] colors={Color.FromArgb(102,78,238),UiGreen,UiRed,Color.FromArgb(43,119,231),UiOrange};individualStatCards=new DashboardStatCard[5];for(int i=0;i<5;i++){individualStatCards[i]=new DashboardStatCard{Location=new Point(8+i*206,59),Size=new Size(198,96),Caption=captions[i],Value="-",Note=notes[i],IconKind=icons[i],Accent=colors[i]};page.Controls.Add(individualStatCards[i]);}
            var filter=(RoundedPanel)Card(8,168,1030,92,Color.White);filter.Radius=13;filter.BorderColor=Color.FromArgb(222,227,241);filter.Controls.Add(new Label{Text="고지년월",Location=new Point(18,16),AutoSize=true,ForeColor=UiText,Font=new Font("맑은 고딕",8.5F,FontStyle.Bold)});individualPeriodLabel=new Label{Text="-",Location=new Point(18,48),AutoSize=true,ForeColor=Color.FromArgb(33,91,235),Font=new Font("맑은 고딕",15F,FontStyle.Bold)};filter.Controls.Add(individualPeriodLabel);filter.Controls.Add(new Panel{Location=new Point(205,16),Size=new Size(1,60),BackColor=UiBorder});filter.Controls.Add(new Label{Text="사업장 관리번호 선택",Location=new Point(232,14),AutoSize=true,ForeColor=UiText,Font=new Font("맑은 고딕",8.5F,FontStyle.Bold)});individualSiteSelector=new ModernSiteSelector{Location=new Point(228,40),Size=new Size(300,42)};individualSiteSelector.SelectedIndexChanged+=(s,e)=>{if(!individualFilterLoading){RebuildIndividualFundChoices(true);UpdateIndividualView();}};filter.Controls.Add(individualSiteSelector);filter.Controls.Add(new Panel{Location=new Point(550,16),Size=new Size(1,60),BackColor=UiBorder});filter.Controls.Add(new Label{Text="재원 선택",Location=new Point(578,14),AutoSize=true,ForeColor=UiText,Font=new Font("맑은 고딕",8.5F,FontStyle.Bold)});individualFundSelector=new ModernSiteSelector{Location=new Point(574,40),Size=new Size(190,42),ShowIcon=false};individualFundSelector.SelectedIndexChanged+=(s,e)=>{if(!individualFilterLoading)UpdateIndividualView();};filter.Controls.Add(individualFundSelector);filter.Controls.Add(new Label{Text="검색",Location=new Point(790,14),AutoSize=true,ForeColor=UiText,Font=new Font("맑은 고딕",8.5F,FontStyle.Bold)});var searchShell=(RoundedPanel)Card(784,39,229,44,Color.White);searchShell.Radius=9;searchShell.BorderColor=Color.FromArgb(157,176,226);searchShell.BorderWidth=2;searchShell.Controls.Add(new Label{Text="⌕",Location=new Point(11,9),AutoSize=true,ForeColor=Color.FromArgb(55,75,192),Font=new Font("맑은 고딕",13F,FontStyle.Bold)});individualSearchBox=new CueTextBox{CueText="이름 또는 금액",Location=new Point(38,12),Size=new Size(174,20),BorderStyle=BorderStyle.None,BackColor=Color.White,ForeColor=Color.FromArgb(31,49,115),Font=new Font("맑은 고딕",8.5F)};individualSearchBox.TextChanged+=(s,e)=>{if(!individualFilterLoading)UpdateIndividualView();};searchShell.Controls.Add(individualSearchBox);filter.Controls.Add(searchShell);page.Controls.Add(filter);
            string[] amountTabs={"개인부담금","기관부담금"};individualModeTabs=new IndividualModeTabButton[2];for(int i=0;i<2;i++){string mode=amountTabs[i];individualModeTabs[i]=new IndividualModeTabButton{Caption=mode,Active=i==0,Accent=UiBlue,Location=new Point(8+i*166,268),Size=new Size(158,34)};individualModeTabs[i].Click+=(s,e)=>SelectIndividualAmountMode(mode);page.Controls.Add(individualModeTabs[i]);}
            individualTable=new IndividualTableControl{Location=new Point(8,306),Size=new Size(1030,248),PageSize=IndividualPageSize,DiscountProvider=SavedDiscount};page.Controls.Add(individualTable);
            var pager=(RoundedPanel)Card(8,560,1030,32,Color.FromArgb(253,254,255));pager.Radius=9;individualRangeLabel=new Label{Text="표시할 데이터가 없습니다.",Location=new Point(14,8),AutoSize=true,ForeColor=UiMuted,Font=new Font("맑은 고딕",7.5F,FontStyle.Bold)};pager.Controls.Add(individualRangeLabel);page.Controls.Add(pager);
            var legend=(RoundedPanel)Card(8,600,1030,34,Color.FromArgb(253,254,255));legend.Radius=9;legend.Controls.Add(new Label{Text="대사결과 안내",Location=new Point(14,9),AutoSize=true,ForeColor=UiText,Font=new Font("맑은 고딕",7.5F,FontStyle.Bold)});AddIndividualLegend(legend,130,"●",UiGreen,"정상");AddIndividualLegend(legend,305,"●",UiRed,"추징 필요");AddIndividualLegend(legend,495,"●",Color.FromArgb(43,119,231),"환급 필요");AddIndividualLegend(legend,690,"△",UiOrange,"확인 필요");page.Controls.Add(legend);
        }
        void AddIndividualLegend(Control parent,int x,string icon,Color color,string text){parent.Controls.Add(new Label{Text=icon,Location=new Point(x,8),AutoSize=true,ForeColor=color,Font=new Font("맑은 고딕",8F,FontStyle.Bold)});parent.Controls.Add(new Label{Text=text,Location=new Point(x+20,9),AutoSize=true,ForeColor=UiMuted,Font=new Font("맑은 고딕",7.5F)});}
        void RefreshIndividualDashboard(){if(validationResult!=null&&!String.IsNullOrWhiteSpace(validationResult.Text)&&File.Exists(validationResult.Text))LoadResultIntoUi(validationResult.Text);else MessageBox.Show("먼저 파일 등록 화면에서 대사 작업을 실행해 주세요.","개인별 내역",MessageBoxButtons.OK,MessageBoxIcon.Information);}
        void SelectIndividualAmountMode(string mode){individualAmountMode=mode=="기관부담금"?"기관부담금":"개인부담금";if(individualModeTabs!=null)foreach(IndividualModeTabButton tab in individualModeTabs){tab.Active=tab.Caption==individualAmountMode;tab.Invalidate();}if(individualTable!=null){individualTable.InstitutionMode=individualAmountMode=="기관부담금";individualTable.ScrollOffset=0;individualTable.Invalidate();}UpdateIndividualView();}
        void BuildAdjustmentScreen(Control page)
        {
            page.Controls.Add(TitleLabel("반환 / 추징",8,10,20F));page.Controls.Add(new Label{Text="보험료 정산 결과에 따라 반환·추징 및 분류 필요 항목을 확인하고 관리합니다.",Location=new Point(177,26),AutoSize=true,ForeColor=UiMuted,Font=new Font("맑은 고딕",8F),Tag="ThemeMuted"});
            string[] captions={"전체 정산 대상","반환 (환급)","추징 (추가 납부)","분류 필요"};string[] notes={"반환·추징·분류 대상","공제액이 고지액보다 큼","고지액이 공제액보다 큼","확인이 필요한 항목"};string[] icons={"people","refund","collection","warning"};Color[] colors={Color.FromArgb(102,78,238),Color.FromArgb(43,119,231),UiRed,UiOrange};adjustmentStatCards=new DashboardStatCard[4];for(int i=0;i<4;i++){adjustmentStatCards[i]=new DashboardStatCard{Location=new Point(8+i*258,59),Size=new Size(249,96),Caption=captions[i],Value="-",Note=notes[i],IconKind=icons[i],Accent=colors[i]};page.Controls.Add(adjustmentStatCards[i]);}
            adjustmentFilterPanel=(RoundedPanel)Card(8,168,1030,84,Color.White);var filter=adjustmentFilterPanel;filter.Radius=13;filter.BorderColor=Color.FromArgb(222,227,241);filter.Controls.Add(new Label{Text="고지년월",Location=new Point(18,14),AutoSize=true,ForeColor=UiText,Font=new Font("맑은 고딕",8.5F,FontStyle.Bold)});adjustmentPeriodLabel=new Label{Text="-",Location=new Point(18,43),AutoSize=true,ForeColor=Color.FromArgb(33,91,235),Font=new Font("맑은 고딕",15F,FontStyle.Bold)};filter.Controls.Add(adjustmentPeriodLabel);filter.Controls.Add(new Panel{Location=new Point(205,14),Size=new Size(1,56),BackColor=UiBorder});filter.Controls.Add(new Label{Text="사업장 관리번호 선택",Location=new Point(235,12),AutoSize=true,ForeColor=UiText,Font=new Font("맑은 고딕",8.5F,FontStyle.Bold)});adjustmentSiteSelector=new ModernSiteSelector{Location=new Point(231,36),Size=new Size(365,42)};adjustmentSiteSelector.SelectedIndexChanged+=(s,e)=>{if(!adjustmentFilterLoading){RebuildAdjustmentFundChoices(true);UpdateAdjustmentView();}};filter.Controls.Add(adjustmentSiteSelector);filter.Controls.Add(new Panel{Location=new Point(624,14),Size=new Size(1,56),BackColor=UiBorder});filter.Controls.Add(new Label{Text="재원 선택",Location=new Point(654,12),AutoSize=true,ForeColor=UiText,Font=new Font("맑은 고딕",8.5F,FontStyle.Bold)});adjustmentFundSelector=new ModernSiteSelector{Location=new Point(650,36),Size=new Size(250,42),ShowIcon=false};adjustmentFundSelector.SelectedIndexChanged+=(s,e)=>{if(!adjustmentFilterLoading)UpdateAdjustmentView();};filter.Controls.Add(adjustmentFundSelector);page.Controls.Add(filter);
            string[] modes={"전체","반환","추징","분류 필요"};Color[] tabColors={UiPurple,Color.FromArgb(43,119,231),UiRed,UiOrange};adjustmentTabs=new AdjustmentTabButton[4];for(int i=0;i<4;i++){int index=i;adjustmentTabs[i]=new AdjustmentTabButton{Location=new Point(8+i*150,262),Size=new Size(142,38),Caption=modes[i],Accent=tabColors[i],Active=i==0};adjustmentTabs[i].Click+=(s,e)=>SelectAdjustmentMode(modes[index]);page.Controls.Add(adjustmentTabs[i]);}
            adjustmentTable=new AdjustmentTableControl{Location=new Point(8,302),Size=new Size(1030,258),Mode="전체",PageSize=4,SelectionKeys=null};page.Controls.Add(adjustmentTable);
            var footer=(RoundedPanel)Card(8,568,1030,66,Color.FromArgb(253,254,255));footer.Radius=10;adjustmentRangeLabel=new Label{Text="표시할 대상이 없습니다.",Location=new Point(16,10),AutoSize=true,ForeColor=UiMuted,Font=new Font("맑은 고딕",7.5F,FontStyle.Bold)};footer.Controls.Add(adjustmentRangeLabel);adjustmentSelectionLabel=new Label{Text="현재 목록 0명",Location=new Point(16,36),AutoSize=true,ForeColor=UiText,Font=new Font("맑은 고딕",8F,FontStyle.Bold)};footer.Controls.Add(adjustmentSelectionLabel);adjustmentAmountLabel=new Label{Text="정산 금액 0원",Location=new Point(135,36),AutoSize=true,ForeColor=UiBlue,Font=new Font("맑은 고딕",8F,FontStyle.Bold)};footer.Controls.Add(adjustmentAmountLabel);adjustmentExcelButton=OutputButton("Excel 생성","excel",724,15,140,38,UiGreen);adjustmentExcelButton.Click+=(s,e)=>ExportAdjustmentExcel();footer.Controls.Add(adjustmentExcelButton);adjustmentPdfButton=OutputButton("PDF 생성","pdf",874,15,140,38,UiRed);adjustmentPdfButton.Click+=(s,e)=>ExportAdjustmentPdf();footer.Controls.Add(adjustmentPdfButton);page.Controls.Add(footer);
        }
        void BuildReviewScreen(Control page)
        {
            page.Controls.Add(TitleLabel("확인 필요",8,10,20F));page.Controls.Add(new Label{Text="자동 대사로 결과 판정이 어려운 항목입니다. 항목을 확인하고 조치 내용을 등록해 주세요.",Location=new Point(166,26),AutoSize=true,ForeColor=UiMuted,Font=new Font("맑은 고딕",8F),Tag="ThemeMuted"});
            string[] captions={"전체 확인 필요 항목","건강보험 (건강+장기요양)","국민연금","고용보험","산재보험"};string[] icons={"warning","health","people","briefcase","warning"};Color[] colors={Color.FromArgb(180,49,42),Color.FromArgb(36,102,225),Color.FromArgb(91,55,238),Color.FromArgb(19,151,78),Color.FromArgb(236,116,36)};reviewStatCards=new DashboardStatCard[5];for(int i=0;i<5;i++){reviewStatCards[i]=new DashboardStatCard{Location=new Point(8+i*206,59),Size=new Size(198,96),Caption=captions[i],Value="-",Note=i==0?"주의가 필요한 항목":"총 0원",IconKind=icons[i],Accent=colors[i]};page.Controls.Add(reviewStatCards[i]);}
            var filter=(RoundedPanel)Card(8,168,1030,84,Color.White);filter.Radius=13;filter.BorderColor=Color.FromArgb(222,227,241);filter.Controls.Add(new Label{Text="고지년월",Location=new Point(18,14),AutoSize=true,ForeColor=UiText,Font=new Font("맑은 고딕",8.5F,FontStyle.Bold)});reviewPeriodLabel=new Label{Text="-",Location=new Point(18,43),AutoSize=true,ForeColor=Color.FromArgb(33,91,235),Font=new Font("맑은 고딕",15F,FontStyle.Bold)};filter.Controls.Add(reviewPeriodLabel);filter.Controls.Add(new Panel{Location=new Point(205,14),Size=new Size(1,56),BackColor=UiBorder});filter.Controls.Add(new Label{Text="사업장 관리번호 선택",Location=new Point(235,12),AutoSize=true,ForeColor=UiText,Font=new Font("맑은 고딕",8.5F,FontStyle.Bold)});reviewSiteSelector=new ModernSiteSelector{Location=new Point(231,36),Size=new Size(365,42)};reviewSiteSelector.SelectedIndexChanged+=(s,e)=>{if(!reviewFilterLoading){RebuildReviewFundChoices(true);UpdateReviewView();}};filter.Controls.Add(reviewSiteSelector);filter.Controls.Add(new Panel{Location=new Point(624,14),Size=new Size(1,56),BackColor=UiBorder});filter.Controls.Add(new Label{Text="재원 선택",Location=new Point(654,12),AutoSize=true,ForeColor=UiText,Font=new Font("맑은 고딕",8.5F,FontStyle.Bold)});reviewFundSelector=new ModernSiteSelector{Location=new Point(650,36),Size=new Size(250,42),ShowIcon=false};reviewFundSelector.SelectedIndexChanged+=(s,e)=>{if(!reviewFilterLoading)UpdateReviewView();};filter.Controls.Add(reviewFundSelector);page.Controls.Add(filter);
            page.Controls.Add(new Label{Text="확인 필요 명단",Location=new Point(10,266),AutoSize=true,ForeColor=UiText,Font=new Font("맑은 고딕",10F,FontStyle.Bold)});var applyPanel=(RoundedPanel)Card(620,253,418,42,Color.FromArgb(255,246,238));applyPanel.Radius=11;applyPanel.BorderWidth=1;applyPanel.BorderColor=Color.FromArgb(239,157,106);applyPanel.Controls.Add(new Label{Text="적용 재원",Location=new Point(16,13),AutoSize=true,ForeColor=Color.FromArgb(165,76,30),BackColor=Color.Transparent,Font=new Font("맑은 고딕",8F,FontStyle.Bold)});reviewApplyFundSelector=new ModernSiteSelector{Location=new Point(118,6),Size=new Size(282,30),ShowIcon=false,Borderless=true};reviewApplyFundSelector.SetItems(new[]{"적용 재원 선택","계약제교원","교특회계","학교회계"});reviewApplyFundSelector.SelectedIndex=0;reviewApplyFundSelector.SelectedIndexChanged+=(s,e)=>ApplyFundToSelectedReviews();applyPanel.Controls.Add(reviewApplyFundSelector);page.Controls.Add(applyPanel);reviewTable=new ReviewTableControl{Location=new Point(8,303),Size=new Size(1030,279),PageSize=5,SelectionKeys=reviewSelections,CheckedKeys=reviewCheckedKeys,FundDrafts=reviewFundDrafts,FundChoices=row=>ReviewFundsForSite202(row.Site)};reviewTable.SelectionChanged+=UpdateReviewSelectionSummary;reviewTable.DetailRequested+=(row,point)=>ShowContributionEditor202(row,point);reviewTable.FundChanged+=(row,fund)=>{reviewFundDrafts[ReviewKey(row)]=fund;reviewTable.Invalidate();UpdateReviewSelectionSummary();};page.Controls.Add(reviewTable);
            var footer=(RoundedPanel)Card(8,590,1030,44,Color.FromArgb(253,254,255));footer.Radius=10;reviewRangeLabel=new Label{Text="표시할 대상이 없습니다.",Location=new Point(14,7),AutoSize=true,ForeColor=UiMuted,Font=new Font("맑은 고딕",7.5F,FontStyle.Bold)};footer.Controls.Add(reviewRangeLabel);reviewSelectionLabel=new Label{Text="선택 0건",Location=new Point(14,25),AutoSize=true,ForeColor=UiText,Font=new Font("맑은 고딕",7.5F,FontStyle.Bold)};footer.Controls.Add(reviewSelectionLabel);reviewAmountLabel=new Label{Text="조정금액 0원",Location=new Point(105,25),AutoSize=true,ForeColor=UiBlue,Font=new Font("맑은 고딕",7.5F,FontStyle.Bold)};footer.Controls.Add(reviewAmountLabel);var checkedButton=ActionButton("조회 상태 일괄 변경",690,3,174,38,Color.FromArgb(65,86,222));checkedButton.Font=new Font("맑은 고딕",8.5F,FontStyle.Bold);checkedButton.Click+=(s,e)=>MarkSelectedReviewsChecked();footer.Controls.Add(checkedButton);var save=OutputButton("수정사항 저장","save",874,3,140,38,UiPurple);save.Click+=(s,e)=>SaveReviewChanges();footer.Controls.Add(save);page.Controls.Add(footer);
        }
        bool IsReviewCompleted(IndividualRowData row){return row!=null&&reviewCheckedKeys.Contains(ReviewKey(row));}
        bool IsReviewRow(IndividualRowData row){return row!=null&&(IsReviewCompleted(row)||row.Fund=="분류필요"||row.Status!="정상"||HasCollectionDirection(row)||HasRefundDirection(row));}
        bool IsPendingReviewRow(IndividualRowData row){return IsReviewRow(row)&&!IsReviewCompleted(row);}
        string ReviewKey(IndividualRowData row){return StablePersonKey(row);}
        string PrimaryReviewInsurance(IndividualRowData row){string reason=row.ReviewReason??"";if(reason.Contains("건강")||reason.Contains("장기"))return "건강보험";if(reason.Contains("국민"))return "국민연금";if(reason.Contains("고용"))return "고용보험";if(reason.Contains("산재"))return "산재보험";decimal[] amounts={Math.Abs(row.HealthDifference),Math.Abs(row.PensionDifference),Math.Abs(row.EmploymentDifference),Math.Abs(row.IndustrialDifference)};int best=0;for(int i=1;i<amounts.Length;i++)if(amounts[i]>amounts[best])best=i;return new[]{"건강보험","국민연금","고용보험","산재보험"}[best];}
        decimal ReviewAmount(IndividualRowData row){return Math.Abs(row.HealthDifference)+Math.Abs(row.PensionDifference)+Math.Abs(row.EmploymentDifference)+Math.Abs(row.IndustrialDifference);}
        string ReviewReasonText(IndividualRowData row){if(!String.IsNullOrWhiteSpace(row.ReviewReason))return row.ReviewReason;string insurance=PrimaryReviewInsurance(row);decimal diff=insurance=="건강보험"?row.HealthDifference:insurance=="국민연금"?row.PensionDifference:insurance=="고용보험"?row.EmploymentDifference:row.IndustrialDifference;if(row.Fund=="분류필요")return "급여대장 재원 분류가 필요합니다.";return insurance+" "+(diff>0?"고지금액 > 급여대장 금액 차이":diff<0?"고지금액 < 급여대장 금액 차이":"세부 내역 확인 필요");}
        void InitializeReviewFilters(){if(reviewSiteSelector==null)return;RememberReviewRows202();string previous=reviewSiteSelector.SelectedIndex>=0&&reviewSiteSelector.SelectedIndex<reviewSiteKeys.Count?reviewSiteKeys[reviewSiteSelector.SelectedIndex]:"";reviewFilterLoading=true;reviewSiteKeys.Clear();var sites=(individualDashboard==null?Enumerable.Empty<IndividualRowData>():individualDashboard.Rows).Where(IsListedReview202).Select(x=>x.Site).Distinct().OrderBy(x=>x).ToList();reviewSiteKeys.AddRange(sites);reviewSiteSelector.SetItems(sites.Select(FormatSite));if(sites.Count>0){int index=sites.IndexOf(previous);reviewSiteSelector.SelectedIndex=index>=0?index:0;}RebuildReviewFundChoices(true);reviewFilterLoading=false;UpdateReviewStatistics();UpdateReviewView();}
        void RebuildReviewFundChoices(bool reset){if(reviewFundSelector==null)return;string previous=reviewFundSelector.SelectedIndex>=0?reviewFundSelector.Items[reviewFundSelector.SelectedIndex]:"전체";IEnumerable<IndividualRowData> rows=individualDashboard==null?Enumerable.Empty<IndividualRowData>():individualDashboard.Rows.Where(IsListedReview202);if(reviewSiteSelector!=null&&reviewSiteSelector.SelectedIndex>=0&&reviewSiteSelector.SelectedIndex<reviewSiteKeys.Count){string site=reviewSiteKeys[reviewSiteSelector.SelectedIndex];rows=rows.Where(x=>x.Site==site);}var choices=new List<string>{"전체"};choices.AddRange(rows.Select(x=>x.Fund=="분류필요"?"분류 필요":x.Fund).Where(x=>!String.IsNullOrWhiteSpace(x)).Distinct().OrderBy(x=>x=="분류 필요"?99:UiFundDisplayOrder(x)));bool old=reviewFilterLoading;reviewFilterLoading=true;reviewFundSelector.SetItems(choices);int index=reset?0:choices.IndexOf(previous);reviewFundSelector.SelectedIndex=index>=0?index:0;reviewFilterLoading=old;}
        List<IndividualRowData> FilteredReviewRows(){IEnumerable<IndividualRowData> rows=individualDashboard==null?Enumerable.Empty<IndividualRowData>():individualDashboard.Rows.Where(IsListedReview202);if(reviewSiteSelector!=null&&reviewSiteSelector.SelectedIndex>=0&&reviewSiteSelector.SelectedIndex<reviewSiteKeys.Count){string site=reviewSiteKeys[reviewSiteSelector.SelectedIndex];rows=rows.Where(x=>x.Site==site);}string fund=reviewFundSelector!=null&&reviewFundSelector.SelectedIndex>=0?reviewFundSelector.Items[reviewFundSelector.SelectedIndex]:"전체";if(fund!="전체")rows=rows.Where(x=>(x.Fund=="분류필요"?"분류 필요":x.Fund)==fund);return rows.OrderBy(x=>x.Name).ToList();}
        void UpdateReviewStatistics(){if(reviewStatCards==null)return;var rows=(individualDashboard==null?new List<IndividualRowData>():individualDashboard.Rows.Where(IsPendingReviewRow).ToList());int total=rows.Count;string[] insurance={"건강보험","국민연금","고용보험","산재보험"};reviewStatCards[0].Value=total+"건";reviewStatCards[0].Note="주의가 필요한 미확인 항목";reviewStatCards[0].Invalidate();for(int i=0;i<4;i++){string kind=insurance[i];var matches=rows.Where(x=>PrimaryReviewInsurance(x)==kind).ToList();reviewStatCards[i+1].Value=matches.Count+"건"+(total>0?" ("+(matches.Count*100.0/total).ToString("0.0")+"%)":"");reviewStatCards[i+1].Note="총 "+UiDrawing.Money(matches.Sum(x=>ReviewAmount(x)))+"원";reviewStatCards[i+1].Invalidate();}reviewPeriodLabel.Text=individualDashboard!=null&&individualDashboard.Year>0?individualDashboard.Year+"년 "+individualDashboard.Month+"월":"-";}
        void UpdateReviewView(){if(reviewTable==null)return;UpdateReviewFundOptions202();List<IndividualRowData> rows=FilteredReviewRows();reviewTable.Rows=rows;reviewTable.ScrollOffset=0;reviewTable.Invalidate();reviewRangeLabel.Text="전체 "+rows.Count+"건"+(rows.Count>0?" 중 1~"+Math.Min(rows.Count,reviewTable.PageSize)+"건 표시":"");UpdateReviewSelectionSummary();}
        void UpdateReviewSelectionSummary(){if(reviewSelectionLabel==null)return;List<IndividualRowData> selected=FilteredReviewRows().Where(x=>reviewSelections.Contains(ReviewKey(x))).ToList();reviewSelectionLabel.Text="선택 "+selected.Count+"건";reviewAmountLabel.Text="조정금액 "+UiDrawing.Money(selected.Sum(x=>ReviewAmount(x)))+"원";if(reviewTable!=null)reviewTable.Invalidate();}
        void ApplyFundToSelectedReviews(){if(reviewFilterLoading||reviewApplyFundSelector==null||reviewApplyFundSelector.SelectedIndex<=0)return;string fund=reviewApplyFundSelector.Items[reviewApplyFundSelector.SelectedIndex];List<IndividualRowData> selected=FilteredReviewRows().Where(x=>reviewSelections.Contains(ReviewKey(x))).ToList();if(selected.Count==0){MessageBox.Show("재원을 적용할 대상을 먼저 체크해 주세요.","적용 재원",MessageBoxButtons.OK,MessageBoxIcon.Information);reviewFilterLoading=true;reviewApplyFundSelector.SelectedIndex=0;reviewFilterLoading=false;return;}foreach(IndividualRowData row in selected)reviewFundDrafts[ReviewKey(row)]=fund;reviewTable.Invalidate();UpdateReviewSelectionSummary();reviewFilterLoading=true;reviewApplyFundSelector.SelectedIndex=0;reviewFilterLoading=false;}
        void RefreshReviewDashboard(){if(validationResult!=null&&!String.IsNullOrWhiteSpace(validationResult.Text)&&File.Exists(validationResult.Text))LoadResultIntoUi(validationResult.Text);else MessageBox.Show("먼저 파일 등록 화면에서 대사 작업을 실행해 주세요.","확인 필요",MessageBoxButtons.OK,MessageBoxIcon.Information);}
        void ExportReviewWorkbook(){if(validationResult==null||String.IsNullOrWhiteSpace(validationResult.Text)||!File.Exists(validationResult.Text)){MessageBox.Show("내보낼 확인 필요 자료가 없습니다.","엑셀 내보내기",MessageBoxButtons.OK,MessageBoxIcon.Information);return;}using(SaveFileDialog d=new SaveFileDialog{Filter="Excel 매크로 통합 문서 (*.xlsm)|*.xlsm",FileName="확인필요_내역_"+DateTime.Now.ToString("yyyyMMdd_HHmmss")+".xlsm"})if(d.ShowDialog()==DialogResult.OK){File.Copy(validationResult.Text,d.FileName,true);MessageBox.Show("확인 필요 내역이 포함된 엑셀 파일을 내보냈습니다.","엑셀 내보내기",MessageBoxButtons.OK,MessageBoxIcon.Information);OpenGeneratedFileIfEnabled(d.FileName);}}
        void ShowReviewDetail(IndividualRowData row,Point screenPoint){if(reviewBubble!=null&&!reviewBubble.IsDisposed)reviewBubble.Close();string draft;reviewFundDrafts.TryGetValue(ReviewKey(row),out draft);reviewBubble=new ReviewDetailBubble(row,String.IsNullOrWhiteSpace(draft)?row.Fund:draft,ReviewReasonText(row),reviewCheckedKeys.Contains(ReviewKey(row)));Rectangle work=Screen.FromControl(this).WorkingArea;int x=Math.Min(work.Right-350,Right-8),y=Math.Max(work.Top+12,Math.Min(work.Bottom-reviewBubble.Height-12,screenPoint.Y-reviewBubble.Height/3));reviewBubble.StartPosition=FormStartPosition.Manual;reviewBubble.Location=new Point(x,y);reviewBubble.Show(this);}
        void MarkSelectedReviewsCheckedLegacy(){List<IndividualRowData> selected=FilteredReviewRows().Where(x=>reviewSelections.Contains(ReviewKey(x))).ToList();if(selected.Count==0){MessageBox.Show("조회 상태를 변경할 대상을 체크해 주세요.","확인 필요",MessageBoxButtons.OK,MessageBoxIcon.Information);return;}foreach(IndividualRowData row in selected)reviewCheckedKeys.Add(ReviewKey(row));NormalizeIndividualStatuses();if(individualDashboard.Rows.All(x=>x.HasSummaryBreakdown))RebuildSummaryDashboardFromIndividuals();RefreshLinkedResultViews();}
        void SaveReviewChanges(){Safe202(()=>{RequireResult202();int count=PersistReviewChangesCore(validationResult.Text);reviewSelectionLabel.Text="저장 완료";reviewAmountLabel.Text=count>0?"재원 "+count+"명 반영":"확인 상태 반영";});}
        int PersistReviewChangesCore(string path){var changed=new List<IndividualRowData>();foreach(IndividualRowData row in individualDashboard.Rows){string fund;if(!reviewFundDrafts.TryGetValue(ReviewKey(row),out fund)||!new[]{"공무원","계약제교원","교특회계","학교회계"}.Contains(fund))continue;row.Fund=fund;row.ReviewReason="재원 확인 완료";changed.Add(row);}NormalizeIndividualStatuses();if(individualDashboard.Rows.All(x=>x.HasSummaryBreakdown))RebuildSummaryDashboardFromIndividuals();PersistReviewState(path);reviewFundDrafts.Clear();reviewSelections.Clear();reconciliationState.Revision++;individualFilterLoading=true;RebuildIndividualFundChoices(false);individualFilterLoading=false;InitializeAdjustmentFilters();InitializeReviewFilters();RefreshLinkedResultViews();return changed.Count;}
        void LoadReviewStateBase(ExcelPackage package){reviewCheckedKeys.Clear();ExcelWorksheet ws=package.Workbook.Worksheets["UI확인상태"];if(ws==null||ws.Dimension==null)return;for(int r=2;r<=ws.Dimension.End.Row;r++)if(!String.IsNullOrWhiteSpace(ws.Cells[r,1].Text))reviewCheckedKeys.Add(ws.Cells[r,1].Text);}
        void PersistReviewStateBase(string path){using(ExcelPackage package=new ExcelPackage(new FileInfo(path))){ExcelWorksheet individual=package.Workbook.Worksheets["UI개인별데이터"];if(individual!=null&&individual.Dimension!=null)for(int r=2;r<=individual.Dimension.End.Row;r++){string key=String.Join("|",new[]{individual.Cells[r,1].Text,individual.Cells[r,3].Text,Regex.Replace(individual.Cells[r,4].Text,"[^0-9]","")});IndividualRowData row=individualDashboard.Rows.FirstOrDefault(x=>StablePersonKey(x)==key);if(row==null)continue;individual.Cells[r,2].Value=row.Fund;individual.Cells[r,6].Value=row.Status;individual.Cells[r,9].Value=row.HealthDifference;individual.Cells[r,12].Value=row.PensionDifference;individual.Cells[r,15].Value=row.EmploymentDifference;individual.Cells[r,18].Value=row.IndustrialDifference;individual.Cells[r,21].Value=row.ReviewReason;}WriteSummaryDashboardSheet(package);ExcelWorksheet old=package.Workbook.Worksheets["UI확인상태"];if(old!=null)package.Workbook.Worksheets.Delete(old);ExcelWorksheet ws=package.Workbook.Worksheets.Add("UI확인상태");ws.Cells[1,1].Value="대상키";ws.Cells[1,2].Value="조회상태";int rowIndex=2;foreach(string key in reviewCheckedKeys.OrderBy(x=>x)){ws.Cells[rowIndex,1].Value=key;ws.Cells[rowIndex,2].Value="확인 완료";rowIndex++;}ws.Hidden=eWorkSheetHidden.Hidden;package.Save();}}
        void BuildDiscountScreen(Control page)
        {
            page.Controls.Add(TitleLabel("감면 적용",8,10,20F));page.Controls.Add(new Label{Text="보험별 기관부담 감면내역을 입력하고 적용 후 기관부담금을 확인해 주세요.",Location=new Point(166,26),AutoSize=true,ForeColor=UiMuted,Font=new Font("맑은 고딕",8F),Tag="ThemeMuted"});
            var filter=(RoundedPanel)Card(8,59,1030,84,Color.White);filter.Radius=13;filter.BorderColor=Color.FromArgb(222,227,241);filter.Controls.Add(new Label{Text="고지년월",Location=new Point(18,14),AutoSize=true,ForeColor=UiText,Font=new Font("맑은 고딕",8.5F,FontStyle.Bold)});discountPeriodLabel=new Label{Text="-",Location=new Point(18,43),AutoSize=true,ForeColor=Color.FromArgb(33,91,235),Font=new Font("맑은 고딕",15F,FontStyle.Bold)};filter.Controls.Add(discountPeriodLabel);filter.Controls.Add(new Panel{Location=new Point(205,14),Size=new Size(1,56),BackColor=UiBorder});filter.Controls.Add(new Label{Text="사업장 관리번호 선택",Location=new Point(235,12),AutoSize=true,ForeColor=UiText,Font=new Font("맑은 고딕",8.5F,FontStyle.Bold)});discountSiteSelector=new ModernSiteSelector{Location=new Point(231,36),Size=new Size(365,42)};discountSiteSelector.SelectedIndexChanged+=(s,e)=>{if(!discountFilterLoading){RebuildDiscountFundChoices(true);UpdateDiscountView();}};filter.Controls.Add(discountSiteSelector);filter.Controls.Add(new Panel{Location=new Point(624,14),Size=new Size(1,56),BackColor=UiBorder});filter.Controls.Add(new Label{Text="재원 선택",Location=new Point(654,12),AutoSize=true,ForeColor=UiText,Font=new Font("맑은 고딕",8.5F,FontStyle.Bold)});discountFundSelector=new ModernSiteSelector{Location=new Point(650,36),Size=new Size(250,42),ShowIcon=false};discountFundSelector.SelectedIndexChanged+=(s,e)=>{if(!discountFilterLoading)UpdateDiscountView();};filter.Controls.Add(discountFundSelector);page.Controls.Add(filter);
            discountBilledTotals=new DiscountTotalsControl{Location=new Point(8,157),Size=new Size(360,190),Title="기관부담 고지금액",Tint=Color.FromArgb(249,251,255),Mode=0};discountAppliedTotals=new DiscountTotalsControl{Location=new Point(380,157),Size=new Size(280,190),Title="기관부담 감면액",Tint=Color.FromArgb(255,252,247),Mode=1};discountAfterTotals=new DiscountTotalsControl{Location=new Point(672,157),Size=new Size(366,190),Title="감면 적용 후 기관부담금",Tint=Color.FromArgb(246,253,248),Mode=2};page.Controls.Add(discountBilledTotals);page.Controls.Add(discountAppliedTotals);page.Controls.Add(discountAfterTotals);
            page.Controls.Add(new Label{Text="기관부담 감면 적용 명단",Location=new Point(10,355),AutoSize=true,ForeColor=UiText,Font=new Font("맑은 고딕",10F,FontStyle.Bold)});discountTable=new DiscountTableControl{Location=new Point(8,380),Size=new Size(1030,208),PageSize=5};discountTable.EntryProvider=row=>EffectiveDiscount(row);discountTable.EntryChanged+=(row,entry)=>{discountDrafts[DiscountKey(row)]=entry;UpdateDiscountView(false);};discountTable.AmountEditRequested+=(row,kind,current)=>EditDiscountAmount(row,kind,current);page.Controls.Add(discountTable);
            var footer=(RoundedPanel)Card(8,592,1030,42,Color.FromArgb(253,254,255));footer.Radius=9;discountRangeLabel=new Label{Text="표시할 대상이 없습니다.",Location=new Point(14,13),AutoSize=true,ForeColor=UiMuted,Font=new Font("맑은 고딕",7.5F,FontStyle.Bold)};footer.Controls.Add(discountRangeLabel);var save=OutputButton("감면내역 저장","save",874,2,140,38,Color.FromArgb(32,84,225));save.Click+=(s,e)=>SaveDiscountChanges();footer.Controls.Add(save);page.Controls.Add(footer);
        }
        void InitializeDiscountFilters(){if(discountSiteSelector==null)return;discountFilterLoading=true;discountSiteKeys.Clear();var sites=(individualDashboard==null?Enumerable.Empty<IndividualRowData>():individualDashboard.Rows).Select(x=>x.Site).Distinct().OrderBy(x=>x).ToList();discountSiteKeys.AddRange(sites);discountSiteSelector.SetItems(sites.Select(FormatSite));if(sites.Count>0)discountSiteSelector.SelectedIndex=0;RebuildDiscountFundChoices(true);discountFilterLoading=false;discountPeriodLabel.Text=individualDashboard!=null&&individualDashboard.Year>0?individualDashboard.Year+"년 "+individualDashboard.Month+"월":"-";UpdateDiscountView();}
        void RebuildDiscountFundChoices(bool reset){if(discountFundSelector==null)return;string previous=discountFundSelector.SelectedIndex>=0?discountFundSelector.Items[discountFundSelector.SelectedIndex]:"전체";IEnumerable<IndividualRowData> rows=individualDashboard==null?Enumerable.Empty<IndividualRowData>():individualDashboard.Rows;if(discountSiteSelector!=null&&discountSiteSelector.SelectedIndex>=0&&discountSiteSelector.SelectedIndex<discountSiteKeys.Count){string site=discountSiteKeys[discountSiteSelector.SelectedIndex];rows=rows.Where(x=>x.Site==site);}var choices=new List<string>{"전체"};choices.AddRange(rows.Select(x=>x.Fund).Where(x=>!String.IsNullOrWhiteSpace(x)&&x!="분류필요").Distinct().OrderBy(UiFundDisplayOrder));bool old=discountFilterLoading;discountFilterLoading=true;discountFundSelector.SetItems(choices);int index=reset?0:choices.IndexOf(previous);discountFundSelector.SelectedIndex=index>=0?index:0;discountFilterLoading=old;}
        List<IndividualRowData> FilteredDiscountRows(){IEnumerable<IndividualRowData> rows=individualDashboard==null?Enumerable.Empty<IndividualRowData>():individualDashboard.Rows;if(discountSiteSelector!=null&&discountSiteSelector.SelectedIndex>=0&&discountSiteSelector.SelectedIndex<discountSiteKeys.Count){string site=discountSiteKeys[discountSiteSelector.SelectedIndex];rows=rows.Where(x=>x.Site==site);}string fund=discountFundSelector!=null&&discountFundSelector.SelectedIndex>=0?discountFundSelector.Items[discountFundSelector.SelectedIndex]:"전체";if(fund!="전체")rows=rows.Where(x=>x.Fund==fund);return rows.OrderBy(x=>UiFundDisplayOrder(x.Fund)).ThenBy(x=>x.Name).ToList();}
        string DiscountKey(IndividualRowData row){return StablePersonKey(row)+"|"+(row.Fund??"")+"|"+(row.ShortTerm?"1":"0");}
        DiscountEntry SavedDiscount(IndividualRowData row){DiscountEntry entry;if(discountSaved.TryGetValue(DiscountKey(row),out entry))return entry;if(discountSaved.TryGetValue(StablePersonKey(row),out entry))return entry;return new DiscountEntry();}
        DiscountEntry EffectiveDiscount(IndividualRowData row){DiscountEntry entry;if(discountDrafts.TryGetValue(DiscountKey(row),out entry))return entry.Clone();return SavedDiscount(row).Clone();}
        List<DiscountAggregateRow> BuildDiscountAggregates(List<IndividualRowData> rows){var result=new List<DiscountAggregateRow>();foreach(var group in rows.GroupBy(x=>x.Fund=="분류필요"?"기타":x.Fund).OrderBy(x=>UiFundDisplayOrder(x.Key))){var a=new DiscountAggregateRow{Fund=group.Key};foreach(IndividualRowData row in group){DiscountEntry saved=SavedDiscount(row),current=EffectiveDiscount(row);decimal[] institution={row.SummaryHealthEmployer+row.SummaryLongTermEmployer,row.SummaryPensionEmployer,row.SummaryEmploymentEmployer,row.SummaryIndustrialEmployer},old={saved.HealthTotal,saved.PensionTotal,saved.EmploymentTotal,saved.IndustrialTotal},next={current.HealthTotal,current.PensionTotal,current.EmploymentTotal,current.IndustrialTotal};for(int i=0;i<4;i++){a.Billed[i]+=institution[i]+old[i];a.Discount[i]+=next[i];a.After[i]+=institution[i]+old[i]-next[i];}}result.Add(a);}result.Add(DiscountAggregateRow.Total(result));return result;}
        void UpdateDiscountView(bool resetScroll=true){if(discountTable==null)return;List<IndividualRowData> rows=FilteredDiscountRows();discountTable.Rows=rows;if(resetScroll)discountTable.ScrollOffset=0;discountTable.Invalidate();List<DiscountAggregateRow> aggregates=BuildDiscountAggregates(rows);discountBilledTotals.Rows=aggregates;discountAppliedTotals.Rows=aggregates;discountAfterTotals.Rows=aggregates;discountBilledTotals.Invalidate();discountAppliedTotals.Invalidate();discountAfterTotals.Invalidate();discountRangeLabel.Text="선택 사업장 · "+(discountFundSelector!=null&&discountFundSelector.SelectedIndex>0?discountFundSelector.Items[discountFundSelector.SelectedIndex]+" · ":"")+"대상 "+rows.Count+"명";}
        void EditDiscountAmount(IndividualRowData row,string kind,decimal current){using(DiscountAmountDialog dialog=new DiscountAmountDialog(kind,current))if(dialog.ShowDialog(this)==DialogResult.OK){DiscountEntry entry=EffectiveDiscount(row);switch(kind){case "건강":entry.Health=dialog.Amount;break;case "국민":entry.Pension=dialog.Amount;break;case "고용":entry.Employment=dialog.Amount;break;case "산재":entry.Industrial=dialog.Amount;break;}discountDrafts[DiscountKey(row)]=entry;UpdateDiscountView(false);}}
        void RefreshDiscountDashboard(){if(validationResult!=null&&!String.IsNullOrWhiteSpace(validationResult.Text)&&File.Exists(validationResult.Text))LoadResultIntoUi(validationResult.Text);else MessageBox.Show("먼저 파일 등록 화면에서 대사 작업을 실행해 주세요.","감면 적용",MessageBoxButtons.OK,MessageBoxIcon.Information);}
        void ExportDiscountWorkbook(){if(validationResult==null||String.IsNullOrWhiteSpace(validationResult.Text)||!File.Exists(validationResult.Text)){MessageBox.Show("내보낼 감면 적용 자료가 없습니다.","엑셀 내보내기",MessageBoxButtons.OK,MessageBoxIcon.Information);return;}using(SaveFileDialog d=new SaveFileDialog{Filter="Excel 매크로 통합 문서 (*.xlsm)|*.xlsm",FileName="감면적용_내역_"+DateTime.Now.ToString("yyyyMMdd_HHmmss")+".xlsm"})if(d.ShowDialog()==DialogResult.OK){File.Copy(validationResult.Text,d.FileName,true);MessageBox.Show("감면 적용 내역이 포함된 엑셀 파일을 내보냈습니다.","엑셀 내보내기",MessageBoxButtons.OK,MessageBoxIcon.Information);OpenGeneratedFileIfEnabled(d.FileName);}}
        void SaveDiscountChanges(){if(validationResult==null||String.IsNullOrWhiteSpace(validationResult.Text)||!File.Exists(validationResult.Text)){MessageBox.Show("먼저 대사 작업을 실행해 주세요.","감면사항 저장",MessageBoxButtons.OK,MessageBoxIcon.Information);return;}try{int changed=PersistDiscountChangesCore(validationResult.Text);MessageBox.Show(changed+"명의 감면사항을 저장하고 모든 대사 결과 화면에 반영했습니다.","감면사항 저장",MessageBoxButtons.OK,MessageBoxIcon.Information);}catch(Exception ex){MessageBox.Show(ex.Message,"감면 저장 오류",MessageBoxButtons.OK,MessageBoxIcon.Error);}}
        int PersistDiscountChangesCore(string path){int changed=0;foreach(IndividualRowData row in individualDashboard.Rows){string key=DiscountKey(row);DiscountEntry draft;if(!discountDrafts.TryGetValue(key,out draft))continue;DiscountEntry saved=SavedDiscount(row);decimal[] old={saved.HealthTotal,saved.PensionTotal,saved.EmploymentTotal,saved.IndustrialTotal},next={draft.HealthTotal,draft.PensionTotal,draft.EmploymentTotal,draft.IndustrialTotal};row.SummaryHealthEmployer+=old[0]-next[0];row.SummaryPensionEmployer+=old[1]-next[1];row.SummaryEmploymentEmployer+=old[2]-next[2];row.SummaryIndustrialEmployer+=old[3]-next[3];discountSaved[key]=draft.Clone();changed++;}discountStateUsesEmployer=true;discountStateUsesSubtraction=true;discountDrafts.Clear();NormalizeIndividualStatuses();if(individualDashboard.Rows.All(x=>x.HasSummaryBreakdown))RebuildSummaryDashboardFromIndividuals();PersistDiscountWorkbook(path);reconciliationState.Revision++;InitializeAdjustmentFilters();InitializeReviewFilters();InitializeDiscountFilters();RefreshLinkedResultViews();return changed;}
        void LoadDiscountState(ExcelPackage package){discountSaved.Clear();discountDrafts.Clear();discountStateUsesEmployer=true;discountStateUsesSubtraction=true;ExcelWorksheet ws=package.Workbook.Worksheets["UI감면상태"];if(ws==null||ws.Dimension==null)return;List<string> standards=Enumerable.Range(2,Math.Max(0,ws.Dimension.End.Row-1)).Select(r=>ws.Cells[r,8].Text).Where(x=>!String.IsNullOrWhiteSpace(x)).ToList();discountStateUsesEmployer=ws.Dimension.End.Column>=8&&String.Equals(ws.Cells[1,8].Text,"적용기준",StringComparison.Ordinal)&&standards.Any(x=>x.StartsWith("기관부담",StringComparison.Ordinal));discountStateUsesSubtraction=standards.Count==0||standards.Any(x=>String.Equals(x,"기관부담차감",StringComparison.Ordinal));for(int r=2;r<=ws.Dimension.End.Row;r++){string key=ws.Cells[r,1].Text;if(String.IsNullOrWhiteSpace(key))continue;discountSaved[key]=new DiscountEntry{AutoEmployment=UiInt(ws.Cells[r,2].Value)>0,AutoIndustrial=UiInt(ws.Cells[r,3].Value)>0,Health=UiDecimal(ws.Cells[r,4].Value),Pension=UiDecimal(ws.Cells[r,5].Value),Employment=UiDecimal(ws.Cells[r,6].Value),Industrial=UiDecimal(ws.Cells[r,7].Value)};}}
        void MigrateLegacyDiscountStateIfNeeded(){if(discountSaved.Count==0||individualDashboard==null)return;if(discountStateUsesEmployer&&discountStateUsesSubtraction)return;foreach(IndividualRowData row in individualDashboard.Rows){DiscountEntry d=SavedDiscount(row);if(d.Total<=0)continue;if(!discountStateUsesEmployer){row.HealthNotice-=d.HealthTotal;row.PensionNotice-=d.PensionTotal;row.EmploymentNotice-=d.EmploymentTotal;row.IndustrialNotice-=d.IndustrialTotal;row.SummaryHealthPersonal-=d.HealthTotal;row.SummaryPensionPersonal-=d.PensionTotal;row.SummaryEmploymentPersonal-=d.EmploymentTotal;row.SummaryHealthEmployer-=d.HealthTotal;row.SummaryPensionEmployer-=d.PensionTotal;row.SummaryEmploymentEmployer-=d.EmploymentTotal;}else{row.SummaryHealthEmployer-=d.HealthTotal*2;row.SummaryPensionEmployer-=d.PensionTotal*2;row.SummaryEmploymentEmployer-=d.EmploymentTotal*2;row.SummaryIndustrialEmployer-=d.IndustrialTotal*2;}row.HealthDifference=row.HealthNotice-row.HealthPayroll;row.PensionDifference=row.PensionNotice-row.PensionPayroll;row.EmploymentDifference=row.EmploymentNotice-row.EmploymentPayroll;row.IndustrialDifference=row.IndustrialNotice-row.IndustrialPayroll;}discountStateUsesEmployer=true;discountStateUsesSubtraction=true;}
        void WriteDiscountStateSheet(ExcelPackage package){ExcelWorksheet old=package.Workbook.Worksheets["UI감면상태"];if(old!=null)package.Workbook.Worksheets.Delete(old);ExcelWorksheet ws=package.Workbook.Worksheets.Add("UI감면상태");string[] heads={"대상키","고용자동","산재자동","건강기타","국민기타","고용기타","산재기타","적용기준"};for(int i=0;i<heads.Length;i++)ws.Cells[1,i+1].Value=heads[i];int r=2;foreach(var item in discountSaved.OrderBy(x=>x.Key)){ws.Cells[r,1].Value=item.Key;ws.Cells[r,2].Value=item.Value.AutoEmployment?1:0;ws.Cells[r,3].Value=item.Value.AutoIndustrial?1:0;ws.Cells[r,4].Value=item.Value.Health;ws.Cells[r,5].Value=item.Value.Pension;ws.Cells[r,6].Value=item.Value.Employment;ws.Cells[r,7].Value=item.Value.Industrial;ws.Cells[r,8].Value="기관부담차감";r++;}ws.Hidden=eWorkSheetHidden.Hidden;}
        void PersistDiscountWorkbook(string path){using(ExcelPackage package=new ExcelPackage(new FileInfo(path))){ExcelWorksheet individual=package.Workbook.Worksheets["UI개인별데이터"];if(individual!=null&&individual.Dimension!=null)for(int r=2;r<=individual.Dimension.End.Row;r++){string stable=String.Join("|",new[]{individual.Cells[r,1].Text,individual.Cells[r,3].Text,Regex.Replace(individual.Cells[r,4].Text,"[^0-9]","")}),fund=individual.Cells[r,2].Text;bool shortTerm=UiInt(individual.Cells[r,34].Value)>0;IndividualRowData row=individualDashboard.Rows.FirstOrDefault(x=>StablePersonKey(x)==stable&&String.Equals(x.Fund,fund,StringComparison.Ordinal)&&x.ShortTerm==shortTerm);if(row==null)continue;individual.Cells[r,6].Value=row.Status;individual.Cells[r,7].Value=row.HealthNotice;individual.Cells[r,9].Value=row.HealthDifference;individual.Cells[r,10].Value=row.PensionNotice;individual.Cells[r,12].Value=row.PensionDifference;individual.Cells[r,13].Value=row.EmploymentNotice;individual.Cells[r,15].Value=row.EmploymentDifference;individual.Cells[r,16].Value=row.IndustrialNotice;individual.Cells[r,18].Value=row.IndustrialDifference;individual.Cells[r,22].Value=row.SummaryHealthPersonal;individual.Cells[r,23].Value=row.SummaryHealthEmployer;individual.Cells[r,26].Value=row.SummaryPensionPersonal;individual.Cells[r,27].Value=row.SummaryPensionEmployer;individual.Cells[r,28].Value=row.SummaryEmploymentPersonal;individual.Cells[r,29].Value=row.SummaryEmploymentEmployer;individual.Cells[r,31].Value=row.SummaryIndustrialEmployer;}WriteSummaryDashboardSheet(package);WriteDiscountStateSheet(package);package.Save();}}
        void BuildSubmissionScreen(Control page)
        {
            page.Controls.Add(TitleLabel("제출서 생성",8,10,20F));page.Controls.Add(new Label{Text="제출서를 생성할 대상과 보험 구분별 기관부담 신청 금액을 확인해 주세요.",Location=new Point(175,26),AutoSize=true,ForeColor=UiMuted,Font=new Font("맑은 고딕",8F),Tag="ThemeMuted"});var folder=OutputButton("저장 폴더","folder",910,8,128,34,UiBlue,false);folder.Click+=(s,e)=>ChooseSubmitFolder();page.Controls.Add(folder);
            var filter=(RoundedPanel)Card(8,59,1030,82,Color.White);filter.Radius=13;filter.BorderColor=Color.FromArgb(220,226,241);filter.Controls.Add(new Label{Text="고지 년월",Location=new Point(20,14),AutoSize=true,ForeColor=UiText,Font=new Font("맑은 고딕",8.5F,FontStyle.Bold)});submissionPeriodLabel=new Label{Text="-",Location=new Point(20,43),AutoSize=true,ForeColor=Color.FromArgb(33,91,235),Font=new Font("맑은 고딕",15F,FontStyle.Bold)};filter.Controls.Add(submissionPeriodLabel);filter.Controls.Add(new Panel{Location=new Point(285,14),Size=new Size(1,54),BackColor=UiBorder});filter.Controls.Add(new Label{Text="사업장 선택",Location=new Point(322,13),AutoSize=true,ForeColor=UiText,Font=new Font("맑은 고딕",8.5F,FontStyle.Bold)});submissionSiteSelector=new ModernSiteSelector{Location=new Point(318,37),Size=new Size(355,40)};submissionSiteSelector.SelectedIndexChanged+=(s,e)=>{if(!submissionFilterLoading)UpdateSubmissionView();};filter.Controls.Add(submissionSiteSelector);submissionSourceLabel=new Label{Text="대사 결과가 자동 연결됩니다.",Location=new Point(705,48),Size=new Size(292,18),AutoEllipsis=true,TextAlign=ContentAlignment.MiddleRight,ForeColor=UiMuted,Font=new Font("맑은 고딕",7.5F)};filter.Controls.Add(submissionSourceLabel);page.Controls.Add(filter);
            var info=(RoundedPanel)Card(8,152,1030,103,Color.FromArgb(253,254,255));info.Radius=13;info.BorderColor=Color.FromArgb(220,226,241);info.Controls.Add(new Label{Text="제출자 입력",Location=new Point(16,11),AutoSize=true,ForeColor=UiText,Font=new Font("맑은 고딕",9F,FontStyle.Bold)});int[] positions={16,126,274,386,530,630};int[] widths={96,136,100,132,88,150};string[] labels={"수신자기호","기관명","담당자","담당자 전화번호","은행명","계좌번호"};TextBox[] fields=new TextBox[6];for(int i=0;i<labels.Length;i++)fields[i]=AddSubmissionInput(info,labels[i],positions[i],widths[i]);recipientCode=fields[0];institutionName=fields[1];institutionName.TextChanged+=(s,e)=>UpdateApprovalView();managerName=fields[2];phone=fields[3];bankName=fields[4];accountNumber=fields[5];industrialRate=AddSubmissionInput(info,"산재보험 요율",796,94);industrialRate.Text="0.008";info.Controls.Add(new Label{Text="제출 차수",Location=new Point(904,35),AutoSize=true,ForeColor=UiText,Font=new Font("맑은 고딕",7.5F,FontStyle.Bold)});submissionRoundSelector=new ModernSiteSelector{Location=new Point(900,57),Size=new Size(110,35),ShowIcon=false};submissionRoundSelector.SetItems(new[]{"1차","2차","3차","4차"});submissionRoundSelector.SelectedIndexChanged+=(s,e)=>{if(submissionRound!=null&&submissionRoundSelector.SelectedIndex>=0)submissionRound.Text=submissionRoundSelector.Items[submissionRoundSelector.SelectedIndex];};info.Controls.Add(submissionRoundSelector);submissionRound=new TextBox{Visible=false};info.Controls.Add(submissionRound);page.Controls.Add(info);
            var worker=(RoundedPanel)Card(8,269,505,350,Color.FromArgb(250,252,255));worker.Radius=13;worker.BorderColor=Color.FromArgb(213,224,247);worker.Controls.Add(new Label{Text="교특회계 교육공무직 기관부담금 신청서",Location=new Point(20,14),AutoSize=true,ForeColor=UiBlue,Font=new Font("맑은 고딕",11F,FontStyle.Bold)});worker.Controls.Add(new Label{Text="무기계약 · 기간제 구분",Location=new Point(20,39),AutoSize=true,ForeColor=UiMuted,Font=new Font("맑은 고딕",7.5F)});submissionWorkerSummary=new SubmissionSummaryControl{Location=new Point(14,61),Size=new Size(477,225),WorkerMode=true,Accent=UiBlue};worker.Controls.Add(submissionWorkerSummary);var wb=OutputButton("제출서 생성","excel",151,296,205,40,UiGreen);wb.Click+=(s,e)=>CreateSubmission(false);worker.Controls.Add(wb);page.Controls.Add(worker);
            var teacher=(RoundedPanel)Card(530,269,508,350,Color.FromArgb(249,254,251));teacher.Radius=13;teacher.BorderColor=Color.FromArgb(210,234,220);teacher.Controls.Add(new Label{Text="계약제교원 인건비 신청서",Location=new Point(20,14),AutoSize=true,ForeColor=UiGreen,Font=new Font("맑은 고딕",11F,FontStyle.Bold)});teacher.Controls.Add(new Label{Text="보험별 기관부담 신청 금액",Location=new Point(20,39),AutoSize=true,ForeColor=UiMuted,Font=new Font("맑은 고딕",7.5F)});submissionTeacherSummary=new SubmissionSummaryControl{Location=new Point(14,61),Size=new Size(480,225),WorkerMode=false,Accent=UiGreen};teacher.Controls.Add(submissionTeacherSummary);var tb=OutputButton("제출서 생성","excel",152,296,205,40,UiGreen);tb.Click+=(s,e)=>CreateSubmission(true);teacher.Controls.Add(tb);page.Controls.Add(teacher);
            validationResult=new TextBox{Visible=false};submitOutput=new TextBox{Visible=false,Text=Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)};submitStatus=new Label{Visible=false};page.Controls.Add(validationResult);page.Controls.Add(submitOutput);page.Controls.Add(submitStatus);
        }
        TextBox AddSubmissionInput(Control parent,string label,int x,int width){parent.Controls.Add(new Label{Text=label,Location=new Point(x,35),AutoSize=true,ForeColor=UiText,Font=new Font("맑은 고딕",7.5F,FontStyle.Bold)});var shell=(RoundedPanel)Card(x,57,width,35,Color.White);shell.Radius=8;shell.BorderColor=Color.FromArgb(216,224,242);var field=new TextBox{Location=new Point(9,9),Size=new Size(width-18,18),BorderStyle=BorderStyle.None,BackColor=Color.White,ForeColor=Color.FromArgb(31,52,126),Font=new Font("맑은 고딕",8F,FontStyle.Bold)};shell.Controls.Add(field);parent.Controls.Add(shell);return field;}
        void InitializeSubmissionView(){if(submissionSiteSelector==null)return;submissionFilterLoading=true;string previous=submissionSiteSelector.SelectedIndex>=0&&submissionSiteSelector.SelectedIndex<submissionSiteKeys.Count?submissionSiteKeys[submissionSiteSelector.SelectedIndex]:"";submissionSiteKeys.Clear();var sites=individualDashboard==null?new List<string>():individualDashboard.Rows.Select(x=>x.Site).Distinct().OrderBy(x=>x).ToList();submissionSiteKeys.AddRange(sites);submissionSiteSelector.SetItems(sites.Select(FormatSite));int index=sites.IndexOf(previous);submissionSiteSelector.SelectedIndex=index>=0?index:(sites.Count>0?0:-1);submissionFilterLoading=false;submissionPeriodLabel.Text=individualDashboard!=null&&individualDashboard.Year>0?individualDashboard.Year+"년 "+individualDashboard.Month+"월":"-";UpdateSubmissionView();}
        string CurrentSubmissionSite(){return submissionSiteSelector!=null&&submissionSiteSelector.SelectedIndex>=0&&submissionSiteSelector.SelectedIndex<submissionSiteKeys.Count?submissionSiteKeys[submissionSiteSelector.SelectedIndex]:"";}
        bool IsFixedTermWorker(IndividualRowData row){return row!=null&&(row.ShortTerm||Regex.IsMatch((row.Job??"").Replace(" ",""),"기간제|대체|일용|단기|계약직",RegexOptions.IgnoreCase));}
        decimal SubmissionInsuranceAmount(IndividualRowData row,int insurance){switch(insurance){case 0:return row.SummaryHealthEmployer+row.SummaryLongTermEmployer;case 1:return row.SummaryPensionEmployer;case 2:return row.SummaryEmploymentEmployer;default:return row.SummaryIndustrialEmployer;}}
        List<SubmissionSummaryRow> BuildSubmissionSummary(IEnumerable<IndividualRowData> source,bool worker){List<IndividualRowData> rows=source.ToList();var result=new List<SubmissionSummaryRow>();string[] names={"건강보험","국민연금","고용보험","산재보험"};for(int i=0;i<4;i++){var row=new SubmissionSummaryRow{Insurance=names[i]};if(worker){List<IndividualRowData> permanent=rows.Where(x=>!IsFixedTermWorker(x)&&SubmissionInsuranceAmount(x,i)!=0).ToList(),fixedTerm=rows.Where(x=>IsFixedTermWorker(x)&&SubmissionInsuranceAmount(x,i)!=0).ToList();row.PrimaryPeople=permanent.Count;row.PrimaryAmount=permanent.Sum(x=>SubmissionInsuranceAmount(x,i));row.SecondaryPeople=fixedTerm.Count;row.SecondaryAmount=fixedTerm.Sum(x=>SubmissionInsuranceAmount(x,i));}else{List<IndividualRowData> insured=rows.Where(x=>SubmissionInsuranceAmount(x,i)!=0).ToList();row.PrimaryPeople=insured.Count;row.PrimaryAmount=insured.Sum(x=>SubmissionInsuranceAmount(x,i));}result.Add(row);}result.Add(SubmissionSummaryRow.Total(result,worker));return result;}
        void UpdateSubmissionView(){if(submissionWorkerSummary==null||submissionTeacherSummary==null)return;IEnumerable<IndividualRowData> siteRows=individualDashboard==null?Enumerable.Empty<IndividualRowData>():individualDashboard.Rows;string site=CurrentSubmissionSite();if(!String.IsNullOrWhiteSpace(site))siteRows=siteRows.Where(x=>x.Site==site);List<IndividualRowData> all=siteRows.ToList();List<IndividualRowData> workers=all.Where(x=>x.Fund=="교특회계"||x.ShortTerm).ToList(),teachers=all.Where(x=>x.Fund=="계약제교원").ToList();submissionWorkerSummary.Rows=BuildSubmissionSummary(workers,true);submissionTeacherSummary.Rows=BuildSubmissionSummary(teachers,false);submissionWorkerSummary.Invalidate();submissionTeacherSummary.Invalidate();}
        void RefreshSubmissionDashboard(){if(validationResult!=null&&!String.IsNullOrWhiteSpace(validationResult.Text)&&File.Exists(validationResult.Text))LoadResultIntoUi(validationResult.Text);else MessageBox.Show("먼저 파일 등록 화면에서 대사 작업을 실행해 주세요.","제출서 생성",MessageBoxButtons.OK,MessageBoxIcon.Information);}
        void ExportSubmissionSource(){if(validationResult==null||String.IsNullOrWhiteSpace(validationResult.Text)||!File.Exists(validationResult.Text)){MessageBox.Show("내보낼 대사 결과가 없습니다.","엑셀 내보내기",MessageBoxButtons.OK,MessageBoxIcon.Information);return;}using(SaveFileDialog d=new SaveFileDialog{Filter="Excel 매크로 통합 문서 (*.xlsm)|*.xlsm",FileName="제출서_생성자료_"+DateTime.Now.ToString("yyyyMMdd_HHmmss")+".xlsm"})if(d.ShowDialog()==DialogResult.OK){File.Copy(validationResult.Text,d.FileName,true);OpenGeneratedFileIfEnabled(d.FileName);}}
        void BuildApprovalScreen(Control page)
        {
            page.Controls.Add(TitleLabel("내부결재자료 생성",8,10,20F));page.Controls.Add(new Label{Text="학교회계 기관부담금 지출을 위한 내부결재 참고자료를 생성합니다.",Location=new Point(242,26),AutoSize=true,ForeColor=UiMuted,Font=new Font("맑은 고딕",8F)});
            var print=OutputButton("인쇄","print",928,8,110,34,UiBlue,false);print.Click+=(s,e)=>PrintApprovalReport();page.Controls.Add(print);
            var filter=(RoundedPanel)Card(8,59,1030,82,Color.White);filter.Radius=13;filter.BorderColor=Color.FromArgb(220,226,241);filter.Controls.Add(new Label{Text="부과 년월",Location=new Point(20,14),AutoSize=true,ForeColor=UiText,Font=new Font("맑은 고딕",8.5F,FontStyle.Bold)});approvalPeriodLabel=new Label{Text="-",Location=new Point(20,43),AutoSize=true,ForeColor=Color.FromArgb(33,91,235),Font=new Font("맑은 고딕",15F,FontStyle.Bold)};filter.Controls.Add(approvalPeriodLabel);filter.Controls.Add(new Panel{Location=new Point(285,14),Size=new Size(1,54),BackColor=UiBorder});filter.Controls.Add(new Label{Text="사업장관리번호 선택",Location=new Point(322,13),AutoSize=true,ForeColor=UiText,Font=new Font("맑은 고딕",8.5F,FontStyle.Bold)});approvalSiteSelector=new ModernSiteSelector{Location=new Point(318,37),Size=new Size(355,40)};approvalSiteSelector.SelectedIndexChanged+=(s,e)=>{if(!approvalFilterLoading)UpdateApprovalView();};filter.Controls.Add(approvalSiteSelector);approvalDescriptionLabel=new Label{Text="학교회계 대상 자료가 자동 연결됩니다.",Location=new Point(700,47),Size=new Size(290,18),AutoEllipsis=true,TextAlign=ContentAlignment.MiddleRight,ForeColor=UiMuted,Font=new Font("맑은 고딕",7.5F)};filter.Controls.Add(approvalDescriptionLabel);page.Controls.Add(filter);
            page.Controls.Add(new Label{Text="내부결재자료 미리보기",Location=new Point(10,151),AutoSize=true,ForeColor=UiBlue,Font=new Font("맑은 고딕",10F,FontStyle.Bold)});page.Controls.Add(new Label{Text="선택한 사업장의 학교회계 기관부담금 지출내역서입니다.",Location=new Point(165,154),AutoSize=true,ForeColor=UiMuted,Font=new Font("맑은 고딕",7.8F)});
            var excelCard=(RoundedPanel)Card(8,178,505,448,Color.FromArgb(250,254,252));excelCard.Radius=13;excelCard.BorderColor=Color.FromArgb(134,203,163);excelCard.Controls.Add(new Label{Text="학교회계 기관부담금 지출내역서 (내부결재용 - Excel)",Location=new Point(48,15),AutoSize=true,ForeColor=UiText,Font=new Font("맑은 고딕",9.3F,FontStyle.Bold)});var excelBadge=OutputButton("","excel",14,9,28,28,UiGreen);excelBadge.Cursor=Cursors.Default;excelCard.Controls.Add(excelBadge);approvalExcelPreview=new ApprovalPreviewControl{Location=new Point(14,49),Size=new Size(477,338),PdfMode=false};excelCard.Controls.Add(approvalExcelPreview);var excelSave=OutputButton("Excel 저장","excel",171,397,164,38,UiGreen,false);excelSave.Click+=(s,e)=>ExportApprovalExcel();excelCard.Controls.Add(excelSave);page.Controls.Add(excelCard);
            var pdfCard=(RoundedPanel)Card(530,178,508,448,Color.FromArgb(255,252,252));pdfCard.Radius=13;pdfCard.BorderColor=Color.FromArgb(245,125,125);pdfCard.Controls.Add(new Label{Text="학교회계 기관부담금 지출내역서 (내부결재용 - PDF)",Location=new Point(48,15),AutoSize=true,ForeColor=UiText,Font=new Font("맑은 고딕",9.3F,FontStyle.Bold)});var pdfBadge=OutputButton("","pdf",14,9,28,28,UiRed);pdfBadge.Cursor=Cursors.Default;pdfCard.Controls.Add(pdfBadge);approvalPdfPreview=new ApprovalPreviewControl{Location=new Point(14,49),Size=new Size(480,338),PdfMode=true};pdfCard.Controls.Add(approvalPdfPreview);var pdfSave=OutputButton("PDF 저장","pdf",172,397,164,38,UiRed,false);pdfSave.Click+=(s,e)=>ExportApprovalPdf();pdfCard.Controls.Add(pdfSave);page.Controls.Add(pdfCard);
        }
        void InitializeApprovalView(){if(approvalSiteSelector==null)return;approvalFilterLoading=true;string previous=approvalSiteSelector.SelectedIndex>=0&&approvalSiteSelector.SelectedIndex<approvalSiteKeys.Count?approvalSiteKeys[approvalSiteSelector.SelectedIndex]:"";approvalSiteKeys.Clear();var sites=individualDashboard==null?new List<string>():individualDashboard.Rows.Select(x=>x.Site).Distinct().OrderBy(x=>x).ToList();approvalSiteKeys.AddRange(sites);approvalSiteSelector.SetItems(sites.Select(FormatSite));int index=sites.IndexOf(previous);approvalSiteSelector.SelectedIndex=index>=0?index:(sites.Count>0?0:-1);approvalFilterLoading=false;UpdateApprovalView();}
        string CurrentApprovalSite(){return approvalSiteSelector!=null&&approvalSiteSelector.SelectedIndex>=0&&approvalSiteSelector.SelectedIndex<approvalSiteKeys.Count?approvalSiteKeys[approvalSiteSelector.SelectedIndex]:"";}
        ApprovalReportData CurrentApprovalData(){string site=CurrentApprovalSite();List<IndividualRowData> rows=(individualDashboard==null?Enumerable.Empty<IndividualRowData>():individualDashboard.Rows).Where(x=>(String.IsNullOrWhiteSpace(site)||x.Site==site)&&x.Fund=="학교회계").OrderBy(x=>x.Name).ThenBy(x=>x.Birth).ToList();return new ApprovalReportData{Year=individualDashboard==null?0:individualDashboard.Year,Month=individualDashboard==null?0:individualDashboard.Month,Site=site,Institution=String.IsNullOrWhiteSpace(institutionName==null?"":institutionName.Text)?"기관명 미입력":institutionName.Text.Trim(),Rows=rows};}
        void UpdateApprovalView(){if(approvalExcelPreview==null||approvalPdfPreview==null)return;ApprovalReportData data=CurrentApprovalData();approvalPeriodLabel.Text=data.Year>0?data.Year+"년 "+data.Month+"월":"-";approvalDescriptionLabel.Text=data.Rows.Count>0?"학교회계 "+data.Rows.Count+"명 · 기관부담금 "+UiDrawing.Money(data.Total)+"원":"선택 사업장에 학교회계 대상이 없습니다.";approvalExcelPreview.Data=data;approvalPdfPreview.Data=data;approvalExcelPreview.Invalidate();approvalPdfPreview.Invalidate();}
        void RefreshApprovalDashboard(){if(validationResult!=null&&!String.IsNullOrWhiteSpace(validationResult.Text)&&File.Exists(validationResult.Text))LoadResultIntoUi(validationResult.Text);else MessageBox.Show("먼저 파일 등록 화면에서 대사 작업을 실행해 주세요.","내부결재자료 생성",MessageBoxButtons.OK,MessageBoxIcon.Information);}
        void ExportApprovalExcel(){ApprovalReportData data=CurrentApprovalData();if(data.Rows.Count==0){MessageBox.Show("선택 사업장의 학교회계 대상자가 없습니다.","Excel 저장",MessageBoxButtons.OK,MessageBoxIcon.Information);return;}string filename="학교회계_기관부담금_지출내역서_"+data.Year+data.Month.ToString("00")+".xlsx";using(SaveFileDialog d=new SaveFileDialog{Filter="Excel 통합 문서 (*.xlsx)|*.xlsx",FileName=filename,InitialDirectory=submitOutput!=null&&Directory.Exists(submitOutput.Text)?submitOutput.Text:""})if(d.ShowDialog()==DialogResult.OK)try{ApprovalReportGenerator.CreateExcel(d.FileName,data);MessageBox.Show("내부결재용 Excel 자료를 생성했습니다.","Excel 저장 완료",MessageBoxButtons.OK,MessageBoxIcon.Information);OpenGeneratedFileIfEnabled(d.FileName);}catch(Exception ex){MessageBox.Show(ex.Message,"Excel 생성 오류",MessageBoxButtons.OK,MessageBoxIcon.Error);}}
        void ExportApprovalPdf(){ApprovalReportData data=CurrentApprovalData();if(data.Rows.Count==0){MessageBox.Show("선택 사업장의 학교회계 대상자가 없습니다.","PDF 저장",MessageBoxButtons.OK,MessageBoxIcon.Information);return;}string filename="학교회계_기관부담금_지출내역서_"+data.Year+data.Month.ToString("00")+".pdf";using(SaveFileDialog d=new SaveFileDialog{Filter="PDF 문서 (*.pdf)|*.pdf",FileName=filename,InitialDirectory=submitOutput!=null&&Directory.Exists(submitOutput.Text)?submitOutput.Text:""})if(d.ShowDialog()==DialogResult.OK)try{ApprovalReportGenerator.CreatePdf(d.FileName,data);MessageBox.Show("내부결재용 PDF 자료를 생성했습니다.","PDF 저장 완료",MessageBoxButtons.OK,MessageBoxIcon.Information);OpenGeneratedFileIfEnabled(d.FileName);}catch(Exception ex){MessageBox.Show(ex.Message,"PDF 생성 오류",MessageBoxButtons.OK,MessageBoxIcon.Error);}}
        void PrintApprovalReport(){ApprovalReportData data=CurrentApprovalData();if(data.Rows.Count==0){MessageBox.Show("인쇄할 학교회계 대상자가 없습니다.","인쇄",MessageBoxButtons.OK,MessageBoxIcon.Information);return;}try{string path=Path.Combine(Path.GetTempPath(),"학교회계_기관부담금_지출내역서_인쇄_"+Guid.NewGuid().ToString("N")+".pdf");ApprovalReportGenerator.CreatePdf(path,data);Process.Start(new ProcessStartInfo(path){UseShellExecute=true});MessageBox.Show("인쇄용 PDF를 열었습니다. PDF 화면에서 인쇄해 주세요.","인쇄",MessageBoxButtons.OK,MessageBoxIcon.Information);}catch(Exception ex){MessageBox.Show(ex.Message,"인쇄 준비 오류",MessageBoxButtons.OK,MessageBoxIcon.Error);}}
        public void ExportApprovalForTest(string result,string xlsx,string pdf,int siteIndex){validationResult.Text=result;LoadResultIntoUi(result);if(siteIndex>=0&&siteIndex<approvalSiteSelector.Items.Count)approvalSiteSelector.SelectedIndex=siteIndex;ApprovalReportData data=CurrentApprovalData();if(data.Rows.Count==0)throw new InvalidOperationException("학교회계 내부결재자료 대상자가 없습니다.");ApprovalReportGenerator.CreateExcel(xlsx,data);ApprovalReportGenerator.CreatePdf(pdf,data);}
        void AddPage(string key,Control page){page.Visible=false;pages[key]=page;contentHost.Controls.Add(page);}
        void AddNav(FlowLayoutPanel nav,string text,bool section)
        {
            if(section){reconciliationSectionButton=new SidebarNavButton{Text=text,IconKind="results",ShowChevron=true,Expanded=true,Width=165,Height=34,Margin=new Padding(0,6,0,2),Font=new Font("맑은 고딕",9F,FontStyle.Bold)};reconciliationSectionButton.Click+=(s,e)=>ToggleReconciliationNavigation();nav.Controls.Add(reconciliationSectionButton);return;}
            if(text=="제출서 생성"||text=="설정")nav.Controls.Add(new Panel{Width=165,Height=1,BackColor=Color.FromArgb(229,232,243),Margin=new Padding(0,9,0,9)});
            bool resultChild=new[]{"총괄표","개인별 내역","반환 / 추징","확인 필요","감면 적용"}.Contains(text);var button=new SidebarNavButton{Text=text,IconKind=NavigationIcon(text),Indent=resultChild?13:0,Width=165,Height=34,Margin=new Padding(0,2,0,2),Font=new Font("맑은 고딕",8.7F)};button.Click+=(s,e)=>ShowPage(text);nav.Controls.Add(button);navigationButtons[text]=button;if(resultChild)reconciliationNavigationItems.Add(button);
        }
        string NavigationIcon(string text)
        {
            switch(text){case "파일 등록":return "folder";case "총괄표":return "summary";case "개인별 내역":return "person";case "반환 / 추징":return "adjustment";case "확인 필요":return "review";case "감면 적용":return "discount";case "제출서 생성":return "document";case "내부결재자료 생성":return "archive";case "설정":return "settings";default:return "dot";}
        }
        void ToggleReconciliationNavigation()
        {
            if(reconciliationAnimationTimer!=null){reconciliationAnimationTimer.Stop();reconciliationAnimationTimer.Dispose();reconciliationAnimationTimer=null;}
            reconciliationExpanded=!reconciliationExpanded;bool expanding=reconciliationExpanded;if(reconciliationSectionButton!=null){reconciliationSectionButton.Expanded=expanding;reconciliationSectionButton.Invalidate();}
            int frame=0,framesPerItem=6,stagger=2,count=reconciliationNavigationItems.Count;
            foreach(Control item in reconciliationNavigationItems){SidebarNavButton nav=item as SidebarNavButton;if(expanding){item.Visible=false;item.Height=8;item.Margin=new Padding(12,0,0,0);if(nav!=null)nav.VisualOpacity=.15F;}else{item.Visible=true;item.Height=34;item.Margin=new Padding(0,2,0,2);if(nav!=null)nav.VisualOpacity=1F;}}
            Timer animation=new Timer{Interval=18};reconciliationAnimationTimer=animation;
            animation.Tick+=(s,e)=>
            {
                navigationPanel.SuspendLayout();
                for(int i=0;i<count;i++)
                {
                    Control item=reconciliationNavigationItems[i];SidebarNavButton nav=item as SidebarNavButton;int delay=expanding?i*stagger:(count-1-i)*stagger;float raw=Math.Max(0F,Math.Min(1F,(frame-delay)/(float)framesPerItem));float eased=raw*raw*(3F-2F*raw);
                    if(expanding)
                    {
                        if(frame>=delay)item.Visible=true;item.Height=Math.Max(8,(int)Math.Round(34*eased));int vertical=(int)Math.Round(2*eased);item.Margin=new Padding((int)Math.Round(12*(1F-eased)),vertical,0,vertical);if(nav!=null)nav.VisualOpacity=.15F+.85F*eased;
                    }
                    else
                    {
                        float remain=1F-eased;item.Height=Math.Max(8,(int)Math.Round(34*remain));int vertical=(int)Math.Round(2*remain);item.Margin=new Padding((int)Math.Round(12*eased),vertical,0,vertical);if(nav!=null)nav.VisualOpacity=.15F+.85F*remain;if(raw>=1F)item.Visible=false;
                    }
                    if(nav!=null)nav.Invalidate();
                }
                navigationPanel.ResumeLayout(true);frame++;
                if(frame>(count-1)*stagger+framesPerItem)
                {
                    animation.Stop();animation.Dispose();if(Object.ReferenceEquals(reconciliationAnimationTimer,animation))reconciliationAnimationTimer=null;
                    foreach(Control item in reconciliationNavigationItems){SidebarNavButton nav=item as SidebarNavButton;item.Height=34;item.Margin=new Padding(0,2,0,2);item.Visible=expanding;if(nav!=null){nav.VisualOpacity=1F;nav.Invalidate();}}
                }
            };
            animation.Start();
        }
        void AnimateNavigationItem(Control item)
        {
            SidebarNavButton navItem=item as SidebarNavButton;item.Visible=true;item.Margin=new Padding(12,2,0,2);if(navItem!=null){navItem.VisualOpacity=.25F;navItem.Invalidate();}int frame=0;var motion=new Timer{Interval=22};motion.Tick+=(s,e)=>{frame++;int left=Math.Max(0,12-frame*3);item.Margin=new Padding(left,2,0,2);if(navItem!=null){navItem.VisualOpacity=Math.Min(1F,.25F+frame*.19F);navItem.Invalidate();}if(frame>=4){item.Margin=new Padding(0,2,0,2);if(navItem!=null){navItem.VisualOpacity=1F;navItem.Invalidate();}motion.Stop();motion.Dispose();}};motion.Start();
        }
        void ShowPage(string key)
        {
            Control page;if(!pages.TryGetValue(key,out page))return;foreach(Control p in pages.Values)p.Visible=false;page.Visible=true;page.BringToFront();footerToFront();
            foreach(var item in navigationButtons){bool active=item.Key==key;SidebarNavButton navButton=item.Value as SidebarNavButton;if(navButton!=null){navButton.Active=active;navButton.Font=new Font("맑은 고딕",8.7F,active?FontStyle.Bold:FontStyle.Regular);navButton.Invalidate();}else{item.Value.BackColor=active?Color.FromArgb(232,235,255):Color.Transparent;item.Value.ForeColor=active?Color.FromArgb(67,56,202):Color.FromArgb(55,60,82);}}
        }
        void footerToFront(){foreach(Control c in contentHost.Controls)if(c.Dock==DockStyle.Bottom)c.BringToFront();}
        public void SavePreview(string path){Show();Application.DoEvents();using(Bitmap image=new Bitmap(Width,Height)){DrawToBitmap(image,new Rectangle(0,0,Width,Height));image.Save(path,System.Drawing.Imaging.ImageFormat.Png);}Close();}
        public void SaveCollapsedPreview(string path){Show();Application.DoEvents();ToggleReconciliationNavigation();DateTime until=DateTime.Now.AddMilliseconds(700);while(DateTime.Now<until){Application.DoEvents();System.Threading.Thread.Sleep(20);}using(Bitmap image=new Bitmap(Width,Height)){DrawToBitmap(image,new Rectangle(0,0,Width,Height));image.Save(path,System.Drawing.Imaging.ImageFormat.Png);}Close();}
        public void SavePagePreview(string path,string page){Show();Application.DoEvents();ShowPage(page);DateTime until=DateTime.Now.AddMilliseconds(400);while(DateTime.Now<until){Application.DoEvents();System.Threading.Thread.Sleep(20);}using(Bitmap image=new Bitmap(Width,Height)){DrawToBitmap(image,new Rectangle(0,0,Width,Height));image.Save(path,System.Drawing.Imaging.ImageFormat.Png);}Close();}
        public void SaveThemePreview(string path,string page,string theme){Show();Application.DoEvents();ApplyTheme(theme,false);ShowPage(page);DateTime until=DateTime.Now.AddMilliseconds(450);while(DateTime.Now<until){Application.DoEvents();System.Threading.Thread.Sleep(20);}using(Bitmap image=new Bitmap(Width,Height)){DrawToBitmap(image,new Rectangle(0,0,Width,Height));image.Save(path,System.Drawing.Imaging.ImageFormat.Png);}Close();}
        public void SaveIndividualModePreview(string path,string result,string mode,string theme){Show();Application.DoEvents();validationResult.Text=result;LoadResultIntoUi(result);ApplyTheme(theme,false);ShowPage("개인별 내역");SelectIndividualAmountMode(mode);DateTime until=DateTime.Now.AddMilliseconds(650);while(DateTime.Now<until){Application.DoEvents();System.Threading.Thread.Sleep(20);}using(Bitmap image=new Bitmap(Width,Height)){DrawToBitmap(image,new Rectangle(0,0,Width,Height));image.Save(path,System.Drawing.Imaging.ImageFormat.Png);}Close();}
        public void SaveResultThemePreview(string path,string result,string page,string theme){Show();Application.DoEvents();validationResult.Text=result;LoadResultIntoUi(result);ApplyTheme(theme,false);ShowPage(page);DateTime until=DateTime.Now.AddMilliseconds(650);while(DateTime.Now<until){Application.DoEvents();System.Threading.Thread.Sleep(20);}using(Bitmap image=new Bitmap(Width,Height)){DrawToBitmap(image,new Rectangle(0,0,Width,Height));image.Save(path,System.Drawing.Imaging.ImageFormat.Png);}Close();}
        public void SaveNavigationCyclePreview(string path,string result,string theme){Show();Application.DoEvents();ApplyTheme(theme,false);ToggleReconciliationNavigation();PumpUi(65);ToggleReconciliationNavigation();PumpUi(65);ToggleReconciliationNavigation();PumpUi(380);ToggleReconciliationNavigation();PumpUi(380);validationResult.Text=result;LoadResultIntoUi(result);ShowPage("총괄표");PumpUi(400);using(Bitmap image=new Bitmap(Width,Height)){DrawToBitmap(image,new Rectangle(0,0,Width,Height));image.Save(path,System.Drawing.Imaging.ImageFormat.Png);}Close();}
        static void PumpUi(int milliseconds){DateTime until=DateTime.Now.AddMilliseconds(milliseconds);while(DateTime.Now<until){Application.DoEvents();System.Threading.Thread.Sleep(15);}}
        public void SaveResultPagePreview(string path,string result,string page,int siteIndex,int fundIndex=-1,string search="",int scrollOffset=0){page=page=="반환-추징"?"반환 / 추징":page=="개인별-내역"?"개인별 내역":page;Show();Application.DoEvents();validationResult.Text=result;LoadResultIntoUi(result);ShowPage(page);if(page=="개인별 내역"&&individualSiteSelector!=null&&siteIndex>=0&&siteIndex<individualSiteSelector.Items.Count){individualSiteSelector.SelectedIndex=siteIndex;if(individualFundSelector!=null&&fundIndex>=0&&fundIndex<individualFundSelector.Items.Count)individualFundSelector.SelectedIndex=fundIndex;if(individualSearchBox!=null)individualSearchBox.Text=search??"";if(individualTable!=null)individualTable.ScrollOffset=Math.Max(0,scrollOffset);}else if(page=="반환 / 추징"&&adjustmentSiteSelector!=null&&siteIndex>=0&&siteIndex<adjustmentSiteSelector.Items.Count){adjustmentSiteSelector.SelectedIndex=siteIndex;if(adjustmentFundSelector!=null&&fundIndex>=0&&fundIndex<adjustmentFundSelector.Items.Count)adjustmentFundSelector.SelectedIndex=fundIndex;if(!String.IsNullOrWhiteSpace(search))SelectAdjustmentMode(search);if(adjustmentTable!=null)adjustmentTable.ScrollOffset=Math.Max(0,scrollOffset);}else if(summarySiteSelector!=null&&siteIndex>=0&&siteIndex<summarySiteSelector.Items.Count)summarySiteSelector.SelectedIndex=siteIndex;DateTime until=DateTime.Now.AddMilliseconds(600);while(DateTime.Now<until){Application.DoEvents();System.Threading.Thread.Sleep(20);}using(Bitmap image=new Bitmap(Width,Height)){DrawToBitmap(image,new Rectangle(0,0,Width,Height));image.Save(path,System.Drawing.Imaging.ImageFormat.Png);}Close();}
        public void SaveReviewDraftPreview(string path,string result,string fund){Show();Application.DoEvents();validationResult.Text=result;LoadResultIntoUi(result);ShowPage("확인 필요");IndividualRowData row=FilteredReviewRows().FirstOrDefault();if(row==null)throw new InvalidOperationException("확인 필요 미리보기 대상이 없습니다.");reviewSelections.Add(ReviewKey(row));int index=reviewApplyFundSelector.Items.IndexOf(fund);if(index<1)throw new InvalidOperationException("적용 재원을 찾을 수 없습니다.");reviewApplyFundSelector.SelectedIndex=index;reviewSelections.Remove(ReviewKey(row));UpdateReviewSelectionSummary();DateTime until=DateTime.Now.AddMilliseconds(500);while(DateTime.Now<until){Application.DoEvents();System.Threading.Thread.Sleep(20);}using(Bitmap image=new Bitmap(Width,Height)){DrawToBitmap(image,new Rectangle(0,0,Width,Height));image.Save(path,System.Drawing.Imaging.ImageFormat.Png);}Close();}
        public void SaveReviewDetailPreview(string path,string result){Show();Application.DoEvents();validationResult.Text=result;LoadResultIntoUi(result);ShowPage("확인 필요");IndividualRowData row=FilteredReviewRows().FirstOrDefault();if(row==null)throw new InvalidOperationException("확인 필요 상세보기 대상이 없습니다.");ShowReviewDetail(row,PointToScreen(new Point(Width-30,350)));DateTime until=DateTime.Now.AddMilliseconds(500);while(DateTime.Now<until){Application.DoEvents();System.Threading.Thread.Sleep(20);}using(Bitmap image=new Bitmap(Width+reviewBubble.Width-12,Math.Max(Height,reviewBubble.Height)))using(Graphics g=Graphics.FromImage(image)){g.Clear(Color.White);using(Bitmap main=new Bitmap(Width,Height)){DrawToBitmap(main,new Rectangle(0,0,Width,Height));g.DrawImageUnscaled(main,0,0);}using(Bitmap detail=new Bitmap(reviewBubble.Width,reviewBubble.Height)){reviewBubble.DrawToBitmap(detail,new Rectangle(0,0,detail.Width,detail.Height));g.DrawImageUnscaled(detail,Width-12,0);}image.Save(path,System.Drawing.Imaging.ImageFormat.Png);}reviewBubble.Close();Close();}
        public void SavePreviewWithFiles(string path,IEnumerable<string> files){Show();Application.DoEvents();AnalyzeRegisteredFiles(files);DateTime until=DateTime.Now.AddSeconds(2);while(DateTime.Now<until){Application.DoEvents();System.Threading.Thread.Sleep(20);}using(Bitmap image=new Bitmap(Width,Height)){DrawToBitmap(image,new Rectangle(0,0,Width,Height));image.Save(path,System.Drawing.Imaging.ImageFormat.Png);}Close();}
        public void SavePreviewWithFilesTheme(string path,string theme,IEnumerable<string> files){Show();Application.DoEvents();ApplyTheme(theme,false);AnalyzeRegisteredFiles(files);DateTime until=DateTime.Now.AddSeconds(2);while(DateTime.Now<until){Application.DoEvents();System.Threading.Thread.Sleep(20);}using(Bitmap image=new Bitmap(Width,Height)){DrawToBitmap(image,new Rectangle(0,0,Width,Height));image.Save(path,System.Drawing.Imaging.ImageFormat.Png);}Close();}
        public void RunRegisteredForTest(string target,IEnumerable<string> files){AnalyzeRegisteredFiles(files);var temp=new List<string>();try{var input=new InputSet{PayrollPackage=PrepareRegisteredInput("급여대장 통합",temp),ShortTermPayroll=PrepareRegisteredInput("단기기간제 근로자",temp),HealthGov=PrepareRegisteredInput("건강보험",temp),Pension=PrepareRegisteredInput("국민연금",temp),Employment=PrepareRegisteredInput("고용보험",temp),Industrial=PrepareRegisteredInput("산재보험",temp)};Processor.Run(input,target);}finally{foreach(string p in temp)try{File.Delete(p);}catch{}}}
        void BuildCalculationTab(Control page)
        {
            page.Controls.Add(new Label{Text="사회보험 재원별 대사 보조 도우미",Font=new Font("맑은 고딕",18F,FontStyle.Bold),AutoSize=true,Location=new Point(24,18)});
            page.Controls.Add(new Label{Text="급여대장 공제액과 4대보험 EDI 부과액을 비교해 검증 결과를 만듭니다.",AutoSize=true,ForeColor=Color.DimGray,Location=new Point(27,57)});
            var payroll=new GroupBox{Text="급여대장 및 단기근로자",Location=new Point(24,88),Size=new Size(445,270)};var insurance=new GroupBox{Text="사대보험",Location=new Point(487,88),Size=new Size(445,270)};page.Controls.Add(payroll);page.Controls.Add(insurance);
            AddFileRow(payroll,"급여대장 통합","급여대장 통합",43);AddFileRow(payroll,"단기기간제 근로자","단기기간제 신청",113);
            payroll.Controls.Add(new Label{Text="급여대장: 단일 Excel 또는 여러 급여대장을 묶은 ZIP",AutoSize=true,ForeColor=Color.DimGray,Location=new Point(16,172)});
            payroll.Controls.Add(new Label{Text="공무원·계약제교원·교특·학회는 파일/시트 내용으로 자동 분류",AutoSize=true,ForeColor=Color.FromArgb(0,102,204),Location=new Point(16,194)});
            payroll.Controls.Add(new Label{Text="단기기간제: 신청서 XLSX 또는 ZIP",AutoSize=true,ForeColor=Color.DimGray,Location=new Point(16,216)});
            AddFileRow(insurance,"건강보험","건강보험",28);AddFileRow(insurance,"국민연금","국민연금",78);AddFileRow(insurance,"고용보험","고용보험",128);AddFileRow(insurance,"산재보험","산재보험",178);
            page.Controls.Add(new Label{Text="사업장이 여러 개일 경우 보험별로 ZIP 파일 압축 후 업로드",AutoSize=true,ForeColor=Color.FromArgb(0,102,204),Font=new Font("맑은 고딕",9F,FontStyle.Bold),Location=new Point(27,377)});
            page.Controls.Add(new Label{Text="※ 교육공무직 제출서는 교특 대상과 단기기간제(학회·일용근로) 대상을 함께 생성합니다.",AutoSize=true,ForeColor=Color.FromArgb(180,90,30),Location=new Point(27,397)});
            page.Controls.Add(new Label{Text="저장 위치",AutoSize=true,Location=new Point(27,443)});
            output=new TextBox{Location=new Point(105,438),Width=704,ReadOnly=true,Text=Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)};page.Controls.Add(output);
            Button choose=new Button{Text="선택",Location=new Point(820,436),Size=new Size(85,29)};choose.Click+=(s,e)=>ChooseFolder();page.Controls.Add(choose);
            Button run=new Button{Text="검증 결과 만들기",Location=new Point(105,488),Size=new Size(335,42),BackColor=Color.FromArgb(79,129,189),ForeColor=Color.White,FlatStyle=FlatStyle.Flat};run.Click+=(s,e)=>Run();page.Controls.Add(run);
            Button open=new Button{Text="저장 위치 열기",Location=new Point(454,488),Size=new Size(335,42)};open.Click+=(s,e)=>OpenFolder();page.Controls.Add(open);
            status=new Label{Text="파일을 선택해 주세요.",AutoSize=true,Location=new Point(105,542),ForeColor=Color.DimGray};page.Controls.Add(status);
        }
        static Icon LoadAppIcon()
        {
            using(Stream s=Assembly.GetExecutingAssembly().GetManifestResourceStream("InsurancePayrollValidator.AppIcon.ico"))return s==null?null:(Icon)new Icon(s).Clone();
        }
        static Image LoadReferenceIcon(){return LoadReferenceIcon(1F);}
        static Image LoadReferenceIcon(float opacity){using(Stream s=Assembly.GetExecutingAssembly().GetManifestResourceStream("InsurancePayrollValidator.ReferenceIcon.png")){if(s==null)return null;using(Bitmap source=new Bitmap(s)){var output=new Bitmap(source.Width,source.Height,System.Drawing.Imaging.PixelFormat.Format32bppArgb);using(Graphics g=Graphics.FromImage(output))using(var attributes=new System.Drawing.Imaging.ImageAttributes()){var matrix=new System.Drawing.Imaging.ColorMatrix();matrix.Matrix33=Math.Max(0F,Math.Min(1F,opacity));attributes.SetColorMatrix(matrix,System.Drawing.Imaging.ColorMatrixFlag.Default,System.Drawing.Imaging.ColorAdjustType.Bitmap);g.DrawImage(source,new Rectangle(0,0,output.Width,output.Height),0,0,source.Width,source.Height,GraphicsUnit.Pixel,attributes);}return output;}}}
        void BuildSubmissionTab(Control page)
        {
            page.Controls.Add(new Label{Text="제출 서식 생성",Font=new Font("맑은 고딕",18F,FontStyle.Bold),AutoSize=true,Location=new Point(24,18)});
            page.Controls.Add(new Label{Text="기관정보와 검증 결과를 한 번 입력하면 두 제출 서식에 함께 반영됩니다.",AutoSize=true,ForeColor=Color.DimGray,Location=new Point(27,57)});
            var org=new GroupBox{Text="기관 및 담당자 정보",Location=new Point(24,84),Size=new Size(908,122)};page.Controls.Add(org);
            AddInfoField(org,"수신자기호",18,28,125,out recipientCode);AddInfoField(org,"기관명",270,28,130,out institutionName);AddInfoField(org,"담당자명",505,28,100,out managerName);org.Controls.Add(new Label{Text="차수",AutoSize=true,Location=new Point(707,34)});submissionRound=new TextBox{Location=new Point(752,28),Width=125};org.Controls.Add(submissionRound);
            AddInfoField(org,"전화번호",18,73,145,out phone);AddInfoField(org,"계좌번호",315,73,95,out bankName);org.Controls.Add(new Label{Text=",",AutoSize=true,Location=new Point(500,80)});accountNumber=new TextBox{Location=new Point(518,72),Width=190};org.Controls.Add(accountNumber);org.Controls.Add(new Label{Text="산재요율",AutoSize=true,Location=new Point(718,79)});industrialRate=new TextBox{Location=new Point(782,72),Width=95,Text="0.008"};org.Controls.Add(industrialRate);
            var box=new GroupBox{Text="검증 결과 선택",Location=new Point(24,216),Size=new Size(908,75)};page.Controls.Add(box);
            validationResult=new TextBox{Location=new Point(22,31),Width=744,ReadOnly=true};box.Controls.Add(validationResult);
            Button select=new Button{Text="파일 선택",Location=new Point(778,28),Size=new Size(105,30)};select.Click+=(s,e)=>ChooseValidationResult();box.Controls.Add(select);
            page.Controls.Add(new Label{Text="저장 위치",AutoSize=true,Location=new Point(27,312)});submitOutput=new TextBox{Location=new Point(105,307),Width=704,ReadOnly=true,Text=output.Text};page.Controls.Add(submitOutput);Button folder=new Button{Text="선택",Location=new Point(820,305),Size=new Size(85,29)};folder.Click+=(s,e)=>ChooseSubmitFolder();page.Controls.Add(folder);
            var worker=new Button{Text="교육공무직원 기관부담금 신청",Location=new Point(105,350),Size=new Size(350,48),BackColor=Color.FromArgb(112,173,71),ForeColor=Color.White,FlatStyle=FlatStyle.Flat};worker.Click+=(s,e)=>CreateSubmission(false);page.Controls.Add(worker);
            var teacher=new Button{Text="계약제교원 인건비(사대보험) 신청",Location=new Point(477,350),Size=new Size(350,48),BackColor=Color.FromArgb(79,129,189),ForeColor=Color.White,FlatStyle=FlatStyle.Flat};teacher.Click+=(s,e)=>CreateSubmission(true);page.Controls.Add(teacher);
            page.Controls.Add(new Label{Text="교특 + 단기기간제 대상 생성",AutoSize=true,ForeColor=Color.DimGray,Location=new Point(205,403)});page.Controls.Add(new Label{Text="계약제교원 대상만 생성",AutoSize=true,ForeColor=Color.DimGray,Location=new Point(583,403)});
            var both=new Button{Text="두 제출서 모두 만들기",Location=new Point(202,433),Size=new Size(255,38)};both.Click+=(s,e)=>{CreateSubmission(false,false);CreateSubmission(true,false);MessageBox.Show("두 제출서를 만들었습니다.","완료",MessageBoxButtons.OK,MessageBoxIcon.Information);};page.Controls.Add(both);
            var open=new Button{Text="저장 위치 열기",Location=new Point(475,433),Size=new Size(255,38)};open.Click+=(s,e)=>OpenFolder();page.Controls.Add(open);
            submitStatus=new Label{Text="검증 결과 파일을 선택해 주세요.",AutoSize=true,Location=new Point(105,485),ForeColor=Color.DimGray};page.Controls.Add(submitStatus);
            page.Controls.Add(new Label{Text="※ 사업장관리번호별로 나누지 않고 제출서 종류마다 전체 사업장을 합친 파일 하나를 만듭니다.",AutoSize=true,ForeColor=Color.FromArgb(180,90,30),Location=new Point(105,522)});
        }
        void AddInfoField(Control parent,string label,int x,int y,int width,out TextBox field){parent.Controls.Add(new Label{Text=label,AutoSize=true,Location=new Point(x,y+6)});field=new TextBox{Location=new Point(x+82,y),Width=width};parent.Controls.Add(field);}
        void LoadSavedSubmissionInfo()
        {
            Dictionary<string,string> s=AppSettings.Load();recipientCode.Text=GetSetting(s,"RecipientCode");institutionName.Text=GetSetting(s,"InstitutionName");managerName.Text=GetSetting(s,"ManagerName");phone.Text=GetSetting(s,"Phone");bankName.Text=GetSetting(s,"BankName");accountNumber.Text=GetSetting(s,"AccountNumber");submissionRound.Text=GetSetting(s,"Round");string rate=GetSetting(s,"IndustrialRate"),folder=GetSetting(s,"OutputFolder"),theme=GetSetting(s,"Theme");openResultAfterSave=GetSetting(s,"OpenResultAfterSave")=="1";automaticUpdateCheck=GetSetting(s,"AutomaticUpdateCheck")!="0";if(!String.IsNullOrWhiteSpace(theme))UiTheme.Set(theme);industrialRate.Text=String.IsNullOrWhiteSpace(rate)?"0.008":rate;if(!String.IsNullOrWhiteSpace(folder)&&Directory.Exists(folder)){submitOutput.Text=folder;output.Text=folder;}if(submissionRoundSelector!=null){int index=submissionRoundSelector.Items.FindIndex(x=>x==submissionRound.Text||x.StartsWith(submissionRound.Text));submissionRoundSelector.SelectedIndex=index>=0?index:0;}
        }
        void SaveSubmissionInfo(){if(recipientCode==null)return;AppSettings.Save(new Dictionary<string,string>{{"RecipientCode",recipientCode.Text},{"InstitutionName",institutionName.Text},{"ManagerName",managerName.Text},{"Phone",phone.Text},{"BankName",bankName.Text},{"AccountNumber",accountNumber.Text},{"Round",submissionRound.Text},{"IndustrialRate",industrialRate.Text},{"OutputFolder",submitOutput==null?"":submitOutput.Text},{"Theme",UiTheme.Name},{"OpenResultAfterSave",openResultAfterSave?"1":"0"},{"AutomaticUpdateCheck",automaticUpdateCheck?"1":"0"}});}
        static string GetSetting(Dictionary<string,string> values,string key){string value;return values.TryGetValue(key,out value)?value:"";}
        void AddFileRow(Control parent,string key,string label,int y)
        {
            parent.Controls.Add(new Label{Text=label,AutoSize=true,Location=new Point(16,y+7)});
            TextBox t=new TextBox{Location=new Point(165,y),Width=184,ReadOnly=true};boxes[key]=t;parent.Controls.Add(t);
            Button b=new Button{Text="선택",Location=new Point(357,y-1),Size=new Size(70,29)};b.Click+=(s,e)=>ChooseFile(t);parent.Controls.Add(b);
        }
        void ChooseFile(TextBox t){using(OpenFileDialog d=new OpenFileDialog{Filter="Excel 또는 ZIP 파일 (*.xlsx;*.xlsm;*.zip)|*.xlsx;*.xlsm;*.zip|Excel 파일 (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|ZIP 파일 (*.zip)|*.zip|모든 파일 (*.*)|*.*"})if(d.ShowDialog()==DialogResult.OK){t.Text=d.FileName;t.SelectionStart=t.TextLength;t.ScrollToCaret();}}
        void ChooseValidationResult(){Safe202(OpenWorkspace202);}
        void ChooseFolder(){using(FolderBrowserDialog d=new FolderBrowserDialog{SelectedPath=output.Text})if(d.ShowDialog()==DialogResult.OK){output.Text=d.SelectedPath;if(submitOutput!=null)submitOutput.Text=d.SelectedPath;}}
        void ChooseSubmitFolder(){using(FolderBrowserDialog d=new FolderBrowserDialog{SelectedPath=submitOutput.Text})if(d.ShowDialog()==DialogResult.OK){submitOutput.Text=d.SelectedPath;output.Text=d.SelectedPath;}}
        void OpenFolder(){if(output!=null&&Directory.Exists(output.Text))Process.Start("explorer.exe",output.Text);}
        void OpenGeneratedFileIfEnabled(string path){if(!openResultAfterSave||String.IsNullOrWhiteSpace(path)||!File.Exists(path))return;try{Process.Start(new ProcessStartInfo(path){UseShellExecute=true});}catch(Exception ex){MessageBox.Show("파일은 정상적으로 저장했지만 자동으로 열지 못했습니다.\r\n"+ex.Message,"파일 자동 열기",MessageBoxButtons.OK,MessageBoxIcon.Warning);}}
        void CleanupTemporaryResult(){if(String.IsNullOrWhiteSpace(temporaryResultPath))return;try{if(File.Exists(temporaryResultPath))File.Delete(temporaryResultPath);}catch{}temporaryResultPath="";}
        string NewTemporaryResultPath(){CleanupTemporaryResult();return Path.Combine(Path.GetTempPath(),"InsurancePayrollValidator_Result_"+Guid.NewGuid().ToString("N")+".xlsm");}
        void Run()
        {
            if(runInProgress)return;
            var prepared=new List<string>();InputSet i;try{i=new InputSet{PayrollPackage=PrepareRegisteredInput("급여대장 통합",prepared),ShortTermPayroll=PrepareRegisteredInput("단기기간제 근로자",prepared),HealthGov=PrepareRegisteredInput("건강보험",prepared),HealthOther="",Pension=PrepareRegisteredInput("국민연금",prepared),Employment=PrepareRegisteredInput("고용보험",prepared),Industrial=PrepareRegisteredInput("산재보험",prepared)};}catch(Exception ex){foreach(string p in prepared)try{File.Delete(p);}catch{}readinessLabel.Text="!  파일 준비 오류";readinessLabel.ForeColor=UiRed;readinessDetail.Text=ex.Message;MessageBox.Show("등록 파일을 준비하지 못했습니다.\r\n\r\n"+ex.Message,"대사 준비 오류",MessageBoxButtons.OK,MessageBoxIcon.Error);return;}
            if(String.IsNullOrWhiteSpace(i.PayrollPackage)&&String.IsNullOrWhiteSpace(i.ShortTermPayroll)){MessageBox.Show("급여대장 또는 단기기간제 신청서를 하나 이상 선택해 주세요.");return;}
            foreach(var x in i.All())if(!String.IsNullOrWhiteSpace(x.Item2)&&!File.Exists(x.Item2)){MessageBox.Show(x.Item1+" 파일을 찾을 수 없습니다.");return;}
            try{runInProgress=true;runButton.Enabled=true;runButton.BackColor=UiPurple;runButton.ForeColor=Color.White;runButton.Text="자료 분석 중...";readinessDetail.Text="급여대장과 사회보험 부과자료를 대사하고 있습니다.";Application.DoEvents();string path=NewTemporaryResultPath();temporaryResultPath=path;Processor.Run(i,path);validationResult.Text=path;LoadResultIntoUi(path);readinessLabel.Text="✓  대사 완료";readinessDetail.Text="대사 결과를 화면에 반영했습니다.";runButton.Text="✓  완료";runButton.ForeColor=Color.White;ShowPage("총괄표");MessageBox.Show("대사 작업이 완료되었습니다.","대사 완료",MessageBoxButtons.OK,MessageBoxIcon.Information);}catch(Exception ex){CleanupTemporaryResult();readinessLabel.Text="!  대사 오류";readinessLabel.ForeColor=UiRed;readinessDetail.Text=ex.Message;runButton.Text="▷  다시 시도";runButton.Enabled=true;runButton.ForeColor=Color.White;MessageBox.Show(ex.Message,"검증 오류",MessageBoxButtons.OK,MessageBoxIcon.Error);}finally{runInProgress=false;foreach(string p in prepared)try{File.Delete(p);}catch{}}
        }
        string PrepareRegisteredInput(string kind,List<string> temporary)
        {
            List<string> files=registeredFiles[kind];if(files.Count==0)return "";if(files.Count==1)return files[0];string zip=Path.Combine(Path.GetTempPath(),"SocialInsuranceUi_"+kind.Replace(" ","")+"_"+Guid.NewGuid().ToString("N")+".zip");var names=new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            Func<string,string> uniqueName=name=>{string unique=Path.GetFileName(name);int n=1;while(!names.Add(unique))unique=Path.GetFileNameWithoutExtension(name)+"_"+(n++)+Path.GetExtension(name);return unique;};
            using(ZipArchive archive=ZipFile.Open(zip,ZipArchiveMode.Create))foreach(string file in files){if(Path.GetExtension(file).Equals(".zip",StringComparison.OrdinalIgnoreCase)){using(ZipArchive source=ZipFile.OpenRead(file))foreach(ZipArchiveEntry item in source.Entries){string ext=Path.GetExtension(item.Name);if(String.IsNullOrWhiteSpace(item.Name)||!(ext.Equals(".xlsx",StringComparison.OrdinalIgnoreCase)||ext.Equals(".xlsm",StringComparison.OrdinalIgnoreCase)))continue;ZipArchiveEntry target=archive.CreateEntry(uniqueName(item.Name),System.IO.Compression.CompressionLevel.Optimal);using(Stream input=item.Open())using(Stream output=target.Open())input.CopyTo(output);}}else archive.CreateEntryFromFile(file,uniqueName(Path.GetFileName(file)),System.IO.Compression.CompressionLevel.Optimal);}temporary.Add(zip);return zip;
        }
        void LoadResultIntoUi(string path)
        {
            try{using(ExcelPackage package=new ExcelPackage(new FileInfo(path))){LoadSummaryDashboard(package);LoadAdjustmentSelections(package);LoadReviewState(package);LoadDiscountState(package);LoadIndividualDashboard(package);MigrateLegacyDiscountStateIfNeeded();InitializeDiscountFilters();InitializeSubmissionView();InitializeApprovalView();NormalizeIndividualStatuses();if(individualDashboard.Rows.All(x=>x.HasSummaryBreakdown))RebuildSummaryDashboardFromIndividuals();RefreshLinkedResultViews();}}catch(Exception ex){MessageBox.Show("결과 파일은 생성했지만 화면 데이터를 읽지 못했습니다.\r\n"+ex.Message,"결과 불러오기",MessageBoxButtons.OK,MessageBoxIcon.Warning);}
        }
        void NormalizeIndividualStatuses(){if(individualDashboard==null)return;foreach(IndividualRowData row in individualDashboard.Rows){row.HealthDifference=row.HealthNotice-row.HealthPayroll;row.PensionDifference=row.PensionNotice-row.PensionPayroll;row.EmploymentDifference=row.EmploymentNotice-row.EmploymentPayroll;row.IndustrialDifference=row.IndustrialNotice-row.IndustrialPayroll;row.SummaryHealthDifference=row.HealthDifference-row.SummaryLongTermDifference;if(IsReviewCompleted(row)){row.Status="정상";continue;}bool collection=row.HealthDifference>.5m||row.PensionDifference>.5m||row.EmploymentDifference>.5m||row.IndustrialDifference>.5m,refund=row.HealthDifference<-.5m||row.PensionDifference<-.5m||row.EmploymentDifference<-.5m||row.IndustrialDifference<-.5m;if(collection&&refund)row.Status="확인 필요";else if(refund)row.Status="환급 필요";else if(collection)row.Status="추징 필요";else if(row.Fund=="분류필요")row.Status="확인 필요";else row.Status="정상";}}
        void LoadSummaryDashboard(ExcelPackage package)
        {
            ExcelWorksheet data=package.Workbook.Worksheets["UI총괄데이터"];var model=new SummaryDashboardModel();ExcelWorksheet info=package.Workbook.Worksheets["제출정보"];if(info!=null){model.Year=UiInt(info.Cells[1,2].Value);model.Month=UiInt(info.Cells[2,2].Value);}ExcelWorksheet rec=package.Workbook.Worksheets["자료인식"];if(rec!=null&&rec.Dimension!=null)model.FileCount=Enumerable.Range(2,Math.Max(0,rec.Dimension.End.Row-1)).Select(r=>rec.Cells[r,2].Text).Where(x=>!String.IsNullOrWhiteSpace(x)).Distinct(StringComparer.OrdinalIgnoreCase).Count();
            if(data!=null&&data.Dimension!=null)for(int r=2;r<=data.Dimension.End.Row;r++)
            {
                string site=data.Cells[r,1].Text,fund=data.Cells[r,2].Text;if(String.IsNullOrWhiteSpace(site)||String.IsNullOrWhiteSpace(fund))continue;SummarySiteData siteData;if(!model.Sites.TryGetValue(site,out siteData)){siteData=new SummarySiteData{Site=site};model.Sites[site]=siteData;}var row=new SummaryFundData{Fund=fund,People=UiInt(data.Cells[r,3].Value),HealthPersonal=UiDecimal(data.Cells[r,4].Value),HealthEmployer=UiDecimal(data.Cells[r,5].Value),LongTermPersonal=UiDecimal(data.Cells[r,6].Value),LongTermEmployer=UiDecimal(data.Cells[r,7].Value),PensionPersonal=UiDecimal(data.Cells[r,8].Value),PensionEmployer=UiDecimal(data.Cells[r,9].Value),EmploymentPersonal=UiDecimal(data.Cells[r,10].Value),EmploymentEmployer=UiDecimal(data.Cells[r,11].Value),IndustrialPersonal=UiDecimal(data.Cells[r,12].Value),IndustrialEmployer=UiDecimal(data.Cells[r,13].Value),InstitutionTotal=UiDecimal(data.Cells[r,14].Value),ReviewCount=UiInt(data.Cells[r,15].Value),ShortTermCount=UiInt(data.Cells[r,16].Value),HealthDifference=UiDecimal(data.Cells[r,19].Value),LongTermDifference=UiDecimal(data.Cells[r,20].Value),PensionDifference=UiDecimal(data.Cells[r,21].Value),EmploymentDifference=UiDecimal(data.Cells[r,22].Value),IndustrialDifference=UiDecimal(data.Cells[r,23].Value)};siteData.Rows.Add(row);if(model.Year==0)model.Year=UiInt(data.Cells[r,17].Value);if(model.Month==0)model.Month=UiInt(data.Cells[r,18].Value);
            }
            summaryDashboard=model;reconciliationState.Summary=model;summaryComboLoading=true;summarySiteKeys.Clear();var displaySites=new List<string>();foreach(string site in model.Sites.Keys.OrderBy(x=>x)){summarySiteKeys.Add(site);displaySites.Add(FormatSite(site));}summarySiteSelector.SetItems(displaySites);if(displaySites.Count>0)summarySiteSelector.SelectedIndex=0;summaryComboLoading=false;UpdateSummaryStatistics();UpdateSummaryForSelectedSite();
        }
        int UiInt(object value){int n;Int32.TryParse(Convert.ToString(value,CultureInfo.InvariantCulture),NumberStyles.Any,CultureInfo.InvariantCulture,out n);return n;}
        decimal UiDecimal(object value){decimal n;Decimal.TryParse(Convert.ToString(value,CultureInfo.InvariantCulture),NumberStyles.Any,CultureInfo.InvariantCulture,out n);return n;}
        void UpdateSummaryStatistics()
        {
            if(summaryStatCards==null)return;int sites=summaryDashboard==null?0:summaryDashboard.Sites.Count,files=summaryDashboard==null?0:summaryDashboard.FileCount;bool hasIndividuals=individualDashboard!=null;int workers=hasIndividuals?individualDashboard.Rows.Count:summaryDashboard==null?0:summaryDashboard.Sites.Values.Sum(x=>x.Rows.Sum(r=>r.People)),shortTerm=hasIndividuals?individualDashboard.Rows.Count(x=>x.ShortTerm):summaryDashboard==null?0:summaryDashboard.Sites.Values.Sum(x=>x.Rows.Sum(r=>r.ShortTermCount)),review=hasIndividuals?individualDashboard.Rows.Count(IsPendingReviewRow):summaryDashboard==null?0:summaryDashboard.Sites.Values.Sum(x=>x.Rows.Sum(r=>r.ReviewCount));string[] values={sites+"개",files+"개",workers+"명",shortTerm+"명",review+"건"};for(int i=0;i<summaryStatCards.Length;i++){summaryStatCards[i].Value=values[i];summaryStatCards[i].Invalidate();}summaryPeriodLabel.Text=summaryDashboard!=null&&summaryDashboard.Year>0?summaryDashboard.Year+"년 "+summaryDashboard.Month+"월":"-";
        }
        void UpdateSummaryForSelectedSite()
        {
            if(summaryDashboard==null||summarySiteSelector==null||summarySiteSelector.SelectedIndex<0||summarySiteSelector.SelectedIndex>=summarySiteKeys.Count){if(summaryTable!=null){summaryTable.Rows=new List<SummaryFundData>();summaryTable.Invalidate();}return;}string siteKey=summarySiteKeys[summarySiteSelector.SelectedIndex];SummarySiteData site;if(!summaryDashboard.Sites.TryGetValue(siteKey,out site))return;summaryPremiumTotals.SiteData=site;summaryPremiumTotals.Invalidate();summaryTable.Rows=site.Rows;summaryTable.Invalidate();
        }
        void LoadIndividualDashboard(ExcelPackage package)
        {
            var model=new IndividualDashboardModel();ExcelWorksheet data=package.Workbook.Worksheets["UI개인별데이터"],info=package.Workbook.Worksheets["제출정보"];if(info!=null){model.Year=UiInt(info.Cells[1,2].Value);model.Month=UiInt(info.Cells[2,2].Value);}if(data!=null&&data.Dimension!=null)for(int r=2;r<=data.Dimension.End.Row;r++){string site=data.Cells[r,1].Text,name=data.Cells[r,3].Text;if(String.IsNullOrWhiteSpace(site)||String.IsNullOrWhiteSpace(name))continue;model.Rows.Add(new IndividualRowData{Site=site,Fund=data.Cells[r,2].Text,Name=name,Birth=data.Cells[r,4].Text,Job=data.Cells[r,5].Text,Status=data.Cells[r,6].Text,HealthNotice=UiDecimal(data.Cells[r,7].Value),HealthPayroll=UiDecimal(data.Cells[r,8].Value),HealthDifference=UiDecimal(data.Cells[r,9].Value),PensionNotice=UiDecimal(data.Cells[r,10].Value),PensionPayroll=UiDecimal(data.Cells[r,11].Value),PensionDifference=UiDecimal(data.Cells[r,12].Value),EmploymentNotice=UiDecimal(data.Cells[r,13].Value),EmploymentPayroll=UiDecimal(data.Cells[r,14].Value),EmploymentDifference=UiDecimal(data.Cells[r,15].Value),IndustrialNotice=UiDecimal(data.Cells[r,16].Value),IndustrialPayroll=UiDecimal(data.Cells[r,17].Value),IndustrialDifference=UiDecimal(data.Cells[r,18].Value),ReviewReason=data.Cells[r,21].Text,SummaryHealthPersonal=UiDecimal(data.Cells[r,22].Value),SummaryHealthEmployer=UiDecimal(data.Cells[r,23].Value),SummaryLongTermPersonal=UiDecimal(data.Cells[r,24].Value),SummaryLongTermEmployer=UiDecimal(data.Cells[r,25].Value),SummaryPensionPersonal=UiDecimal(data.Cells[r,26].Value),SummaryPensionEmployer=UiDecimal(data.Cells[r,27].Value),SummaryEmploymentPersonal=UiDecimal(data.Cells[r,28].Value),SummaryEmploymentEmployer=UiDecimal(data.Cells[r,29].Value),SummaryIndustrialPersonal=UiDecimal(data.Cells[r,30].Value),SummaryIndustrialEmployer=UiDecimal(data.Cells[r,31].Value),SummaryHealthDifference=UiDecimal(data.Cells[r,32].Value),SummaryLongTermDifference=UiDecimal(data.Cells[r,33].Value),ShortTerm=UiInt(data.Cells[r,34].Value)>0,HasSummaryBreakdown=UiInt(data.Cells[r,35].Value)>0});if(model.Year==0)model.Year=UiInt(data.Cells[r,19].Value);if(model.Month==0)model.Month=UiInt(data.Cells[r,20].Value);}individualDashboard=model;reconciliationState.Individuals=model;reconciliationState.Revision++;individualFilterLoading=true;individualSiteKeys.Clear();var sites=model.Rows.Select(x=>x.Site).Distinct().OrderBy(x=>x).ToList();foreach(string site in sites)individualSiteKeys.Add(site);individualSiteSelector.SetItems(sites.Select(FormatSite));if(sites.Count>0)individualSiteSelector.SelectedIndex=0;RebuildIndividualFundChoices(true);if(individualSearchBox!=null)individualSearchBox.Text="";individualFilterLoading=false;InitializeAdjustmentFilters();InitializeReviewFilters();RefreshLinkedResultViews();
        }
        void RebuildIndividualFundChoices(bool resetToAll)
        {
            if(individualFundSelector==null)return;string previous=individualFundSelector.SelectedIndex>=0&&individualFundSelector.SelectedIndex<individualFundSelector.Items.Count?individualFundSelector.Items[individualFundSelector.SelectedIndex]:"전체";IEnumerable<IndividualRowData> siteRows=individualDashboard==null?Enumerable.Empty<IndividualRowData>():individualDashboard.Rows;if(individualSiteSelector!=null&&individualSiteSelector.SelectedIndex>=0&&individualSiteSelector.SelectedIndex<individualSiteKeys.Count){string site=individualSiteKeys[individualSiteSelector.SelectedIndex];siteRows=siteRows.Where(x=>x.Site==site);}var choices=new List<string>{"전체"};choices.AddRange(siteRows.Select(x=>x.Fund).Where(x=>!String.IsNullOrWhiteSpace(x)&&x!="분류필요").Distinct().OrderBy(UiFundDisplayOrder));if(siteRows.Any(IsPendingReviewRow))choices.Add("확인 필요");bool oldLoading=individualFilterLoading;individualFilterLoading=true;individualFundSelector.SetItems(choices);int selected=!resetToAll?choices.IndexOf(previous):0;individualFundSelector.SelectedIndex=selected>=0?selected:0;individualFilterLoading=oldLoading;
        }
        void RefreshLinkedResultViews(){UpdateSummaryStatistics();UpdateSummaryForSelectedSite();UpdateIndividualStatistics();UpdateIndividualView();UpdateAdjustmentStatistics();UpdateAdjustmentView();UpdateReviewStatistics();UpdateReviewView();UpdateDiscountView(false);UpdateSubmissionView();UpdateApprovalView();}
        int UiFundDisplayOrder(string fund){switch(fund){case "공무원":return 0;case "계약제교원":return 1;case "교특회계":return 2;case "학교회계":return 3;case "분류필요":return 4;default:return 5;}}
        void UpdateIndividualStatistics()
        {
            if(individualStatCards==null)return;var rows=individualDashboard==null?new List<IndividualRowData>():individualDashboard.Rows;int total=rows.Count,review=rows.Count(IsPendingReviewRow),normal=total-review,collection=rows.Count(HasCollectionDirection),refund=rows.Count(HasRefundDirection);Func<int,string> withPercent=n=>n+"명"+(total>0?" ("+(n*100.0/total).ToString("0.0")+"%)":"");string[] values={total+"명",withPercent(normal),withPercent(collection),withPercent(refund),review+"건"};for(int i=0;i<individualStatCards.Length;i++){individualStatCards[i].Value=values[i];individualStatCards[i].Invalidate();}individualPeriodLabel.Text=individualDashboard!=null&&individualDashboard.Year>0?individualDashboard.Year+"년 "+individualDashboard.Month+"월":"-";
        }
        bool HasCollectionDirection(IndividualRowData row){return !IsReviewCompleted(row)&&row.Fund!="분류필요"&&(row.HealthDifference>.5m||row.PensionDifference>.5m||row.EmploymentDifference>.5m||row.IndustrialDifference>.5m);}
        bool HasRefundDirection(IndividualRowData row){return !IsReviewCompleted(row)&&row.Fund!="분류필요"&&(row.HealthDifference<-.5m||row.PensionDifference<-.5m||row.EmploymentDifference<-.5m||row.IndustrialDifference<-.5m);}
        List<IndividualRowData> FilteredIndividualRows()
        {
            if(individualDashboard==null)return new List<IndividualRowData>();IEnumerable<IndividualRowData> query=individualDashboard.Rows;if(individualSiteSelector!=null&&individualSiteSelector.SelectedIndex>=0&&individualSiteSelector.SelectedIndex<individualSiteKeys.Count){string site=individualSiteKeys[individualSiteSelector.SelectedIndex];query=query.Where(x=>x.Site==site);}if(individualFundSelector!=null&&individualFundSelector.SelectedIndex>0){string fund=individualFundSelector.Items[individualFundSelector.SelectedIndex];query=fund=="확인 필요"?query.Where(IsPendingReviewRow):query.Where(x=>x.Fund==fund);}string search=individualSearchBox==null?"":individualSearchBox.Text.Trim();if(search.Length>0)query=query.Where(x=>IndividualSearchMatch(x,search));return query.OrderBy(x=>UiFundDisplayOrder(x.Fund)).ThenBy(x=>x.Name).ThenBy(x=>x.Birth).ToList();
        }
        bool IndividualSearchMatch(IndividualRowData row,string search)
        {
            if((row.Name??"").IndexOf(search,StringComparison.CurrentCultureIgnoreCase)>=0)return true;string digits=Regex.Replace(search,"[^0-9]","");if(digits.Length==0)return false;decimal[] amounts;if(individualAmountMode=="기관부담금"){DiscountEntry d=SavedDiscount(row);decimal[] after={row.SummaryHealthEmployer+row.SummaryLongTermEmployer,row.SummaryPensionEmployer,row.SummaryEmploymentEmployer,row.SummaryIndustrialEmployer},discount={d.HealthTotal,d.PensionTotal,d.EmploymentTotal,d.IndustrialTotal};amounts=Enumerable.Range(0,4).SelectMany(i=>new[]{after[i]+discount[i],discount[i],after[i]}).ToArray();}else amounts=new[]{row.HealthNotice,row.HealthPayroll,row.HealthDifference,row.PensionNotice,row.PensionPayroll,row.PensionDifference,row.EmploymentNotice,row.EmploymentPayroll,row.EmploymentDifference,row.IndustrialNotice,row.IndustrialPayroll,row.IndustrialDifference};return amounts.Any(x=>Regex.Replace(Math.Abs(x).ToString("#,##0",CultureInfo.InvariantCulture),"[^0-9]","").Contains(digits));
        }
        void UpdateIndividualView()
        {
            if(individualTable==null)return;List<IndividualRowData> filtered=FilteredIndividualRows();individualTable.InstitutionMode=individualAmountMode=="기관부담금";individualTable.Rows=filtered;individualTable.ScrollOffset=0;individualTable.Invalidate();IEnumerable<IndividualRowData> siteRows=individualDashboard==null?Enumerable.Empty<IndividualRowData>():individualDashboard.Rows;if(individualSiteSelector!=null&&individualSiteSelector.SelectedIndex>=0&&individualSiteSelector.SelectedIndex<individualSiteKeys.Count){string site=individualSiteKeys[individualSiteSelector.SelectedIndex];siteRows=siteRows.Where(x=>x.Site==site);}int siteTotal=siteRows.Count();string fund=individualFundSelector!=null&&individualFundSelector.SelectedIndex>=0?individualFundSelector.Items[individualFundSelector.SelectedIndex]:"전체";int fundCount=fund=="전체"?siteTotal:fund=="확인 필요"?siteRows.Count(IsPendingReviewRow):siteRows.Count(x=>x.Fund==fund);string countText=fund=="전체"?"사업장 전체 "+siteTotal+"명":"사업장 전체 "+siteTotal+"명 중 "+fund+" "+fundCount+"명";string search=individualSearchBox==null?"":individualSearchBox.Text.Trim();if(search.Length>0)countText+="  ·  검색 결과 "+filtered.Count+"명";individualRangeLabel.Text=siteTotal==0?"표시할 데이터가 없습니다.":countText+"  ·  "+individualAmountMode;
        }
        void LoadAdjustmentSelections(ExcelPackage package)
        {
            adjustmentSelections.Clear();ExcelWorksheet ws=package.Workbook.Worksheets["UI정산선택"];if(ws==null||ws.Dimension==null)return;for(int r=2;r<=ws.Dimension.End.Row;r++){string key=ws.Cells[r,1].Text;if(!String.IsNullOrWhiteSpace(key))adjustmentSelections.Add(key);}
        }
        void InitializeAdjustmentFilters()
        {
            if(adjustmentSiteSelector==null)return;adjustmentFilterLoading=true;adjustmentSiteKeys.Clear();var sites=individualDashboard==null?new List<string>():individualDashboard.Rows.Select(x=>x.Site).Distinct().OrderBy(x=>x).ToList();adjustmentSiteKeys.AddRange(sites);adjustmentSiteSelector.SetItems(sites.Select(FormatSite));if(sites.Count>0)adjustmentSiteSelector.SelectedIndex=0;adjustmentMode="전체";if(adjustmentTabs!=null)for(int i=0;i<adjustmentTabs.Length;i++){adjustmentTabs[i].Active=i==0;adjustmentTabs[i].Invalidate();}RebuildAdjustmentFundChoices(true);adjustmentFilterLoading=false;
        }
        void RebuildAdjustmentFundChoices(bool resetToAll)
        {
            if(adjustmentFundSelector==null)return;string previous=adjustmentFundSelector.SelectedIndex>=0&&adjustmentFundSelector.SelectedIndex<adjustmentFundSelector.Items.Count?adjustmentFundSelector.Items[adjustmentFundSelector.SelectedIndex]:"전체";IEnumerable<IndividualRowData> rows=individualDashboard==null?Enumerable.Empty<IndividualRowData>():individualDashboard.Rows;if(adjustmentSiteSelector!=null&&adjustmentSiteSelector.SelectedIndex>=0&&adjustmentSiteSelector.SelectedIndex<adjustmentSiteKeys.Count){string site=adjustmentSiteKeys[adjustmentSiteSelector.SelectedIndex];rows=rows.Where(x=>x.Site==site);}var choices=new List<string>{"전체"};choices.AddRange(rows.Select(x=>x.Fund).Where(x=>!String.IsNullOrWhiteSpace(x)&&x!="분류필요").Distinct().OrderBy(UiFundDisplayOrder));if(rows.Any(IsClassificationNeeded))choices.Add("분류 필요");bool old=adjustmentFilterLoading;adjustmentFilterLoading=true;adjustmentFundSelector.SetItems(choices);int selected=resetToAll?0:choices.IndexOf(previous);adjustmentFundSelector.SelectedIndex=selected>=0?selected:0;adjustmentFilterLoading=old;
        }
        IEnumerable<IndividualRowData> CurrentAdjustmentBaseRows()
        {
            IEnumerable<IndividualRowData> rows=individualDashboard==null?Enumerable.Empty<IndividualRowData>():individualDashboard.Rows;if(adjustmentSiteSelector!=null&&adjustmentSiteSelector.SelectedIndex>=0&&adjustmentSiteSelector.SelectedIndex<adjustmentSiteKeys.Count){string site=adjustmentSiteKeys[adjustmentSiteSelector.SelectedIndex];rows=rows.Where(x=>x.Site==site);}if(adjustmentFundSelector!=null&&adjustmentFundSelector.SelectedIndex>0){string fund=adjustmentFundSelector.Items[adjustmentFundSelector.SelectedIndex];rows=fund=="분류 필요"?rows.Where(IsClassificationNeeded):rows.Where(x=>x.Fund==fund);}return rows;
        }
        IEnumerable<IndividualRowData> AdjustmentRowsForMode(IEnumerable<IndividualRowData> rows,string mode)
        {
            if(mode=="반환")return rows.Where(HasActualRefund202);if(mode=="추징")return rows.Where(HasActualCollection202);if(mode=="분류 필요")return rows.Where(IsClassificationNeeded);return rows.Where(x=>HasActualRefund202(x)||HasActualCollection202(x)||IsClassificationNeeded(x));
        }
        List<IndividualRowData> FilteredAdjustmentRows(string mode=null){return AdjustmentRowsForMode(CurrentAdjustmentBaseRows(),mode??adjustmentMode).OrderBy(x=>UiFundDisplayOrder(x.Fund)).ThenBy(x=>x.Name).ThenBy(x=>x.Birth).ToList();}
        bool IsClassificationNeeded(IndividualRowData row){return row!=null&&row.Fund=="분류필요"&&!IsReviewCompleted(row);}
        static string AdjustmentKey(IndividualRowData row){return String.Join("|",new[]{row.Site??"",row.Fund??"",row.Name??"",Regex.Replace(row.Birth??"","[^0-9]","")});}
        decimal AdjustmentAmount(IndividualRowData row,string mode)
        {
            decimal[] diffs={row.HealthDifference,row.PensionDifference,row.EmploymentDifference,row.IndustrialDifference};if(mode=="반환")return diffs.Where(x=>x<-.5m).Sum(x=>Math.Abs(x));if(mode=="추징")return diffs.Where(x=>x>.5m).Sum();return diffs.Where(x=>Math.Abs(x)>.5m).Sum(x=>Math.Abs(x));
        }
        void SelectAdjustmentMode(string mode)
        {
            adjustmentMode=mode;if(adjustmentTabs!=null)foreach(AdjustmentTabButton tab in adjustmentTabs){tab.Active=tab.Caption==mode;tab.Invalidate();}UpdateAdjustmentView();
        }
        void UpdateAdjustmentStatistics()
        {
            if(adjustmentStatCards==null)return;var rows=individualDashboard==null?new List<IndividualRowData>():individualDashboard.Rows;int refund=rows.Count(HasActualRefund202),collection=rows.Count(HasActualCollection202),classification=rows.Count(IsClassificationNeeded),total=rows.Count(x=>HasActualRefund202(x)||HasActualCollection202(x)||IsClassificationNeeded(x));decimal refundAmount=rows.Sum(x=>AdjustmentAmount(x,"반환")),collectionAmount=rows.Sum(x=>AdjustmentAmount(x,"추징"));string[] values={total+"명",refund+"명",collection+"명",classification+"명"};string[] notes={"반환 "+refund+"명 · 추징 "+collection+"명 · 분류 "+classification+"명","총 "+UiDrawing.Money(refundAmount)+"원","총 "+UiDrawing.Money(collectionAmount)+"원","확인이 필요한 항목"};for(int i=0;i<adjustmentStatCards.Length;i++){adjustmentStatCards[i].Value=values[i];adjustmentStatCards[i].Note=notes[i];adjustmentStatCards[i].Invalidate();}adjustmentPeriodLabel.Text=individualDashboard!=null&&individualDashboard.Year>0?individualDashboard.Year+"년 "+individualDashboard.Month+"월":"-";
        }
        void UpdateAdjustmentView()
        {
            if(adjustmentTable==null)return;if(adjustmentFilterPanel!=null)adjustmentFilterPanel.Height=84;if(adjustmentTabs!=null)foreach(AdjustmentTabButton tab in adjustmentTabs)tab.Top=262;adjustmentTable.Top=302;adjustmentTable.Height=258;adjustmentTable.PageSize=4;List<IndividualRowData> rows=FilteredAdjustmentRows();adjustmentTable.Rows=rows;adjustmentTable.Mode=adjustmentMode;adjustmentTable.ScrollOffset=0;IEnumerable<IndividualRowData> baseRows=CurrentAdjustmentBaseRows();if(adjustmentTabs!=null)foreach(AdjustmentTabButton tab in adjustmentTabs){tab.Count=AdjustmentRowsForMode(baseRows,tab.Caption).Count();tab.Invalidate();}string fund=adjustmentFundSelector!=null&&adjustmentFundSelector.SelectedIndex>=0?adjustmentFundSelector.Items[adjustmentFundSelector.SelectedIndex]:"전체";adjustmentRangeLabel.Text=(fund=="전체"?"선택 사업장":"선택 사업장 · "+fund)+"  ·  "+adjustmentMode+" 대상 "+rows.Count+"명";if(adjustmentExcelButton!=null)adjustmentExcelButton.Enabled=true;if(adjustmentPdfButton!=null)adjustmentPdfButton.Enabled=true;UpdateAdjustmentSelectionSummary();
        }
        void UpdateAdjustmentSelectionSummary()
        {
            if(adjustmentSelectionLabel==null)return;List<IndividualRowData> rows=FilteredAdjustmentRows();adjustmentSelectionLabel.Text="현재 목록 "+rows.Count+"명";adjustmentAmountLabel.Text="정산 금액 "+UiDrawing.Money(rows.Sum(x=>AdjustmentAmount(x,adjustmentMode)))+"원";adjustmentTable.Invalidate();
        }
        List<IndividualRowData> RowsForAdjustmentExport()
        {
            if(adjustmentMode=="분류 필요")return new List<IndividualRowData>();return FilteredAdjustmentRows().Where(x=>HasActualRefund202(x)||HasActualCollection202(x)).ToList();
        }
        string CurrentAdjustmentSite(){return adjustmentSiteSelector!=null&&adjustmentSiteSelector.SelectedIndex>=0&&adjustmentSiteSelector.SelectedIndex<adjustmentSiteKeys.Count?adjustmentSiteKeys[adjustmentSiteSelector.SelectedIndex]:"전체 사업장";}
        void RefreshAdjustmentDashboard(){if(validationResult!=null&&!String.IsNullOrWhiteSpace(validationResult.Text)&&File.Exists(validationResult.Text))LoadResultIntoUi(validationResult.Text);else MessageBox.Show("먼저 파일 등록 화면에서 대사 작업을 실행해 주세요.","반환 / 추징",MessageBoxButtons.OK,MessageBoxIcon.Information);}
        void ExportAdjustmentExcel()
        {
            List<IndividualRowData> rows=RowsForAdjustmentExport();if(rows.Count==0){MessageBox.Show("반환 또는 추징 대상이 없습니다.\r\n분류 필요 항목은 자료 생성 대상에서 제외됩니다.","자료 생성",MessageBoxButtons.OK,MessageBoxIcon.Information);return;}string filename="보험료_"+(adjustmentMode=="전체"?"반환추징":adjustmentMode)+"_내역_"+DateTime.Now.ToString("yyyyMMdd")+".xlsx";using(SaveFileDialog d=new SaveFileDialog{Filter="Excel 통합 문서 (*.xlsx)|*.xlsx",FileName=filename})if(d.ShowDialog()==DialogResult.OK)try{AdjustmentReportGenerator.CreateExcel(d.FileName,rows,adjustmentMode,individualDashboard.Year,individualDashboard.Month,FormatSite(CurrentAdjustmentSite()));MessageBox.Show("엑셀 자료를 생성했습니다.","자료 생성 완료",MessageBoxButtons.OK,MessageBoxIcon.Information);OpenGeneratedFileIfEnabled(d.FileName);}catch(Exception ex){MessageBox.Show(ex.Message,"엑셀 생성 오류",MessageBoxButtons.OK,MessageBoxIcon.Error);}
        }
        void ExportAdjustmentPdf()
        {
            List<IndividualRowData> rows=RowsForAdjustmentExport();if(rows.Count==0){MessageBox.Show("반환 또는 추징 대상이 없습니다.\r\n분류 필요 항목은 자료 생성 대상에서 제외됩니다.","자료 생성",MessageBoxButtons.OK,MessageBoxIcon.Information);return;}string filename="보험료_"+(adjustmentMode=="전체"?"반환추징":adjustmentMode)+"_내역_"+DateTime.Now.ToString("yyyyMMdd")+".pdf";using(SaveFileDialog d=new SaveFileDialog{Filter="PDF 문서 (*.pdf)|*.pdf",FileName=filename})if(d.ShowDialog()==DialogResult.OK)try{AdjustmentReportGenerator.CreatePdf(d.FileName,rows,adjustmentMode,individualDashboard.Year,individualDashboard.Month,FormatSite(CurrentAdjustmentSite()));MessageBox.Show("PDF 자료를 생성했습니다.","자료 생성 완료",MessageBoxButtons.OK,MessageBoxIcon.Information);OpenGeneratedFileIfEnabled(d.FileName);}catch(Exception ex){MessageBox.Show(ex.Message,"PDF 생성 오류",MessageBoxButtons.OK,MessageBoxIcon.Error);}
        }
        void SaveAdjustmentChanges()
        {
            if(validationResult==null||String.IsNullOrWhiteSpace(validationResult.Text)||!File.Exists(validationResult.Text)){MessageBox.Show("먼저 대사 작업을 실행해 주세요.","수정사항 저장",MessageBoxButtons.OK,MessageBoxIcon.Information);return;}try{PersistAdjustmentSelections(validationResult.Text);reconciliationState.Revision++;MessageBox.Show("선택한 정산 대상과 수정사항을 대사 결과에 저장했습니다.\r\n분류 필요자의 재원 지정은 확인 필요 메뉴에서 처리할 예정입니다.","수정사항 저장",MessageBoxButtons.OK,MessageBoxIcon.Information);}catch(Exception ex){MessageBox.Show(ex.Message,"저장 오류",MessageBoxButtons.OK,MessageBoxIcon.Error);}
        }
        int ApplySelectedFundAssignments(string path,string targetFund)
        {
            if(!new[]{"계약제교원","교특회계","학교회계"}.Contains(targetFund))return 0;List<IndividualRowData> selected=FilteredAdjustmentRows("분류 필요").Where(x=>adjustmentSelections.Contains(AdjustmentKey(x))).ToList();if(selected.Count==0)return 0;var stableKeys=new HashSet<string>(selected.Select(StablePersonKey),StringComparer.Ordinal);foreach(IndividualRowData row in selected){adjustmentSelections.Remove(AdjustmentKey(row));row.Fund=targetFund;row.Status=ReclassifiedStatus(row);row.ReviewReason=row.Status=="정상"?"":row.Status=="확인 필요"?"보험별 추징·환급 혼재":row.Status;}
            bool exact=individualDashboard.Rows.All(x=>x.HasSummaryBreakdown);if(exact)RebuildSummaryDashboardFromIndividuals();else UpdateLegacySummaryClassification(selected.Count,targetFund);PersistFundAssignments(path,stableKeys);individualFilterLoading=true;RebuildIndividualFundChoices(false);individualFilterLoading=false;adjustmentFilterLoading=true;RebuildAdjustmentFundChoices(false);adjustmentFilterLoading=false;RefreshLinkedResultViews();return selected.Count;
        }
        static string StablePersonKey(IndividualRowData row){return String.Join("|",new[]{row.Site??"",row.Name??"",Regex.Replace(row.Birth??"","[^0-9]","")});}
        string ReclassifiedStatus(IndividualRowData row){bool positive=row.HealthDifference>.5m||row.PensionDifference>.5m||row.EmploymentDifference>.5m||row.IndustrialDifference>.5m,negative=row.HealthDifference<-.5m||row.PensionDifference<-.5m||row.EmploymentDifference<-.5m||row.IndustrialDifference<-.5m;return positive&&negative?"확인 필요":positive?"추징 필요":negative?"환급 필요":"정상";}
        void RebuildSummaryDashboardFromIndividuals()
        {
            int fileCount=summaryDashboard==null?0:summaryDashboard.FileCount;var rebuilt=new SummaryDashboardModel{Year=individualDashboard.Year,Month=individualDashboard.Month,FileCount=fileCount};foreach(var siteGroup in individualDashboard.Rows.GroupBy(x=>x.Site))
            {
                var site=new SummarySiteData{Site=siteGroup.Key};foreach(var fundGroup in siteGroup.GroupBy(x=>x.Fund=="분류필요"?"기타":x.Fund).OrderBy(x=>UiFundDisplayOrder(x.Key))){var row=new SummaryFundData{Fund=fundGroup.Key,People=fundGroup.Count(),HealthPersonal=fundGroup.Sum(x=>x.SummaryHealthPersonal),HealthEmployer=fundGroup.Sum(x=>x.SummaryHealthEmployer),LongTermPersonal=fundGroup.Sum(x=>x.SummaryLongTermPersonal),LongTermEmployer=fundGroup.Sum(x=>x.SummaryLongTermEmployer),PensionPersonal=fundGroup.Sum(x=>x.SummaryPensionPersonal),PensionEmployer=fundGroup.Sum(x=>x.SummaryPensionEmployer),EmploymentPersonal=fundGroup.Sum(x=>x.SummaryEmploymentPersonal),EmploymentEmployer=fundGroup.Sum(x=>x.SummaryEmploymentEmployer),IndustrialPersonal=fundGroup.Sum(x=>x.SummaryIndustrialPersonal),IndustrialEmployer=fundGroup.Sum(x=>x.SummaryIndustrialEmployer),ReviewCount=fundGroup.Count(x=>x.Status!="정상"),ShortTermCount=fundGroup.Count(x=>x.ShortTerm),HealthDifference=fundGroup.Where(x=>!IsReviewCompleted(x)).Sum(x=>x.SummaryHealthDifference),LongTermDifference=fundGroup.Where(x=>!IsReviewCompleted(x)).Sum(x=>x.SummaryLongTermDifference),PensionDifference=fundGroup.Where(x=>!IsReviewCompleted(x)).Sum(x=>x.PensionDifference),EmploymentDifference=fundGroup.Where(x=>!IsReviewCompleted(x)).Sum(x=>x.EmploymentDifference),IndustrialDifference=fundGroup.Where(x=>!IsReviewCompleted(x)).Sum(x=>x.IndustrialDifference)};row.InstitutionTotal=row.HealthEmployer+row.LongTermEmployer+row.PensionEmployer+row.EmploymentEmployer+row.IndustrialEmployer;site.Rows.Add(row);}rebuilt.Sites[siteGroup.Key]=site;
            }summaryDashboard=rebuilt;reconciliationState.Summary=rebuilt;
        }
        void UpdateLegacySummaryClassification(int count,string targetFund)
        {
            if(summaryDashboard==null||adjustmentSiteSelector==null||adjustmentSiteSelector.SelectedIndex<0||adjustmentSiteSelector.SelectedIndex>=adjustmentSiteKeys.Count)return;SummarySiteData site;if(!summaryDashboard.Sites.TryGetValue(adjustmentSiteKeys[adjustmentSiteSelector.SelectedIndex],out site))return;SummaryFundData old=site.Rows.FirstOrDefault(x=>x.Fund=="기타"||x.Fund=="분류필요"),target=site.Rows.FirstOrDefault(x=>x.Fund==targetFund);if(target==null){target=new SummaryFundData{Fund=targetFund};site.Rows.Add(target);}target.People+=count;if(old!=null){old.People=Math.Max(0,old.People-count);old.ReviewCount=Math.Max(0,old.ReviewCount-count);if(old.People==0&&old.InsuranceTotal==0)site.Rows.Remove(old);}
        }
        void PersistFundAssignments(string path,HashSet<string> stableKeys)
        {
            using(ExcelPackage package=new ExcelPackage(new FileInfo(path))){ExcelWorksheet individual=package.Workbook.Worksheets["UI개인별데이터"];if(individual!=null&&individual.Dimension!=null)for(int r=2;r<=individual.Dimension.End.Row;r++){string key=String.Join("|",new[]{individual.Cells[r,1].Text,individual.Cells[r,3].Text,Regex.Replace(individual.Cells[r,4].Text,"[^0-9]","")});if(!stableKeys.Contains(key))continue;IndividualRowData model=individualDashboard.Rows.FirstOrDefault(x=>StablePersonKey(x)==key);if(model==null)continue;individual.Cells[r,2].Value=model.Fund;individual.Cells[r,6].Value=model.Status;individual.Cells[r,21].Value=model.ReviewReason;}WriteSummaryDashboardSheet(package);WriteAdjustmentSelectionSheet(package);package.Save();}
        }
        void WriteSummaryDashboardSheet(ExcelPackage package)
        {
            ExcelWorksheet ws=package.Workbook.Worksheets["UI총괄데이터"];if(ws==null)ws=package.Workbook.Worksheets.Add("UI총괄데이터");else ws.Cells.Clear();string[] headers={"사업장관리번호","재원","인원","건강개인","건강기관","장기요양개인","장기요양기관","국민개인","국민기관","고용개인","고용기관","산재개인","산재기관","기관부담계","확인필요","대체근로자","연도","월","건강차액","장기요양차액","국민차액","고용차액","산재차액"};for(int c=0;c<headers.Length;c++)ws.Cells[1,c+1].Value=headers[c];int r=2;foreach(SummarySiteData site in summaryDashboard.Sites.Values.OrderBy(x=>x.Site))foreach(SummaryFundData row in site.Rows.OrderBy(x=>UiFundDisplayOrder(x.Fund))){object[] values={site.Site,row.Fund,row.People,row.HealthPersonal,row.HealthEmployer,row.LongTermPersonal,row.LongTermEmployer,row.PensionPersonal,row.PensionEmployer,row.EmploymentPersonal,row.EmploymentEmployer,row.IndustrialPersonal,row.IndustrialEmployer,row.InstitutionTotal,row.ReviewCount,row.ShortTermCount,summaryDashboard.Year,summaryDashboard.Month,row.HealthDifference,row.LongTermDifference,row.PensionDifference,row.EmploymentDifference,row.IndustrialDifference};for(int c=0;c<values.Length;c++)ws.Cells[r,c+1].Value=values[c];r++;}ws.Hidden=eWorkSheetHidden.Hidden;
        }
        void WriteAdjustmentSelectionSheet(ExcelPackage package){ExcelWorksheet old=package.Workbook.Worksheets["UI정산선택"];if(old!=null)package.Workbook.Worksheets.Delete(old);ExcelWorksheet ws=package.Workbook.Worksheets.Add("UI정산선택");ws.Cells[1,1].Value="선택키";ws.Cells[1,2].Value="저장시각";int r=2;foreach(string key in adjustmentSelections.OrderBy(x=>x)){ws.Cells[r,1].Value=key;ws.Cells[r,2].Value=DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");r++;}ws.Hidden=eWorkSheetHidden.Hidden;}
        void PersistAdjustmentSelections(string path){using(ExcelPackage package=new ExcelPackage(new FileInfo(path))){WriteAdjustmentSelectionSheet(package);package.Save();}}
        public void ExportAdjustmentForTest(string result,string xlsx,string pdf,int siteIndex,int fundIndex,string mode)
        {
            validationResult.Text=result;LoadResultIntoUi(result);if(siteIndex>=0&&siteIndex<adjustmentSiteSelector.Items.Count)adjustmentSiteSelector.SelectedIndex=siteIndex;if(fundIndex>=0&&fundIndex<adjustmentFundSelector.Items.Count)adjustmentFundSelector.SelectedIndex=fundIndex;SelectAdjustmentMode(mode);List<IndividualRowData> rows=RowsForAdjustmentExport();if(rows.Count==0)throw new InvalidOperationException("반환 또는 추징 대상이 없습니다.");AdjustmentReportGenerator.CreateExcel(xlsx,rows,adjustmentMode,individualDashboard.Year,individualDashboard.Month,FormatSite(CurrentAdjustmentSite()));AdjustmentReportGenerator.CreatePdf(pdf,rows,adjustmentMode,individualDashboard.Year,individualDashboard.Month,FormatSite(CurrentAdjustmentSite()));
        }
        public void SaveAdjustmentForTest(string result){validationResult.Text=result;LoadResultIntoUi(result);IndividualRowData first=FilteredAdjustmentRows().FirstOrDefault();if(first==null)throw new InvalidOperationException("저장 검증 대상이 없습니다.");adjustmentSelections.Add(AdjustmentKey(first));PersistAdjustmentSelections(result);}
        public void ApplyFundForTest(string result,string fund){validationResult.Text=result;LoadResultIntoUi(result);SelectAdjustmentMode("분류 필요");for(int i=0;i<adjustmentSiteSelector.Items.Count&&FilteredAdjustmentRows().Count==0;i++)adjustmentSiteSelector.SelectedIndex=i;IndividualRowData first=FilteredAdjustmentRows().FirstOrDefault();if(first==null)throw new InvalidOperationException("분류 필요 검증 대상이 없습니다.");adjustmentSelections.Add(AdjustmentKey(first));if(ApplySelectedFundAssignments(result,fund)!=1)throw new InvalidOperationException("적용 재원 저장에 실패했습니다.");}
        public void SaveReviewFundForTest(string result,string fund){validationResult.Text=result;LoadResultIntoUi(result);IndividualRowData first=FilteredReviewRows().FirstOrDefault();if(first==null)throw new InvalidOperationException("확인 필요 저장 검증 대상이 없습니다.");reviewSelections.Add(ReviewKey(first));reviewFundDrafts[ReviewKey(first)]=fund;reviewSelections.Remove(ReviewKey(first));if(!reviewFundDrafts.ContainsKey(ReviewKey(first)))throw new InvalidOperationException("체크 해제 후 재원 적용 상태가 유지되지 않았습니다.");if(PersistReviewChangesCore(result)!=1)throw new InvalidOperationException("확인 필요 재원 저장에 실패했습니다.");}
        public void SaveReviewCompletionForTest(string result)
        {
            validationResult.Text=result;LoadResultIntoUi(result);List<IndividualRowData> rows=FilteredReviewRows();if(rows.Count==0)throw new InvalidOperationException("확인 완료 검증 대상이 없습니다.");reviewTable.ToggleAllSelection();if(rows.Any(x=>!reviewSelections.Contains(ReviewKey(x))))throw new InvalidOperationException("확인 명단 머리글 전체 선택이 동작하지 않습니다.");reviewTable.ToggleAllSelection();if(rows.Any(x=>reviewSelections.Contains(ReviewKey(x))))throw new InvalidOperationException("확인 명단 머리글 전체 선택 해제가 동작하지 않습니다.");IndividualRowData target=rows.FirstOrDefault(x=>HasCollectionDirection(x)||HasRefundDirection(x))??rows[0];string key=ReviewKey(target);bool hadCollection=HasCollectionDirection(target),hadRefund=HasRefundDirection(target);decimal[] differences={target.HealthDifference,target.PensionDifference,target.EmploymentDifference,target.IndustrialDifference};reviewSelections.Add(key);MarkSelectedReviewsChecked();if(!IsReviewCompleted(target)||target.Status!="정상"||FilteredReviewRows().Any(x=>ReviewKey(x)==key))throw new InvalidOperationException("확인 완료 대상이 확인 필요 명단에서 즉시 제외되지 않았습니다.");if(HasCollectionDirection(target)!=hadCollection||HasRefundDirection(target)!=hadRefund||!differences.SequenceEqual(new[]{target.HealthDifference,target.PensionDifference,target.EmploymentDifference,target.IndustrialDifference}))throw new InvalidOperationException("확인 완료 처리 중 반환·추징 차액이 변경되었습니다.");PersistReviewChangesCore(result);LoadResultIntoUi(result);IndividualRowData reloaded=individualDashboard.Rows.FirstOrDefault(x=>ReviewKey(x)==key);if(reloaded==null||!IsReviewCompleted(reloaded)||reloaded.Status!="정상"||FilteredReviewRows().Any(x=>ReviewKey(x)==key))throw new InvalidOperationException("확인 완료 상태가 저장 후 유지되지 않았습니다.");if(HasCollectionDirection(reloaded)!=hadCollection||HasRefundDirection(reloaded)!=hadRefund||!differences.SequenceEqual(new[]{reloaded.HealthDifference,reloaded.PensionDifference,reloaded.EmploymentDifference,reloaded.IndustrialDifference}))throw new InvalidOperationException("저장 후 반환·추징 차액이 유지되지 않았습니다.");if((hadCollection||hadRefund)&&!FilteredAdjustmentRows().Any(x=>ReviewKey(x)==key))throw new InvalidOperationException("확인 완료 대상이 반환·추징 명단에서 사라졌습니다.");
        }
        public void SaveDiscountForTest(string result){validationResult.Text=result;LoadResultIntoUi(result);IndividualRowData first=FilteredDiscountRows().FirstOrDefault(x=>x.Status=="정상")??FilteredDiscountRows().FirstOrDefault();if(first==null)throw new InvalidOperationException("감면 저장 검증 대상이 없습니다.");string key=DiscountKey(first);decimal healthNotice=first.HealthNotice,healthPersonal=first.SummaryHealthPersonal,healthEmployer=first.SummaryHealthEmployer,employmentEmployer=first.SummaryEmploymentEmployer;DiscountEntry before=EffectiveDiscount(first),draft=before.Clone();draft.AutoEmployment=true;draft.Health=before.Health+1230m;decimal expectedHealthDelta=draft.HealthTotal-before.HealthTotal,expectedEmploymentDelta=draft.EmploymentTotal-before.EmploymentTotal;discountDrafts[key]=draft;if(PersistDiscountChangesCore(result)!=1)throw new InvalidOperationException("감면 적용 저장에 실패했습니다.");LoadResultIntoUi(result);IndividualRowData reloaded=individualDashboard.Rows.FirstOrDefault(x=>DiscountKey(x)==key);if(reloaded==null)throw new InvalidOperationException("저장한 감면 대상자를 다시 찾을 수 없습니다.");DiscountEntry saved=EffectiveDiscount(reloaded);if(!saved.AutoEmployment||saved.Health!=draft.Health)throw new InvalidOperationException("감면 적용값이 결과 파일에 유지되지 않았습니다.");if(reloaded.Status!="정상")throw new InvalidOperationException("기관부담 감면액으로 개인 대사 판정이 변경되었습니다.");if(reloaded.HealthNotice!=healthNotice||reloaded.SummaryHealthPersonal!=healthPersonal)throw new InvalidOperationException("기관부담 감면액이 개인부담금에 반영되었습니다.");if(reloaded.SummaryHealthEmployer!=healthEmployer-expectedHealthDelta||reloaded.SummaryEmploymentEmployer!=employmentEmployer-expectedEmploymentDelta)throw new InvalidOperationException("감면액이 기관부담금에서 정확히 차감되지 않았습니다.");}
        public void SaveSchoolDiscountForTest(string result){validationResult.Text=result;LoadResultIntoUi(result);List<IndividualRowData> school=individualDashboard.Rows.Where(x=>x.Fund=="학교회계").Take(4).ToList();if(school.Count==0)throw new InvalidOperationException("학교회계 감면 검증 대상이 없습니다.");decimal before=school.Sum(x=>x.SummaryHealthEmployer+x.SummaryLongTermEmployer+x.SummaryPensionEmployer+x.SummaryEmploymentEmployer+x.SummaryIndustrialEmployer),discount=0;foreach(IndividualRowData row in school){DiscountEntry saved=EffectiveDiscount(row),draft=saved.Clone();draft.AutoEmployment=true;discount+=draft.EmploymentTotal-saved.EmploymentTotal;discountDrafts[DiscountKey(row)]=draft;}if(PersistDiscountChangesCore(result)!=school.Count)throw new InvalidOperationException("학교회계 자동이체 감면 저장 건수가 일치하지 않습니다.");LoadResultIntoUi(result);decimal after=individualDashboard.Rows.Where(x=>x.Fund=="학교회계").Sum(x=>x.SummaryHealthEmployer+x.SummaryLongTermEmployer+x.SummaryPensionEmployer+x.SummaryEmploymentEmployer+x.SummaryIndustrialEmployer);if(after!=before-discount)throw new InvalidOperationException("학교회계 기관부담 합계에서 자동이체 감면액이 정확히 차감되지 않았습니다.");HashSet<string> keys=new HashSet<string>(school.Select(DiscountKey));if(individualDashboard.Rows.Where(x=>keys.Contains(DiscountKey(x))).Any(x=>x.Status!="정상"))throw new InvalidOperationException("학교회계 감면 대상의 정상 판정이 변경되었습니다.");}
        public void AuditDashboardStatsForTest(string result,string report)
        {
            validationResult.Text=result;LoadResultIntoUi(result);var rows=individualDashboard==null?new List<IndividualRowData>():individualDashboard.Rows;int sites=summaryDashboard==null?0:summaryDashboard.Sites.Count,files=summaryDashboard==null?0:summaryDashboard.FileCount,total=rows.Count,shortTerm=rows.Count(x=>x.ShortTerm),review=rows.Count(IsPendingReviewRow),normal=total-review,collection=rows.Count(HasCollectionDirection),refund=rows.Count(HasRefundDirection),classification=rows.Count(IsClassificationNeeded),adjustmentTotal=rows.Count(x=>HasRefundDirection(x)||HasCollectionDirection(x)||IsClassificationNeeded(x));Func<int,string> withPercent=n=>n+"명"+(total>0?" ("+(n*100.0/total).ToString("0.0")+"%)":"");ExpectDashboardStat(summaryStatCards[0],sites+"개","총괄표 · 총 사업장 수");ExpectDashboardStat(summaryStatCards[1],files+"개","총괄표 · 처리 완료 파일");ExpectDashboardStat(summaryStatCards[2],total+"명","총괄표 · 총 근로자 수");ExpectDashboardStat(summaryStatCards[3],shortTerm+"명","총괄표 · 대체근로자 수");ExpectDashboardStat(summaryStatCards[4],review+"건","총괄표 · 확인 필요 항목");ExpectDashboardStat(individualStatCards[0],total+"명","개인별 · 전체 인원");ExpectDashboardStat(individualStatCards[1],withPercent(normal),"개인별 · 정상");ExpectDashboardStat(individualStatCards[2],withPercent(collection),"개인별 · 추징");ExpectDashboardStat(individualStatCards[3],withPercent(refund),"개인별 · 환급");ExpectDashboardStat(individualStatCards[4],review+"건","개인별 · 확인 필요");ExpectDashboardStat(adjustmentStatCards[0],adjustmentTotal+"명","반환/추징 · 전체");ExpectDashboardStat(adjustmentStatCards[1],refund+"명","반환/추징 · 반환");ExpectDashboardStat(adjustmentStatCards[2],collection+"명","반환/추징 · 추징");ExpectDashboardStat(adjustmentStatCards[3],classification+"명","반환/추징 · 분류 필요");ExpectDashboardStat(reviewStatCards[0],review+"건","확인 필요 · 전체");string[] insurance={"건강보험","국민연금","고용보험","산재보험"};int insuranceTotal=0;var lines=new List<string>{"상단 상황판 집계 회귀 검사: 통과","결과 파일: "+Path.GetFileName(result),"","총괄표: 사업장 "+sites+"개 / 파일 "+files+"개 / 근로자 "+total+"명 / 대체근로자 "+shortTerm+"명 / 확인 필요 "+review+"건","개인별: 전체 "+total+"명 / 정상 "+normal+"명 / 추징 "+collection+"명 / 환급 "+refund+"명 / 확인 필요 "+review+"건","반환·추징: 전체 "+adjustmentTotal+"명 / 반환 "+refund+"명 / 추징 "+collection+"명 / 분류 필요 "+classification+"명"};for(int i=0;i<insurance.Length;i++){int count=rows.Count(x=>IsPendingReviewRow(x)&&PrimaryReviewInsurance(x)==insurance[i]);insuranceTotal+=count;string expected=count+"건"+(review>0?" ("+(count*100.0/review).ToString("0.0")+"%)":"");ExpectDashboardStat(reviewStatCards[i+1],expected,"확인 필요 · "+insurance[i]);lines.Add("확인 필요 "+insurance[i]+": "+count+"건");}if(insuranceTotal!=review)throw new InvalidOperationException("확인 필요 보험별 합계("+insuranceTotal+")가 전체("+review+")와 다릅니다.");lines.Add("확인 필요 보험별 합계: "+insuranceTotal+"건");lines.Add("");lines.Add("판정 기준: 확인 완료는 정상으로 제외, 추징·환급·분류 필요는 확인 필요 전체에 포함, 혼재 방향은 반환·추징 양쪽 세부 인원에 각각 포함");File.WriteAllLines(report,lines,Encoding.UTF8);
        }
        static void ExpectDashboardStat(DashboardStatCard card,string expected,string label){if(card==null||card.Value!=expected)throw new InvalidOperationException(label+" 집계가 일치하지 않습니다. 기대값: "+expected+", 표시값: "+(card==null?"(없음)":card.Value));}
        void FillGridFromSheet(DataGridView grid,ExcelWorksheet ws,int maxRows)
        {
            if(grid==null||ws.Dimension==null)return;int header=1,best=0;for(int r=1;r<=Math.Min(15,ws.Dimension.End.Row);r++){int count=0;for(int c=1;c<=Math.Min(18,ws.Dimension.End.Column);c++)if(!String.IsNullOrWhiteSpace(ws.Cells[r,c].Text))count++;if(count>best){best=count;header=r;}}int lastCol=Math.Min(14,ws.Dimension.End.Column);grid.Columns.Clear();grid.Rows.Clear();for(int c=1;c<=lastCol;c++){string name=ws.Cells[header,c].Text;if(String.IsNullOrWhiteSpace(name))name="열 "+c;grid.Columns.Add("C"+c,name);}for(int r=header+1;r<=Math.Min(ws.Dimension.End.Row,header+maxRows);r++){object[] values=new object[lastCol];bool any=false;for(int c=1;c<=lastCol;c++){string value=ws.Cells[r,c].Text;values[c-1]=value;if(!String.IsNullOrWhiteSpace(value))any=true;}if(any)grid.Rows.Add(values);}grid.AutoSizeColumnsMode=DataGridViewAutoSizeColumnsMode.DisplayedCells;grid.AutoSizeColumnsMode=DataGridViewAutoSizeColumnsMode.Fill;
        }
        void CreateSubmission(bool teacher,bool showMessage=true)
        {
            if(String.IsNullOrWhiteSpace(validationResult.Text)||!File.Exists(validationResult.Text)){MessageBox.Show("먼저 검증 결과 파일을 선택해 주세요.");return;}
            decimal rate;if(!Decimal.TryParse((industrialRate.Text??"").Trim(),NumberStyles.Any,CultureInfo.InvariantCulture,out rate)&&!Decimal.TryParse((industrialRate.Text??"").Trim(),NumberStyles.Any,CultureInfo.CurrentCulture,out rate)||rate<=0||rate>=1){MessageBox.Show("산재보험 요율을 0보다 크고 1보다 작은 소수로 입력해 주세요.\r\n예: 0.008","산재보험 요율 확인",MessageBoxButtons.OK,MessageBoxIcon.Warning);industrialRate.Focus();return;}industrialRate.Text=rate.ToString("0.########",CultureInfo.InvariantCulture);if(!new[]{"1차","2차","3차","4차"}.Contains(submissionRound.Text))submissionRound.Text="1차";
            SubmissionInfo info=new SubmissionInfo{RecipientCode=recipientCode.Text.Trim(),InstitutionName=institutionName.Text.Trim(),ManagerName=managerName.Text.Trim(),Phone=phone.Text.Trim(),BankName=bankName.Text.Trim(),AccountNumber=accountNumber.Text.Trim(),Round=submissionRound.Text.Trim(),IndustrialRate=industrialRate.Text.Trim(),Site=CurrentSubmissionSite()};
            if(!Preflight202(teacher,info))return;
            try{string folder=submitOutput==null||String.IsNullOrWhiteSpace(submitOutput.Text)?Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory):submitOutput.Text;Directory.CreateDirectory(folder);string path=SubmissionGenerator.Create(validationResult.Text,folder,teacher,info);submitStatus.Text="완료: "+path;if(showMessage)MessageBox.Show("제출 서식을 만들었습니다.","완료",MessageBoxButtons.OK,MessageBoxIcon.Information);OpenGeneratedFileIfEnabled(path);}catch(Exception ex){submitStatus.Text="오류가 발생했습니다.";MessageBox.Show(ex.Message,"제출 생성 오류",MessageBoxButtons.OK,MessageBoxIcon.Error);}
        }
    }

    class ReconciliationUiState{public SummaryDashboardModel Summary;public IndividualDashboardModel Individuals;public int Revision;}
    class SummaryDashboardModel{public int Year,Month,FileCount;public readonly Dictionary<string,SummarySiteData> Sites=new Dictionary<string,SummarySiteData>();}
    class SummarySiteData{public string Site;public readonly List<SummaryFundData> Rows=new List<SummaryFundData>();}
    class IndividualDashboardModel{public int Year,Month;public readonly List<IndividualRowData> Rows=new List<IndividualRowData>();}
    class DiscountEntry
    {
        public bool AutoEmployment,AutoIndustrial;public decimal Health,Pension,Employment,Industrial;public decimal HealthTotal{get{return Math.Max(0,Health);}}public decimal PensionTotal{get{return Math.Max(0,Pension);}}public decimal EmploymentTotal{get{return Math.Max(0,Employment)+(AutoEmployment?250m:0m);}}public decimal IndustrialTotal{get{return Math.Max(0,Industrial)+(AutoIndustrial?250m:0m);}}public decimal Total{get{return HealthTotal+PensionTotal+EmploymentTotal+IndustrialTotal;}}public DiscountEntry Clone(){return new DiscountEntry{AutoEmployment=AutoEmployment,AutoIndustrial=AutoIndustrial,Health=Health,Pension=Pension,Employment=Employment,Industrial=Industrial};}
    }
    class DiscountAggregateRow
    {
        public string Fund;public readonly decimal[] Billed=new decimal[4],Discount=new decimal[4],After=new decimal[4];public static DiscountAggregateRow Total(IEnumerable<DiscountAggregateRow> rows){var total=new DiscountAggregateRow{Fund="합계"};foreach(DiscountAggregateRow row in rows)for(int i=0;i<4;i++){total.Billed[i]+=row.Billed[i];total.Discount[i]+=row.Discount[i];total.After[i]+=row.After[i];}return total;}
    }
    class IndividualRowData
    {
        public string Site,Fund,Name,Birth,Job,Status,ReviewReason;public decimal HealthNotice,HealthPayroll,HealthDifference,PensionNotice,PensionPayroll,PensionDifference,EmploymentNotice,EmploymentPayroll,EmploymentDifference,IndustrialNotice,IndustrialPayroll,IndustrialDifference,SummaryHealthPersonal,SummaryHealthEmployer,SummaryLongTermPersonal,SummaryLongTermEmployer,SummaryPensionPersonal,SummaryPensionEmployer,SummaryEmploymentPersonal,SummaryEmploymentEmployer,SummaryIndustrialPersonal,SummaryIndustrialEmployer,SummaryHealthDifference,SummaryLongTermDifference;public bool ShortTerm,HasSummaryBreakdown;
    }
    class SummaryFundData
    {
        public string Fund;public int People,ReviewCount,ShortTermCount;public decimal HealthPersonal,HealthEmployer,LongTermPersonal,LongTermEmployer,PensionPersonal,PensionEmployer,EmploymentPersonal,EmploymentEmployer,IndustrialPersonal,IndustrialEmployer,InstitutionTotal,HealthDifference,LongTermDifference,PensionDifference,EmploymentDifference,IndustrialDifference;
        public decimal InsuranceTotal{get{return HealthPersonal+HealthEmployer+LongTermPersonal+LongTermEmployer+PensionPersonal+PensionEmployer+EmploymentPersonal+EmploymentEmployer+IndustrialPersonal+IndustrialEmployer;}}
        public decimal OverallDifference{get{return HealthDifference+LongTermDifference+PensionDifference+EmploymentDifference+IndustrialDifference;}}
        public static SummaryFundData Total(IEnumerable<SummaryFundData> source){var t=new SummaryFundData{Fund="합계"};foreach(var r in source){t.People+=r.People;t.ReviewCount+=r.ReviewCount;t.ShortTermCount+=r.ShortTermCount;t.HealthPersonal+=r.HealthPersonal;t.HealthEmployer+=r.HealthEmployer;t.LongTermPersonal+=r.LongTermPersonal;t.LongTermEmployer+=r.LongTermEmployer;t.PensionPersonal+=r.PensionPersonal;t.PensionEmployer+=r.PensionEmployer;t.EmploymentPersonal+=r.EmploymentPersonal;t.EmploymentEmployer+=r.EmploymentEmployer;t.IndustrialPersonal+=r.IndustrialPersonal;t.IndustrialEmployer+=r.IndustrialEmployer;t.InstitutionTotal+=r.InstitutionTotal;t.HealthDifference+=r.HealthDifference;t.LongTermDifference+=r.LongTermDifference;t.PensionDifference+=r.PensionDifference;t.EmploymentDifference+=r.EmploymentDifference;t.IndustrialDifference+=r.IndustrialDifference;}return t;}
    }

    static class UiDrawing
    {
        public static GraphicsPath Rounded(RectangleF r,float radius){float d=radius*2;var p=new GraphicsPath();p.AddArc(r.Left,r.Top,d,d,180,90);p.AddArc(r.Right-d,r.Top,d,d,270,90);p.AddArc(r.Right-d,r.Bottom-d,d,d,0,90);p.AddArc(r.Left,r.Bottom-d,d,d,90,90);p.CloseFigure();return p;}
        public static void Text(Graphics g,string text,Font font,Color color,Rectangle bounds,ContentAlignment align){TextFormatFlags flags=TextFormatFlags.EndEllipsis|TextFormatFlags.VerticalCenter|TextFormatFlags.SingleLine;if(align==ContentAlignment.MiddleCenter)flags|=TextFormatFlags.HorizontalCenter;else if(align==ContentAlignment.MiddleRight)flags|=TextFormatFlags.Right;else flags|=TextFormatFlags.Left;TextRenderer.DrawText(g,text,font,bounds,color,flags);}
        public static string Money(decimal value){return value.ToString("#,##0;−#,##0;0",CultureInfo.InvariantCulture);}
        public static Color StatusColor(decimal difference){if(Math.Abs(difference)<=.5m)return Color.FromArgb(22,151,74);return difference>0?Color.FromArgb(239,63,63):Color.FromArgb(43,102,224);}
    }

    class SubmissionSummaryRow
    {
        public string Insurance;public int PrimaryPeople,SecondaryPeople;public decimal PrimaryAmount,SecondaryAmount;
        public int TotalPeople{get{return PrimaryPeople+SecondaryPeople;}}public decimal TotalAmount{get{return PrimaryAmount+SecondaryAmount;}}
        public static SubmissionSummaryRow Total(IEnumerable<SubmissionSummaryRow> rows,bool worker){var total=new SubmissionSummaryRow{Insurance="합계"};foreach(SubmissionSummaryRow row in rows){total.PrimaryAmount+=row.PrimaryAmount;total.SecondaryAmount+=row.SecondaryAmount;}total.PrimaryPeople=rows.Max(x=>x.PrimaryPeople);total.SecondaryPeople=worker?rows.Max(x=>x.SecondaryPeople):0;return total;}
    }

    class SubmissionSummaryControl : Control
    {
        public bool WorkerMode;public Color Accent=Color.FromArgb(32,84,225);public List<SubmissionSummaryRow> Rows=new List<SubmissionSummaryRow>();
        public SubmissionSummaryControl(){SetStyle(ControlStyles.UserPaint|ControlStyles.OptimizedDoubleBuffer|ControlStyles.AllPaintingInWmPaint,true);BackColor=Color.White;}
        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.SmoothingMode=SmoothingMode.AntiAlias;e.Graphics.Clear(Parent==null?UiTheme.Page:Parent.BackColor);RectangleF outer=new RectangleF(.5F,.5F,Width-1.5F,Height-1.5F);using(GraphicsPath path=UiDrawing.Rounded(outer,8))using(SolidBrush fill=new SolidBrush(UiTheme.Card))using(Pen border=new Pen(UiTheme.Border)){e.Graphics.FillPath(fill,path);e.Graphics.DrawPath(border,path);}int head=WorkerMode?54:42,rowHeight=(Height-head-2)/5;int[] xs=WorkerMode?new[]{1,(int)(Width*.20),(int)(Width*.33),(int)(Width*.47),(int)(Width*.60),(int)(Width*.74),(int)(Width*.87),Width-1}:new[]{1,(int)(Width*.34),(int)(Width*.61),Width-1};using(SolidBrush hf=new SolidBrush(UiTheme.Header))e.Graphics.FillRectangle(hf,1,1,Width-2,head);using(Font header=new Font("맑은 고딕",6.8F,FontStyle.Bold),cell=new Font("맑은 고딕",7.2F,FontStyle.Bold),money=new Font("Segoe UI",7F,FontStyle.Bold))using(Pen grid=new Pen(UiTheme.Grid))
            {
                if(WorkerMode){UiDrawing.Text(e.Graphics,"보험 구분",header,UiTheme.Text,new Rectangle(xs[0],1,xs[1]-xs[0],head),ContentAlignment.MiddleCenter);string[] groups={"무기계약","기간제","합계"};for(int g=0;g<3;g++){int start=1+g*2;UiDrawing.Text(e.Graphics,groups[g],header,UiTheme.Dark?Color.FromArgb(174,194,255):Color.FromArgb(34,65,160),new Rectangle(xs[start],1,xs[start+2]-xs[start],26),ContentAlignment.MiddleCenter);UiDrawing.Text(e.Graphics,"인원",header,UiTheme.Muted,new Rectangle(xs[start],27,xs[start+1]-xs[start],27),ContentAlignment.MiddleCenter);UiDrawing.Text(e.Graphics,"신청액",header,UiTheme.Muted,new Rectangle(xs[start+1],27,xs[start+2]-xs[start+1],27),ContentAlignment.MiddleCenter);}}else{string[] heads={"보험 구분","인원","신청 금액(원)"};for(int i=0;i<3;i++)UiDrawing.Text(e.Graphics,heads[i],header,UiTheme.Dark?Color.FromArgb(133,222,169):Color.FromArgb(28,104,65),new Rectangle(xs[i],1,xs[i+1]-xs[i],head),ContentAlignment.MiddleCenter);}List<SubmissionSummaryRow> data=Rows??new List<SubmissionSummaryRow>();for(int r=0;r<5;r++){int y=head+r*rowHeight;SubmissionSummaryRow row=r<data.Count?data[r]:new SubmissionSummaryRow{Insurance=r==4?"합계":new[]{"건강보험","국민연금","고용보험","산재보험"}[Math.Min(r,3)]};if(row.Insurance=="합계")using(SolidBrush totalFill=new SolidBrush(UiTheme.Surface))e.Graphics.FillRectangle(totalFill,1,y,Width-2,rowHeight);UiDrawing.Text(e.Graphics,row.Insurance,cell,row.Insurance=="합계"?Accent:UiTheme.Text,new Rectangle(xs[0],y,xs[1]-xs[0],rowHeight),ContentAlignment.MiddleCenter);if(WorkerMode){object[] values={row.PrimaryPeople+"명",UiDrawing.Money(row.PrimaryAmount),row.SecondaryPeople+"명",UiDrawing.Money(row.SecondaryAmount),row.TotalPeople+"명",UiDrawing.Money(row.TotalAmount)};for(int i=0;i<6;i++)UiDrawing.Text(e.Graphics,Convert.ToString(values[i]),i%2==0?cell:money,row.Insurance=="합계"?Accent:UiTheme.Text,new Rectangle(xs[i+1],y,xs[i+2]-xs[i+1],rowHeight),ContentAlignment.MiddleCenter);}else{UiDrawing.Text(e.Graphics,row.PrimaryPeople+"명",cell,row.Insurance=="합계"?Accent:UiTheme.Text,new Rectangle(xs[1],y,xs[2]-xs[1],rowHeight),ContentAlignment.MiddleCenter);UiDrawing.Text(e.Graphics,UiDrawing.Money(row.PrimaryAmount),money,row.Insurance=="합계"?Accent:UiTheme.Text,new Rectangle(xs[2],y,xs[3]-xs[2],rowHeight),ContentAlignment.MiddleCenter);}e.Graphics.DrawLine(grid,1,y+rowHeight,Width-1,y+rowHeight);}for(int i=1;i<xs.Length-1;i++)e.Graphics.DrawLine(grid,xs[i],1,xs[i],Height-1);e.Graphics.DrawLine(grid,1,head,Width-1,head);}
        }
    }

    class ApprovalReportData
    {
        public int Year,Month;public string Site,Institution;public List<IndividualRowData> Rows=new List<IndividualRowData>();
        public decimal Amount(IndividualRowData row,int insurance){if(row==null)return 0;switch(insurance){case 0:return row.SummaryHealthEmployer;case 1:return row.SummaryLongTermEmployer;case 2:return row.SummaryPensionEmployer;case 3:return row.SummaryEmploymentEmployer;default:return row.SummaryIndustrialEmployer;}}
        public decimal InsuranceTotal(int insurance){return Rows.Sum(x=>Amount(x,insurance));}public int InsurancePeople(int insurance){return Rows.Count(x=>Math.Abs(Amount(x,insurance))>.5m);}public decimal Total{get{return Enumerable.Range(0,5).Sum(i=>InsuranceTotal(i));}}
    }

    class ApprovalPreviewControl : Control
    {
        public ApprovalReportData Data;public bool PdfMode;
        public ApprovalPreviewControl(){BackColor=Color.White;SetStyle(ControlStyles.UserPaint|ControlStyles.OptimizedDoubleBuffer|ControlStyles.AllPaintingInWmPaint,true);}
        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.SmoothingMode=SmoothingMode.AntiAlias;e.Graphics.Clear(Parent==null?Color.White:Parent.BackColor);RectangleF page=new RectangleF(1,1,Width-3,Height-3);using(GraphicsPath p=UiDrawing.Rounded(page,7))using(SolidBrush b=new SolidBrush(Color.White))using(Pen border=new Pen(Color.FromArgb(215,222,238))){e.Graphics.FillPath(b,p);e.Graphics.DrawPath(border,p);}ApprovalReportData d=Data??new ApprovalReportData();Color accent=PdfMode?Color.FromArgb(214,48,48):Color.FromArgb(18,139,72),ink=Color.FromArgb(38,44,65),grid=Color.FromArgb(169,176,197);using(Font title=new Font("맑은 고딕",10.5F,FontStyle.Bold),sub=new Font("맑은 고딕",6.7F,FontStyle.Bold),section=new Font("맑은 고딕",6.8F,FontStyle.Bold),cell=new Font("맑은 고딕",6.1F),bold=new Font("맑은 고딕",6.1F,FontStyle.Bold))using(Pen gp=new Pen(grid,.65F))
            {
                UiDrawing.Text(e.Graphics,"학교회계 기관부담금 지출내역서",title,ink,new Rectangle(20,10,Width-40,25),ContentAlignment.MiddleCenter);UiDrawing.Text(e.Graphics,"("+(d.Year>0?d.Year+"년 "+d.Month+"월분":"부과년월 미확인")+")",sub,ink,new Rectangle(20,34,Width-40,16),ContentAlignment.MiddleCenter);int x=22,w=Width-44,y=56;UiDrawing.Text(e.Graphics,"1. 개요",section,ink,new Rectangle(x,y,w,17),ContentAlignment.MiddleLeft);y+=18;string[,] overview={{"사업장관리번호",FormatSite(d.Site)},{"사업장명",d.Institution??""},{"재원구분","학교회계"},{"부과년월",d.Year>0?d.Year+"년 "+d.Month+"월":"-"},{"부과인원",d.Rows.Count+"명"}};int labelW=106,rowH=17;for(int r=0;r<overview.GetLength(0);r++){e.Graphics.DrawRectangle(gp,x,y+r*rowH,labelW,rowH);e.Graphics.DrawRectangle(gp,x+labelW,y+r*rowH,w-labelW,rowH);UiDrawing.Text(e.Graphics,overview[r,0],bold,ink,new Rectangle(x+3,y+r*rowH,labelW-6,rowH),ContentAlignment.MiddleLeft);UiDrawing.Text(e.Graphics,overview[r,1],cell,ink,new Rectangle(x+labelW+4,y+r*rowH,w-labelW-8,rowH),ContentAlignment.MiddleLeft);}y+=overview.GetLength(0)*rowH+9;UiDrawing.Text(e.Graphics,"2. 보험별 기관부담금 현황",section,ink,new Rectangle(x,y,w,17),ContentAlignment.MiddleLeft);y+=18;string[] insurance={"건강보험","장기요양보험","국민연금","고용보험","산재보험","합계"};int[] widths={(int)(w*.34),(int)(w*.28),w-(int)(w*.34)-(int)(w*.28)};string[] heads={"구분","부과인원","기관부담금(원)"};int xx=x;for(int c=0;c<3;c++){e.Graphics.DrawRectangle(gp,xx,y,widths[c],19);UiDrawing.Text(e.Graphics,heads[c],bold,ink,new Rectangle(xx,y,widths[c],19),ContentAlignment.MiddleCenter);xx+=widths[c];}y+=19;for(int r=0;r<insurance.Length;r++){xx=x;bool total=r==5;if(total)using(SolidBrush fill=new SolidBrush(PdfMode?Color.FromArgb(255,242,242):Color.FromArgb(240,250,244)))e.Graphics.FillRectangle(fill,x,y,w,rowH);string[] values={insurance[r],(total?d.Rows.Count:d.InsurancePeople(r))+"명",UiDrawing.Money(total?d.Total:d.InsuranceTotal(r))};for(int c=0;c<3;c++){e.Graphics.DrawRectangle(gp,xx,y,widths[c],rowH);UiDrawing.Text(e.Graphics,values[c],total?bold:cell,total?accent:ink,new Rectangle(xx+2,y,widths[c]-4,rowH),ContentAlignment.MiddleCenter);xx+=widths[c];}y+=rowH;}y+=8;UiDrawing.Text(e.Graphics,"3. 개인별 내역",section,ink,new Rectangle(x,y,w,17),ContentAlignment.MiddleLeft);y+=18;int[] detail={(int)(w*.08),(int)(w*.16),(int)(w*.22),(int)(w*.09),(int)(w*.09),(int)(w*.09),(int)(w*.09),(int)(w*.09)};detail=detail.Concat(new[]{w-detail.Sum()}).ToArray();string[] dh={"No.","성명","주민등록번호","건강","장기","국민","고용","산재","합계"};xx=x;for(int c=0;c<dh.Length;c++){e.Graphics.DrawRectangle(gp,xx,y,detail[c],20);UiDrawing.Text(e.Graphics,dh[c],bold,ink,new Rectangle(xx,y,detail[c],20),ContentAlignment.MiddleCenter);xx+=detail[c];}y+=20;List<IndividualRowData> rows=d.Rows.Take(4).ToList();for(int r=0;r<rows.Count;r++){IndividualRowData person=rows[r];decimal[] a=Enumerable.Range(0,5).Select(i=>d.Amount(person,i)).ToArray();string[] values={(r+1).ToString(),person.Name,MaskBirth(person.Birth),UiDrawing.Money(a[0]),UiDrawing.Money(a[1]),UiDrawing.Money(a[2]),UiDrawing.Money(a[3]),UiDrawing.Money(a[4]),UiDrawing.Money(a.Sum())};xx=x;for(int c=0;c<values.Length;c++){e.Graphics.DrawRectangle(gp,xx,y,detail[c],rowH);UiDrawing.Text(e.Graphics,values[c],cell,ink,new Rectangle(xx+1,y,detail[c]-2,rowH),ContentAlignment.MiddleCenter);xx+=detail[c];}y+=rowH;}if(d.Rows.Count>4){e.Graphics.DrawRectangle(gp,x,y,w,rowH);UiDrawing.Text(e.Graphics,"… 외 "+(d.Rows.Count-4)+"명",cell,Color.FromArgb(102,111,142),new Rectangle(x,y,w,rowH),ContentAlignment.MiddleCenter);y+=rowH;}if(d.Rows.Count==0)UiDrawing.Text(e.Graphics,"대사 결과를 생성하면 학교회계 대상 내역이 표시됩니다.",cell,Color.FromArgb(115,124,151),new Rectangle(x,y,w,34),ContentAlignment.MiddleCenter);
            }
        }
        static string FormatSite(string site){string d=Regex.Replace(site??"","[^0-9]","");return d.Length==11?d.Substring(0,3)+"-"+d.Substring(3,2)+"-"+d.Substring(5,6):(String.IsNullOrWhiteSpace(site)?"-":site);}static string MaskBirth(string value){string d=Regex.Replace(value??"","[^0-9]","");return d.Length>=7?d.Substring(0,6)+"-"+d.Substring(6,1)+"******":d.Length>=6?d.Substring(0,6)+"-*******":value??"";}
    }

    class CueTextBox : TextBox
    {
        const int EmSetCueBanner=0x1501;string cueText="";[DllImport("user32.dll",CharSet=CharSet.Unicode)]static extern IntPtr SendMessage(IntPtr hWnd,int msg,IntPtr wParam,string lParam);public string CueText{get{return cueText;}set{cueText=value??"";ApplyCue();}}protected override void OnHandleCreated(EventArgs e){base.OnHandleCreated(e);ApplyCue();}void ApplyCue(){if(IsHandleCreated)SendMessage(Handle,EmSetCueBanner,(IntPtr)1,cueText);}
    }

    class ModernSiteSelector : Control
    {
        public readonly List<string> Items=new List<string>();public bool ShowIcon=true,AccentBackground=false,Borderless=false;int selectedIndex=-1;bool hovered;ContextMenuStrip choicesMenu;public event EventHandler SelectedIndexChanged;
        public int SelectedIndex{get{return selectedIndex;}set{int next=value>=0&&value<Items.Count?value:-1;if(selectedIndex==next)return;selectedIndex=next;Invalidate();if(SelectedIndexChanged!=null)SelectedIndexChanged(this,EventArgs.Empty);}}
        public ModernSiteSelector(){Cursor=Cursors.Hand;Font=new Font("맑은 고딕",9F,FontStyle.Bold);SetStyle(ControlStyles.UserPaint|ControlStyles.OptimizedDoubleBuffer|ControlStyles.AllPaintingInWmPaint,true);}
        public void SetItems(IEnumerable<string> values){Items.Clear();Items.AddRange(values);selectedIndex=-1;Invalidate();}
        protected override void OnMouseEnter(EventArgs e){hovered=true;Invalidate();base.OnMouseEnter(e);}protected override void OnMouseLeave(EventArgs e){hovered=false;Invalidate();base.OnMouseLeave(e);}protected override void OnClick(EventArgs e){base.OnClick(e);ShowChoices();}
        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.SmoothingMode=SmoothingMode.AntiAlias;Color parentBack=Parent==null?UiTheme.Card:Parent.BackColor;e.Graphics.Clear(parentBack);RectangleF rect=new RectangleF(1,1,Width-3,Height-3);Color normalFill=Borderless?parentBack:AccentBackground?(UiTheme.Dark?Color.FromArgb(48,58,80):Color.FromArgb(226,233,255)):UiTheme.Input,hoverFill=Borderless?(UiTheme.Dark?Color.FromArgb(72,52,43):Color.FromArgb(255,238,225)):UiTheme.Surface,normalBorder=AccentBackground?UiTheme.Accent:UiTheme.Border;using(GraphicsPath path=UiDrawing.Rounded(rect,9))using(SolidBrush fill=new SolidBrush(hovered?hoverFill:normalFill)){e.Graphics.FillPath(fill,path);if(!Borderless)using(Pen border=new Pen(hovered?UiTheme.Accent:normalBorder,AccentBackground?1.4F:1F))e.Graphics.DrawPath(border,path);}Color blue=Borderless?Color.FromArgb(221,113,61):UiTheme.Accent;int textX=14;if(ShowIcon){using(SolidBrush tile=new SolidBrush(UiTheme.Dark?Color.FromArgb(53,63,84):Color.FromArgb(234,238,255)))e.Graphics.FillEllipse(tile,10,8,26,26);using(Pen icon=new Pen(blue,1.4F)){icon.LineJoin=LineJoin.Round;e.Graphics.DrawRectangle(icon,17,15,5,11);e.Graphics.DrawRectangle(icon,24,12,5,14);e.Graphics.DrawLine(icon,19,18,20,18);e.Graphics.DrawLine(icon,26,16,27,16);e.Graphics.DrawLine(icon,26,19,27,19);}textX=44;}UiDrawing.Text(e.Graphics,selectedIndex>=0?Items[selectedIndex]:(ShowIcon?"사업장을 선택하세요":"전체"),Font,selectedIndex>=0?UiTheme.Text:UiTheme.Muted,new Rectangle(textX,0,Width-textX-29,Height),ContentAlignment.MiddleLeft);float arrowY=Math.Max(11,(Height-10)/2F);using(Pen arrow=new Pen(blue,1.7F)){arrow.StartCap=LineCap.Round;arrow.EndCap=LineCap.Round;e.Graphics.DrawLines(arrow,new[]{new PointF(Width-24,arrowY),new PointF(Width-19,arrowY+5),new PointF(Width-14,arrowY)});}
        }
        void ShowChoices()
        {
            if(Items.Count==0)return;if(choicesMenu!=null){choicesMenu.Dispose();choicesMenu=null;}choicesMenu=new ContextMenuStrip{ShowImageMargin=false,ShowCheckMargin=false,BackColor=UiTheme.Card,Padding=new Padding(5),AutoSize=false,Size=new Size(Width,Items.Count*38+10),DropShadowEnabled=true,Renderer=new SiteMenuRenderer()};ContextMenuStrip menu=choicesMenu;for(int i=0;i<Items.Count;i++){int index=i;var item=new ToolStripMenuItem{Text=(i==selectedIndex?"✓  ":"     ")+Items[i],Checked=i==selectedIndex,AutoSize=false,Size=new Size(Width-12,36),Font=new Font("맑은 고딕",9F,i==selectedIndex?FontStyle.Bold:FontStyle.Regular),ForeColor=UiTheme.Text,Tag=i};item.Click+=(s,e)=>SelectedIndex=index;menu.Items.Add(item);}menu.Opened+=(s,e)=>{using(GraphicsPath p=UiDrawing.Rounded(new RectangleF(0,0,menu.Width-1,menu.Height-1),10))menu.Region=new Region(p);};menu.Closed+=(s,e)=>{hovered=false;Invalidate();};menu.Show(this,new Point(0,Height+3));
        }
        protected override void Dispose(bool disposing){if(disposing&&choicesMenu!=null){choicesMenu.Dispose();choicesMenu=null;}base.Dispose(disposing);}
    }

    class SiteMenuRenderer : ToolStripProfessionalRenderer
    {
        protected override void OnRenderToolStripBackground(ToolStripRenderEventArgs e){e.Graphics.Clear(UiTheme.Card);}protected override void OnRenderToolStripBorder(ToolStripRenderEventArgs e){using(Pen p=new Pen(UiTheme.Border))e.Graphics.DrawRectangle(p,0,0,e.ToolStrip.Width-1,e.ToolStrip.Height-1);}
        protected override void OnRenderMenuItemBackground(ToolStripItemRenderEventArgs e){ToolStripMenuItem item=e.Item as ToolStripMenuItem;Color back=item!=null&&item.Checked?UiTheme.Surface:e.Item.Selected?UiTheme.Surface:UiTheme.Card;RectangleF r=new RectangleF(2,2,e.Item.Width-4,e.Item.Height-4);using(GraphicsPath path=UiDrawing.Rounded(r,8))using(SolidBrush b=new SolidBrush(back))e.Graphics.FillPath(b,path);}
        protected override void OnRenderItemText(ToolStripItemTextRenderEventArgs e){ToolStripMenuItem item=e.Item as ToolStripMenuItem;e.TextColor=item!=null&&(item.Checked||item.Selected)?UiTheme.Accent:UiTheme.Text;base.OnRenderItemText(e);}
    }

    class OutputActionButton : Button
    {
        public string IconKind="excel";public Color Accent=Color.FromArgb(24,164,91);public bool Filled=true;bool hovered,pressed;
        public OutputActionButton(){FlatStyle=FlatStyle.Flat;FlatAppearance.BorderSize=0;BackColor=Color.Transparent;Cursor=Cursors.Hand;TabStop=false;Font=new Font("맑은 고딕",8.5F,FontStyle.Bold);SetStyle(ControlStyles.UserPaint|ControlStyles.OptimizedDoubleBuffer|ControlStyles.AllPaintingInWmPaint,true);}
        protected override void OnMouseEnter(EventArgs e){hovered=true;Invalidate();base.OnMouseEnter(e);}protected override void OnMouseLeave(EventArgs e){hovered=false;pressed=false;Invalidate();base.OnMouseLeave(e);}protected override void OnMouseDown(MouseEventArgs e){pressed=true;Invalidate();base.OnMouseDown(e);}protected override void OnMouseUp(MouseEventArgs e){pressed=false;Invalidate();base.OnMouseUp(e);}
        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.SmoothingMode=SmoothingMode.AntiAlias;e.Graphics.Clear(Parent==null?UiTheme.Page:Parent.BackColor);Color accent=Enabled?Accent:Color.FromArgb(170,177,198),fill=Filled?(pressed?Darken(accent,25):hovered?Darken(accent,9):accent):(hovered?UiTheme.Surface:UiTheme.Card),ink=Filled?Color.White:accent,border=Filled?fill:UiTheme.Border;RectangleF r=new RectangleF(1,1,Width-3,Height-3);using(GraphicsPath p=UiDrawing.Rounded(r,8))using(SolidBrush b=new SolidBrush(fill))using(Pen pen=new Pen(border)){e.Graphics.FillPath(b,p);e.Graphics.DrawPath(pen,p);}bool iconOnly=String.IsNullOrWhiteSpace(Text);Rectangle iconRect=iconOnly?new Rectangle((Width-22)/2,(Height-22)/2,22,22):new Rectangle(11,(Height-22)/2,22,22);DrawIcon(e.Graphics,iconRect,accent,ink,Filled);if(!iconOnly)UiDrawing.Text(e.Graphics,Text,Font,ink,new Rectangle(39,0,Width-45,Height),ContentAlignment.MiddleCenter);
        }
        static Color Darken(Color c,int amount){return Color.FromArgb(c.A,Math.Max(0,c.R-amount),Math.Max(0,c.G-amount),Math.Max(0,c.B-amount));}
        void DrawIcon(Graphics g,Rectangle r,Color accent,Color ink,bool filled)
        {
            Color page=filled?Color.White:accent,mark=filled?accent:Color.White;using(SolidBrush b=new SolidBrush(page))using(Pen p=new Pen(page,1.6F))
            {
                p.StartCap=LineCap.Round;p.EndCap=LineCap.Round;
                if(IconKind=="folder"){g.DrawLine(p,r.X+2,r.Y+7,r.X+8,r.Y+7);g.DrawLine(p,r.X+8,r.Y+7,r.X+10,r.Y+4);g.DrawLine(p,r.X+10,r.Y+4,r.X+18,r.Y+4);g.DrawLine(p,r.X+18,r.Y+4,r.X+20,r.Y+8);g.DrawRectangle(p,r.X+2,r.Y+8,18,11);return;}
                if(IconKind=="print"){g.DrawRectangle(p,r.X+5,r.Y+2,12,7);g.DrawRectangle(p,r.X+3,r.Y+8,16,8);g.DrawRectangle(p,r.X+6,r.Y+13,10,7);g.DrawEllipse(p,r.X+15,r.Y+10,1.5F,1.5F);return;}
                if(IconKind=="save"){g.DrawRectangle(p,r.X+3,r.Y+2,16,18);g.DrawRectangle(p,r.X+6,r.Y+3,8,6);g.DrawRectangle(p,r.X+6,r.Y+13,10,7);g.DrawLine(p,r.X+15,r.Y+4,r.X+15,r.Y+8);return;}
                if(IconKind=="refresh"){g.DrawArc(p,r.X+3,r.Y+3,16,16,35,285);g.DrawLine(p,r.X+16,r.Y+2,r.X+20,r.Y+4);g.DrawLine(p,r.X+20,r.Y+4,r.X+18,r.Y+8);return;}
                using(GraphicsPath pagePath=UiDrawing.Rounded(new RectangleF(r.X+2,r.Y+1,18,20),3))g.FillPath(b,pagePath);
                using(SolidBrush mb=new SolidBrush(mark))using(Font f=new Font("Segoe UI",IconKind=="pdf"?5.2F:10F,FontStyle.Bold))
                {
                    string label=IconKind=="pdf"?"PDF":"X";StringFormat sf=new StringFormat{Alignment=StringAlignment.Center,LineAlignment=StringAlignment.Center};g.DrawString(label,f,mb,new RectangleF(r.X+1,r.Y+1,20,20),sf);
                }
            }
        }
    }

    class SoftToolbarButton : Button
    {
        public string IconKind="refresh";public Color Accent=Color.FromArgb(48,63,220);bool hovered;
        public SoftToolbarButton(){FlatStyle=System.Windows.Forms.FlatStyle.Flat;FlatAppearance.BorderSize=0;BackColor=Color.Transparent;Cursor=Cursors.Hand;TabStop=false;Font=new Font("맑은 고딕",8.5F,FontStyle.Bold);SetStyle(ControlStyles.UserPaint|ControlStyles.OptimizedDoubleBuffer|ControlStyles.AllPaintingInWmPaint,true);}protected override void OnMouseEnter(EventArgs e){hovered=true;Invalidate();base.OnMouseEnter(e);}protected override void OnMouseLeave(EventArgs e){hovered=false;Invalidate();base.OnMouseLeave(e);}
        protected override void OnPaint(PaintEventArgs e){e.Graphics.SmoothingMode=SmoothingMode.AntiAlias;e.Graphics.Clear(Parent==null?UiTheme.Page:Parent.BackColor);RectangleF r=new RectangleF(1,1,Width-3,Height-3);using(GraphicsPath p=UiDrawing.Rounded(r,8))using(SolidBrush b=new SolidBrush(hovered?UiTheme.Surface:UiTheme.Card))using(Pen pen=new Pen(UiTheme.Border)){e.Graphics.FillPath(b,p);e.Graphics.DrawPath(pen,p);}if(IconKind=="refresh"){using(Pen p=new Pen(Accent,1.55F)){p.StartCap=LineCap.Round;p.EndCap=LineCap.Round;e.Graphics.DrawArc(p,10,9,14,14,35,275);e.Graphics.DrawLine(p,21,8,24,9);e.Graphics.DrawLine(p,24,9,23,12);}}else{Color iconColor=IconKind=="pdf"?Color.FromArgb(226,38,38):Accent;using(SolidBrush b=new SolidBrush(iconColor))using(GraphicsPath p=UiDrawing.Rounded(new RectangleF(9,7,17,20),3))e.Graphics.FillPath(b,p);using(SolidBrush b=new SolidBrush(Color.White))using(Font f=new Font("Segoe UI",IconKind=="pdf"?4.8F:8F,FontStyle.Bold)){StringFormat sf=new StringFormat{Alignment=StringAlignment.Center,LineAlignment=StringAlignment.Center};e.Graphics.DrawString(IconKind=="pdf"?"PDF":"X",f,b,new RectangleF(9,7,17,20),sf);}}UiDrawing.Text(e.Graphics,Text,Font,Accent,new Rectangle(31,0,Width-35,Height),ContentAlignment.MiddleLeft);}
    }

    class DashboardStatCard : Control
    {
        public string Caption="",Value="-",Note="",IconKind="";public Color Accent=Color.RoyalBlue;
        public DashboardStatCard(){SetStyle(ControlStyles.UserPaint|ControlStyles.OptimizedDoubleBuffer|ControlStyles.AllPaintingInWmPaint,true);BackColor=UiTheme.Card;}
        protected override void OnPaint(PaintEventArgs e){e.Graphics.SmoothingMode=SmoothingMode.AntiAlias;e.Graphics.Clear(Parent==null?UiTheme.Page:Parent.BackColor);RectangleF r=new RectangleF(1,1,Width-3,Height-3);using(GraphicsPath path=UiDrawing.Rounded(r,11))using(SolidBrush b=new SolidBrush(UiTheme.Card))using(Pen pen=new Pen(UiTheme.Dark?UiTheme.Border:Color.FromArgb(Accent.A,Math.Min(255,(Accent.R+245)/2),Math.Min(255,(Accent.G+245)/2),Math.Min(255,(Accent.B+245)/2)))){e.Graphics.FillPath(b,path);e.Graphics.DrawPath(pen,path);}float valueSize=(Value??"").Length>9?10.5F:(Value??"").Length>6?12F:16F;using(Font caption=new Font("맑은 고딕",8.5F,FontStyle.Bold),value=new Font("맑은 고딕",valueSize,FontStyle.Bold),note=new Font("맑은 고딕",7.5F)){UiDrawing.Text(e.Graphics,Caption,caption,UiTheme.Text,new Rectangle(14,9,125,20),ContentAlignment.MiddleLeft);UiDrawing.Text(e.Graphics,Value,value,Accent,new Rectangle(14,31,128,29),ContentAlignment.MiddleLeft);UiDrawing.Text(e.Graphics,Note,note,UiTheme.Muted,new Rectangle(14,65,150,18),ContentAlignment.MiddleLeft);}using(SolidBrush circle=new SolidBrush(Color.FromArgb(34,Accent)))e.Graphics.FillEllipse(circle,145,27,40,40);DrawStatIcon(e.Graphics,new Rectangle(157,39,16,16));}
        void DrawStatIcon(Graphics g,Rectangle r){using(Pen p=new Pen(Accent,1.7F)){p.StartCap=LineCap.Round;p.EndCap=LineCap.Round;if(IconKind=="building"){g.DrawRectangle(p,r.X+1,r.Y+4,6,11);g.DrawRectangle(p,r.X+9,r.Y+1,6,14);for(int y=0;y<3;y++){g.DrawLine(p,r.X+11,r.Y+4+y*3,r.X+13,r.Y+4+y*3);}}else if(IconKind=="files"){g.DrawRectangle(p,r.X+1,r.Y+3,12,11);g.DrawLine(p,r.X+4,r.Y+6,r.X+10,r.Y+6);g.DrawLine(p,r.X+4,r.Y+9,r.X+9,r.Y+9);g.DrawEllipse(p,r.X+10,r.Y+9,6,6);}else if(IconKind=="warning"){g.DrawPolygon(p,new[]{new Point(r.X+8,r.Y),new Point(r.X+16,r.Y+15),new Point(r.X,r.Y+15)});g.DrawLine(p,r.X+8,r.Y+5,r.X+8,r.Y+9);g.DrawEllipse(p,r.X+7.5F,r.Y+12,1,1);}else if(IconKind=="normal"){g.DrawEllipse(p,r.X,r.Y,16,16);g.DrawLine(p,r.X+4,r.Y+8,r.X+7,r.Y+11);g.DrawLine(p,r.X+7,r.Y+11,r.X+12,r.Y+5);}else if(IconKind=="collection"){g.DrawEllipse(p,r.X,r.Y,16,16);g.DrawLine(p,r.X+8,r.Y+4,r.X+8,r.Y+12);g.DrawLine(p,r.X+4,r.Y+8,r.X+12,r.Y+8);}else if(IconKind=="refund"){g.DrawEllipse(p,r.X,r.Y,16,16);g.DrawLine(p,r.X+8,r.Y+3,r.X+8,r.Y+11);g.DrawLine(p,r.X+4,r.Y+8,r.X+8,r.Y+12);g.DrawLine(p,r.X+12,r.Y+8,r.X+8,r.Y+12);}else{g.DrawEllipse(p,r.X+5,r.Y,6,6);g.DrawArc(p,r.X+1,r.Y+7,14,10,190,160);if(IconKind=="short")g.DrawEllipse(p,r.X+12,r.Y+1,3,3);}}}
    }

    class PremiumTotalsControl : Control
    {
        public SummarySiteData SiteData;public PremiumTotalsControl(){SetStyle(ControlStyles.UserPaint|ControlStyles.OptimizedDoubleBuffer|ControlStyles.AllPaintingInWmPaint,true);BackColor=UiTheme.Card;}
        protected override void OnPaint(PaintEventArgs e){e.Graphics.SmoothingMode=SmoothingMode.AntiAlias;e.Graphics.Clear(UiTheme.Card);using(Font title=new Font("맑은 고딕",8.5F,FontStyle.Bold),head=new Font("맑은 고딕",7.5F,FontStyle.Bold),value=new Font("맑은 고딕",10F,FontStyle.Bold)){UiDrawing.Text(e.Graphics,"사대보험 총 고지금액",title,UiTheme.Text,new Rectangle(0,0,Width,20),ContentAlignment.MiddleLeft);string[] names={"건강·장기요양","국민연금","고용보험","산재보험"};SummaryFundData total=SummaryFundData.Total(SiteData==null?new List<SummaryFundData>():SiteData.Rows);decimal[] values={total.HealthPersonal+total.HealthEmployer+total.LongTermPersonal+total.LongTermEmployer,total.PensionPersonal+total.PensionEmployer,total.EmploymentPersonal+total.EmploymentEmployer,total.IndustrialEmployer},differences={total.HealthDifference+total.LongTermDifference,total.PensionDifference,total.EmploymentDifference,total.IndustrialDifference};int cell=Width/4;for(int i=0;i<4;i++){Rectangle rect=new Rectangle(i*cell,24,cell,55);if(i>0)using(Pen separator=new Pen(UiTheme.Border))e.Graphics.DrawLine(separator,rect.Left,30,rect.Left,72);UiDrawing.Text(e.Graphics,names[i],head,UiTheme.Muted,new Rectangle(rect.X,24,rect.Width,22),ContentAlignment.MiddleCenter);UiDrawing.Text(e.Graphics,UiDrawing.Money(values[i]),value,UiDrawing.StatusColor(differences[i]),new Rectangle(rect.X,48,rect.Width,24),ContentAlignment.MiddleCenter);}}}
    }

    class SummaryTableControl : Control
    {
        public List<SummaryFundData> Rows=new List<SummaryFundData>();public SummaryTableControl(){SetStyle(ControlStyles.UserPaint|ControlStyles.OptimizedDoubleBuffer|ControlStyles.AllPaintingInWmPaint,true);BackColor=UiTheme.Card;}
        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.SmoothingMode=SmoothingMode.AntiAlias;e.Graphics.Clear(Parent==null?UiTheme.Page:Parent.BackColor);RectangleF outer=new RectangleF(.5F,.5F,Width-1.5F,Height-1.5F);
            using(GraphicsPath path=UiDrawing.Rounded(outer,11))using(SolidBrush fill=new SolidBrush(UiTheme.Card))using(Pen border=new Pen(UiTheme.Border)){e.Graphics.FillPath(fill,path);e.Graphics.DrawPath(border,path);}
            using(Font title=new Font("맑은 고딕",9F,FontStyle.Bold),header=new Font("맑은 고딕",8F,FontStyle.Bold),cellFont=new Font("맑은 고딕",8.2F,FontStyle.Bold),bold=new Font("맑은 고딕",8.6F,FontStyle.Bold))
            {
                UiDrawing.Text(e.Graphics,"재원별 보험료 대조 결과 (고지금액 기준)",title,UiTheme.Text,new Rectangle(14,3,500,27),ContentAlignment.MiddleLeft);UiDrawing.Text(e.Graphics,"금액 단위 : 원",header,UiTheme.Muted,new Rectangle(850,3,165,27),ContentAlignment.MiddleRight);
                int top=31,head1=29,head2=27,rowH=36;int[] widths={82,55,100,81,81,81,81,81,81,81,81,102};int[] xs=new int[widths.Length+1];for(int i=0;i<widths.Length;i++)xs[i+1]=xs[i]+widths[i];float scale=(Width-2)/(float)xs[xs.Length-1];for(int i=0;i<xs.Length;i++)xs[i]=(int)(xs[i]*scale)+1;
                using(SolidBrush hf=new SolidBrush(UiTheme.Header))e.Graphics.FillRectangle(hf,1,top,Width-2,head1+head2);using(SolidBrush peach=new SolidBrush(UiTheme.Dark?Color.FromArgb(66,49,40):UiTheme.Name=="회색"?Color.FromArgb(92,97,106):Color.FromArgb(255,248,239)))e.Graphics.FillRectangle(peach,xs[2],top,xs[3]-xs[2],head1+head2+rowH*6);
                using(Pen grid=new Pen(UiTheme.Grid))
                {
                    string[] fixedHeads={"재원","인원","기관부담금 계"};for(int i=0;i<3;i++)UiDrawing.Text(e.Graphics,fixedHeads[i],header,UiTheme.Text,new Rectangle(xs[i],top,xs[i+1]-xs[i],head1+head2),ContentAlignment.MiddleCenter);string[] groups={"건강보험","장기요양보험","국민연금","고용보험"};for(int g=0;g<4;g++){int c=3+g*2;UiDrawing.Text(e.Graphics,groups[g],header,UiTheme.Text,new Rectangle(xs[c],top,xs[c+2]-xs[c],head1),ContentAlignment.MiddleCenter);UiDrawing.Text(e.Graphics,"개인부담",header,UiTheme.Muted,new Rectangle(xs[c],top+head1,xs[c+1]-xs[c],head2),ContentAlignment.MiddleCenter);UiDrawing.Text(e.Graphics,"기관부담",header,UiTheme.Muted,new Rectangle(xs[c+1],top+head1,xs[c+2]-xs[c+1],head2),ContentAlignment.MiddleCenter);e.Graphics.DrawLine(grid,xs[c+1],top+head1,xs[c+1],top+head1+head2);}int industrialColumn=11;UiDrawing.Text(e.Graphics,"산재보험",header,UiTheme.Text,new Rectangle(xs[industrialColumn],top,xs[industrialColumn+1]-xs[industrialColumn],head1),ContentAlignment.MiddleCenter);UiDrawing.Text(e.Graphics,"기관부담",header,UiTheme.Muted,new Rectangle(xs[industrialColumn],top+head1,xs[industrialColumn+1]-xs[industrialColumn],head2),ContentAlignment.MiddleCenter);e.Graphics.DrawLine(grid,1,top+head1+head2,Width-1,top+head1+head2);
                    List<SummaryFundData> data=Rows==null?new List<SummaryFundData>():Rows.ToList();data.Add(SummaryFundData.Total(data));while(data.Count<5)data.Insert(data.Count-1,new SummaryFundData{Fund="-"});for(int r=0;r<Math.Min(6,data.Count);r++){int y=top+head1+head2+r*rowH;if(r==data.Count-1)using(SolidBrush totalFill=new SolidBrush(UiTheme.Surface))e.Graphics.FillRectangle(totalFill,1,y,Width-2,rowH);SummaryFundData d=data[r];object[] values={d.Fund,d.People,d.InstitutionTotal,d.HealthPersonal,d.HealthEmployer,d.LongTermPersonal,d.LongTermEmployer,d.PensionPersonal,d.PensionEmployer,d.EmploymentPersonal,d.EmploymentEmployer,d.IndustrialEmployer};for(int c=0;c<values.Length;c++){string text=c==0?Convert.ToString(values[c]):c==1?Convert.ToInt32(values[c]).ToString()+"명":UiDrawing.Money(Convert.ToDecimal(values[c]));decimal difference=c==2?d.OverallDifference:c<=4?d.HealthDifference:c<=6?d.LongTermDifference:c<=8?d.PensionDifference:c<=10?d.EmploymentDifference:d.IndustrialDifference;Color ink=c<2?UiTheme.Text:UiDrawing.StatusColor(difference);UiDrawing.Text(e.Graphics,text,r==data.Count-1?bold:cellFont,ink,new Rectangle(xs[c],y,xs[c+1]-xs[c],rowH),ContentAlignment.MiddleCenter);}e.Graphics.DrawLine(grid,1,y+rowH,Width-1,y+rowH);}for(int c=1;c<xs.Length-1;c++)e.Graphics.DrawLine(grid,xs[c],top,xs[c],top+head1+head2+rowH*Math.Min(6,data.Count));
                }
            }
        }
    }

    class IndividualTableControl : Control
    {
        public List<IndividualRowData> Rows=new List<IndividualRowData>();public int ScrollOffset,PageSize=6;public bool InstitutionMode;public Func<IndividualRowData,DiscountEntry> DiscountProvider;bool draggingScroll;int dragStartY,dragStartOffset;public IndividualTableControl(){SetStyle(ControlStyles.UserPaint|ControlStyles.OptimizedDoubleBuffer|ControlStyles.AllPaintingInWmPaint|ControlStyles.Selectable,true);BackColor=UiTheme.Card;TabStop=true;}
        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.SmoothingMode=SmoothingMode.AntiAlias;e.Graphics.Clear(Parent==null?UiTheme.Page:Parent.BackColor);RectangleF outer=new RectangleF(.5F,.5F,Width-1.5F,Height-1.5F);using(GraphicsPath path=UiDrawing.Rounded(outer,11))using(SolidBrush fill=new SolidBrush(UiTheme.Card))using(Pen border=new Pen(UiTheme.Border)){e.Graphics.FillPath(fill,path);e.Graphics.DrawPath(border,path);}using(Font title=new Font("맑은 고딕",8.5F,FontStyle.Bold),header=new Font("맑은 고딕",7F,FontStyle.Bold),cell=new Font("맑은 고딕",7.2F,FontStyle.Bold),small=new Font("맑은 고딕",6.8F,FontStyle.Bold))
            {
                UiDrawing.Text(e.Graphics,InstitutionMode?"개인별 기관부담금 내역 (감면 적용 기준)":"개인별 개인부담금 대사 내역 (고지금액 기준)",title,UiTheme.Text,new Rectangle(14,2,600,27),ContentAlignment.MiddleLeft);UiDrawing.Text(e.Graphics,"금액 단위 : 원",header,UiTheme.Muted,new Rectangle(860,2,155,27),ContentAlignment.MiddleRight);
                int top=30,head1=29,head2=25,rowH=33,contentRight=Width-14;int[] widths={31,58,57,74,88,67,51,51,51,51,51,51,51,51,51,51,51,51};int[] xs=new int[widths.Length+1];for(int i=0;i<widths.Length;i++)xs[i+1]=xs[i]+widths[i];float scale=(contentRight-1)/(float)xs[xs.Length-1];for(int i=0;i<xs.Length;i++)xs[i]=(int)(xs[i]*scale)+1;
                using(SolidBrush hf=new SolidBrush(UiTheme.Header))e.Graphics.FillRectangle(hf,1,top,contentRight-1,head1+head2);using(Pen grid=new Pen(UiTheme.Grid))
                {
                    string[] fixedHeads={"No.","재원","이름","주민/사번","직종명",InstitutionMode?"감면상태":"대사결과"};for(int i=0;i<fixedHeads.Length;i++)UiDrawing.Text(e.Graphics,fixedHeads[i],header,UiTheme.Text,new Rectangle(xs[i],top,xs[i+1]-xs[i],head1+head2),ContentAlignment.MiddleCenter);string[] groups={"건강보험","국민연금","고용보험","산재보험"};for(int g=0;g<4;g++){int c=6+g*3;string groupText=g==0?"건강보험 (건강+장기)":groups[g];UiDrawing.Text(e.Graphics,groupText,header,UiTheme.Text,new Rectangle(xs[c],top,xs[c+3]-xs[c],head1),ContentAlignment.MiddleCenter);string[] subs=InstitutionMode?new[]{"감면전","감면액","감면후"}:new[]{"고지금액","급여대장","차액"};for(int s=0;s<3;s++)UiDrawing.Text(e.Graphics,subs[s],small,UiTheme.Muted,new Rectangle(xs[c+s],top+head1,xs[c+s+1]-xs[c+s],head2),ContentAlignment.MiddleCenter);e.Graphics.DrawLine(grid,xs[c+1],top+head1,xs[c+1],top+head1+head2);e.Graphics.DrawLine(grid,xs[c+2],top+head1,xs[c+2],top+head1+head2);}e.Graphics.DrawLine(grid,1,top+head1,contentRight,top+head1);e.Graphics.DrawLine(grid,1,top+head1+head2,contentRight,top+head1+head2);
                    SetScrollOffset(ScrollOffset,false);List<IndividualRowData> visible=(Rows??new List<IndividualRowData>()).Skip(ScrollOffset).Take(PageSize).ToList();if(visible.Count==0)UiDrawing.Text(e.Graphics,"대사 결과를 생성하면 개인별 내역이 표시됩니다.",cell,UiTheme.Muted,new Rectangle(1,top+head1+head2,contentRight-1,Height-top-head1-head2),ContentAlignment.MiddleCenter);for(int r=0;r<visible.Count;r++){int y=top+head1+head2+r*rowH;IndividualRowData d=visible[r];if(r%2==1)using(SolidBrush alt=new SolidBrush(UiTheme.Dark?UiTheme.Surface:Color.FromArgb(252,253,255)))e.Graphics.FillRectangle(alt,1,y,contentRight-1,rowH);DrawCell(e.Graphics,""+(ScrollOffset+r+1),cell,UiTheme.Text,xs,0,y,rowH);DrawCell(e.Graphics,d.Fund,small,UiTheme.Text,xs,1,y,rowH);DrawCell(e.Graphics,d.Name,cell,UiTheme.Dark?Color.FromArgb(174,194,255):Color.FromArgb(28,52,137),xs,2,y,rowH);DrawCell(e.Graphics,MaskBirth(d.Birth),small,UiTheme.Text,xs,3,y,rowH);DrawCell(e.Graphics,d.Job,small,UiTheme.Muted,xs,4,y,rowH);DiscountEntry discount=DiscountProvider==null?new DiscountEntry():DiscountProvider(d);if(InstitutionMode)DrawDiscountStatus(e.Graphics,discount.Total,xs[5],y,xs[6]-xs[5],rowH,small);else DrawStatus(e.Graphics,d.Status,xs[5],y,xs[6]-xs[5],rowH,small);decimal[] amounts;if(InstitutionMode){decimal[] after={d.SummaryHealthEmployer+d.SummaryLongTermEmployer,d.SummaryPensionEmployer,d.SummaryEmploymentEmployer,d.SummaryIndustrialEmployer},discounts={discount.HealthTotal,discount.PensionTotal,discount.EmploymentTotal,discount.IndustrialTotal};amounts=Enumerable.Range(0,4).SelectMany(i=>new[]{after[i]+discounts[i],discounts[i],after[i]}).ToArray();}else amounts=new[]{d.HealthNotice,d.HealthPayroll,d.HealthDifference,d.PensionNotice,d.PensionPayroll,d.PensionDifference,d.EmploymentNotice,d.EmploymentPayroll,d.EmploymentDifference,d.IndustrialNotice,d.IndustrialPayroll,d.IndustrialDifference};for(int c=0;c<amounts.Length;c++){bool diff=!InstitutionMode&&c%3==2,discountCell=InstitutionMode&&c%3==1;Color ink=diff?UiDrawing.StatusColor(amounts[c]):discountCell&&amounts[c]>0?Color.FromArgb(235,116,45):UiTheme.Text;string text=diff?Difference(amounts[c]):UiDrawing.Money(amounts[c]);DrawCell(e.Graphics,text,small,ink,xs,6+c,y,rowH);}e.Graphics.DrawLine(grid,1,y+rowH,contentRight,y+rowH);}int lineBottom=top+head1+head2+rowH*Math.Max(visible.Count,1);for(int c=1;c<xs.Length-1;c++)e.Graphics.DrawLine(grid,xs[c],top,xs[c],Math.Min(Height-1,lineBottom));DrawScrollBar(e.Graphics,top+head1+head2);
                }
            }
        }
        int MaxOffset{get{return Math.Max(0,(Rows==null?0:Rows.Count)-PageSize);}}
        void SetScrollOffset(int value,bool repaint=true){int next=Math.Max(0,Math.Min(MaxOffset,value));if(next==ScrollOffset)return;ScrollOffset=next;if(repaint)Invalidate();}
        Rectangle ScrollTrack(int bodyTop){return new Rectangle(Width-10,bodyTop+4,5,Math.Max(20,Height-bodyTop-9));}
        Rectangle ScrollThumb(int bodyTop){Rectangle track=ScrollTrack(bodyTop);if(MaxOffset==0)return track;int thumbH=Math.Max(25,(int)Math.Round(track.Height*Math.Min(1.0,PageSize/(double)Math.Max(PageSize,Rows.Count)))),travel=Math.Max(1,track.Height-thumbH),y=track.Y+(int)Math.Round(travel*ScrollOffset/(double)MaxOffset);return new Rectangle(track.X-1,y,7,thumbH);}
        void DrawScrollBar(Graphics g,int bodyTop){Rectangle track=ScrollTrack(bodyTop);using(SolidBrush b=new SolidBrush(UiTheme.Dark?UiTheme.Surface:Color.FromArgb(244,246,252)))g.FillRectangle(b,track);if(MaxOffset>0){Rectangle thumb=ScrollThumb(bodyTop);using(GraphicsPath p=UiDrawing.Rounded(thumb,3))using(SolidBrush b=new SolidBrush(UiTheme.Dark?Color.FromArgb(112,125,153):Color.FromArgb(170,180,211)))g.FillPath(b,p);}}
        protected override void OnMouseEnter(EventArgs e){Focus();base.OnMouseEnter(e);}protected override void OnMouseWheel(MouseEventArgs e){SetScrollOffset(ScrollOffset+(e.Delta<0?2:-2));base.OnMouseWheel(e);}protected override void OnMouseDown(MouseEventArgs e){int bodyTop=84;if(e.X>=Width-16&&MaxOffset>0){Rectangle thumb=ScrollThumb(bodyTop);if(thumb.Contains(e.Location)){draggingScroll=true;dragStartY=e.Y;dragStartOffset=ScrollOffset;Capture=true;}else{Rectangle track=ScrollTrack(bodyTop);int target=(int)Math.Round((e.Y-track.Y)/(double)Math.Max(1,track.Height)*MaxOffset);SetScrollOffset(target);}}base.OnMouseDown(e);}protected override void OnMouseMove(MouseEventArgs e){if(draggingScroll){Rectangle track=ScrollTrack(84),thumb=ScrollThumb(84);int travel=Math.Max(1,track.Height-thumb.Height);SetScrollOffset(dragStartOffset+(int)Math.Round((e.Y-dragStartY)*MaxOffset/(double)travel));}base.OnMouseMove(e);}protected override void OnMouseUp(MouseEventArgs e){draggingScroll=false;Capture=false;base.OnMouseUp(e);}
        static void DrawCell(Graphics g,string text,Font font,Color ink,int[] xs,int c,int y,int h){UiDrawing.Text(g,text??"",font,ink,new Rectangle(xs[c]+1,y,xs[c+1]-xs[c]-2,h),ContentAlignment.MiddleCenter);}
        static string Difference(decimal value){if(Math.Abs(value)<=.5m)return "0";return value>0?"+"+UiDrawing.Money(value):UiDrawing.Money(value);}
        static string MaskBirth(string value){string digits=new string((value??"").Where(Char.IsDigit).ToArray());if(digits.Length>=13)return digits.Substring(0,6)+"-"+digits.Substring(6,1)+"******";if(digits.Length>=7)return digits.Substring(0,6)+"-"+digits.Substring(6,1)+"******";return value??"";}
        static void DrawStatus(Graphics g,string status,int x,int y,int w,int h,Font font){Color ink,back;if(status=="정상"){ink=Color.FromArgb(22,139,69);back=Color.FromArgb(235,248,238);}else if(status=="추징 필요"){ink=Color.FromArgb(229,55,55);back=Color.FromArgb(255,237,237);}else if(status=="환급 필요"){ink=Color.FromArgb(39,102,218);back=Color.FromArgb(235,243,255);}else{ink=Color.FromArgb(224,101,23);back=Color.FromArgb(255,244,232);}RectangleF pill=new RectangleF(x+4,y+7,Math.Max(20,w-8),h-14);using(GraphicsPath p=UiDrawing.Rounded(pill,6))using(SolidBrush b=new SolidBrush(back))g.FillPath(b,p);UiDrawing.Text(g,status??"확인 필요",font,ink,new Rectangle(x+2,y,w-4,h),ContentAlignment.MiddleCenter);}
        static void DrawDiscountStatus(Graphics g,decimal amount,int x,int y,int w,int h,Font font){bool applied=amount>0;Color ink=applied?Color.FromArgb(226,103,35):UiTheme.Dark?Color.FromArgb(144,221,172):Color.FromArgb(22,139,69),back=UiTheme.Dark?(applied?Color.FromArgb(79,52,37):Color.FromArgb(35,72,52)):(applied?Color.FromArgb(255,243,232):Color.FromArgb(235,248,238));RectangleF pill=new RectangleF(x+4,y+7,Math.Max(20,w-8),h-14);using(GraphicsPath p=UiDrawing.Rounded(pill,6))using(SolidBrush b=new SolidBrush(back))g.FillPath(b,p);UiDrawing.Text(g,applied?"감면 적용":"해당 없음",font,ink,new Rectangle(x+2,y,w-4,h),ContentAlignment.MiddleCenter);}
    }

    class DiscountTotalsControl : Control
    {
        public string Title="";
        public Color Tint=Color.White;
        public int Mode;
        public List<DiscountAggregateRow> Rows=new List<DiscountAggregateRow>();

        public DiscountTotalsControl()
        {
            SetStyle(ControlStyles.UserPaint|ControlStyles.OptimizedDoubleBuffer|ControlStyles.AllPaintingInWmPaint,true);
            BackColor=Color.White;
        }

        static void DrawMoneyFit(Graphics graphics,decimal value,Color color,Rectangle bounds)
        {
            string text=UiDrawing.Money(value);
            float size=7.1F;
            Font font=null;
            try
            {
                while(size>=4.3F)
                {
                    if(font!=null)font.Dispose();
                    font=new Font("Segoe UI",size,FontStyle.Bold);
                    if(graphics.MeasureString(text,font).Width<=bounds.Width-3)break;
                    size-=.3F;
                }
                UiDrawing.Text(graphics,text,font,color,bounds,ContentAlignment.MiddleCenter);
            }
            finally{if(font!=null)font.Dispose();}
        }

        static void DrawFundFit(Graphics graphics,string text,Color color,Rectangle bounds)
        {
            float size=6.8F;Font font=null;
            try
            {
                while(size>=5.2F)
                {
                    if(font!=null)font.Dispose();
                    font=new Font("맑은 고딕",size,FontStyle.Bold);
                    if(graphics.MeasureString(text??"",font).Width<=bounds.Width-3)break;
                    size-=.25F;
                }
                UiDrawing.Text(graphics,text??"",font,color,bounds,ContentAlignment.MiddleCenter);
            }
            finally{if(font!=null)font.Dispose();}
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.SmoothingMode=SmoothingMode.AntiAlias;
            e.Graphics.Clear(Parent==null?UiTheme.Page:Parent.BackColor);
            using(Font title=new Font("맑은 고딕",9.5F,FontStyle.Bold),unit=new Font("맑은 고딕",6.7F,FontStyle.Bold),header=new Font("맑은 고딕",6.6F,FontStyle.Bold),cell=new Font("맑은 고딕",6.8F,FontStyle.Bold))
            {
                UiDrawing.Text(e.Graphics,Title,title,Mode==2?Color.FromArgb(25,122,70):Color.FromArgb(36,65,177),new Rectangle(2,0,Width-90,25),ContentAlignment.MiddleLeft);
                UiDrawing.Text(e.Graphics,"금액 단위 : 원",unit,UiTheme.Muted,new Rectangle(Width-95,2,92,22),ContentAlignment.MiddleRight);
                RectangleF outer=new RectangleF(.5F,27.5F,Width-1.5F,Height-28.5F);
                using(GraphicsPath p=UiDrawing.Rounded(outer,8))using(SolidBrush b=new SolidBrush(UiTheme.Dark||UiTheme.Name=="회색"?UiTheme.Card:Tint))using(Pen pen=new Pen(UiTheme.Border)){e.Graphics.FillPath(b,p);e.Graphics.DrawPath(pen,p);}
                int top=28,head=27,rowH=25;
                int[] widths={64,61,61,61,61,68};
                int[] xs=new int[7];
                for(int i=0;i<6;i++)xs[i+1]=xs[i]+widths[i];
                float scale=(Width-2)/(float)xs[6];
                for(int i=0;i<7;i++)xs[i]=(int)(xs[i]*scale)+1;
                using(SolidBrush hf=new SolidBrush(UiTheme.Header))e.Graphics.FillRectangle(hf,1,top,Width-2,head);
                string[] heads={"재원","건강","국민","고용","산재","합계"};
                using(Pen grid=new Pen(UiTheme.Grid))
                {
                    for(int i=0;i<6;i++)UiDrawing.Text(e.Graphics,heads[i],header,UiTheme.Text,new Rectangle(xs[i],top,xs[i+1]-xs[i],head),ContentAlignment.MiddleCenter);
                    List<DiscountAggregateRow> data=Rows??new List<DiscountAggregateRow>();
                    int count=Math.Min(5,data.Count);
                    for(int r=0;r<count;r++)
                    {
                        DiscountAggregateRow row=data[r];int y=top+head+r*rowH;
                        if(row.Fund=="합계")using(SolidBrush b=new SolidBrush(UiTheme.Surface))e.Graphics.FillRectangle(b,1,y,Width-2,rowH);
                        decimal[] values=Mode==0?row.Billed:Mode==1?row.Discount:row.After;
                        DrawFundFit(e.Graphics,row.Fund,UiTheme.Text,new Rectangle(xs[0]+1,y,xs[1]-xs[0]-2,rowH));
                        decimal sum=0;
                        for(int i=0;i<4;i++){sum+=values[i];DrawMoneyFit(e.Graphics,values[i],Mode==1&&values[i]>0?Color.FromArgb(242,129,76):UiTheme.Text,new Rectangle(xs[i+1]+1,y,xs[i+2]-xs[i+1]-2,rowH));}
                        DrawMoneyFit(e.Graphics,sum,UiTheme.Text,new Rectangle(xs[5]+1,y,xs[6]-xs[5]-2,rowH));
                        e.Graphics.DrawLine(grid,1,y+rowH,Width-1,y+rowH);
                    }
                    for(int i=1;i<6;i++)e.Graphics.DrawLine(grid,xs[i],top,xs[i],Math.Min(Height-1,top+head+rowH*count));
                    e.Graphics.DrawLine(grid,1,top+head,Width-1,top+head);
                }
            }
        }
    }

    class DiscountTableControl : Control
    {
        public List<IndividualRowData> Rows=new List<IndividualRowData>();public int ScrollOffset,PageSize=5;public Func<IndividualRowData,DiscountEntry> EntryProvider;public event Action<IndividualRowData,DiscountEntry> EntryChanged;public event Action<IndividualRowData,string,decimal> AmountEditRequested;int[] xs;int bodyTop=52,rowHeight=30;bool dragging;int dragY,dragOffset;
        public DiscountTableControl(){SetStyle(ControlStyles.UserPaint|ControlStyles.OptimizedDoubleBuffer|ControlStyles.AllPaintingInWmPaint|ControlStyles.Selectable,true);BackColor=Color.White;TabStop=true;}
        static string Mask(string value){string d=new string((value??"").Where(Char.IsDigit).ToArray());return d.Length>=7?d.Substring(0,6)+"-"+d.Substring(6,1)+"******":value??"";}
        protected override void OnPaint(PaintEventArgs e){e.Graphics.SmoothingMode=SmoothingMode.AntiAlias;e.Graphics.Clear(Parent==null?UiTheme.Page:Parent.BackColor);RectangleF outer=new RectangleF(.5F,.5F,Width-1.5F,Height-1.5F);using(GraphicsPath p=UiDrawing.Rounded(outer,9))using(SolidBrush b=new SolidBrush(UiTheme.Card))using(Pen pen=new Pen(UiTheme.Border)){e.Graphics.FillPath(b,p);e.Graphics.DrawPath(pen,p);}using(Font header=new Font("맑은 고딕",6.8F,FontStyle.Bold),cell=new Font("맑은 고딕",7F,FontStyle.Bold),small=new Font("맑은 고딕",6.5F,FontStyle.Bold))using(Pen grid=new Pen(UiTheme.Grid)){int contentRight=Width-14;int[] widths={36,82,105,78,88,88,95,95,95,95,76};xs=new int[widths.Length+1];for(int i=0;i<widths.Length;i++)xs[i+1]=xs[i]+widths[i];float scale=(contentRight-1)/(float)xs[xs.Length-1];for(int i=0;i<xs.Length;i++)xs[i]=(int)(xs[i]*scale)+1;using(SolidBrush hf=new SolidBrush(UiTheme.Header))e.Graphics.FillRectangle(hf,1,1,contentRight-1,bodyTop-1);string[] fixedHeads={"No.","이름","주민번호","재원"};for(int i=0;i<4;i++)UiDrawing.Text(e.Graphics,fixedHeads[i],header,UiTheme.Text,new Rectangle(xs[i],1,xs[i+1]-xs[i],bodyTop-1),ContentAlignment.MiddleCenter);UiDrawing.Text(e.Graphics,"자동이체 할인 (250원 고정)",header,UiTheme.Dark?Color.FromArgb(132,162,255):Color.FromArgb(39,83,196),new Rectangle(xs[4],1,xs[6]-xs[4],24),ContentAlignment.MiddleCenter);UiDrawing.Text(e.Graphics,"고용",header,UiTheme.Muted,new Rectangle(xs[4],25,xs[5]-xs[4],26),ContentAlignment.MiddleCenter);UiDrawing.Text(e.Graphics,"산재",header,UiTheme.Muted,new Rectangle(xs[5],25,xs[6]-xs[5],26),ContentAlignment.MiddleCenter);UiDrawing.Text(e.Graphics,"기타 할인 (금액 입력)",header,UiTheme.Dark?Color.FromArgb(188,156,255):Color.FromArgb(88,61,191),new Rectangle(xs[6],1,xs[10]-xs[6],24),ContentAlignment.MiddleCenter);string[] kinds={"건강","국민","고용","산재"};for(int i=0;i<4;i++)UiDrawing.Text(e.Graphics,kinds[i],header,UiTheme.Muted,new Rectangle(xs[6+i],25,xs[7+i]-xs[6+i],26),ContentAlignment.MiddleCenter);UiDrawing.Text(e.Graphics,"감면 총액",header,UiTheme.Text,new Rectangle(xs[10],1,xs[11]-xs[10],bodyTop-1),ContentAlignment.MiddleCenter);e.Graphics.DrawLine(grid,1,bodyTop,contentRight,bodyTop);SetOffset(ScrollOffset,false);List<IndividualRowData> visible=(Rows??new List<IndividualRowData>()).Skip(ScrollOffset).Take(PageSize).ToList();if(visible.Count==0)UiDrawing.Text(e.Graphics,"현재 조건에 해당하는 대상이 없습니다.",cell,UiTheme.Muted,new Rectangle(1,bodyTop,contentRight-1,Height-bodyTop),ContentAlignment.MiddleCenter);for(int r=0;r<visible.Count;r++){int y=bodyTop+r*rowHeight;IndividualRowData row=visible[r];DiscountEntry entry=EntryProvider==null?new DiscountEntry():EntryProvider(row);if(r%2==1)using(SolidBrush alt=new SolidBrush(UiTheme.Surface))e.Graphics.FillRectangle(alt,1,y,contentRight-1,rowHeight);DrawText(e.Graphics,(ScrollOffset+r+1).ToString(),cell,xs,0,y);DrawText(e.Graphics,row.Name,cell,xs,1,y);DrawText(e.Graphics,Mask(row.Birth),small,xs,2,y);DrawText(e.Graphics,row.Fund,small,xs,3,y);DrawToggle(e.Graphics,xs[4],xs[5],y,entry.AutoEmployment,entry.AutoEmployment?250m:0m,small);DrawToggle(e.Graphics,xs[5],xs[6],y,entry.AutoIndustrial,entry.AutoIndustrial?250m:0m,small);decimal[] other={entry.Health,entry.Pension,entry.Employment,entry.Industrial};for(int i=0;i<4;i++)DrawToggle(e.Graphics,xs[6+i],xs[7+i],y,other[i]>0,other[i],small);UiDrawing.Text(e.Graphics,UiDrawing.Money(entry.Total),cell,UiTheme.Accent,new Rectangle(xs[10],y,xs[11]-xs[10],rowHeight),ContentAlignment.MiddleCenter);e.Graphics.DrawLine(grid,1,y+rowHeight,contentRight,y+rowHeight);}int bottom=Math.Min(Height-1,bodyTop+rowHeight*Math.Max(1,visible.Count));for(int i=1;i<xs.Length-1;i++)e.Graphics.DrawLine(grid,xs[i],1,xs[i],bottom);DrawScroll(e.Graphics);}}
        static void DrawText(Graphics g,string text,Font f,int[] xs,int c,int y){UiDrawing.Text(g,text??"",f,UiTheme.Text,new Rectangle(xs[c]+2,y,xs[c+1]-xs[c]-4,30),ContentAlignment.MiddleCenter);}static void DrawToggle(Graphics g,int left,int right,int y,bool enabled,decimal amount,Font font){int w=right-left;RectangleF box=new RectangleF(left+5,y+5,w-10,20);using(GraphicsPath p=UiDrawing.Rounded(box,5))using(SolidBrush b=new SolidBrush(enabled?UiTheme.Surface:UiTheme.Input))using(Pen pen=new Pen(UiTheme.Border)){g.FillPath(b,p);g.DrawPath(pen,p);}RectangleF check=new RectangleF(left+10,y+9,10,10);using(GraphicsPath p=UiDrawing.Rounded(check,2))using(SolidBrush b=new SolidBrush(enabled?UiTheme.Accent:UiTheme.Input))using(Pen pen=new Pen(enabled?UiTheme.Accent:UiTheme.Border)){g.FillPath(b,p);g.DrawPath(pen,p);}if(enabled)using(Pen p=new Pen(Color.White,1.2F)){g.DrawLine(p,check.X+2,check.Y+5,check.X+4,check.Y+7);g.DrawLine(p,check.X+4,check.Y+7,check.X+8,check.Y+2);}if(amount>0)UiDrawing.Text(g,UiDrawing.Money(amount),font,UiTheme.Accent,new Rectangle(left+22,y,w-27,30),ContentAlignment.MiddleCenter);}
        int MaxOffset{get{return Math.Max(0,(Rows==null?0:Rows.Count)-PageSize);}}void SetOffset(int n,bool repaint=true){int v=Math.Max(0,Math.Min(MaxOffset,n));if(v==ScrollOffset)return;ScrollOffset=v;if(repaint)Invalidate();}Rectangle Track(){return new Rectangle(Width-10,bodyTop+4,5,Math.Max(20,Height-bodyTop-9));}Rectangle Thumb(){Rectangle t=Track();if(MaxOffset==0)return t;int h=Math.Max(25,(int)Math.Round(t.Height*Math.Min(1.0,PageSize/(double)Math.Max(PageSize,Rows.Count)))),travel=Math.Max(1,t.Height-h),y=t.Y+(int)Math.Round(travel*ScrollOffset/(double)MaxOffset);return new Rectangle(t.X-1,y,7,h);}void DrawScroll(Graphics g){Rectangle t=Track();using(SolidBrush b=new SolidBrush(UiTheme.Surface))g.FillRectangle(b,t);if(MaxOffset>0){Rectangle th=Thumb();using(GraphicsPath p=UiDrawing.Rounded(th,3))using(SolidBrush b=new SolidBrush(UiTheme.Dark?Color.FromArgb(112,125,153):Color.FromArgb(170,180,211)))g.FillPath(b,p);}}
        protected override void OnMouseEnter(EventArgs e){Focus();base.OnMouseEnter(e);}protected override void OnMouseWheel(MouseEventArgs e){SetOffset(ScrollOffset+(e.Delta<0?2:-2));base.OnMouseWheel(e);}protected override void OnMouseDown(MouseEventArgs e){if(e.X>=Width-16&&MaxOffset>0){Rectangle th=Thumb();if(th.Contains(e.Location)){dragging=true;dragY=e.Y;dragOffset=ScrollOffset;Capture=true;}else SetOffset((int)Math.Round((e.Y-Track().Y)/(double)Math.Max(1,Track().Height)*MaxOffset));}else if(xs!=null&&e.Y>=bodyTop&&e.Y<bodyTop+PageSize*rowHeight){int index=ScrollOffset+(e.Y-bodyTop)/rowHeight;if(Rows==null||index<0||index>=Rows.Count)return;IndividualRowData row=Rows[index];DiscountEntry entry=EntryProvider==null?new DiscountEntry():EntryProvider(row);if(e.X>=xs[4]&&e.X<xs[5]){entry.AutoEmployment=!entry.AutoEmployment;if(EntryChanged!=null)EntryChanged(row,entry);}else if(e.X>=xs[5]&&e.X<xs[6]){entry.AutoIndustrial=!entry.AutoIndustrial;if(EntryChanged!=null)EntryChanged(row,entry);}else for(int i=0;i<4;i++)if(e.X>=xs[6+i]&&e.X<xs[7+i]&&AmountEditRequested!=null){decimal current=i==0?entry.Health:i==1?entry.Pension:i==2?entry.Employment:entry.Industrial;AmountEditRequested(row,new[]{"건강","국민","고용","산재"}[i],current);break;}}base.OnMouseDown(e);}protected override void OnMouseMove(MouseEventArgs e){if(dragging){int travel=Math.Max(1,Track().Height-Thumb().Height);SetOffset(dragOffset+(int)Math.Round((e.Y-dragY)*MaxOffset/(double)travel));}base.OnMouseMove(e);}protected override void OnMouseUp(MouseEventArgs e){dragging=false;Capture=false;base.OnMouseUp(e);}
    }

    class DiscountAmountDialog : Form
    {
        readonly TextBox input;public decimal Amount{get;private set;}public DiscountAmountDialog(string kind,decimal current){Width=340;Height=178;FormBorderStyle=FormBorderStyle.None;StartPosition=FormStartPosition.CenterParent;ShowInTaskbar=false;BackColor=Color.White;Font=new Font("맑은 고딕",9F);Controls.Add(new Label{Text=kind+"보험 기타 감면액",Location=new Point(22,18),AutoSize=true,ForeColor=Color.FromArgb(30,43,91),Font=new Font("맑은 고딕",12F,FontStyle.Bold)});Controls.Add(new Label{Text="금액을 원 단위로 입력해 주세요.",Location=new Point(24,52),AutoSize=true,ForeColor=Color.FromArgb(99,108,140),Font=new Font("맑은 고딕",8F)});input=new TextBox{Location=new Point(24,79),Size=new Size(292,28),Text=current>0?Math.Round(current).ToString():"0",TextAlign=HorizontalAlignment.Right,Font=new Font("맑은 고딕",10F,FontStyle.Bold)};Controls.Add(input);var cancel=new Button{Text="취소",Location=new Point(154,125),Size=new Size(76,32),FlatStyle=FlatStyle.Flat,BackColor=Color.White,ForeColor=Color.FromArgb(53,67,127)};cancel.FlatAppearance.BorderColor=Color.FromArgb(214,221,240);cancel.Click+=(s,e)=>{DialogResult=DialogResult.Cancel;Close();};Controls.Add(cancel);var ok=new Button{Text="적용",Location=new Point(238,125),Size=new Size(78,32),FlatStyle=FlatStyle.Flat,BackColor=Color.FromArgb(55,88,225),ForeColor=Color.White};ok.FlatAppearance.BorderSize=0;ok.Click+=(s,e)=>ApplyValue();Controls.Add(ok);AcceptButton=ok;CancelButton=cancel;Resize+=(s,e)=>SetRound();SetRound();}
        void SetRound(){using(GraphicsPath p=UiDrawing.Rounded(new RectangleF(0,0,Width-1,Height-1),12))Region=new Region(p);}void ApplyValue(){decimal value;if(!Decimal.TryParse((input.Text??"").Replace(",",""),NumberStyles.Any,CultureInfo.InvariantCulture,out value)||value<0){MessageBox.Show("0 이상의 숫자를 입력해 주세요.","감면액 입력",MessageBoxButtons.OK,MessageBoxIcon.Information);return;}Amount=Math.Round(value);DialogResult=DialogResult.OK;Close();}protected override void OnPaint(PaintEventArgs e){e.Graphics.SmoothingMode=SmoothingMode.AntiAlias;using(GraphicsPath p=UiDrawing.Rounded(new RectangleF(.5F,.5F,Width-1.5F,Height-1.5F),12))using(Pen pen=new Pen(Color.FromArgb(207,216,238),1.2F))e.Graphics.DrawPath(pen,p);base.OnPaint(e);}
    }

    class ReviewTableControl : Control
    {
        public Func<IndividualRowData,IEnumerable<string>> FundChoices;
        public List<IndividualRowData> Rows=new List<IndividualRowData>();public int ScrollOffset,PageSize=5;public HashSet<string> SelectionKeys,CheckedKeys;public Dictionary<string,string> FundDrafts;public event Action SelectionChanged;public event Action<IndividualRowData,Point> DetailRequested;public event Action<IndividualRowData,string> FundChanged;int[] lastXs;int bodyTop=43,rowHeight=42;bool dragging;int dragY,dragOffset;
        public ReviewTableControl(){SetStyle(ControlStyles.UserPaint|ControlStyles.OptimizedDoubleBuffer|ControlStyles.AllPaintingInWmPaint|ControlStyles.Selectable,true);BackColor=Color.White;TabStop=true;}
        static string Key(IndividualRowData row){return String.Join("|",new[]{row.Site??"",row.Name??"",Regex.Replace(row.Birth??"","[^0-9]","")});}
        static string Mask(string value){string d=new string((value??"").Where(Char.IsDigit).ToArray());return d.Length>=7?d.Substring(0,6)+"-"+d.Substring(6,1)+"******":value??"";}
        static string Result(IndividualRowData row){bool up=row.HealthDifference>.5m||row.PensionDifference>.5m||row.EmploymentDifference>.5m||row.IndustrialDifference>.5m,down=row.HealthDifference<-.5m||row.PensionDifference<-.5m||row.EmploymentDifference<-.5m||row.IndustrialDifference<-.5m;if(row.Fund=="분류필요"||up&&down)return "분류필요";if(up)return "추징";if(down)return "환급";return "확인필요";}
        static string Reason(IndividualRowData row){if(!String.IsNullOrWhiteSpace(row.ReviewReason))return row.ReviewReason;if(row.Fund=="분류필요")return "급여대장 재원 분류가 필요합니다.";decimal[] a={Math.Abs(row.HealthDifference),Math.Abs(row.PensionDifference),Math.Abs(row.EmploymentDifference),Math.Abs(row.IndustrialDifference)};int best=0;for(int i=1;i<a.Length;i++)if(a[i]>a[best])best=i;decimal[] d={row.HealthDifference,row.PensionDifference,row.EmploymentDifference,row.IndustrialDifference};return new[]{"건강보험","국민연금","고용보험","산재보험"}[best]+(d[best]>0?" 고지금액 > 급여대장 금액 차이":d[best]<0?" 고지금액 < 급여대장 금액 차이":" 세부 내역 확인 필요");}
        protected override void OnPaint(PaintEventArgs e){e.Graphics.SmoothingMode=SmoothingMode.AntiAlias;e.Graphics.Clear(Parent==null?UiTheme.Page:Parent.BackColor);RectangleF outer=new RectangleF(.5F,.5F,Width-1.5F,Height-1.5F);using(GraphicsPath p=UiDrawing.Rounded(outer,10))using(SolidBrush b=new SolidBrush(UiTheme.Card))using(Pen pen=new Pen(UiTheme.Border)){e.Graphics.FillPath(b,p);e.Graphics.DrawPath(pen,p);}using(Font header=new Font("맑은 고딕",7.2F,FontStyle.Bold),cell=new Font("맑은 고딕",7.2F,FontStyle.Bold),small=new Font("맑은 고딕",6.8F,FontStyle.Bold))using(Pen grid=new Pen(UiTheme.Grid)){int contentRight=Width-14;int[] widths={30,38,82,105,250,125,98,105,74};int[] xs=new int[widths.Length+1];for(int i=0;i<widths.Length;i++)xs[i+1]=xs[i]+widths[i];float scale=(contentRight-1)/(float)xs[xs.Length-1];for(int i=0;i<xs.Length;i++)xs[i]=(int)(xs[i]*scale)+1;lastXs=xs;using(SolidBrush hf=new SolidBrush(UiTheme.Header))e.Graphics.FillRectangle(hf,1,1,contentRight-1,bodyTop-1);string[] heads={"□","No.","이름","주민번호","사유","재원 선택","대사 결과","조회 상태","상세"};for(int i=0;i<heads.Length;i++)UiDrawing.Text(e.Graphics,heads[i],header,UiTheme.Text,new Rectangle(xs[i],1,xs[i+1]-xs[i],bodyTop-1),ContentAlignment.MiddleCenter);e.Graphics.DrawLine(grid,1,bodyTop,contentRight,bodyTop);SetOffset(ScrollOffset,false);List<IndividualRowData> visible=(Rows??new List<IndividualRowData>()).Skip(ScrollOffset).Take(PageSize).ToList();if(visible.Count==0)UiDrawing.Text(e.Graphics,"현재 조건에 해당하는 확인 필요 대상이 없습니다.",cell,UiTheme.Muted,new Rectangle(1,bodyTop,contentRight-1,Height-bodyTop),ContentAlignment.MiddleCenter);for(int r=0;r<visible.Count;r++){int y=bodyTop+r*rowHeight;IndividualRowData row=visible[r];string key=Key(row),draft="";bool selected=SelectionKeys!=null&&SelectionKeys.Contains(key),checkedDone=CheckedKeys!=null&&CheckedKeys.Contains(key),hasDraft=FundDrafts!=null&&FundDrafts.TryGetValue(key,out draft);if(r%2==1)using(SolidBrush alt=new SolidBrush(UiTheme.Surface))e.Graphics.FillRectangle(alt,1,y,contentRight-1,rowHeight);DrawCheck(e.Graphics,new Rectangle(xs[0],y,xs[1]-xs[0],rowHeight),selected);DrawText(e.Graphics,(ScrollOffset+r+1).ToString(),cell,UiTheme.Text,xs,1,y);DrawText(e.Graphics,row.Name,cell,UiTheme.Dark?Color.FromArgb(174,194,255):Color.FromArgb(28,52,137),xs,2,y);DrawText(e.Graphics,Mask(row.Birth),small,UiTheme.Text,xs,3,y);DrawText(e.Graphics,Reason(row),small,UiTheme.Muted,xs,4,y);string fund=hasDraft?draft:row.Fund=="분류필요"?"선택 필요":row.Fund;DrawPill(e.Graphics,fund,xs[5],y,xs[6]-xs[5],rowHeight,hasDraft?Color.FromArgb(49,77,218):Color.FromArgb(54,72,142),hasDraft?Color.FromArgb(232,237,255):Color.FromArgb(247,249,254),small);string result=checkedDone?"정상":hasDraft?draft:Result(row);Color resultInk=result=="정상"?Color.FromArgb(22,139,69):result=="추징"?Color.FromArgb(226,56,56):result=="환급"?Color.FromArgb(42,103,222):result=="분류필요"?Color.FromArgb(98,65,220):Color.FromArgb(224,101,23);DrawPill(e.Graphics,result,xs[6],y,xs[7]-xs[6],rowHeight,resultInk,Color.FromArgb(245,246,255),small);DrawPill(e.Graphics,checkedDone?"확인 완료":"미확인",xs[7],y,xs[8]-xs[7],rowHeight,checkedDone?Color.FromArgb(22,139,69):Color.FromArgb(224,111,24),checkedDone?Color.FromArgb(235,248,238):Color.FromArgb(255,245,232),small);DrawPill(e.Graphics,"보기 ›",xs[8],y,xs[9]-xs[8],rowHeight,Color.FromArgb(41,85,216),Color.White,small,true);e.Graphics.DrawLine(grid,1,y+rowHeight,contentRight,y+rowHeight);}int bottom=Math.Min(Height-1,bodyTop+rowHeight*Math.Max(1,visible.Count));for(int i=1;i<xs.Length-1;i++)e.Graphics.DrawLine(grid,xs[i],1,xs[i],bottom);DrawScroll(e.Graphics);} }
        void DrawText(Graphics g,string text,Font f,Color color,int[] xs,int c,int y){UiDrawing.Text(g,text??"",f,color,new Rectangle(xs[c]+3,y,xs[c+1]-xs[c]-6,rowHeight),c==4?ContentAlignment.MiddleLeft:ContentAlignment.MiddleCenter);}static void DrawCheck(Graphics g,Rectangle cell,bool selected){RectangleF r=new RectangleF(cell.X+(cell.Width-14)/2F,cell.Y+(cell.Height-14)/2F,14,14);using(GraphicsPath p=UiDrawing.Rounded(r,3))using(SolidBrush b=new SolidBrush(selected?UiTheme.Accent:UiTheme.Input))using(Pen pen=new Pen(selected?UiTheme.Accent:UiTheme.Border)){g.FillPath(b,p);g.DrawPath(pen,p);}if(selected)using(Pen p=new Pen(Color.White,1.6F)){p.StartCap=LineCap.Round;p.EndCap=LineCap.Round;g.DrawLine(p,r.X+3,r.Y+7,r.X+6,r.Y+10);g.DrawLine(p,r.X+6,r.Y+10,r.X+11,r.Y+4);}}static void DrawPill(Graphics g,string text,int x,int y,int w,int h,Color ink,Color back,Font font,bool border=false){if(UiTheme.Dark){back=Color.FromArgb(Math.Min(78,34+ink.R/10),Math.Min(78,38+ink.G/10),Math.Min(88,46+ink.B/10));if((ink.R+ink.G+ink.B)/3<120)ink=Color.FromArgb(Math.Min(255,ink.R+105),Math.Min(255,ink.G+105),Math.Min(255,ink.B+80));}RectangleF r=new RectangleF(x+7,y+8,Math.Max(22,w-14),h-16);using(GraphicsPath p=UiDrawing.Rounded(r,6))using(SolidBrush b=new SolidBrush(back))using(Pen pen=new Pen(border?UiTheme.Border:back)){g.FillPath(b,p);g.DrawPath(pen,p);}UiDrawing.Text(g,text??"",font,ink,new Rectangle(x+2,y,w-4,h),ContentAlignment.MiddleCenter);}
        int MaxOffset{get{return Math.Max(0,(Rows==null?0:Rows.Count)-PageSize);}}void SetOffset(int n,bool repaint=true){int v=Math.Max(0,Math.Min(MaxOffset,n));if(v==ScrollOffset)return;ScrollOffset=v;if(repaint)Invalidate();}Rectangle Track(){return new Rectangle(Width-10,bodyTop+4,5,Math.Max(20,Height-bodyTop-9));}Rectangle Thumb(){Rectangle t=Track();if(MaxOffset==0)return t;int h=Math.Max(28,(int)Math.Round(t.Height*Math.Min(1.0,PageSize/(double)Math.Max(PageSize,Rows.Count)))),travel=Math.Max(1,t.Height-h),y=t.Y+(int)Math.Round(travel*ScrollOffset/(double)MaxOffset);return new Rectangle(t.X-1,y,7,h);}void DrawScroll(Graphics g){Rectangle t=Track();using(SolidBrush b=new SolidBrush(UiTheme.Surface))g.FillRectangle(b,t);if(MaxOffset>0){Rectangle th=Thumb();using(GraphicsPath p=UiDrawing.Rounded(th,3))using(SolidBrush b=new SolidBrush(UiTheme.Dark?Color.FromArgb(112,125,153):Color.FromArgb(170,180,211)))g.FillPath(b,p);}}
        void ShowFundMenu(IndividualRowData row,Point at){var menu=new ContextMenuStrip{ShowImageMargin=false,ShowCheckMargin=false,BackColor=UiTheme.Card,Padding=new Padding(5),Renderer=new SiteMenuRenderer()};foreach(string fund in (FundChoices==null?new[]{"공무원","계약제교원","교특회계","학교회계"}:FundChoices(row))){string choice=fund;var item=new ToolStripMenuItem{Text=choice,AutoSize=false,Size=new Size(160,34),Font=new Font("맑은 고딕",8.5F),ForeColor=UiTheme.Text};item.Click+=(s,e)=>{if(FundChanged!=null)FundChanged(row,choice);};menu.Items.Add(item);}menu.Show(this,at);}
        public void ToggleAllSelection(){if(SelectionKeys==null)return;bool all=Rows!=null&&Rows.Count>0&&Rows.All(x=>SelectionKeys.Contains(Key(x)));if(Rows!=null)foreach(IndividualRowData row in Rows){string key=Key(row);if(all)SelectionKeys.Remove(key);else SelectionKeys.Add(key);}Invalidate();if(SelectionChanged!=null)SelectionChanged();}
        protected override void OnMouseEnter(EventArgs e){Focus();base.OnMouseEnter(e);}protected override void OnMouseWheel(MouseEventArgs e){SetOffset(ScrollOffset+(e.Delta<0?2:-2));base.OnMouseWheel(e);}protected override void OnMouseDown(MouseEventArgs e){if(lastXs!=null&&e.Y<bodyTop&&e.X>=lastXs[0]&&e.X<lastXs[1]&&SelectionKeys!=null){ToggleAllSelection();}else if(e.X>=Width-16&&MaxOffset>0){Rectangle th=Thumb();if(th.Contains(e.Location)){dragging=true;dragY=e.Y;dragOffset=ScrollOffset;Capture=true;}else SetOffset((int)Math.Round((e.Y-Track().Y)/(double)Math.Max(1,Track().Height)*MaxOffset));}else if(lastXs!=null&&e.Y>=bodyTop&&e.Y<bodyTop+PageSize*rowHeight){int index=ScrollOffset+(e.Y-bodyTop)/rowHeight;if(Rows==null||index<0||index>=Rows.Count)return;IndividualRowData row=Rows[index];if(e.X>=lastXs[0]&&e.X<lastXs[1]&&SelectionKeys!=null){string key=Key(row);if(!SelectionKeys.Add(key))SelectionKeys.Remove(key);Invalidate();if(SelectionChanged!=null)SelectionChanged();}else if(e.X>=lastXs[5]&&e.X<lastXs[6])ShowFundMenu(row,new Point(lastXs[5],e.Y+12));else if(((e.X>=lastXs[2]&&e.X<lastXs[3])||(e.X>=lastXs[8]&&e.X<lastXs[9])||e.Clicks==2)&&DetailRequested!=null)DetailRequested(row,PointToScreen(e.Location));}base.OnMouseDown(e);}protected override void OnMouseMove(MouseEventArgs e){if(dragging){int travel=Math.Max(1,Track().Height-Thumb().Height);SetOffset(dragOffset+(int)Math.Round((e.Y-dragY)*MaxOffset/(double)travel));}base.OnMouseMove(e);}protected override void OnMouseUp(MouseEventArgs e){dragging=false;Capture=false;base.OnMouseUp(e);}
    }

    class ReviewDetailBubble : Form
    {
        readonly IndividualRowData row;readonly string fund,reason;readonly bool checkedDone;readonly Rectangle closeButton=new Rectangle(258,548,64,34);protected override CreateParams CreateParams{get{CreateParams cp=base.CreateParams;cp.ClassStyle|=0x00020000;return cp;}}
        public ReviewDetailBubble(IndividualRowData item,string appliedFund,string reviewReason,bool done){row=item;fund=appliedFund=="분류필요"?"분류 필요":appliedFund;reason=reviewReason;checkedDone=done;Width=340;Height=596;FormBorderStyle=FormBorderStyle.None;ShowInTaskbar=false;BackColor=UiTheme.Card;DoubleBuffered=true;TopMost=true;Font=new Font("맑은 고딕",8F);Resize+=(s,e)=>SetBubbleRegion();SetBubbleRegion();}
        void SetBubbleRegion(){using(GraphicsPath p=BubblePath())Region=new Region(p);}GraphicsPath BubblePath(){GraphicsPath p=UiDrawing.Rounded(new RectangleF(12,.5F,Width-13,Height-1.5F),13);p.AddPolygon(new[]{new PointF(12,155),new PointF(1,165),new PointF(12,175)});return p;}static string Mask(string value){string d=new string((value??"").Where(Char.IsDigit).ToArray());return d.Length>=7?d.Substring(0,6)+"-"+d.Substring(6,1)+"******":value??"";}static string Money(decimal v){return Math.Round(v).ToString("#,##0");}static string Diff(decimal v){return Math.Abs(v)<=.5m?"0":v>0?"+"+Money(v):Money(v);}static string Result(IndividualRowData r){bool up=r.HealthDifference>.5m||r.PensionDifference>.5m||r.EmploymentDifference>.5m||r.IndustrialDifference>.5m,down=r.HealthDifference<-.5m||r.PensionDifference<-.5m||r.EmploymentDifference<-.5m||r.IndustrialDifference<-.5m;return r.Fund=="분류필요"||up&&down?"분류필요":up?"추징":down?"환급":"확인필요";}
        protected override void OnPaint(PaintEventArgs e){e.Graphics.SmoothingMode=SmoothingMode.AntiAlias;using(GraphicsPath p=BubblePath())using(SolidBrush b=new SolidBrush(UiTheme.Card))using(Pen pen=new Pen(UiTheme.Border,1.2F)){e.Graphics.FillPath(b,p);e.Graphics.DrawPath(pen,p);}using(Font title=new Font("맑은 고딕",13F,FontStyle.Bold),section=new Font("맑은 고딕",8.3F,FontStyle.Bold),label=new Font("맑은 고딕",7.5F,FontStyle.Bold),cell=new Font("맑은 고딕",7.3F,FontStyle.Bold)){UiDrawing.Text(e.Graphics,"상세 보기",title,UiTheme.Text,new Rectangle(30,14,220,32),ContentAlignment.MiddleLeft);UiDrawing.Text(e.Graphics,"×",title,UiTheme.Text,new Rectangle(294,12,28,32),ContentAlignment.MiddleCenter);UiDrawing.Text(e.Graphics,"기본 정보",section,UiTheme.Dark?Color.FromArgb(160,182,255):Color.FromArgb(44,66,157),new Rectangle(30,54,200,22),ContentAlignment.MiddleLeft);DrawInfo(e.Graphics,label,cell,80,"이름",row.Name);DrawInfo(e.Graphics,label,cell,114,"주민번호",Mask(row.Birth));DrawInfo(e.Graphics,label,cell,148,"재원",fund);DrawInfo(e.Graphics,label,cell,182,"사유",reason);DrawInfo(e.Graphics,label,cell,230,"대사 결과",checkedDone?"정상":Result(row));DrawInfo(e.Graphics,label,cell,264,"조회 상태",checkedDone?"확인 완료":"미확인");string insurance=Primary();UiDrawing.Text(e.Graphics,insurance+(insurance=="건강보험"?" (건강+장기요양)":""),section,UiTheme.Dark?Color.FromArgb(160,182,255):Color.FromArgb(44,66,157),new Rectangle(30,308,250,22),ContentAlignment.MiddleLeft);UiDrawing.Text(e.Graphics,"금액 단위 : 원",label,UiTheme.Muted,new Rectangle(220,308,95,22),ContentAlignment.MiddleRight);decimal notice=insurance=="건강보험"?row.HealthNotice:insurance=="국민연금"?row.PensionNotice:insurance=="고용보험"?row.EmploymentNotice:row.IndustrialNotice,payroll=insurance=="건강보험"?row.HealthPayroll:insurance=="국민연금"?row.PensionPayroll:insurance=="고용보험"?row.EmploymentPayroll:row.IndustrialPayroll,difference=insurance=="건강보험"?row.HealthDifference:insurance=="국민연금"?row.PensionDifference:insurance=="고용보험"?row.EmploymentDifference:row.IndustrialDifference;DrawCompare(e.Graphics,label,cell,338,"합계",payroll,notice,difference);UiDrawing.Text(e.Graphics,"보험별 상세",section,UiTheme.Dark?Color.FromArgb(160,182,255):Color.FromArgb(44,66,157),new Rectangle(30,404,200,22),ContentAlignment.MiddleLeft);DrawDetail(e.Graphics,label,cell,432,"건강",row.HealthPayroll,row.HealthNotice,row.HealthDifference);DrawDetail(e.Graphics,label,cell,458,"국민",row.PensionPayroll,row.PensionNotice,row.PensionDifference);DrawDetail(e.Graphics,label,cell,484,"고용",row.EmploymentPayroll,row.EmploymentNotice,row.EmploymentDifference);DrawDetail(e.Graphics,label,cell,510,"산재",row.IndustrialPayroll,row.IndustrialNotice,row.IndustrialDifference);using(GraphicsPath p=UiDrawing.Rounded(closeButton,8))using(SolidBrush b=new SolidBrush(UiTheme.Surface))using(Pen pen=new Pen(UiTheme.Border)){e.Graphics.FillPath(b,p);e.Graphics.DrawPath(pen,p);}UiDrawing.Text(e.Graphics,"닫기",section,UiTheme.Text,closeButton,ContentAlignment.MiddleCenter);}}
        string Primary(){decimal[] a={Math.Abs(row.HealthDifference),Math.Abs(row.PensionDifference),Math.Abs(row.EmploymentDifference),Math.Abs(row.IndustrialDifference)};int b=0;for(int i=1;i<a.Length;i++)if(a[i]>a[b])b=i;return new[]{"건강보험","국민연금","고용보험","산재보험"}[b];}void DrawInfo(Graphics g,Font label,Font cell,int y,string name,string value){Rectangle r=new Rectangle(30,y,292,name=="사유"?44:30);using(SolidBrush b=new SolidBrush(UiTheme.Surface))g.FillRectangle(b,r);using(Pen p=new Pen(UiTheme.Grid))g.DrawRectangle(p,r);UiDrawing.Text(g,name,label,UiTheme.Muted,new Rectangle(r.X+8,r.Y,72,r.Height),ContentAlignment.MiddleLeft);UiDrawing.Text(g,value??"",cell,UiTheme.Text,new Rectangle(r.X+82,r.Y,r.Width-88,r.Height),ContentAlignment.MiddleLeft);}void DrawCompare(Graphics g,Font label,Font cell,int y,string kind,decimal payroll,decimal notice,decimal diff){string[] h={"구분","급여대장","고지금액","차액"},v={kind,Money(payroll),Money(notice),Diff(diff)};int w=73;for(int i=0;i<4;i++){Rectangle hr=new Rectangle(30+i*w,y,w,25),vr=new Rectangle(30+i*w,y+25,w,31);using(SolidBrush b=new SolidBrush(UiTheme.Header))g.FillRectangle(b,hr);using(Pen p=new Pen(UiTheme.Grid)){g.DrawRectangle(p,hr);g.DrawRectangle(p,vr);}UiDrawing.Text(g,h[i],label,UiTheme.Muted,hr,ContentAlignment.MiddleCenter);UiDrawing.Text(g,v[i],cell,i==3?(diff>0?Color.Red:diff<0?Color.RoyalBlue:Color.Green):UiTheme.Text,vr,ContentAlignment.MiddleCenter);}}void DrawDetail(Graphics g,Font label,Font cell,int y,string kind,decimal payroll,decimal notice,decimal diff){string[] v={kind,Money(payroll),Money(notice),Diff(diff)};int w=73;for(int i=0;i<4;i++){Rectangle r=new Rectangle(30+i*w,y,w,26);using(Pen p=new Pen(UiTheme.Grid))g.DrawRectangle(p,r);UiDrawing.Text(g,v[i],i==0?label:cell,i==3?(diff>0?Color.Red:diff<0?Color.RoyalBlue:Color.Green):UiTheme.Text,r,ContentAlignment.MiddleCenter);}}protected override void OnMouseDown(MouseEventArgs e){if(new Rectangle(294,12,28,32).Contains(e.Location)||closeButton.Contains(e.Location))Close();base.OnMouseDown(e);}
    }

    class ThemeChoiceButton : Button
    {
        public string ThemeName="",Description="";public Color Accent=Color.RoyalBlue;public bool DarkPreview,Active;bool hovered;
        public ThemeChoiceButton(){FlatStyle=FlatStyle.Flat;FlatAppearance.BorderSize=0;Cursor=Cursors.Hand;TabStop=false;SetStyle(ControlStyles.UserPaint|ControlStyles.OptimizedDoubleBuffer|ControlStyles.AllPaintingInWmPaint,true);}
        protected override void OnMouseEnter(EventArgs e){hovered=true;Invalidate();base.OnMouseEnter(e);}protected override void OnMouseLeave(EventArgs e){hovered=false;Invalidate();base.OnMouseLeave(e);}
        protected override void OnPaint(PaintEventArgs e){e.Graphics.SmoothingMode=SmoothingMode.AntiAlias;e.Graphics.Clear(Parent==null?UiTheme.Page:Parent.BackColor);bool gray=UiTheme.Name=="회색";Color preview=gray?(DarkPreview?Color.FromArgb(65,69,76):Active?Color.FromArgb(94,100,110):Color.FromArgb(84,89,98)):UiTheme.Dark?(DarkPreview?Color.FromArgb(24,29,40):Color.FromArgb(35,42,57)):(DarkPreview?Color.FromArgb(25,30,42):Color.White),border=Active?Accent:hovered?Color.FromArgb(145,155,181):UiTheme.Border;RectangleF outer=new RectangleF(1,1,Width-3,Height-3);using(GraphicsPath p=UiDrawing.Rounded(outer,11))using(SolidBrush b=new SolidBrush(preview))using(Pen pen=new Pen(border,Active?2F:1F)){e.Graphics.FillPath(b,p);e.Graphics.DrawPath(pen,p);}using(SolidBrush b=new SolidBrush(Accent))e.Graphics.FillEllipse(b,15,17,24,24);using(Pen p=new Pen(gray||UiTheme.Dark||DarkPreview?Color.FromArgb(235,239,247):Color.FromArgb(42,51,83),2F)){p.StartCap=LineCap.Round;p.EndCap=LineCap.Round;e.Graphics.DrawLine(p,20,28,26,34);e.Graphics.DrawLine(p,26,34,35,23);}Color text=gray?UiTheme.Text:UiTheme.Dark||DarkPreview?Color.FromArgb(239,243,253):Color.FromArgb(32,43,83),muted=gray?UiTheme.Muted:UiTheme.Dark||DarkPreview?Color.FromArgb(175,185,208):Color.FromArgb(103,112,142);using(Font title=new Font("맑은 고딕",10F,FontStyle.Bold),note=new Font("맑은 고딕",7.2F)){UiDrawing.Text(e.Graphics,ThemeName,title,text,new Rectangle(50,11,Width-65,35),ContentAlignment.MiddleLeft);UiDrawing.Text(e.Graphics,Description,note,muted,new Rectangle(15,54,Width-30,42),ContentAlignment.MiddleLeft);}if(Active)using(SolidBrush b=new SolidBrush(Accent))e.Graphics.FillEllipse(b,Width-27,14,12,12);}
    }

    class IndividualModeTabButton : Control
    {
        public string Caption="";public Color Accent=Color.RoyalBlue;public bool Active;bool hovered;
        public IndividualModeTabButton(){Cursor=Cursors.Hand;SetStyle(ControlStyles.UserPaint|ControlStyles.OptimizedDoubleBuffer|ControlStyles.AllPaintingInWmPaint,true);}protected override void OnMouseEnter(EventArgs e){hovered=true;Invalidate();base.OnMouseEnter(e);}protected override void OnMouseLeave(EventArgs e){hovered=false;Invalidate();base.OnMouseLeave(e);}protected override void OnMouseUp(MouseEventArgs e){if(e.Button==MouseButtons.Left)OnClick(EventArgs.Empty);base.OnMouseUp(e);}protected override void OnPaint(PaintEventArgs e){e.Graphics.SmoothingMode=SmoothingMode.AntiAlias;e.Graphics.Clear(Parent==null?UiTheme.Page:Parent.BackColor);RectangleF r=new RectangleF(1,1,Width-3,Height-2);Color fill=Active?(UiTheme.Dark?Color.FromArgb(45,55,76):Color.FromArgb(244,247,255)):hovered?UiTheme.Surface:UiTheme.Card;using(GraphicsPath p=UiDrawing.Rounded(r,9))using(SolidBrush b=new SolidBrush(fill))using(Pen pen=new Pen(Active?UiTheme.Accent:UiTheme.Border,Active?1.6F:1F)){e.Graphics.FillPath(b,p);e.Graphics.DrawPath(pen,p);}using(Font font=new Font("맑은 고딕",8.5F,FontStyle.Bold))UiDrawing.Text(e.Graphics,Caption,font,Active?UiTheme.Accent:UiTheme.Text,new Rectangle(0,0,Width,Height),ContentAlignment.MiddleCenter);}
    }

    class AdjustmentTabButton : Control
    {
        public string Caption="";public int Count;public Color Accent=Color.RoyalBlue;public bool Active;bool hovered;
        public AdjustmentTabButton(){Cursor=Cursors.Hand;SetStyle(ControlStyles.UserPaint|ControlStyles.OptimizedDoubleBuffer|ControlStyles.AllPaintingInWmPaint,true);}
        protected override void OnMouseEnter(EventArgs e){hovered=true;Invalidate();base.OnMouseEnter(e);}protected override void OnMouseLeave(EventArgs e){hovered=false;Invalidate();base.OnMouseLeave(e);}protected override void OnMouseUp(MouseEventArgs e){if(e.Button==MouseButtons.Left)OnClick(EventArgs.Empty);base.OnMouseUp(e);}
        protected override void OnPaint(PaintEventArgs e){e.Graphics.SmoothingMode=SmoothingMode.AntiAlias;e.Graphics.Clear(Parent==null?UiTheme.Page:Parent.BackColor);RectangleF r=new RectangleF(1,1,Width-3,Height-2);Color fill=Active?(UiTheme.Dark?Color.FromArgb(45,55,76):Color.FromArgb(247,249,255)):hovered?UiTheme.Surface:UiTheme.Card;using(GraphicsPath p=UiDrawing.Rounded(r,9))using(SolidBrush b=new SolidBrush(fill))using(Pen pen=new Pen(Active?Accent:UiTheme.Border,Active?1.5F:1F)){e.Graphics.FillPath(b,p);e.Graphics.DrawPath(pen,p);}using(Font font=new Font("맑은 고딕",8.5F,FontStyle.Bold))UiDrawing.Text(e.Graphics,Caption+" ("+Count+")",font,Active?Accent:UiTheme.Text,new Rectangle(0,0,Width,Height),ContentAlignment.MiddleCenter);}
    }

    class AdjustmentTableControl : Control
    {
        public List<IndividualRowData> Rows=new List<IndividualRowData>();public string Mode="전체",PreviewFund="";public int ScrollOffset,PageSize=4;public HashSet<string> SelectionKeys;public event Action SelectionChanged;bool draggingScroll;int dragStartY,dragStartOffset;int[] lastXs;int bodyTop=106,rowHeight=35;
        public AdjustmentTableControl(){SetStyle(ControlStyles.UserPaint|ControlStyles.OptimizedDoubleBuffer|ControlStyles.AllPaintingInWmPaint|ControlStyles.Selectable,true);BackColor=Color.White;TabStop=true;}
        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.SmoothingMode=SmoothingMode.AntiAlias;e.Graphics.Clear(Parent==null?UiTheme.Page:Parent.BackColor);RectangleF outer=new RectangleF(.5F,.5F,Width-1.5F,Height-1.5F);using(GraphicsPath path=UiDrawing.Rounded(outer,11))using(SolidBrush fill=new SolidBrush(UiTheme.Card))using(Pen border=new Pen(UiTheme.Border)){e.Graphics.FillPath(fill,path);e.Graphics.DrawPath(border,path);}using(Font title=new Font("맑은 고딕",8.7F,FontStyle.Bold),notice=new Font("맑은 고딕",7.1F),header=new Font("맑은 고딕",6.8F,FontStyle.Bold),cell=new Font("맑은 고딕",7F,FontStyle.Bold),small=new Font("맑은 고딕",6.5F,FontStyle.Bold))
            {
                string titleText=Mode=="전체"?"반환·추징·분류 필요 대상 목록":Mode+" 대상 목록";UiDrawing.Text(e.Graphics,titleText,title,UiTheme.Text,new Rectangle(14,2,500,25),ContentAlignment.MiddleLeft);UiDrawing.Text(e.Graphics,"금액 단위 : 원",header,UiTheme.Muted,new Rectangle(860,2,155,25),ContentAlignment.MiddleRight);Rectangle noticeRect=new Rectangle(10,27,Width-20,25);using(GraphicsPath p=UiDrawing.Rounded(noticeRect,6))using(SolidBrush b=new SolidBrush(UiTheme.Surface))e.Graphics.FillPath(b,p);string note=Mode=="반환"?"ⓘ 급여 공제액이 보험료 고지액보다 큰 항목입니다.":Mode=="추징"?"ⓘ 보험료 고지액이 급여 공제액보다 큰 항목입니다.":Mode=="분류 필요"?"ⓘ 자동 분류가 어렵거나 보험별 방향이 혼재된 항목입니다.":"ⓘ 반환·추징·분류 필요 관리 대상만 표시하며 정상 대상은 제외합니다.";UiDrawing.Text(e.Graphics,note,notice,UiTheme.Dark?Color.FromArgb(150,178,255):Color.FromArgb(54,91,188),new Rectangle(18,27,Width-36,25),ContentAlignment.MiddleLeft);
                int top=54,head1=28,head2=24,contentRight=Width-14;bodyTop=top+head1+head2;rowHeight=35;int[] widths={26,30,56,55,70,79,61,45,45,45,45,45,45,45,45,45,45,45,45,72};int[] xs=new int[widths.Length+1];for(int i=0;i<widths.Length;i++)xs[i+1]=xs[i]+widths[i];float scale=(contentRight-1)/(float)xs[xs.Length-1];for(int i=0;i<xs.Length;i++)xs[i]=(int)(xs[i]*scale)+1;lastXs=xs;
                using(SolidBrush hf=new SolidBrush(UiTheme.Header))e.Graphics.FillRectangle(hf,1,top,contentRight-1,head1+head2);using(Pen grid=new Pen(UiTheme.Grid))
                {
                    string[] fixedHeads={"","No.","재원","이름","주민/사번","직종명","대사결과"};for(int i=0;i<fixedHeads.Length;i++)UiDrawing.Text(e.Graphics,fixedHeads[i],header,UiTheme.Text,new Rectangle(xs[i],top,xs[i+1]-xs[i],head1+head2),ContentAlignment.MiddleCenter);string[] groups={"건강보험","국민연금","고용보험","산재보험"};for(int g=0;g<4;g++){int c=7+g*3;UiDrawing.Text(e.Graphics,groups[g],header,UiTheme.Text,new Rectangle(xs[c],top,xs[c+3]-xs[c],head1),ContentAlignment.MiddleCenter);string[] subs={"고지","급여","차액"};for(int s=0;s<3;s++)UiDrawing.Text(e.Graphics,subs[s],small,UiTheme.Muted,new Rectangle(xs[c+s],top+head1,xs[c+s+1]-xs[c+s],head2),ContentAlignment.MiddleCenter);e.Graphics.DrawLine(grid,xs[c+1],top+head1,xs[c+1],bodyTop);e.Graphics.DrawLine(grid,xs[c+2],top+head1,xs[c+2],bodyTop);}UiDrawing.Text(e.Graphics,Mode=="반환"?"환급액":Mode=="추징"?"추징액":"조정금액",header,UiTheme.Text,new Rectangle(xs[19],top,xs[20]-xs[19],head1+head2),ContentAlignment.MiddleCenter);e.Graphics.DrawLine(grid,1,top+head1,contentRight,top+head1);e.Graphics.DrawLine(grid,1,bodyTop,contentRight,bodyTop);
                    SetScrollOffset(ScrollOffset,false);List<IndividualRowData> visible=(Rows??new List<IndividualRowData>()).Skip(ScrollOffset).Take(PageSize).ToList();if(visible.Count==0)UiDrawing.Text(e.Graphics,"현재 조건에 해당하는 대상이 없습니다.",cell,UiTheme.Muted,new Rectangle(1,bodyTop,contentRight-1,Height-bodyTop),ContentAlignment.MiddleCenter);for(int r=0;r<visible.Count;r++){int y=bodyTop+r*rowHeight;IndividualRowData d=visible[r];bool staged=Mode=="분류 필요"&&SelectionKeys!=null&&SelectionKeys.Contains(Key(d))&&!String.IsNullOrWhiteSpace(PreviewFund);string displayFund=staged?PreviewFund:d.Fund;string kind=staged?PreviewFund:Kind(d,Mode);if(r%2==1)using(SolidBrush alt=new SolidBrush(UiTheme.Surface))e.Graphics.FillRectangle(alt,1,y,contentRight-1,rowHeight);DrawCheck(e.Graphics,new Rectangle(xs[0],y,xs[1]-xs[0],rowHeight),SelectionKeys!=null&&SelectionKeys.Contains(Key(d)));DrawCell(e.Graphics,""+(ScrollOffset+r+1),cell,UiTheme.Text,xs,1,y,rowHeight);DrawCell(e.Graphics,displayFund,small,staged?UiTheme.Accent:UiTheme.Text,xs,2,y,rowHeight);DrawCell(e.Graphics,d.Name,cell,UiTheme.Dark?Color.FromArgb(174,194,255):Color.FromArgb(28,52,137),xs,3,y,rowHeight);DrawCell(e.Graphics,MaskBirth(d.Birth),small,UiTheme.Text,xs,4,y,rowHeight);DrawCell(e.Graphics,d.Job,small,UiTheme.Muted,xs,5,y,rowHeight);DrawStatus(e.Graphics,kind,xs[6],y,xs[7]-xs[6],rowHeight,small);decimal[] amounts={d.HealthNotice,d.HealthPayroll,d.HealthDifference,d.PensionNotice,d.PensionPayroll,d.PensionDifference,d.EmploymentNotice,d.EmploymentPayroll,d.EmploymentDifference,d.IndustrialNotice,d.IndustrialPayroll,d.IndustrialDifference};for(int c=0;c<amounts.Length;c++){bool diff=c%3==2;Color ink=diff?UiDrawing.StatusColor(amounts[c]):UiTheme.Text;DrawCell(e.Graphics,diff?Difference(amounts[c]):UiDrawing.Money(amounts[c]),small,ink,xs,7+c,y,rowHeight);}decimal action=Amount(d,Mode);DrawCell(e.Graphics,UiDrawing.Money(action),cell,kind=="반환"?Color.FromArgb(43,102,224):kind=="추징"?Color.FromArgb(239,63,63):staged?Color.FromArgb(49,77,218):Color.FromArgb(226,112,30),xs,19,y,rowHeight);e.Graphics.DrawLine(grid,1,y+rowHeight,contentRight,y+rowHeight);}int lineBottom=Math.Min(Height-1,bodyTop+rowHeight*Math.Max(visible.Count,1));for(int c=1;c<xs.Length-1;c++)e.Graphics.DrawLine(grid,xs[c],top,xs[c],lineBottom);DrawScrollBar(e.Graphics,bodyTop);
                }
            }
        }
        int MaxOffset{get{return Math.Max(0,(Rows==null?0:Rows.Count)-PageSize);}}void SetScrollOffset(int value,bool repaint=true){int next=Math.Max(0,Math.Min(MaxOffset,value));if(next==ScrollOffset)return;ScrollOffset=next;if(repaint)Invalidate();}Rectangle ScrollTrack(){return new Rectangle(Width-10,bodyTop+4,5,Math.Max(20,Height-bodyTop-9));}Rectangle ScrollThumb(){Rectangle track=ScrollTrack();if(MaxOffset==0)return track;int thumbH=Math.Max(25,(int)Math.Round(track.Height*Math.Min(1.0,PageSize/(double)Math.Max(PageSize,Rows.Count)))),travel=Math.Max(1,track.Height-thumbH),y=track.Y+(int)Math.Round(travel*ScrollOffset/(double)MaxOffset);return new Rectangle(track.X-1,y,7,thumbH);}void DrawScrollBar(Graphics g,int top){Rectangle track=ScrollTrack();using(SolidBrush b=new SolidBrush(UiTheme.Surface))g.FillRectangle(b,track);if(MaxOffset>0){Rectangle thumb=ScrollThumb();using(GraphicsPath p=UiDrawing.Rounded(thumb,3))using(SolidBrush b=new SolidBrush(UiTheme.Dark?Color.FromArgb(112,125,153):Color.FromArgb(170,180,211)))g.FillPath(b,p);}}
        protected override void OnMouseEnter(EventArgs e){Focus();base.OnMouseEnter(e);}protected override void OnMouseWheel(MouseEventArgs e){SetScrollOffset(ScrollOffset+(e.Delta<0?2:-2));base.OnMouseWheel(e);}protected override void OnMouseDown(MouseEventArgs e){if(e.X>=Width-16&&MaxOffset>0){Rectangle thumb=ScrollThumb();if(thumb.Contains(e.Location)){draggingScroll=true;dragStartY=e.Y;dragStartOffset=ScrollOffset;Capture=true;}else{Rectangle track=ScrollTrack();SetScrollOffset((int)Math.Round((e.Y-track.Y)/(double)Math.Max(1,track.Height)*MaxOffset));}}else if(lastXs!=null&&e.Y>=bodyTop&&e.Y<bodyTop+PageSize*rowHeight&&e.X>=lastXs[0]&&e.X<lastXs[1]){int index=ScrollOffset+(e.Y-bodyTop)/rowHeight;if(Rows!=null&&index>=0&&index<Rows.Count&&SelectionKeys!=null){string key=Key(Rows[index]);if(!SelectionKeys.Add(key))SelectionKeys.Remove(key);Invalidate();if(SelectionChanged!=null)SelectionChanged();}}base.OnMouseDown(e);}protected override void OnMouseMove(MouseEventArgs e){if(draggingScroll){Rectangle track=ScrollTrack(),thumb=ScrollThumb();int travel=Math.Max(1,track.Height-thumb.Height);SetScrollOffset(dragStartOffset+(int)Math.Round((e.Y-dragStartY)*MaxOffset/(double)travel));}base.OnMouseMove(e);}protected override void OnMouseUp(MouseEventArgs e){draggingScroll=false;Capture=false;base.OnMouseUp(e);}
        static string Key(IndividualRowData row){return String.Join("|",new[]{row.Site??"",row.Fund??"",row.Name??"",Regex.Replace(row.Birth??"","[^0-9]","")});}static bool Refund(IndividualRowData d){return d.Fund!="분류필요"&&(d.HealthDifference<-.5m||d.PensionDifference<-.5m||d.EmploymentDifference<-.5m||d.IndustrialDifference<-.5m);}static bool Collection(IndividualRowData d){return d.Fund!="분류필요"&&(d.HealthDifference>.5m||d.PensionDifference>.5m||d.EmploymentDifference>.5m||d.IndustrialDifference>.5m);}static bool Classification(IndividualRowData d){return d.Fund=="분류필요";}static string Kind(IndividualRowData d,string mode){if(mode!="전체")return mode;bool refund=Refund(d),collection=Collection(d);return refund&&collection?"혼재":refund?"반환":collection?"추징":Classification(d)?"분류 필요":"정상";}static decimal Amount(IndividualRowData d,string mode){decimal[] values={d.HealthDifference,d.PensionDifference,d.EmploymentDifference,d.IndustrialDifference};if(mode=="반환")return values.Where(x=>x<-.5m).Sum(x=>Math.Abs(x));if(mode=="추징")return values.Where(x=>x>.5m).Sum();return values.Where(x=>Math.Abs(x)>.5m).Sum(x=>Math.Abs(x));}
        static void DrawCheck(Graphics g,Rectangle cell,bool selected){}
        static void DrawCell(Graphics g,string text,Font font,Color ink,int[] xs,int c,int y,int h){UiDrawing.Text(g,text??"",font,ink,new Rectangle(xs[c]+1,y,xs[c+1]-xs[c]-2,h),ContentAlignment.MiddleCenter);}static string Difference(decimal value){if(Math.Abs(value)<=.5m)return "0";return value>0?"+"+UiDrawing.Money(value):UiDrawing.Money(value);}static string MaskBirth(string value){string digits=new string((value??"").Where(Char.IsDigit).ToArray());if(digits.Length>=7)return digits.Substring(0,6)+"-"+digits.Substring(6,1)+"******";return value??"";}
        static void DrawStatus(Graphics g,string status,int x,int y,int w,int h,Font font){bool staged=status=="계약제교원"||status=="교특회계"||status=="학교회계";Color ink=staged?Color.FromArgb(49,77,218):status=="반환"?Color.FromArgb(39,102,218):status=="추징"?Color.FromArgb(229,55,55):status=="혼재"?Color.FromArgb(139,76,204):Color.FromArgb(224,101,23);Color back=UiTheme.Dark?(staged?Color.FromArgb(43,53,82):status=="반환"?Color.FromArgb(35,54,83):status=="추징"?Color.FromArgb(78,40,47):status=="혼재"?Color.FromArgb(63,43,82):Color.FromArgb(78,55,38)):(staged?Color.FromArgb(232,237,255):status=="반환"?Color.FromArgb(235,243,255):status=="추징"?Color.FromArgb(255,237,237):status=="혼재"?Color.FromArgb(246,237,255):Color.FromArgb(255,244,232));RectangleF pill=new RectangleF(x+4,y+8,Math.Max(20,w-8),h-16);using(GraphicsPath p=UiDrawing.Rounded(pill,6))using(SolidBrush b=new SolidBrush(back))g.FillPath(b,p);UiDrawing.Text(g,status,font,ink,new Rectangle(x+2,y,w-4,h),ContentAlignment.MiddleCenter);}
    }

    static class ApprovalReportGenerator
    {
        static readonly Color Navy=Color.FromArgb(31,51,125),Green=Color.FromArgb(20,139,72),Red=Color.FromArgb(218,48,48),Grid=Color.FromArgb(188,197,218),Soft=Color.FromArgb(246,249,255);
        public static void CreateExcel(string path,ApprovalReportData data)
        {
            if(data==null||data.Rows==null||data.Rows.Count==0)throw new InvalidOperationException("학교회계 내부결재자료 대상자가 없습니다.");Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path)));using(ExcelPackage package=new ExcelPackage()){ExcelWorksheet ws=package.Workbook.Worksheets.Add("학교회계 기관부담금 지출내역서");ws.View.ShowGridLines=false;ws.Cells.Style.Font.Name="맑은 고딕";ws.Cells.Style.Font.Size=10;ws.Cells[1,1,1,9].Merge=true;ws.Cells[1,1].Value="학교회계 기관부담금 지출내역서";ws.Cells[1,1].Style.Font.Size=20;ws.Cells[1,1].Style.Font.Bold=true;ws.Cells[1,1].Style.Font.Color.SetColor(Navy);ws.Cells[1,1].Style.HorizontalAlignment=ExcelHorizontalAlignment.Center;ws.Row(1).Height=34;ws.Cells[2,1,2,9].Merge=true;ws.Cells[2,1].Value="("+data.Year+"년 "+data.Month+"월분)";ws.Cells[2,1].Style.Font.Bold=true;ws.Cells[2,1].Style.HorizontalAlignment=ExcelHorizontalAlignment.Center;ws.Row(2).Height=21;
                Section(ws,4,"1. 개요");string[,] overview={{"사업장관리번호",FormatSite(data.Site)},{"사업장명",data.Institution??""},{"재원구분","학교회계"},{"부과년월",data.Year+"년 "+data.Month+"월"},{"부과인원",data.Rows.Count+"명"}};for(int r=0;r<5;r++){int rr=5+r;ws.Cells[rr,1,rr,2].Merge=true;ws.Cells[rr,1].Value=overview[r,0];ws.Cells[rr,3,rr,9].Merge=true;ws.Cells[rr,3].Value=overview[r,1];ws.Cells[rr,1,rr,2].Style.Font.Bold=true;ws.Cells[rr,1,rr,2].Style.Fill.PatternType=ExcelFillStyle.Solid;ws.Cells[rr,1,rr,2].Style.Fill.BackgroundColor.SetColor(Soft);Border(ws.Cells[rr,1,rr,9]);ws.Row(rr).Height=22;}
                Section(ws,11,"2. 보험별 기관부담금 현황");ws.Cells[12,1,12,3].Merge=true;ws.Cells[12,4,12,5].Merge=true;ws.Cells[12,6,12,9].Merge=true;ws.Cells[12,1].Value="구분";ws.Cells[12,4].Value="부과인원";ws.Cells[12,6].Value="기관부담금(원)";Header(ws.Cells[12,1,12,9]);string[] insurance={"건강보험","장기요양보험","국민연금","고용보험","산재보험"};for(int i=0;i<5;i++){int rr=13+i;ws.Cells[rr,1,rr,3].Merge=true;ws.Cells[rr,4,rr,5].Merge=true;ws.Cells[rr,6,rr,9].Merge=true;ws.Cells[rr,1].Value=insurance[i];ws.Cells[rr,4].Formula="COUNTIF("+ExcelCellAddress.GetColumnLetter(4+i)+"22:"+ExcelCellAddress.GetColumnLetter(4+i)+(21+data.Rows.Count)+",\">0\")";ws.Cells[rr,6].Formula="SUM("+ExcelCellAddress.GetColumnLetter(4+i)+"22:"+ExcelCellAddress.GetColumnLetter(4+i)+(21+data.Rows.Count)+")";ws.Cells[rr,6].Style.Numberformat.Format="#,##0";Border(ws.Cells[rr,1,rr,9]);ws.Row(rr).Height=22;}int summaryTotal=18;ws.Cells[summaryTotal,1,summaryTotal,3].Merge=true;ws.Cells[summaryTotal,4,summaryTotal,5].Merge=true;ws.Cells[summaryTotal,6,summaryTotal,9].Merge=true;ws.Cells[summaryTotal,1].Value="합계";ws.Cells[summaryTotal,4].Value=data.Rows.Count;ws.Cells[summaryTotal,6].Formula="SUM(F13:F17)";ws.Cells[summaryTotal,6].Style.Numberformat.Format="#,##0";TotalStyle(ws.Cells[summaryTotal,1,summaryTotal,9],Green);
                Section(ws,20,"3. 개인별 내역");string[] headers={"No.","성명","주민등록번호(뒷자리)","건강보험","장기요양보험","국민연금","고용보험","산재보험","합계"};for(int c=0;c<headers.Length;c++)ws.Cells[21,c+1].Value=headers[c];Header(ws.Cells[21,1,21,9]);ws.Row(21).Height=28;int row=22,index=1;foreach(IndividualRowData person in data.Rows){ws.Cells[row,1].Value=index++;ws.Cells[row,2].Value=person.Name;ws.Cells[row,3].Value=MaskBirth(person.Birth);for(int i=0;i<5;i++){ws.Cells[row,4+i].Value=data.Amount(person,i);ws.Cells[row,4+i].Style.Numberformat.Format="#,##0;[Red]-#,##0;0";}ws.Cells[row,9].Formula="SUM(D"+row+":H"+row+")";ws.Cells[row,9].Style.Numberformat.Format="#,##0;[Red]-#,##0;0";if(row%2==1){ws.Cells[row,1,row,9].Style.Fill.PatternType=ExcelFillStyle.Solid;ws.Cells[row,1,row,9].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(250,252,255));}Border(ws.Cells[row,1,row,9]);ws.Row(row).Height=22;row++;}int totalRow=row;ws.Cells[totalRow,1,totalRow,3].Merge=true;ws.Cells[totalRow,1].Value="합계";for(int c=4;c<=9;c++){ws.Cells[totalRow,c].Formula="SUM("+ExcelCellAddress.GetColumnLetter(c)+"22:"+ExcelCellAddress.GetColumnLetter(c)+(totalRow-1)+")";ws.Cells[totalRow,c].Style.Numberformat.Format="#,##0";}TotalStyle(ws.Cells[totalRow,1,totalRow,9],Green);ws.Cells[totalRow+2,1,totalRow+2,9].Merge=true;ws.Cells[totalRow+2,1].Value="※ 본 자료는 내부결재 참고용 보조자료입니다. 원자료와 최종 납부금액을 반드시 확인해 주세요.";ws.Cells[totalRow+2,1].Style.Font.Size=9;ws.Cells[totalRow+2,1].Style.Font.Color.SetColor(Color.FromArgb(126,32,32));ws.Row(totalRow+2).Height=24;
                ws.Cells[5,1,totalRow,9].Style.VerticalAlignment=ExcelVerticalAlignment.Center;ws.Cells[5,1,totalRow,9].Style.HorizontalAlignment=ExcelHorizontalAlignment.Center;ws.Cells[5,3,9,9].Style.HorizontalAlignment=ExcelHorizontalAlignment.Left;double[] widths={7,13,20,14,15,14,14,14,16};for(int c=1;c<=9;c++)ws.Column(c).Width=widths[c-1];ws.PrinterSettings.PaperSize=ePaperSize.A4;ws.PrinterSettings.Orientation=eOrientation.Portrait;ws.PrinterSettings.FitToPage=true;ws.PrinterSettings.FitToWidth=1;ws.PrinterSettings.FitToHeight=0;ws.PrinterSettings.LeftMargin=.35M;ws.PrinterSettings.RightMargin=.35M;ws.PrinterSettings.TopMargin=.4M;ws.PrinterSettings.BottomMargin=.4M;ws.PrinterSettings.RepeatRows=new ExcelAddress("21:21");ws.PrinterSettings.PrintArea=ws.Cells[1,1,totalRow+2,9];ws.HeaderFooter.OddFooter.CenteredText="학교회계 기관부담금 지출내역서";ws.HeaderFooter.OddFooter.RightAlignedText="Page &P / &N";package.SaveAs(new FileInfo(path));}
        }
        public static void CreatePdf(string path,ApprovalReportData data)
        {
            if(data==null||data.Rows==null||data.Rows.Count==0)throw new InvalidOperationException("학교회계 내부결재자료 대상자가 없습니다.");Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path)));const int firstPageRows=15,nextPageRows=27;int pages=data.Rows.Count<=firstPageRows?1:1+(int)Math.Ceiling((data.Rows.Count-firstPageRows)/(double)nextPageRows);var images=new List<byte[]>();for(int page=0;page<pages;page++){int offset=page==0?0:firstPageRows+(page-1)*nextPageRows,count=page==0?firstPageRows:nextPageRows;using(Bitmap bmp=PrintQuality202.CreateBitmap(1240,1754))using(Graphics g=Graphics.FromImage(bmp)){g.SmoothingMode=SmoothingMode.AntiAlias;g.TextRenderingHint=System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;g.Clear(Color.White);g.ScaleTransform(4F,4F);DrawPdfPage(g,data,data.Rows.Skip(offset).Take(count).ToList(),page+1,pages,offset);using(MemoryStream image=new MemoryStream()){PrintQuality202.WriteLossless(bmp,image);images.Add(image.ToArray());}}}WriteImagePdf(path,images,4960,7016);
        }
        static void DrawPdfPage(Graphics g,ApprovalReportData data,List<IndividualRowData> rows,int page,int pages,int offset)
        {
            using(Font title=new Font("맑은 고딕",25F,FontStyle.Bold),subtitle=new Font("맑은 고딕",12F,FontStyle.Bold),section=new Font("맑은 고딕",12F,FontStyle.Bold),head=new Font("맑은 고딕",10F,FontStyle.Bold),cell=new Font("맑은 고딕",9.5F),bold=new Font("맑은 고딕",9.5F,FontStyle.Bold),foot=new Font("맑은 고딕",8.5F))using(Pen grid=new Pen(Grid,1F))
            {
                Draw(g,"학교회계 기관부담금 지출내역서",title,Navy,new RectangleF(60,42,1120,48),StringAlignment.Center);Draw(g,"("+data.Year+"년 "+data.Month+"월분)",subtitle,Navy,new RectangleF(60,91,1120,30),StringAlignment.Center);int y=145;if(page==1){Draw(g,"1. 개요",section,Navy,new RectangleF(70,y,1080,28),StringAlignment.Near);y+=34;string[,] overview={{"사업장관리번호",FormatSite(data.Site)},{"사업장명",data.Institution??""},{"재원구분","학교회계"},{"부과년월",data.Year+"년 "+data.Month+"월"},{"부과인원",data.Rows.Count+"명"}};for(int r=0;r<5;r++){using(SolidBrush fill=new SolidBrush(Soft))g.FillRectangle(fill,70,y,250,35);g.DrawRectangle(grid,70,y,250,35);g.DrawRectangle(grid,320,y,850,35);Draw(g,overview[r,0],bold,Navy,new RectangleF(80,y,230,35),StringAlignment.Near);Draw(g,overview[r,1],cell,Navy,new RectangleF(335,y,820,35),StringAlignment.Near);y+=35;}y+=24;Draw(g,"2. 보험별 기관부담금 현황",section,Navy,new RectangleF(70,y,1080,28),StringAlignment.Near);y+=34;int[] sw={360,250,490};string[] sh={"구분","부과인원","기관부담금(원)"};DrawHeader(g,70,y,sw,sh,head);y+=42;string[] names={"건강보험","장기요양보험","국민연금","고용보험","산재보험","합계"};for(int r=0;r<6;r++){bool total=r==5;if(total)using(SolidBrush fill=new SolidBrush(Color.FromArgb(240,249,244)))g.FillRectangle(fill,70,y,1100,36);string[] v={names[r],(total?data.Rows.Count:data.InsurancePeople(r))+"명",UiDrawing.Money(total?data.Total:data.InsuranceTotal(r))};DrawRow(g,70,y,sw,v,total?bold:cell,total?Green:Navy);y+=36;}y+=28;Draw(g,"3. 개인별 내역",section,Navy,new RectangleF(70,y,1080,28),StringAlignment.Near);y+=35;}else{Draw(g,"3. 개인별 내역 (계속)",section,Navy,new RectangleF(70,y,1080,28),StringAlignment.Near);y+=35;}
                int[] widths={55,110,190,120,120,120,120,120,145};string[] heads={"No.","성명","주민등록번호\n(뒷자리)","건강보험","장기요양","국민연금","고용보험","산재보험","합계"};DrawHeader(g,50,y,widths,heads,head);y+=48;for(int r=0;r<rows.Count;r++){IndividualRowData person=rows[r];decimal[] a=Enumerable.Range(0,5).Select(i=>data.Amount(person,i)).ToArray();if(r%2==1)using(SolidBrush fill=new SolidBrush(Color.FromArgb(250,252,255)))g.FillRectangle(fill,50,y,widths.Sum(),39);string[] values={(offset+r+1).ToString(),person.Name,MaskBirth(person.Birth),UiDrawing.Money(a[0]),UiDrawing.Money(a[1]),UiDrawing.Money(a[2]),UiDrawing.Money(a[3]),UiDrawing.Money(a[4]),UiDrawing.Money(a.Sum())};DrawRow(g,50,y,widths,values,cell,Navy);y+=39;}if(page==pages){decimal[] totals=Enumerable.Range(0,5).Select(data.InsuranceTotal).ToArray();string[] values={"합계","",data.Rows.Count+"명",UiDrawing.Money(totals[0]),UiDrawing.Money(totals[1]),UiDrawing.Money(totals[2]),UiDrawing.Money(totals[3]),UiDrawing.Money(totals[4]),UiDrawing.Money(totals.Sum())};using(SolidBrush fill=new SolidBrush(Color.FromArgb(240,249,244)))g.FillRectangle(fill,50,y,widths.Sum(),42);DrawRow(g,50,y,widths,values,bold,Green);}using(SolidBrush notice=new SolidBrush(Color.FromArgb(255,248,248)))g.FillRoundedRectangle(notice,new RectangleF(60,1642,1120,44),9);Draw(g,"※ 내부결재 참고용 보조자료입니다. 원자료와 최종 납부금액을 반드시 확인해 주세요.",foot,Color.FromArgb(126,32,32),new RectangleF(78,1645,1050,38),StringAlignment.Near);Draw(g,"생성일 "+DateTime.Now.ToString("yyyy-MM-dd HH:mm")+"   |   "+page+" / "+pages,foot,Color.FromArgb(105,114,143),new RectangleF(60,1691,1120,25),StringAlignment.Far);
            }
        }
        static void DrawHeader(Graphics g,int x,int y,int[] widths,string[] texts,Font font){int xx=x;using(SolidBrush fill=new SolidBrush(Navy))g.FillRectangle(fill,x,y,widths.Sum(),texts.Any(t=>t.Contains("\n"))?48:42);int h=texts.Any(t=>t.Contains("\n"))?48:42;for(int i=0;i<texts.Length;i++){Draw(g,texts[i],font,Color.White,new RectangleF(xx+2,y,widths[i]-4,h),StringAlignment.Center);xx+=widths[i];}}
        static void DrawRow(Graphics g,int x,int y,int[] widths,string[] values,Font font,Color color){int xx=x,h=values.Length>0&&values[0]=="합계"?42:values.Length==3?36:39;using(Pen p=new Pen(Grid))for(int i=0;i<values.Length;i++){g.DrawRectangle(p,xx,y,widths[i],h);Draw(g,values[i],font,color,new RectangleF(xx+3,y,widths[i]-6,h),StringAlignment.Center);xx+=widths[i];}}
        static void Draw(Graphics g,string text,Font font,Color color,RectangleF rect,StringAlignment align){using(SolidBrush b=new SolidBrush(color))using(StringFormat f=new StringFormat{Alignment=align,LineAlignment=StringAlignment.Center,Trimming=StringTrimming.EllipsisCharacter})g.DrawString(text??"",font,b,rect,f);}
        static void Section(ExcelWorksheet ws,int row,string text){ws.Cells[row,1,row,9].Merge=true;ws.Cells[row,1].Value=text;ws.Cells[row,1].Style.Font.Bold=true;ws.Cells[row,1].Style.Font.Size=12;ws.Cells[row,1].Style.Font.Color.SetColor(Navy);ws.Row(row).Height=24;}
        static void Header(ExcelRange range){range.Style.Font.Bold=true;range.Style.Font.Color.SetColor(Color.White);range.Style.Fill.PatternType=ExcelFillStyle.Solid;range.Style.Fill.BackgroundColor.SetColor(Navy);range.Style.HorizontalAlignment=ExcelHorizontalAlignment.Center;range.Style.VerticalAlignment=ExcelVerticalAlignment.Center;Border(range);}
        static void TotalStyle(ExcelRange range,Color accent){range.Style.Font.Bold=true;range.Style.Font.Color.SetColor(accent);range.Style.Fill.PatternType=ExcelFillStyle.Solid;range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(240,249,244));range.Style.HorizontalAlignment=ExcelHorizontalAlignment.Center;Border(range);}
        static void Border(ExcelRange range){range.Style.Border.Top.Style=range.Style.Border.Bottom.Style=range.Style.Border.Left.Style=range.Style.Border.Right.Style=ExcelBorderStyle.Thin;range.Style.Border.Top.Color.SetColor(Grid);range.Style.Border.Bottom.Color.SetColor(Grid);range.Style.Border.Left.Color.SetColor(Grid);range.Style.Border.Right.Color.SetColor(Grid);}
        static string FormatSite(string site){string d=Regex.Replace(site??"","[^0-9]","");return d.Length==11?d.Substring(0,3)+"-"+d.Substring(3,2)+"-"+d.Substring(5,6):(String.IsNullOrWhiteSpace(site)?"-":site);}static string MaskBirth(string value){string d=Regex.Replace(value??"","[^0-9]","");return d.Length>=7?d.Substring(0,6)+"-"+d.Substring(6,1)+"******":d.Length>=6?d.Substring(0,6)+"-*******":value??"";}
        static void WriteImagePdf(string path,List<byte[]> images,int width,int height){int objectCount=2+images.Count*3;long[] offsets=new long[objectCount+1];using(FileStream stream=new FileStream(path,FileMode.Create,FileAccess.Write)){Ascii(stream,"%PDF-1.4\n%");stream.Write(new byte[]{0xE2,0xE3,0xCF,0xD3},0,4);Ascii(stream,"\n");Action<int,string> writeObject=(number,body)=>{offsets[number]=stream.Position;Ascii(stream,number+" 0 obj\n"+body+"\nendobj\n");};writeObject(1,"<< /Type /Catalog /Pages 2 0 R >>");string kids=String.Join(" ",Enumerable.Range(0,images.Count).Select(i=>(3+i*3)+" 0 R"));writeObject(2,"<< /Type /Pages /Kids ["+kids+"] /Count "+images.Count+" >>");for(int i=0;i<images.Count;i++){int page=3+i*3,image=page+1,content=page+2;writeObject(page,"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << /XObject << /Im0 "+image+" 0 R >> >> /Contents "+content+" 0 R >>");offsets[image]=stream.Position;Ascii(stream,image+" 0 obj\n<< /Type /XObject /Subtype /Image /Width "+width+" /Height "+height+" /ColorSpace /DeviceRGB /BitsPerComponent 8 /Interpolate false /Filter /FlateDecode /Length "+images[i].Length+" >>\nstream\n");stream.Write(images[i],0,images[i].Length);Ascii(stream,"\nendstream\nendobj\n");string draw="q\n595 0 0 842 0 0 cm\n/Im0 Do\nQ\n";writeObject(content,"<< /Length "+Encoding.ASCII.GetByteCount(draw)+" >>\nstream\n"+draw+"endstream");}long xref=stream.Position;Ascii(stream,"xref\n0 "+(objectCount+1)+"\n0000000000 65535 f \n");for(int i=1;i<=objectCount;i++)Ascii(stream,offsets[i].ToString("0000000000",CultureInfo.InvariantCulture)+" 00000 n \n");Ascii(stream,"trailer\n<< /Size "+(objectCount+1)+" /Root 1 0 R >>\nstartxref\n"+xref+"\n%%EOF");}}
        static void Ascii(Stream stream,string text){byte[] bytes=Encoding.ASCII.GetBytes(text);stream.Write(bytes,0,bytes.Length);}
    }

    static class AdjustmentReportGenerator
    {
        static readonly Color Navy=Color.FromArgb(31,51,125),Blue=Color.FromArgb(43,102,224),Red=Color.FromArgb(229,55,55),Green=Color.FromArgb(22,139,69),Grid=Color.FromArgb(218,225,241);
        public static void CreateExcel(string path,List<IndividualRowData> rows,string mode,int year,int month,string site)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path)));using(ExcelPackage package=new ExcelPackage()){ExcelWorksheet ws=package.Workbook.Worksheets.Add(mode=="전체"?"반환추징내역":mode+"내역");ws.Cells[1,1].Value="사회보험료 "+(mode=="전체"?"반환·추징":mode)+" 대상 내역";ws.Cells[1,1,1,20].Merge=true;ws.Cells[1,1].Style.Font.Size=18;ws.Cells[1,1].Style.Font.Bold=true;ws.Cells[1,1].Style.Font.Color.SetColor(Navy);ws.Cells[1,1].Style.HorizontalAlignment=ExcelHorizontalAlignment.Left;ws.Row(1).Height=31;ws.Cells[2,1].Value="고지년월";ws.Cells[2,2].Value=year+"년 "+month+"월";ws.Cells[2,4].Value="사업장";ws.Cells[2,5].Value=site;ws.Cells[2,9].Value="생성일";ws.Cells[2,10].Value=DateTime.Now.ToString("yyyy-MM-dd");ws.Cells[3,1].Value="※ 본 자료는 반환 또는 추징 대상자 확인용 보조자료입니다. 원자료와 최종 금액을 반드시 확인해 주세요.";ws.Cells[3,1,3,20].Merge=true;ws.Cells[3,1].Style.Font.Color.SetColor(Color.FromArgb(126,32,32));
                string[] headers={"No.","구분","사업장 관리번호","재원","이름","주민/사번","직종명","건강 고지","건강 급여","건강 차액","국민 고지","국민 급여","국민 차액","고용 고지","고용 급여","고용 차액","산재 고지","산재 급여","산재 차액","조정 금액"};for(int c=0;c<headers.Length;c++)ws.Cells[5,c+1].Value=headers[c];using(ExcelRange h=ws.Cells[5,1,5,20]){h.Style.Font.Bold=true;h.Style.Font.Color.SetColor(Color.White);h.Style.Fill.PatternType=ExcelFillStyle.Solid;h.Style.Fill.BackgroundColor.SetColor(Navy);h.Style.HorizontalAlignment=ExcelHorizontalAlignment.Center;h.Style.VerticalAlignment=ExcelVerticalAlignment.Center;}ws.Row(5).Height=27;
                int r=6,index=1;foreach(IndividualRowData d in rows){string kind=Kind(d,mode);decimal[] values={d.HealthNotice,d.HealthPayroll,d.HealthDifference,d.PensionNotice,d.PensionPayroll,d.PensionDifference,d.EmploymentNotice,d.EmploymentPayroll,d.EmploymentDifference,d.IndustrialNotice,d.IndustrialPayroll,d.IndustrialDifference};object[] fixedValues={index++,kind,d.Site,d.Fund,d.Name,MaskBirth(d.Birth),d.Job};for(int c=0;c<fixedValues.Length;c++)ws.Cells[r,c+1].Value=fixedValues[c];for(int c=0;c<values.Length;c++){ws.Cells[r,8+c].Value=values[c];ws.Cells[r,8+c].Style.Numberformat.Format="#,##0;[Red]-#,##0;0";if(c%3==2)ws.Cells[r,8+c].Style.Font.Color.SetColor(values[c]>.5m?Red:values[c]<-.5m?Blue:Green);}ws.Cells[r,20].Value=Amount(d,mode);ws.Cells[r,20].Style.Numberformat.Format="#,##0";ws.Cells[r,20].Style.Font.Bold=true;ws.Cells[r,20].Style.Font.Color.SetColor(kind=="반환"?Blue:kind=="추징"?Red:Color.FromArgb(139,76,204));if(r%2==1){ws.Cells[r,1,r,20].Style.Fill.PatternType=ExcelFillStyle.Solid;ws.Cells[r,1,r,20].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(248,250,255));}r++;}
                int last=Math.Max(6,r-1);using(ExcelRange body=ws.Cells[5,1,last,20]){body.Style.Border.Top.Style=body.Style.Border.Bottom.Style=body.Style.Border.Left.Style=body.Style.Border.Right.Style=ExcelBorderStyle.Thin;body.Style.Border.Top.Color.SetColor(Grid);body.Style.Border.Bottom.Color.SetColor(Grid);body.Style.Border.Left.Color.SetColor(Grid);body.Style.Border.Right.Color.SetColor(Grid);body.Style.VerticalAlignment=ExcelVerticalAlignment.Center;body.Style.HorizontalAlignment=ExcelHorizontalAlignment.Center;}ws.View.FreezePanes(6,8);ws.Column(1).Width=6;ws.Column(2).Width=12;ws.Column(3).Width=18;ws.Column(4).Width=14;ws.Column(5).Width=12;ws.Column(6).Width=17;ws.Column(7).Width=24;for(int c=8;c<=20;c++)ws.Column(c).Width=13;ws.PrinterSettings.Orientation=eOrientation.Landscape;ws.PrinterSettings.FitToPage=true;ws.PrinterSettings.FitToWidth=1;ws.PrinterSettings.FitToHeight=0;ws.PrinterSettings.RepeatRows=new ExcelAddress("5:5");package.SaveAs(new FileInfo(path));}
        }
        public static void CreatePdf(string path,List<IndividualRowData> rows,string mode,int year,int month,string site)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path)));const int perPage=18;var pages=new List<byte[]>();int totalCount=rows.Count,pageCount=Math.Max(1,(int)Math.Ceiling(totalCount/(double)perPage));for(int page=0;page<pageCount;page++){List<IndividualRowData> slice=rows.Skip(page*perPage).Take(perPage).ToList();using(Bitmap bmp=PrintQuality202.CreateBitmap(1754,1240))using(Graphics g=Graphics.FromImage(bmp)){g.SmoothingMode=SmoothingMode.AntiAlias;g.TextRenderingHint=System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;g.Clear(Color.White);g.ScaleTransform(4F,4F);DrawPdfPage(g,slice,mode,year,month,site,page+1,pageCount,page*perPage,totalCount);using(MemoryStream image=new MemoryStream()){PrintQuality202.WriteLossless(bmp,image);pages.Add(image.ToArray());}}}WriteImagePdf(path,pages,7016,4960);
        }
        static void DrawPdfPage(Graphics g,List<IndividualRowData> rows,string mode,int year,int month,string site,int page,int pages,int offset,int totalCount)
        {
            using(Font title=new Font("맑은 고딕",27F,FontStyle.Bold),subtitle=new Font("맑은 고딕",11F),meta=new Font("맑은 고딕",12F,FontStyle.Bold),head=new Font("맑은 고딕",10F,FontStyle.Bold),cell=new Font("맑은 고딕",9F,FontStyle.Bold),small=new Font("맑은 고딕",8.5F),foot=new Font("맑은 고딕",8.5F))
            {
                string reportMode=mode=="전체"?"반환·추징":mode;Draw(g,"사회보험료 "+reportMode+" 대상 내역",title,Navy,new RectangleF(48,35,1200,45),StringAlignment.Near);Draw(g,"급여 공제액과 보험료 고지액의 차이를 기준으로 작성한 업무 보조자료",subtitle,Color.FromArgb(90,101,137),new RectangleF(50,83,1200,30),StringAlignment.Near);using(SolidBrush b=new SolidBrush(Color.FromArgb(247,249,255)))g.FillRoundedRectangle(b,new RectangleF(48,124,1658,72),12);Draw(g,"고지년월  "+year+"년 "+month+"월",meta,Navy,new RectangleF(70,136,330,46),StringAlignment.Near);Draw(g,"사업장  "+site,meta,Navy,new RectangleF(410,136,800,46),StringAlignment.Near);Draw(g,"대상 "+totalCount+"명",meta,mode=="반환"?Blue:mode=="추징"?Red:Navy,new RectangleF(1420,136,250,46),StringAlignment.Far);
                string[] heads={"No.","구분","재원","이름","주민/사번","직종명","건강 차액","국민 차액","고용 차액","산재 차액","조정 금액"};int[] widths={48,88,108,105,137,230,160,160,160,160,172};int x=48,y=220,headerH=58,rowH=45;using(SolidBrush hb=new SolidBrush(Navy))g.FillRectangle(hb,x,y,widths.Sum(),headerH);for(int c=0;c<heads.Length;c++){Draw(g,heads[c],head,Color.White,new RectangleF(x,y,widths[c],headerH),StringAlignment.Center);x+=widths[c];}y+=headerH;for(int r=0;r<rows.Count;r++){IndividualRowData d=rows[r];x=48;if(r%2==1)using(SolidBrush alt=new SolidBrush(Color.FromArgb(249,251,255)))g.FillRectangle(alt,x,y,widths.Sum(),rowH);string kind=Kind(d,mode);string[] texts={""+(offset+r+1),kind,d.Fund,d.Name,MaskBirth(d.Birth),d.Job,Signed(d.HealthDifference),Signed(d.PensionDifference),Signed(d.EmploymentDifference),Signed(d.IndustrialDifference),UiDrawing.Money(Amount(d,mode))};for(int c=0;c<texts.Length;c++){Color ink=c>=6&&c<=9?DiffColor(c==6?d.HealthDifference:c==7?d.PensionDifference:c==8?d.EmploymentDifference:d.IndustrialDifference):c==1?kind=="반환"?Blue:kind=="추징"?Red:Color.FromArgb(139,76,204):Navy;Draw(g,texts[c],c==5?small:cell,ink,new RectangleF(x+3,y,widths[c]-6,rowH),StringAlignment.Center);using(Pen p=new Pen(Grid))g.DrawRectangle(p,x,y,widths[c],rowH);x+=widths[c];}y+=rowH;}
                using(Pen p=new Pen(Grid))g.DrawRectangle(p,48,220,widths.Sum(),headerH+rows.Count*rowH);using(SolidBrush notice=new SolidBrush(Color.FromArgb(255,248,248)))g.FillRoundedRectangle(notice,new RectangleF(48,1100,1658,55),10);Draw(g,"※ 본 자료는 업무 확인용 보조자료입니다. 반환·추징 처리 전 원자료와 최종 금액을 반드시 확인해 주세요.",small,Color.FromArgb(126,32,32),new RectangleF(70,1105,1550,45),StringAlignment.Near);Draw(g,"생성일 "+DateTime.Now.ToString("yyyy-MM-dd HH:mm")+"   |   "+page+" / "+pages,foot,Color.FromArgb(105,114,143),new RectangleF(48,1173,1658,28),StringAlignment.Far);
            }
        }
        static void Draw(Graphics g,string text,Font font,Color color,RectangleF rect,StringAlignment align){using(SolidBrush b=new SolidBrush(color))using(StringFormat f=new StringFormat{Alignment=align,LineAlignment=StringAlignment.Center,Trimming=StringTrimming.EllipsisCharacter,FormatFlags=StringFormatFlags.NoWrap})g.DrawString(text??"",font,b,rect,f);}static Color DiffColor(decimal d){return d>.5m?Red:d<-.5m?Blue:Green;}static string Signed(decimal value){if(Math.Abs(value)<=.5m)return "0";return value>0?"+"+UiDrawing.Money(value):UiDrawing.Money(value);}static string MaskBirth(string value){string digits=Regex.Replace(value??"","[^0-9]","");return digits.Length>=7?digits.Substring(0,6)+"-"+digits.Substring(6,1)+"******":value??"";}static bool Refund(IndividualRowData d){return d.Fund!="분류필요"&&(d.HealthDifference<-.5m||d.PensionDifference<-.5m||d.EmploymentDifference<-.5m||d.IndustrialDifference<-.5m);}static bool Collection(IndividualRowData d){return d.Fund!="분류필요"&&(d.HealthDifference>.5m||d.PensionDifference>.5m||d.EmploymentDifference>.5m||d.IndustrialDifference>.5m);}static string Kind(IndividualRowData d,string mode){if(mode=="반환"||mode=="추징")return mode;return Refund(d)&&Collection(d)?"반환·추징":Refund(d)?"반환":"추징";}static decimal Amount(IndividualRowData d,string mode){decimal[] diffs={d.HealthDifference,d.PensionDifference,d.EmploymentDifference,d.IndustrialDifference};if(mode=="반환")return diffs.Where(x=>x<-.5m).Sum(x=>Math.Abs(x));if(mode=="추징")return diffs.Where(x=>x>.5m).Sum();return diffs.Where(x=>Math.Abs(x)>.5m).Sum(x=>Math.Abs(x));}
        static void WriteImagePdf(string path,List<byte[]> images,int width,int height)
        {
            int objectCount=2+images.Count*3;long[] offsets=new long[objectCount+1];using(FileStream stream=new FileStream(path,FileMode.Create,FileAccess.Write)){Ascii(stream,"%PDF-1.4\n%");stream.Write(new byte[]{0xE2,0xE3,0xCF,0xD3},0,4);Ascii(stream,"\n");Action<int,string> writeObject=(number,body)=>{offsets[number]=stream.Position;Ascii(stream,number+" 0 obj\n"+body+"\nendobj\n");};writeObject(1,"<< /Type /Catalog /Pages 2 0 R >>");string kids=String.Join(" ",Enumerable.Range(0,images.Count).Select(i=>(3+i*3)+" 0 R"));writeObject(2,"<< /Type /Pages /Kids ["+kids+"] /Count "+images.Count+" >>");for(int i=0;i<images.Count;i++){int page=3+i*3,image=page+1,content=page+2;writeObject(page,"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 842 595] /Resources << /XObject << /Im0 "+image+" 0 R >> >> /Contents "+content+" 0 R >>");offsets[image]=stream.Position;Ascii(stream,image+" 0 obj\n<< /Type /XObject /Subtype /Image /Width "+width+" /Height "+height+" /ColorSpace /DeviceRGB /BitsPerComponent 8 /Interpolate false /Filter /FlateDecode /Length "+images[i].Length+" >>\nstream\n");stream.Write(images[i],0,images[i].Length);Ascii(stream,"\nendstream\nendobj\n");string draw="q\n842 0 0 595 0 0 cm\n/Im0 Do\nQ\n";writeObject(content,"<< /Length "+Encoding.ASCII.GetByteCount(draw)+" >>\nstream\n"+draw+"endstream");}long xref=stream.Position;Ascii(stream,"xref\n0 "+(objectCount+1)+"\n0000000000 65535 f \n");for(int i=1;i<=objectCount;i++)Ascii(stream,offsets[i].ToString("0000000000",CultureInfo.InvariantCulture)+" 00000 n \n");Ascii(stream,"trailer\n<< /Size "+(objectCount+1)+" /Root 1 0 R >>\nstartxref\n"+xref+"\n%%EOF");}
        }
        static void Ascii(Stream stream,string text){byte[] bytes=Encoding.ASCII.GetBytes(text);stream.Write(bytes,0,bytes.Length);}
    }

    static class ReportGraphicsExtensions
    {
        public static void FillRoundedRectangle(this Graphics g,Brush brush,RectangleF rect,float radius){using(GraphicsPath p=UiDrawing.Rounded(rect,radius))g.FillPath(brush,p);}
    }

    class SidebarNavButton : Button
    {
        public string IconKind="dot";public bool Active,ShowChevron,Expanded=true;public int Indent;public float VisualOpacity=1F;bool hovered;
        public SidebarNavButton(){FlatStyle=System.Windows.Forms.FlatStyle.Flat;FlatAppearance.BorderSize=0;BackColor=Color.Transparent;ForeColor=Color.FromArgb(43,55,103);Cursor=Cursors.Hand;TabStop=false;SetStyle(ControlStyles.AllPaintingInWmPaint|ControlStyles.OptimizedDoubleBuffer|ControlStyles.UserPaint,true);}
        protected override void OnMouseEnter(EventArgs e){hovered=true;Invalidate();base.OnMouseEnter(e);}protected override void OnMouseLeave(EventArgs e){hovered=false;Invalidate();base.OnMouseLeave(e);}
        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.SmoothingMode=SmoothingMode.AntiAlias;e.Graphics.TextRenderingHint=System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;e.Graphics.Clear(Parent==null?UiTheme.Sidebar:Parent.BackColor);
            if(Width<24||Height<8)return;
            Rectangle bounds=new Rectangle(1,1,Width-3,Height-3);using(GraphicsPath path=RoundRect(bounds,Math.Max(2,Math.Min(10,(bounds.Height-1)/2))))
            {
                if(Active)using(LinearGradientBrush fill=new LinearGradientBrush(bounds,UiTheme.Dark?Color.FromArgb(47,57,78):Color.FromArgb(229,234,255),UiTheme.Dark?Color.FromArgb(54,50,76):Color.FromArgb(238,237,255),0F))e.Graphics.FillPath(fill,path);
                else if(hovered)using(SolidBrush fill=new SolidBrush(UiTheme.Surface))e.Graphics.FillPath(fill,path);
            }
            int alpha=Math.Max(0,Math.Min(255,(int)(255F*VisualOpacity)));Color baseColor=Active?UiTheme.Accent:UiTheme.Text;Color ink=Color.FromArgb(alpha,baseColor);int iconX=13+Indent,iconY=(Height-18)/2;DrawIcon(e.Graphics,IconKind,new Rectangle(iconX,iconY,18,18),ink,Active);
            float textWidth=Width-iconX-(ShowChevron?50:22);using(SolidBrush textBrush=new SolidBrush(ink))using(StringFormat format=new StringFormat{Alignment=StringAlignment.Near,LineAlignment=StringAlignment.Center,Trimming=StringTrimming.EllipsisCharacter,FormatFlags=StringFormatFlags.NoWrap})e.Graphics.DrawString(Text,Font,textBrush,new RectangleF(iconX+29,0,textWidth,Height),format);
            if(ShowChevron){using(Pen pen=new Pen(ink,1.7F)){pen.StartCap=LineCap.Round;pen.EndCap=LineCap.Round;PointF[] arrow=Expanded?new[]{new PointF(145,19),new PointF(150,14),new PointF(155,19)}:new[]{new PointF(145,14),new PointF(150,19),new PointF(155,14)};e.Graphics.DrawLines(pen,arrow);}}
        }
        static GraphicsPath RoundRect(Rectangle r,int radius){int d=radius*2;var p=new GraphicsPath();p.AddArc(r.Left,r.Top,d,d,180,90);p.AddArc(r.Right-d,r.Top,d,d,270,90);p.AddArc(r.Right-d,r.Bottom-d,d,d,0,90);p.AddArc(r.Left,r.Bottom-d,d,d,90,90);p.CloseFigure();return p;}
        static void DrawIcon(Graphics g,string kind,Rectangle r,Color color,bool active)
        {
            float x=r.X,y=r.Y,w=r.Width,h=r.Height;using(Pen p=new Pen(color,1.45F)){p.StartCap=LineCap.Round;p.EndCap=LineCap.Round;p.LineJoin=LineJoin.Round;
                if(kind=="folder"||kind=="archive")
                {
                    using(GraphicsPath folder=new GraphicsPath()){folder.AddLine(x+1,y+5,x+7,y+5);folder.AddLine(x+7,y+5,x+9,y+7);folder.AddLine(x+9,y+7,x+17,y+7);folder.AddLine(x+17,y+7,x+17,y+16);folder.AddLine(x+17,y+16,x+1,y+16);folder.CloseFigure();if(active)using(SolidBrush b=new SolidBrush(Color.FromArgb(Math.Min(80,(int)color.A),color)))g.FillPath(b,folder);g.DrawPath(p,folder);}if(kind=="archive")g.DrawLine(p,x+4,y+11,x+14,y+11);return;
                }
                if(kind=="results")
                {
                    g.DrawRoundedRectangle(p,new RectangleF(x+1,y+3,w-2,h-4),2);g.DrawLine(p,x+1,y+7,x+17,y+7);g.DrawLine(p,x+5,y+1,x+5,y+5);g.DrawLine(p,x+13,y+1,x+13,y+5);for(int rr=0;rr<2;rr++)for(int cc=0;cc<3;cc++)g.FillEllipse(p.Brush,x+4+cc*4.5F,y+10+rr*3.4F,1.5F,1.5F);return;
                }
                if(kind=="summary")
                {
                    g.DrawRoundedRectangle(p,new RectangleF(x+2,y+2,14,14),2);g.DrawLine(p,x+6,y+6,x+13,y+6);g.DrawLine(p,x+6,y+9,x+13,y+9);g.DrawLine(p,x+6,y+12,x+11,y+12);return;
                }
                if(kind=="person")
                {
                    g.DrawEllipse(p,x+6,y+1,6,6);g.DrawArc(p,x+2,y+8,14,11,195,150);return;
                }
                if(kind=="adjustment")
                {
                    g.DrawEllipse(p,x+1,y+1,16,16);g.DrawLine(p,x+5,y+7,x+13,y+7);g.DrawLine(p,x+5,y+7,x+7,y+5);g.DrawLine(p,x+5,y+7,x+7,y+9);g.DrawLine(p,x+13,y+11,x+5,y+11);g.DrawLine(p,x+13,y+11,x+11,y+9);g.DrawLine(p,x+13,y+11,x+11,y+13);return;
                }
                if(kind=="review")
                {
                    g.DrawRoundedRectangle(p,new RectangleF(x+2,y+3,14,14),2);g.DrawRoundedRectangle(p,new RectangleF(x+6,y+1,6,4),1);g.DrawLine(p,x+5,y+10,x+8,y+13);g.DrawLine(p,x+8,y+13,x+13,y+7);return;
                }
                if(kind=="discount")
                {
                    using(GraphicsPath ticket=new GraphicsPath()){ticket.AddPolygon(new[]{new PointF(x+1,y+7),new PointF(x+7,y+1),new PointF(x+11,y+5),new PointF(x+17,y+11),new PointF(x+11,y+17),new PointF(x+5,y+11)});g.DrawPath(p,ticket);}g.DrawEllipse(p,x+5,y+5,2,2);g.DrawEllipse(p,x+11,y+11,2,2);g.DrawLine(p,x+7,y+11,x+11,y+7);return;
                }
                if(kind=="document")
                {
                    using(GraphicsPath page=new GraphicsPath()){page.AddLine(x+3,y+1,x+11,y+1);page.AddLine(x+11,y+1,x+16,y+6);page.AddLine(x+16,y+6,x+16,y+17);page.AddLine(x+16,y+17,x+3,y+17);page.CloseFigure();g.DrawPath(p,page);}g.DrawLine(p,x+11,y+1,x+11,y+6);g.DrawLine(p,x+11,y+6,x+16,y+6);g.DrawLine(p,x+6,y+10,x+13,y+10);g.DrawLine(p,x+6,y+13,x+13,y+13);return;
                }
                if(kind=="settings")
                {
                    PointF[] nut={new PointF(x+9,y+1),new PointF(x+16,y+5),new PointF(x+16,y+13),new PointF(x+9,y+17),new PointF(x+2,y+13),new PointF(x+2,y+5)};
                    using(GraphicsPath hex=new GraphicsPath()){hex.AddPolygon(nut);if(active)using(SolidBrush b=new SolidBrush(Color.FromArgb(Math.Min(70,(int)color.A),color)))g.FillPath(b,hex);g.DrawPath(p,hex);}
                    g.DrawEllipse(p,x+6,y+6,6,6);return;
                }
                g.FillEllipse(p.Brush,x+7,y+7,4,4);
            }
        }
    }

    static class GraphicsExtensions
    {
        public static void DrawRoundedRectangle(this Graphics g,Pen pen,RectangleF r,float radius){float d=radius*2;using(GraphicsPath p=new GraphicsPath()){p.AddArc(r.Left,r.Top,d,d,180,90);p.AddArc(r.Right-d,r.Top,d,d,270,90);p.AddArc(r.Right-d,r.Bottom-d,d,d,0,90);p.AddArc(r.Left,r.Bottom-d,d,d,90,90);p.CloseFigure();g.DrawPath(pen,p);}}
    }

    class BufferedFlowLayoutPanel : FlowLayoutPanel
    {
        public BufferedFlowLayoutPanel(){DoubleBuffered=true;ResizeRedraw=true;SetStyle(ControlStyles.AllPaintingInWmPaint|ControlStyles.OptimizedDoubleBuffer|ControlStyles.UserPaint,true);}
    }

    class RoundedPanel : Panel
    {
        public int Radius=14;public Color BorderColor=Color.FromArgb(226,230,242);public int BorderWidth=1;public System.Drawing.Drawing2D.DashStyle BorderDashStyle=System.Drawing.Drawing2D.DashStyle.Solid;
        protected override void OnResize(EventArgs e){base.OnResize(e);using(GraphicsPath p=Path())Region=new Region(p);}
        GraphicsPath Path(){int d=Math.Max(2,Radius*2),w=Math.Max(1,Width-1),h=Math.Max(1,Height-1);var p=new GraphicsPath();p.AddArc(0,0,d,d,180,90);p.AddArc(w-d,0,d,d,270,90);p.AddArc(w-d,h-d,d,d,0,90);p.AddArc(0,h-d,d,d,90,90);p.CloseFigure();return p;}
        protected override void OnPaint(PaintEventArgs e){base.OnPaint(e);e.Graphics.SmoothingMode=SmoothingMode.AntiAlias;float inset=Math.Max(.5F,BorderWidth/2F);RectangleF bounds=new RectangleF(inset,inset,Math.Max(1,Width-1-2*inset),Math.Max(1,Height-1-2*inset));using(GraphicsPath p=UiDrawing.Rounded(bounds,Math.Max(2,Radius-inset)))using(Pen pen=new Pen(BorderColor,BorderWidth)){pen.DashStyle=BorderDashStyle;e.Graphics.DrawPath(pen,p);}}
    }

    class Person
    {
        public string Key, Name, Birth, Source, Job, Fund, Reason; public decimal Health, Pension, Employment, Industrial, HealthSettlement, PensionSettlement, EmploymentSettlement, IndustrialSettlement;
        public decimal HealthCurrent, LongTermCurrent, HealthSettlementHealth, LongTermSettlement; public bool HasHealthComponents;
    }
    class Charge
    {
        public string Key, Name, Birth, Insurance, Source, WorkplaceNumber; public decimal CompareAmount, EmployerAmount, SettlementPersonal, SettlementEmployer;
        public decimal HealthCurrent, LongTermCurrent, HealthSettlementHealth, LongTermSettlement, EmployerHealthCurrent, EmployerLongTermCurrent, EmployerSettlementHealth, EmployerSettlementLongTerm; public bool HasHealthComponents;
    }
    class ResultRow
    {
        public string Fund, Name, Birth, Source, Job, Reason, Insurance, ChargeSource, WorkplaceNumber, Status, Note; public decimal Deduction, DeductionSettlement, Charge, Difference, Employer, SettlementPersonal, SettlementEmployer;
        public decimal DeductionHealth, DeductionLongTerm, DeductionSettlementHealth, DeductionSettlementLongTerm, ChargeHealth, ChargeLongTerm, SettlementPersonalHealth, SettlementPersonalLongTerm, EmployerHealth, EmployerLongTerm, SettlementEmployerHealth, SettlementEmployerLongTerm; public bool HasHealthComponents;
    }
    class SiteSheetSet
    {
        public string Site;public ExcelWorksheet Summary,Individual,Review;public int IndividualLast,ReviewLast;
    }
    class ReviewSheetResult
    {
        public int Last;public Dictionary<string,int> Rows=new Dictionary<string,int>();
    }
    class InputFile { public string Kind, Path; }
    class Recognition { public string Kind, File, Sheet, HeaderRow, Rows, State, Detail; }
    class BillingPeriod { public int Year, Month; }
    class SubmissionInfo { public string RecipientCode, InstitutionName, ManagerName, Phone, BankName, AccountNumber, Round, IndustrialRate, Site; }
    static class AppSettings
    {
        static string SettingsPath{get{return TestStore202.FilePath("settings.ini");}}
        public static Dictionary<string,string> Load(){var result=new Dictionary<string,string>();try{string path=SettingsPath;if(!File.Exists(path)&&String.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("SOCIAL_INSURANCE_TEST_HOME")))path=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"살구아빠","4대보험급여검증기","settings.ini");if(!File.Exists(path))return result;foreach(string line in File.ReadAllLines(path,Encoding.UTF8)){int p=line.IndexOf('=');if(p<=0)continue;try{result[line.Substring(0,p)]=Encoding.UTF8.GetString(Convert.FromBase64String(line.Substring(p+1)));}catch{}}}catch{}return result;}
        public static void Save(Dictionary<string,string> values){try{File.WriteAllLines(SettingsPath,values.Select(x=>x.Key+"="+Convert.ToBase64String(Encoding.UTF8.GetBytes(x.Value??""))).ToArray(),Encoding.UTF8);}catch{}}
    }

    static class Processor
    {
        static readonly string[] NameAliases={"성명","가입자명","근로자명","피보험자명","이름"};
        static readonly string[] BirthAliases={"생년월일","주민등록번호","외국인등록번호","주민번호"};
        static readonly string[] JobAliases={"직종","공무직급여직종","시도직종","직급","직종명","고용형태"};
        public static void Run(InputSet input,string output)
        {
            string tempRoot=Path.Combine(Path.GetTempPath(),"InsurancePayrollValidator_"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(tempRoot);
            try
            {
                List<Recognition> log=new List<Recognition>();List<InputFile> originals=new List<InputFile>();
                List<InputFile> payroll=ExpandInput(input.PayrollPackage,tempRoot,"급여대장통합",log),gov=ExpandInput(input.GovPayroll,tempRoot,"공무원급여",log),special=ExpandInput(input.WorkerPayrollSpecial,tempRoot,"교특급여",log),school=ExpandInput(input.WorkerPayrollSchool,tempRoot,"학회급여",log),shortTerm=ExpandInput(input.ShortTermPayroll,tempRoot,"단기기간제",log);
                List<InputFile> health=ExpandInput(input.HealthGov,tempRoot,"건강보험",log);health.AddRange(ExpandInput(input.HealthOther,tempRoot,"건강보험",log));List<InputFile> pension=ExpandInput(input.Pension,tempRoot,"국민연금",log),employment=ExpandInput(input.Employment,tempRoot,"고용보험",log),industrial=ExpandInput(input.Industrial,tempRoot,"산재보험",log);
                originals.AddRange(payroll);originals.AddRange(gov);originals.AddRange(special);originals.AddRange(school);originals.AddRange(shortTerm);originals.AddRange(health);originals.AddRange(pension);originals.AddRange(employment);originals.AddRange(industrial);
                BillingPeriod period=DetectBillingPeriod(health.Concat(pension).Concat(employment).Concat(industrial).Select(x=>x.Path));
                Dictionary<string,Person> people=new Dictionary<string,Person>();
                foreach(InputFile f in payroll)ReadAutoClassifiedPayroll(f,people,log);
                foreach(InputFile f in gov)ReadPayroll(f.Path,"공무원","공무원",people,log);foreach(InputFile f in special)ReadPayroll(f.Path,"교육공무직(교특)","교특",people,log);foreach(InputFile f in school)ReadPayroll(f.Path,"교육공무직(학회)","학회(교육공무직)",people,log);foreach(InputFile f in shortTerm)ReadShortTermPayroll(f.Path,people,log);
                Dictionary<string,Charge> charges=new Dictionary<string,Charge>();
                foreach(InputFile f in health)ReadCharges(f.Path,"건강보험","건강보험",charges,log);
                foreach(InputFile f in pension)ReadCharges(f.Path,"국민연금","국민연금",charges,log);foreach(InputFile f in employment)ReadCharges(f.Path,"고용보험","고용보험",charges,log);foreach(InputFile f in industrial)ReadCharges(f.Path,"산재보험","산재보험",charges,log);
                InferMissingWorkplaceNumbers(charges,log);
                List<ResultRow> results=Compare(people,charges);AssignMissingResultWorkplaces(results,log);WriteOutput(output,results,log,people.Count,charges.Count,originals,period);
            }
            finally{try{Directory.Delete(tempRoot,true);}catch{}}
        }

        static void ReadAutoClassifiedPayroll(InputFile file,Dictionary<string,Person> people,List<Recognition> log)
        {
            string source,defaultFund,detail;
            if(!ClassifyPayroll(file.Path,out source,out defaultFund,out detail))
            {
                log.Add(new Recognition{Kind="급여대장 자동분류",File=Path.GetFileName(file.Path),State="확인필요",Detail=detail});return;
            }
            file.Kind=source;log.Add(new Recognition{Kind="급여대장 자동분류",File=Path.GetFileName(file.Path),State="정상",Detail=source+"으로 분류: "+detail});
            ReadPayroll(file.Path,source,defaultFund,people,log);
        }

        static bool ClassifyPayroll(string path,out string source,out string defaultFund,out string detail)
        {
            source="";defaultFund="";detail="";
            try
            {
                using(ExcelPackage p=new ExcelPackage(new FileInfo(path)))
                {
                    string context=Path.GetFileNameWithoutExtension(path)+" "+String.Join(" ",p.Workbook.Worksheets.Select(x=>x.Name));
                    foreach(ExcelWorksheet ws in p.Workbook.Worksheets)
                    {
                        if(ws.Dimension==null)continue;int mr=Math.Min(ws.Dimension.End.Row,25),mc=Math.Min(ws.Dimension.End.Column,35);
                        for(int r=1;r<=mr;r++)for(int c=1;c<=mc;c++){string value=Text(ws.Cells[r,c].Value);if(value.Length>0)context+=" "+value;}
                    }
                    string n=Norm(context);
                    if(n.Contains("학교회계")||n.Contains("학회")){source="교육공무직(학회)";defaultFund="학회(교육공무직)";detail="학회/학교회계 표기";return true;}
                    if(n.Contains("공무원")||n.Contains("계약제교원")||n.Contains("기간제교원")){source="공무원";defaultFund="공무원";detail="공무원/계약제교원 표기";return true;}
                    if(n.Contains("교육비특별회계")||n.Contains("교특")||n.Contains("교육공무직")||n.Contains("공무직급여대장")||n.Contains("공무직원급여대장")){source="교육공무직(교특)";defaultFund="교특";detail="교특/교육공무직 급여대장 표기";return true;}
                    SheetInfo si=FindSheet(p,NameAliases,BirthAliases);
                    if(si!=null)
                    {
                        int jc=FindCol(si,JobAliases);if(jc>0){string jobs="";for(int r=si.HeaderRow+1;r<=Math.Min(si.Sheet.Dimension.End.Row,si.HeaderRow+80);r++)jobs+=" "+Text(si.Sheet.Cells[r,jc].Value);string j=Norm(jobs);if(j.Contains("교원")||j.Contains("교사")||j.Contains("시간강사")){source="공무원";defaultFund="공무원";detail="직종 표본으로 공무원 급여대장 판정";return true;}}
                    }
                }
                detail="파일명·시트명·머리글에서 공무원/교특/학회 구분 근거를 찾지 못함";return false;
            }
            catch(Exception ex){detail="급여대장 자동분류 실패: "+ex.Message;return false;}
        }

        static List<InputFile> ExpandInput(string path,string tempRoot,string kind,List<Recognition> log)
        {
            var result=new List<InputFile>();if(String.IsNullOrWhiteSpace(path))return result;
            if(!String.Equals(Path.GetExtension(path),".zip",StringComparison.OrdinalIgnoreCase)){result.Add(new InputFile{Kind=kind,Path=path});return result;}
            string folder=Path.Combine(tempRoot,SafeSheetName(kind)+"_"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(folder);int index=0;
            try
            {
                using(ZipArchive zip=ZipFile.OpenRead(path))foreach(ZipArchiveEntry entry in zip.Entries)
                {
                    string ext=Path.GetExtension(entry.Name);if(String.IsNullOrWhiteSpace(entry.Name)||entry.Name.StartsWith("~$")||!(ext.Equals(".xlsx",StringComparison.OrdinalIgnoreCase)||ext.Equals(".xlsm",StringComparison.OrdinalIgnoreCase)))continue;
                    string name=(++index).ToString("000")+"_"+Path.GetFileName(entry.Name);string target=Path.Combine(folder,name);using(Stream source=entry.Open())using(FileStream dest=new FileStream(target,FileMode.Create,FileAccess.Write,FileShare.None)){source.CopyTo(dest);}result.Add(new InputFile{Kind=kind,Path=target});
                }
                if(result.Count==0)log.Add(new Recognition{Kind=kind+" ZIP",File=Path.GetFileName(path),State="확인필요",Detail="ZIP 안에서 .xlsx 또는 .xlsm 파일을 찾지 못함"});
                else log.Add(new Recognition{Kind=kind+" ZIP",File=Path.GetFileName(path),Rows=result.Count.ToString(),State="정상",Detail="압축 해제한 엑셀 "+result.Count+"개"});
            }
            catch(Exception ex){log.Add(new Recognition{Kind=kind+" ZIP",File=Path.GetFileName(path),State="확인필요",Detail="ZIP 압축 해제 실패: "+ex.Message});}
            return result;
        }

        static BillingPeriod DetectBillingPeriod(IEnumerable<string> inputPaths)
        {
            var paths=inputPaths.Where(x=>!String.IsNullOrWhiteSpace(x)&&File.Exists(x)).ToArray();
            foreach(string path in paths)
            {
                BillingPeriod fromName=ParsePeriod(Path.GetFileNameWithoutExtension(path));if(fromName!=null)return fromName;
                try
                {
                    using(ExcelPackage p=new ExcelPackage(new FileInfo(path)))foreach(var ws in p.Workbook.Worksheets)
                    {
                        if(ws.Dimension==null)continue;int mr=Math.Min(ws.Dimension.End.Row,35),mc=Math.Min(ws.Dimension.End.Column,45);
                        for(int r=1;r<=mr;r++)for(int c=1;c<=mc;c++){object v=ws.Cells[r,c].Value;if(v is DateTime){DateTime d=(DateTime)v;if(d.Year>=2020&&d.Year<=2100)return new BillingPeriod{Year=d.Year,Month=d.Month};}BillingPeriod found=ParsePeriod(Text(v));if(found!=null)return found;}
                    }
                }catch{}
            }
            DateTime now=DateTime.Now;return new BillingPeriod{Year=now.Year,Month=now.Month};
        }
        static BillingPeriod ParsePeriod(string text)
        {
            if(String.IsNullOrWhiteSpace(text))return null;Match m=Regex.Match(text,@"(20\d{2})\s*[년./_-]\s*(0?[1-9]|1[0-2])\s*월?");if(m.Success)return new BillingPeriod{Year=Int32.Parse(m.Groups[1].Value),Month=Int32.Parse(m.Groups[2].Value)};
            m=Regex.Match(text,@"(?<!\d)(20\d{2})(0[1-9]|1[0-2])(?!\d)");if(m.Success)return new BillingPeriod{Year=Int32.Parse(m.Groups[1].Value),Month=Int32.Parse(m.Groups[2].Value)};return null;
        }

        static void ReadPayroll(string path,string source,string defaultFund,Dictionary<string,Person> people,List<Recognition> log)
        {
            if(String.IsNullOrWhiteSpace(path)){log.Add(new Recognition{Kind=source+" 급여대장",State="미선택",Detail="선택하지 않음"});return;}
            using(ExcelPackage p=new ExcelPackage(new FileInfo(path)))
            {
                SheetInfo si=FindSheet(p,NameAliases,BirthAliases); if(si==null){log.Add(Fail(source,path,"성명·생년월일 머리글을 찾지 못함"));return;}
                int nc=FindCol(si,NameAliases),bc=FindCol(si,BirthAliases),jc=FindCol(si,JobAliases);
                int hc=FindCol(si,new[]{"건강보험"}),lc=FindCol(si,new[]{"노인장기요양보험","장기요양보험"});
                int hh=FindCol(si,new[]{"건강보험휴직정산","건강보험휴직자정산"}),lh=FindCol(si,new[]{"장기요양휴직정산","노인장기요양보험휴직정산","장기요양보험휴직정산"});
                int hsum=FindCol(si,new[]{"건강보험기타정산합계","건강보험정산합계"}),lsum=FindCol(si,new[]{"장기요양보험기타정산합계","장기요양기타정산합계","장기요양보험정산합계"});
                int hy=FindCol(si,new[]{"건강보험연말정산"}),ly=FindCol(si,new[]{"장기요양연말정산","장기요양보험연말정산"});
                int pc=FindCol(si,new[]{"국민연금"}),ps=FindCol(si,new[]{"국민연금정산","국민연금소급","국민연금추납"});
                int ec=FindCol(si,new[]{"고용보험"}),eret=FindCol(si,new[]{"고용보험퇴직정산"}),ey=FindCol(si,new[]{"고용보험연말정산"}),es=FindCol(si,new[]{"고용보험정산","고용보험휴직정산","고용보험소급"});
                int ic=FindCol(si,new[]{"산재보험"}),iret=FindCol(si,new[]{"산재보험퇴직정산"}),iy=FindCol(si,new[]{"산재보험연말정산"}),iset=FindCol(si,new[]{"산재보험정산","산재보험휴직정산","산재보험소급"});
                int count=0;
                for(int r=si.HeaderRow+1;r<=si.Sheet.Dimension.End.Row;r++)
                {
                    string name=CleanName(si.Sheet.Cells[r,nc].Value); if(name.Length==0)continue; string birth=Birth6(si.Sheet.Cells[r,bc].Value);
                    string key=Key(name,birth); Person x;
                    string job=jc>0?Text(si.Sheet.Cells[r,jc].Value):"", newFund=FundOf(defaultFund,job);
                    if(!people.TryGetValue(key,out x)){x=new Person{Key=key,Name=name,Birth=birth,Source=source,Job=job,Fund=newFund};people[key]=x;}
                    else if(!x.Source.Contains(source)){x.Source += ", "+source;if(x.Fund!=newFund)x.Fund="분류필요";}
                    decimal healthCurrent=Num(si.Sheet,r,hc),longTermCurrent=Num(si.Sheet,r,lc);
                    decimal healthSettlementHealth=Num(si.Sheet,r,hh)+Num(si.Sheet,r,hsum)+Num(si.Sheet,r,hy),longTermSettlement=Num(si.Sheet,r,lh)+Num(si.Sheet,r,lsum)+Num(si.Sheet,r,ly);
                    decimal healthSet=healthSettlementHealth+longTermSettlement;
                    decimal pensionSet=Num(si.Sheet,r,ps),employmentSet=Num(si.Sheet,r,eret)+Num(si.Sheet,r,ey)+Num(si.Sheet,r,es),industrialSet=Num(si.Sheet,r,iret)+Num(si.Sheet,r,iy)+Num(si.Sheet,r,iset);
                    x.HealthCurrent+=healthCurrent;x.LongTermCurrent+=longTermCurrent;x.HealthSettlementHealth+=healthSettlementHealth;x.LongTermSettlement+=longTermSettlement;x.HasHealthComponents|=hc>0||lc>0||hh>0||lh>0||hsum>0||lsum>0||hy>0||ly>0;
                    x.Health+=healthCurrent+longTermCurrent+healthSet;x.HealthSettlement+=healthSet;
                    x.Pension+=Num(si.Sheet,r,pc)+pensionSet;x.PensionSettlement+=pensionSet;
                    x.Employment+=Num(si.Sheet,r,ec)+employmentSet;x.EmploymentSettlement+=employmentSet;
                    x.Industrial+=Num(si.Sheet,r,ic)+industrialSet;x.IndustrialSettlement+=industrialSet;count++;
                }
                log.Add(Ok(source+" 급여대장",path,si,count,"공제열: 건강 "+Cols(hc,lc,hh,lh,hsum,lsum,hy,ly)+", 국민 "+Cols(pc,ps)+", 고용 "+Cols(ec,eret,ey,es)+", 산재 "+Cols(ic,iret,iy,iset)));
            }
        }

        static void ReadShortTermPayroll(string path,Dictionary<string,Person> people,List<Recognition> log)
        {
            if(String.IsNullOrWhiteSpace(path)){log.Add(new Recognition{Kind="단기기간제 근로자 신청",State="미선택",Detail="선택하지 않음"});return;}
            using(ExcelPackage p=new ExcelPackage(new FileInfo(path)))
            {
                SheetInfo si=FindSheet(p,NameAliases,JobAliases);
                if(si==null){log.Add(Fail("단기기간제 근로자 신청",path,"성명·직종 머리글을 찾지 못함"));return;}
                int nc=FindCol(si,NameAliases),jc=FindCol(si,JobAliases),codeCol=FindCol(si,new[]{"수신자기호"}),reasonCol=FindCol(si,new[]{"기간제근무사유","근무사유원근로자현황","대체근무사유","대체기간"});
                int hc=FindCol(si,new[]{"개인부담금건강보험","건강보험개인부담금"}),lc=FindCol(si,new[]{"개인부담금노인장기요양보험","개인부담금장기요양보험"});
                int pc=FindCol(si,new[]{"개인부담금국민연금","국민연금개인부담금"}),ec=FindCol(si,new[]{"개인부담금고용보험필수공제","개인부담금고용보험","고용보험개인부담금"});
                if(nc==0||jc==0){log.Add(Fail("단기기간제 근로자 신청",path,"성명 또는 직종 열을 찾지 못함"));return;}
                int count=0,blankRun=0,maxRow=Math.Min(si.Sheet.Dimension.End.Row,si.HeaderRow+5000);
                for(int r=si.HeaderRow+1;r<=maxRow;r++)
                {
                    string name=CleanName(si.Sheet.Cells[r,nc].Value);
                    if(name.Length==0){if(++blankRun>=200)break;continue;}blankRun=0;
                    string code=codeCol>0?Text(si.Sheet.Cells[r,codeCol].Value):"";
                    if(Norm(code).Contains("예시")||Norm(name)=="성명"||Norm(name)=="합계"||Norm(name)=="총계")continue;
                    string job=jc>0?Text(si.Sheet.Cells[r,jc].Value):"단기기간제 근로자",key=Key(name,"");Person x;
                    if(!people.TryGetValue(key,out x)){x=new Person{Key=key,Name=name,Birth="",Source="단기기간제 근로자",Job=job,Fund="교특(일용)"};people[key]=x;}
                    else{x.Source=x.Source.Contains("단기기간제 근로자")?x.Source:x.Source+", 단기기간제 근로자";x.Fund="교특(일용)";if(String.IsNullOrWhiteSpace(x.Job))x.Job=job;}
                    string reason=reasonCol>0?Text(si.Sheet.Cells[r,reasonCol].Value):"";if(reason.Length>0&&!String.Equals(x.Reason,reason,StringComparison.Ordinal)){if(String.IsNullOrWhiteSpace(x.Reason))x.Reason=reason;else if(!x.Reason.Contains(reason))x.Reason+=" / "+reason;}
                    decimal healthCurrent=Num(si.Sheet,r,hc),longTermCurrent=Num(si.Sheet,r,lc);x.HealthCurrent+=healthCurrent;x.LongTermCurrent+=longTermCurrent;x.HasHealthComponents|=hc>0||lc>0;
                    x.Health+=healthCurrent+longTermCurrent;x.Pension+=Num(si.Sheet,r,pc);x.Employment+=Num(si.Sheet,r,ec);count++;
                }
                log.Add(Ok("단기기간제 근로자 신청",path,si,count,"교특(일용) 분류 / 근무사유 "+Cols(reasonCol)+" / 개인공제열: 건강 "+Cols(hc,lc)+", 국민 "+Cols(pc)+", 고용 "+Cols(ec)));
            }
        }

        static void ReadCharges(string path,string insurance,string source,Dictionary<string,Charge> charges,List<Recognition> log)
        {
            if(String.IsNullOrWhiteSpace(path)){log.Add(new Recognition{Kind=source,State="미선택",Detail="선택하지 않음"});return;}
            using(ExcelPackage p=new ExcelPackage(new FileInfo(path)))
            {
                SheetInfo si=FindSheet(p,NameAliases,BirthAliases);if(si==null){log.Add(Fail(source,path,"성명·주민등록번호 머리글을 찾지 못함"));return;}
                string workplaceNumber=DetectWorkplaceNumber(p,si,path);
                int nc=FindCol(si,NameAliases),bc=FindCol(si,BirthAliases),accounting=FindCol(si,new[]{"회계","회계구분"}),compare=0,employer=0,extra=0,settlementPersonal=0,settlementEmployer=0,settlementEmployerExtra=0;
                int healthCurrent=0,longTermCurrent=0,healthSettlement=0,healthYearEnd=0,healthInterest=0,longTermSettlement=0,longTermYearEnd=0,longTermInterest=0;
                if(insurance=="건강보험")
                {
                    compare=FindCol(si,new[]{"가입자총납부할보험료","가입자부담금","가입자보험료"});settlementPersonal=FindCol(si,new[]{"정산보험료계(건강+요양)","정산보험료계건강요양"});employer=source=="건강보험(공무원)"?0:compare;settlementEmployer=source=="건강보험(공무원)"?0:settlementPersonal;
                    healthCurrent=FindCol(si,new[]{"고지금액"});if(healthCurrent==0)healthCurrent=FindCol(si,new[]{"산출보험료"});
                    longTermCurrent=FindCol(si,new[]{"요양고지보험료"});if(longTermCurrent==0)longTermCurrent=FindCol(si,new[]{"요양산출보험료"});
                    healthSettlement=FindCol(si,new[]{"정산금액"});healthYearEnd=FindCol(si,new[]{"연말정산"});healthInterest=FindCol(si,new[]{"건강환급금이자"});
                    longTermSettlement=FindCol(si,new[]{"요양정산보험료"});longTermYearEnd=FindCol(si,new[]{"요양연말정산보험료"});longTermInterest=FindCol(si,new[]{"요양환급금이자"});
                }
                if(insurance=="국민연금"){compare=FindCol(si,new[]{"총부담금계본인기여금원","본인기여금원","총부담금계본인기여금당월분만표기","당월분본인기여금"});employer=FindCol(si,new[]{"총부담금계사용자부담금원","사용자부담금원","총부담금계사용자부담금당월분만표기","당월분사용자부담금"});settlementPersonal=FindCol(si,new[]{"정산보험료본인기여금원","정산본인기여금원","소급분본인기여금"});settlementEmployer=FindCol(si,new[]{"정산보험료사용자부담금원","정산사용자부담금원","소급분사용자부담금"});}
                if(insurance=="고용보험"){compare=FindCol(si,new[]{"보험료합계①②③근로자실업급여보험료"});employer=FindCol(si,new[]{"보험료합계①②③사업주실업급여보험료"});extra=FindCol(si,new[]{"보험료합계①②③사업주고안직능보험료"});settlementPersonal=FindCol(si,new[]{"정산보험료③근로자실업급여보험료"});settlementEmployer=FindCol(si,new[]{"정산보험료③사업주실업급여보험료"});settlementEmployerExtra=FindCol(si,new[]{"정산보험료③사업주고안직능보험료"});}
                if(insurance=="산재보험"){employer=FindCol(si,new[]{"보험료합계①②③","보험료합계","산재보험료","사업주부담보험료"});settlementEmployer=FindCol(si,new[]{"정산보험료③","정산보험료"});compare=0;}
                bool employerRequired=!(insurance=="건강보험"&&source=="건강보험(공무원)");
                if((insurance!="산재보험"&&compare==0)||(employerRequired&&employer==0)){log.Add(Fail(source,path,"보험료 금액 열을 찾지 못함. 인식한 머리글: "+String.Join(", ",si.Headers.Values.Take(20))));return;}
                int count=0;
                for(int r=si.HeaderRow+1;r<=si.Sheet.Dimension.End.Row;r++)
                {
                    string name=CleanName(si.Sheet.Cells[r,nc].Value);if(name.Length==0||Norm(name)=="합계"||Norm(name)=="총계")continue;string birth=Birth6(si.Sheet.Cells[r,bc].Value);string rowSource=source;
                    if(insurance=="건강보험"&&source=="건강보험"){string accountCode=accounting>0?Regex.Replace(Text(si.Sheet.Cells[r,accounting].Value),"[^0-9]",""):"";rowSource=accountCode=="95"?"건강보험(공무원)":accountCode=="00"||accountCode=="0"?"건강보험(비공무원)":"건강보험";}
                    string personKey=Key(name,birth), key=personKey+"|"+insurance+"|"+workplaceNumber+"|"+rowSource;Charge x;
                    if(!charges.TryGetValue(key,out x)){x=new Charge{Key=personKey,Name=name,Birth=birth,Insurance=insurance,Source=rowSource,WorkplaceNumber=workplaceNumber};charges[key]=x;}
                    x.CompareAmount+=Num(si.Sheet,r,compare);x.EmployerAmount+=(rowSource=="건강보험(공무원)"?0:Num(si.Sheet,r,employer))+Num(si.Sheet,r,extra);
                    if(insurance=="건강보험"&&(healthCurrent>0||longTermCurrent>0))
                    {
                        decimal hc=Num(si.Sheet,r,healthCurrent),lc=Num(si.Sheet,r,longTermCurrent),hs=Num(si.Sheet,r,healthSettlement)+Num(si.Sheet,r,healthYearEnd)+Num(si.Sheet,r,healthInterest),ls=Num(si.Sheet,r,longTermSettlement)+Num(si.Sheet,r,longTermYearEnd)+Num(si.Sheet,r,longTermInterest);
                        x.HealthCurrent+=hc;x.LongTermCurrent+=lc;x.HealthSettlementHealth+=hs;x.LongTermSettlement+=ls;x.HasHealthComponents=true;x.SettlementPersonal+=hs+ls;
                        if(rowSource!="건강보험(공무원)"){x.EmployerHealthCurrent+=hc;x.EmployerLongTermCurrent+=lc;x.EmployerSettlementHealth+=hs;x.EmployerSettlementLongTerm+=ls;x.SettlementEmployer+=hs+ls;}
                    }
                    else{x.SettlementPersonal+=Num(si.Sheet,r,settlementPersonal);x.SettlementEmployer+=(rowSource=="건강보험(공무원)"?0:Num(si.Sheet,r,settlementEmployer))+Num(si.Sheet,r,settlementEmployerExtra);}
                    count++;
                }
                log.Add(Ok(source,path,si,count,"사업장번호 "+workplaceNumber+" / 비교금액 "+Cols(compare)+", 기관부담 "+Cols(employer,extra)+", 정산 개인 "+Cols(settlementPersonal)+", 정산 기관 "+Cols(settlementEmployer,settlementEmployerExtra)+(insurance=="건강보험"?" / 건강·장기요양 원자료 분리열 "+Cols(healthCurrent,longTermCurrent,healthSettlement,healthYearEnd,healthInterest,longTermSettlement,longTermYearEnd,longTermInterest):"")));
            }
        }

        static void InferMissingWorkplaceNumbers(Dictionary<string,Charge> charges,List<Recognition> log)
        {
            int linked=0,ambiguous=0,unresolved=0;
            foreach(var personGroup in charges.Values.GroupBy(x=>x.Key))
            {
                string[] known=personGroup.Where(x=>!IsMissingWorkplace(x.WorkplaceNumber)).Select(x=>x.WorkplaceNumber).Distinct().ToArray();
                foreach(Charge charge in personGroup.Where(x=>IsMissingWorkplace(x.WorkplaceNumber)&&(x.Insurance=="고용보험"||x.Insurance=="산재보험")))
                {
                    string[] preferred=personGroup.Where(x=>!IsMissingWorkplace(x.WorkplaceNumber)&&(x.Insurance=="건강보험"||x.Insurance=="국민연금"||x.Insurance=="고용보험")).Select(x=>x.WorkplaceNumber).Distinct().ToArray();
                    string[] candidates=preferred.Length>0?preferred:known;
                    if(candidates.Length==1){charge.WorkplaceNumber=candidates[0];linked++;}
                    else if(candidates.Length>1)ambiguous++;else unresolved++;
                }
            }
            log.Add(new Recognition{Kind="고용·산재 사업장번호 보완",Rows=linked.ToString(),State=(ambiguous+unresolved)==0?"정상":"확인필요",Detail="건강·국민연금·고용의 동일인 사업장번호로 "+linked+"건 보완 / 복수후보 "+ambiguous+"건 / 미확인 "+unresolved+"건"});
        }
        static bool IsMissingWorkplace(string value){return String.IsNullOrWhiteSpace(value)||value=="미확인";}

        static void AssignMissingResultWorkplaces(List<ResultRow> rows,List<Recognition> log)
        {
            int linkedByPerson=0,linkedByFund=0;
            foreach(var person in rows.GroupBy(ResultIdentityKey))
            {
                string[] known=person.Where(x=>!IsMissingWorkplace(x.WorkplaceNumber)).Select(x=>x.WorkplaceNumber).Distinct().ToArray();
                if(known.Length!=1)continue;
                foreach(ResultRow row in person.Where(x=>IsMissingWorkplace(x.WorkplaceNumber))){row.WorkplaceNumber=known[0];linkedByPerson++;}
            }
            string[] publicSites=rows.Where(x=>!IsMissingWorkplace(x.WorkplaceNumber)&&(x.Fund=="공무원"||x.ChargeSource=="건강보험(공무원)")).Select(x=>x.WorkplaceNumber).Distinct().ToArray();
            string[] otherSites=rows.Where(x=>!IsMissingWorkplace(x.WorkplaceNumber)&&x.Fund!="공무원"&&x.ChargeSource!="건강보험(공무원)").Select(x=>x.WorkplaceNumber).Distinct().ToArray();
            foreach(ResultRow row in rows.Where(x=>IsMissingWorkplace(x.WorkplaceNumber)))
            {
                string[] candidates=row.Fund=="공무원"?publicSites:otherSites;
                if(candidates.Length==1){row.WorkplaceNumber=candidates[0];linkedByFund++;}
            }
            int unresolved=rows.Count(x=>IsMissingWorkplace(x.WorkplaceNumber));
            log.Add(new Recognition{Kind="결과 사업장번호 보완",Rows=(linkedByPerson+linkedByFund).ToString(),State=unresolved==0?"정상":"확인필요",Detail="동일인 보험자료로 "+linkedByPerson+"건, 공무원/그 외 단일 사업장 기준으로 "+linkedByFund+"건 보완 / 미확인 "+unresolved+"건"});
        }

        static List<ResultRow> Compare(Dictionary<string,Person> people,Dictionary<string,Charge> charges)
        {
            List<ResultRow> rows=new List<ResultRow>(); HashSet<string> used=new HashSet<string>();
            foreach(Person p in people.Values.OrderBy(x=>x.Fund).ThenBy(x=>x.Name))
            {
                if(p.Source.Split(',').Any(x=>x.Trim()=="공무원"))
                {
                    bool hasOther=p.Pension!=0||p.Employment!=0;
                    foreach(string other in new[]{"국민연금","고용보험","산재보험"})if(charges.Values.Any(x=>x.Key==p.Key&&x.Insurance==other&&(x.CompareAmount!=0||x.EmployerAmount!=0))){hasOther=true;break;}
                    p.Fund=hasOther?"계약제교원":"공무원";
                }
                foreach(string ins in new[]{"건강보험","국민연금","고용보험","산재보험"})
                {
                    decimal ded=ins=="건강보험"?p.Health:ins=="국민연금"?p.Pension:ins=="고용보험"?p.Employment:p.Industrial;
                    decimal dedSet=ins=="건강보험"?p.HealthSettlement:ins=="국민연금"?p.PensionSettlement:ins=="고용보험"?p.EmploymentSettlement:p.IndustrialSettlement;
                    List<KeyValuePair<string,Charge>> matches=charges.Where(x=>x.Value.Insurance==ins&&(x.Value.Key==p.Key||(String.IsNullOrWhiteSpace(p.Birth)&&CleanName(x.Value.Name)==CleanName(p.Name)))).OrderBy(x=>x.Value.WorkplaceNumber).ToList();
                    if(matches.Count==0)
                    {
                        if(ded==0)continue;var missing=new ResultRow{Fund=p.Fund,Name=p.Name,Birth=p.Birth,Source=p.Source,Job=p.Job,Reason=p.Reason,Insurance=ins,WorkplaceNumber="미확인",Deduction=ded,DeductionSettlement=dedSet,Status="고지누락",Note="급여대장에는 있으나 보험 고지자료에서 찾지 못함"};
                        if(ins=="건강보험")SetPayrollHealthComponents(missing,p,p.HealthCurrent,p.LongTermCurrent,p.HealthSettlementHealth,p.LongTermSettlement,ded,dedSet);rows.Add(missing);continue;
                    }
                    decimal remainingDed=ded,remainingSet=dedSet,remainingDedHealth=p.HealthCurrent,remainingDedLongTerm=p.LongTermCurrent,remainingDedSettlementHealth=p.HealthSettlementHealth,remainingDedSettlementLongTerm=p.LongTermSettlement,totalWeight=matches.Sum(x=>Math.Abs(x.Value.CompareAmount));
                    for(int m=0;m<matches.Count;m++)
                    {
                        string dictionaryKey=matches[m].Key;Charge c=matches[m].Value;used.Add(dictionaryKey);
                        decimal allocatedDed=AllocateAcrossCharges(ded,ref remainingDed,m,matches.Count,c.CompareAmount,totalWeight),allocatedSet=AllocateAcrossCharges(dedSet,ref remainingSet,m,matches.Count,c.CompareAmount,totalWeight);
                        decimal allocatedDedHealth=0,allocatedDedLongTerm=0,allocatedDedSettlementHealth=0,allocatedDedSettlementLongTerm=0;
                        if(ins=="건강보험"&&p.HasHealthComponents)
                        {
                            allocatedDedHealth=AllocateAcrossCharges(p.HealthCurrent,ref remainingDedHealth,m,matches.Count,c.CompareAmount,totalWeight);allocatedDedLongTerm=AllocateAcrossCharges(p.LongTermCurrent,ref remainingDedLongTerm,m,matches.Count,c.CompareAmount,totalWeight);
                            allocatedDedSettlementHealth=AllocateAcrossCharges(p.HealthSettlementHealth,ref remainingDedSettlementHealth,m,matches.Count,c.CompareAmount,totalWeight);allocatedDedSettlementLongTerm=AllocateAcrossCharges(p.LongTermSettlement,ref remainingDedSettlementLongTerm,m,matches.Count,c.CompareAmount,totalWeight);
                        }
                        List<Person> sameChargePeople=people.Values.Where(x=>ChargeMatchesPerson(c,x)&&PersonHasInsuranceAmount(x,ins)).OrderBy(x=>IsShortTermPerson(x)?1:0).ThenBy(x=>x.Key).ToList();
                        decimal charge,employerAmount,settlementPersonalAmount,settlementEmployerAmount,chargeHealth=0,chargeLongTerm=0,settlementPersonalHealth=0,settlementPersonalLongTerm=0,employerHealth=0,employerLongTerm=0,settlementEmployerHealth=0,settlementEmployerLongTerm=0;
                        if(ins=="건강보험"&&c.HasHealthComponents)
                        {
                            chargeHealth=AllocateHealthComponent(c.HealthCurrent,c,p,sameChargePeople,0);chargeLongTerm=AllocateHealthComponent(c.LongTermCurrent,c,p,sameChargePeople,1);settlementPersonalHealth=AllocateHealthComponent(c.HealthSettlementHealth,c,p,sameChargePeople,2);settlementPersonalLongTerm=AllocateHealthComponent(c.LongTermSettlement,c,p,sameChargePeople,3);
                            employerHealth=AllocateHealthComponent(c.EmployerHealthCurrent,c,p,sameChargePeople,0);employerLongTerm=AllocateHealthComponent(c.EmployerLongTermCurrent,c,p,sameChargePeople,1);settlementEmployerHealth=AllocateHealthComponent(c.EmployerSettlementHealth,c,p,sameChargePeople,2);settlementEmployerLongTerm=AllocateHealthComponent(c.EmployerSettlementLongTerm,c,p,sameChargePeople,3);
                            charge=chargeHealth+chargeLongTerm+settlementPersonalHealth+settlementPersonalLongTerm;settlementPersonalAmount=settlementPersonalHealth+settlementPersonalLongTerm;employerAmount=employerHealth+employerLongTerm+settlementEmployerHealth+settlementEmployerLongTerm;settlementEmployerAmount=settlementEmployerHealth+settlementEmployerLongTerm;
                        }
                        else{charge=AllocateChargeValue(c.CompareAmount,c,p,sameChargePeople,ins,false);employerAmount=AllocateChargeValue(c.EmployerAmount,c,p,sameChargePeople,ins,false);settlementPersonalAmount=AllocateChargeValue(c.SettlementPersonal,c,p,sameChargePeople,ins,true);settlementEmployerAmount=AllocateChargeValue(c.SettlementEmployer,c,p,sameChargePeople,ins,true);}
                        decimal diff=charge-allocatedDed;string chargeSource=c.Source;
                        if(ins=="건강보험"&&c.Source=="건강보험"){chargeSource=p.Fund=="공무원"?"건강보험(공무원)":"건강보험(비공무원)";if(p.Fund=="공무원"){employerAmount=0;settlementEmployerAmount=0;employerHealth=0;employerLongTerm=0;settlementEmployerHealth=0;settlementEmployerLongTerm=0;}}
                        if(ins=="산재보험"){charge=0;diff=0;}string status=ins=="산재보험"?"부과확인":diff==0?"정상":diff>0?"추납":"환급";
                        string allocationNote=sameChargePeople.Count>1?"동일인이 일반 급여와 단기기간제 급여에 함께 있어 단기근로분과 일반·정산분을 분리 배분":matches.Count>1?"동일 보험의 여러 사업장 고지액 비율로 급여공제액 배분":"";
                        var result=new ResultRow{Fund=p.Fund,Name=p.Name,Birth=String.IsNullOrWhiteSpace(p.Birth)?c.Birth:p.Birth,Source=p.Source,Job=p.Job,Reason=p.Reason,Insurance=ins,ChargeSource=chargeSource,WorkplaceNumber=c.WorkplaceNumber,Deduction=allocatedDed,DeductionSettlement=allocatedSet,Charge=charge,Difference=diff,Employer=employerAmount,SettlementPersonal=settlementPersonalAmount,SettlementEmployer=settlementEmployerAmount,Status=status,Note=allocationNote};
                        if(ins=="건강보험")
                        {
                            SetPayrollHealthComponents(result,p,allocatedDedHealth,allocatedDedLongTerm,allocatedDedSettlementHealth,allocatedDedSettlementLongTerm,allocatedDed,allocatedSet);
                            result.HasHealthComponents|=c.HasHealthComponents;if(c.HasHealthComponents){result.ChargeHealth=chargeHealth;result.ChargeLongTerm=chargeLongTerm;result.SettlementPersonalHealth=settlementPersonalHealth;result.SettlementPersonalLongTerm=settlementPersonalLongTerm;result.EmployerHealth=employerHealth;result.EmployerLongTerm=employerLongTerm;result.SettlementEmployerHealth=settlementEmployerHealth;result.SettlementEmployerLongTerm=settlementEmployerLongTerm;}else SetChargeHealthComponents(result,charge,settlementPersonalAmount,employerAmount,settlementEmployerAmount);
                        }
                        rows.Add(result);
                    }
                }
            }
            foreach(var kv in charges)if(!used.Contains(kv.Key))
            {
                Charge c=kv.Value;var missing=new ResultRow{Fund="분류필요",Name=c.Name,Birth=c.Birth,Source=c.Source,Insurance=c.Insurance,ChargeSource=c.Source,WorkplaceNumber=c.WorkplaceNumber,Charge=c.CompareAmount,Difference=c.CompareAmount,Employer=c.EmployerAmount,SettlementPersonal=c.SettlementPersonal,SettlementEmployer=c.SettlementEmployer,Status="급여대장누락",Note="보험 고지자료에는 있으나 급여대장에서 찾지 못함"};
                if(c.Insurance=="건강보험"){missing.HasHealthComponents=c.HasHealthComponents;if(c.HasHealthComponents){missing.ChargeHealth=c.HealthCurrent;missing.ChargeLongTerm=c.LongTermCurrent;missing.SettlementPersonalHealth=c.HealthSettlementHealth;missing.SettlementPersonalLongTerm=c.LongTermSettlement;missing.EmployerHealth=c.EmployerHealthCurrent;missing.EmployerLongTerm=c.EmployerLongTermCurrent;missing.SettlementEmployerHealth=c.EmployerSettlementHealth;missing.SettlementEmployerLongTerm=c.EmployerSettlementLongTerm;}else SetChargeHealthComponents(missing,c.CompareAmount,c.SettlementPersonal,c.EmployerAmount,c.SettlementEmployer);}rows.Add(missing);
            }
            return rows;
        }

        static decimal AllocateAcrossCharges(decimal total,ref decimal remaining,int index,int count,decimal itemWeight,decimal totalWeight)
        {
            if(index==count-1)return remaining;decimal value=totalWeight>0?Math.Round(total*Math.Abs(itemWeight)/totalWeight,0,MidpointRounding.AwayFromZero):(index==0?total:0);remaining-=value;return value;
        }

        static void SetPayrollHealthComponents(ResultRow row,Person person,decimal healthCurrent,decimal longTermCurrent,decimal healthSettlement,decimal longTermSettlement,decimal total,decimal settlementTotal)
        {
            row.HasHealthComponents|=person.HasHealthComponents;if(person.HasHealthComponents){row.DeductionHealth=healthCurrent;row.DeductionLongTerm=longTermCurrent;row.DeductionSettlementHealth=healthSettlement;row.DeductionSettlementLongTerm=longTermSettlement;return;}
            SplitCombinedHealth(total-settlementTotal,out row.DeductionHealth,out row.DeductionLongTerm);SplitCombinedHealth(settlementTotal,out row.DeductionSettlementHealth,out row.DeductionSettlementLongTerm);
        }

        static void SetChargeHealthComponents(ResultRow row,decimal total,decimal settlementTotal,decimal employerTotal,decimal employerSettlementTotal)
        {
            SplitCombinedHealth(total-settlementTotal,out row.ChargeHealth,out row.ChargeLongTerm);SplitCombinedHealth(settlementTotal,out row.SettlementPersonalHealth,out row.SettlementPersonalLongTerm);SplitCombinedHealth(employerTotal-employerSettlementTotal,out row.EmployerHealth,out row.EmployerLongTerm);SplitCombinedHealth(employerSettlementTotal,out row.SettlementEmployerHealth,out row.SettlementEmployerLongTerm);
        }

        static bool ChargeMatchesPerson(Charge charge,Person person){return charge.Key==person.Key||(String.IsNullOrWhiteSpace(person.Birth)&&CleanName(charge.Name)==CleanName(person.Name));}
        static bool IsShortTermPerson(Person person){return person.Fund=="교특(일용)"||person.Fund=="학회(일용근로)"||(!String.IsNullOrWhiteSpace(person.Source)&&person.Source.Contains("단기기간제 근로자"));}
        static string ResultIdentityKey(ResultRow row){bool shortTerm=row.Fund=="교특(일용)"||row.Fund=="학회(일용근로)"||(!String.IsNullOrWhiteSpace(row.Source)&&row.Source.Contains("단기기간제 근로자"));return Key(row.Name,row.Birth)+(shortTerm?"|단기":"|일반");}
        static bool PersonHasInsuranceAmount(Person person,string insurance){return PersonInsuranceWeight(person,insurance)!=0||PersonSettlementAmount(person,insurance)!=0;}
        static decimal PersonInsuranceAmount(Person person,string insurance){return insurance=="건강보험"?person.Health:insurance=="국민연금"?person.Pension:insurance=="고용보험"?person.Employment:person.Industrial;}
        static decimal PersonInsuranceWeight(Person person,string insurance){decimal value=PersonInsuranceAmount(person,insurance);return insurance=="산재보험"&&value==0?person.Employment:value;}
        static decimal PersonSettlementAmount(Person person,string insurance){return insurance=="건강보험"?person.HealthSettlement:insurance=="국민연금"?person.PensionSettlement:insurance=="고용보험"?person.EmploymentSettlement:person.IndustrialSettlement;}
        static decimal AllocateChargeValue(decimal total,Charge charge,Person current,List<Person> people,string insurance,bool settlement)
        {
            if(people.Count<=1||total==0)return total;
            List<Person> eligible=settlement?people.Where(x=>!IsShortTermPerson(x)).ToList():people;
            if(eligible.Count==0)eligible=people;
            int currentIndex=eligible.FindIndex(x=>Object.ReferenceEquals(x,current));if(currentIndex<0)return 0;
            decimal weightTotal=eligible.Sum(x=>Math.Abs(settlement?PersonSettlementAmount(x,insurance):PersonInsuranceWeight(x,insurance)));
            decimal used=0;
            for(int i=0;i<eligible.Count;i++)
            {
                decimal share;
                if(i==eligible.Count-1)share=total-used;
                else if(weightTotal>0){decimal weight=Math.Abs(settlement?PersonSettlementAmount(eligible[i],insurance):PersonInsuranceWeight(eligible[i],insurance));share=Math.Round(total*weight/weightTotal,0,MidpointRounding.AwayFromZero);used+=share;}
                else{share=i==0?total:0;if(i==0)used=share;}
                if(i==currentIndex)return share;
            }
            return 0;
        }

        static decimal AllocateHealthComponent(decimal total,Charge charge,Person current,List<Person> people,int component)
        {
            if(people.Count<=1||total==0)return total;bool settlement=component>=2;List<Person> eligible=settlement?people.Where(x=>!IsShortTermPerson(x)).ToList():people;if(eligible.Count==0)eligible=people;int currentIndex=eligible.FindIndex(x=>Object.ReferenceEquals(x,current));if(currentIndex<0)return 0;
            Func<Person,decimal> weight=x=>component==0?x.HealthCurrent:component==1?x.LongTermCurrent:component==2?x.HealthSettlementHealth:x.LongTermSettlement;decimal weightTotal=eligible.Sum(x=>Math.Abs(weight(x))),used=0;
            for(int i=0;i<eligible.Count;i++){decimal share;if(i==eligible.Count-1)share=total-used;else if(weightTotal>0){share=Math.Round(total*Math.Abs(weight(eligible[i]))/weightTotal,0,MidpointRounding.AwayFromZero);used+=share;}else{share=i==0?total:0;if(i==0)used=share;}if(i==currentIndex)return share;}return 0;
        }

        static void WriteOutput(string path,List<ResultRow> rows,List<Recognition> log,int personCount,int chargeCount,List<InputFile> originals,BillingPeriod period)
        {
            using(Stream template=Assembly.GetExecutingAssembly().GetManifestResourceStream("InsurancePayrollValidator.ValidationTemplate.xlsx"))
            {
                if(template==null)throw new InvalidOperationException("내장 검증 결과 양식을 불러오지 못했습니다.");
                using(ExcelPackage p=new ExcelPackage(template))
                {
                    ExcelWorksheet summaryTemplate=p.Workbook.Worksheets.FirstOrDefault(x=>x.Name.StartsWith("검증결과(",StringComparison.OrdinalIgnoreCase)),individualTemplate=p.Workbook.Worksheets.FirstOrDefault(x=>x.Name.StartsWith("개인별내역(",StringComparison.OrdinalIgnoreCase)),reviewTemplate=p.Workbook.Worksheets.FirstOrDefault(x=>x.Name.StartsWith("확인명단(",StringComparison.OrdinalIgnoreCase));
                    if(summaryTemplate==null||individualTemplate==null||reviewTemplate==null)throw new InvalidOperationException("새 검증 결과 양식의 필수 탭 3개를 찾지 못했습니다.");
                    foreach(string supportName in new[]{"누락자","_MissingData","근무자별 부담금","UI총괄데이터","UI개인별데이터","선택목록","제출정보","자료인식"}){ExcelWorksheet old=p.Workbook.Worksheets[supportName];if(old!=null)p.Workbook.Worksheets.Delete(old);}
                    var miss=p.Workbook.Worksheets.Add("_MissingData");Dictionary<string,int> missingRows=WriteMissingSheet(miss,rows);
                    var person=p.Workbook.Worksheets.Add("근무자별 부담금");int personLast=WritePersonSheet(person,rows,missingRows);var uiSummary=p.Workbook.Worksheets.Add("UI총괄데이터");WriteUiSummarySheet(uiSummary,rows,period);var uiIndividual=p.Workbook.Worksheets.Add("UI개인별데이터");WriteUiIndividualSheet(uiIndividual,rows,period);
                    List<string> sites=rows.Where(x=>!IsMissingWorkplace(x.WorkplaceNumber)).Select(x=>x.WorkplaceNumber).Distinct().OrderBy(x=>x).ToList();if(sites.Count==0)sites.Add("미확인");
                    var sheetSets=new List<SiteSheetSet>();
                    for(int i=0;i<sites.Count;i++)
                    {
                        string site=sites[i];ExcelWorksheet summary,individual,review;
                        if(i==0){summary=summaryTemplate;individual=individualTemplate;review=reviewTemplate;}
                        else{summary=p.Workbook.Worksheets.Add(SiteSheetName("검증결과",site),summaryTemplate);individual=p.Workbook.Worksheets.Add(SiteSheetName("개인별내역",site),individualTemplate);review=p.Workbook.Worksheets.Add(SiteSheetName("확인명단",site),reviewTemplate);}
                        if(i==0){summary.Name=SiteSheetName("검증결과",site);individual.Name=SiteSheetName("개인별내역",site);review.Name=SiteSheetName("확인명단",site);}
                        sheetSets.Add(new SiteSheetSet{Site=site,Summary=summary,Individual=individual,Review=review});
                    }
                    List<ResultRow> unresolved=rows.Where(x=>IsMissingWorkplace(x.WorkplaceNumber)).ToList();
                    for(int i=0;i<sheetSets.Count;i++)
                    {
                        SiteSheetSet set=sheetSets[i];List<ResultRow> siteRows=set.Site=="미확인"?unresolved:rows.Where(x=>x.WorkplaceNumber==set.Site).ToList();
                        List<ResultRow> reviewRows=i==0&&set.Site!="미확인"&&unresolved.Count>0?siteRows.Concat(unresolved).ToList():siteRows;ReviewSheetResult reviewResult=WriteNewReviewSheet(set.Review,reviewRows);set.ReviewLast=reviewResult.Last;set.IndividualLast=WriteNewIndividualSheet(set.Individual,siteRows,set.Review,reviewResult.Rows,set.Summary);WriteNewSummarySheet(set.Summary,set.Individual,set.IndividualLast,siteRows,period);
                    }
                    WriteOverallSummary(p.Workbook.Worksheets["총괄"],sheetSets,period);
                    var info=p.Workbook.Worksheets.Add("제출정보");info.Cells[1,1].Value="보험 부과연도";info.Cells[1,2].Value=period.Year;info.Cells[2,1].Value="보험 부과월";info.Cells[2,2].Value=period.Month;info.Cells[3,1].Value="제출서 생성 기준";info.Cells[3,2].Value="전체 사업장 통합 / 교특·일용근로 / 계약제교원";
                    CopyInputSheets(p,originals,log);
                    var rec=p.Workbook.Worksheets.Add("자료인식");string[] rh={"자료종류","파일명","인식시트","머리글행","인식건수","상태","상세"};WriteHeader(rec,rh);int r=2;foreach(Recognition x in log){object[] v={x.Kind,x.File,x.Sheet,x.HeaderRow,x.Rows,x.State,x.Detail};for(int c=0;c<v.Length;c++)rec.Cells[r,c+1].Value=v[c];if(x.State!="정상"){rec.Cells[r,1,r,7].Style.Fill.PatternType=ExcelFillStyle.Solid;rec.Cells[r,1,r,7].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255,235,156));}r++;}Finish(rec,rh.Length,r-1);rec.Column(7).Width=65;
                    foreach(ExcelWorksheet ws in p.Workbook.Worksheets.Where(x=>x.Name=="_MissingData"||x.Name=="근무자별 부담금"||x.Name=="UI총괄데이터"||x.Name=="UI개인별데이터"||x.Name=="제출정보"))ws.Hidden=eWorkSheetHidden.Hidden;rec.Hidden=log.Any(x=>x.State!="정상")?eWorkSheetHidden.Visible:eWorkSheetHidden.Hidden;
                    AddTemplateInteractivity(p,sheetSets);
                    ApplyWorksheetTabColors(p);
                    ApplyShrinkToFit(p);
                    p.Workbook.CalcMode=ExcelCalcMode.Automatic;p.Workbook.FullCalcOnLoad=true;
                    p.SaveAs(new FileInfo(path));
                }
            }
        }

        static string SiteSheetName(string prefix,string site)
        {
            string clean=Regex.Replace(DisplaySiteNumber(site),@"[\\/:*?\[\]]","_");string name=prefix+"("+clean+")";return name.Length<=31?name:name.Substring(0,31);
        }

        static string DisplaySiteNumber(string site)
        {
            string digits=Regex.Replace(site??"","[^0-9]","");
            if(digits.Length>=11)return digits.Substring(5,6);
            if(digits.Length>6)return digits.Substring(digits.Length-6);
            return digits.Length>0?digits:"미확인";
        }

        static string FormattedSiteNumber(string site)
        {
            string digits=Regex.Replace(site??"","[^0-9]","");
            if(digits.Length>=11)return digits.Substring(0,3)+"-0"+digits.Substring(4,1)+"-"+digits.Substring(5,6);
            return digits.Length>0?digits:"미확인";
        }

        static void WriteOverallSummary(ExcelWorksheet ws,List<SiteSheetSet> sets,BillingPeriod period)
        {
            if(ws==null)return;ws.Cells[3,2].Value=period.Year+"년 "+period.Month+"월 고지분 사대사회보험 기관부담금 총괄표";
            foreach(string merged in ws.MergedCells.Where(x=>{var a=new ExcelAddress(x);return a.Start.Row>=8;}).ToList())ws.Cells[merged].Merge=false;
            int requiredLast=7+sets.Count*3;if(requiredLast>13)ws.InsertRow(14,requiredLast-13,8);for(int clearRow=8;clearRow<=Math.Max(13,requiredLast);clearRow++)for(int clearCol=2;clearCol<=10;clearCol++){ws.Cells[clearRow,clearCol].Value=null;ws.Cells[clearRow,clearCol].Formula="";}
            string[] funds={"교특","학회","계약제교원"};int row=8;
            foreach(SiteSheetSet set in sets)
            {
                int first=row,last=Math.Max(10,set.IndividualLast);string ir=SheetReference(set.Individual),fundRange=ir+"$C$10:$C$"+last;
                ws.Cells[row,2,row+2,2].Merge=true;ws.Cells[row,2].Style.Numberformat.Format="@";ws.Cells[row,2].Formula="=\""+FormattedSiteNumber(set.Site)+"\"";
                for(int i=0;i<3;i++,row++)
                {
                    for(int c=2;c<=10;c++)CopyBasicStyle(ws.Cells[8,c],ws.Cells[row,c]);
                    string fund=funds[i],criterion=(fund=="학회"||fund=="교특")?fund+"*":fund;ws.Cells[row,3].Formula="COUNTIF("+fundRange+",\""+criterion+"\")";ws.Cells[row,4].Value=fund;
                    ws.Cells[row,5].Formula="SUMIF("+fundRange+",\""+criterion+"\","+ir+"$H$10:$H$"+last+")+SUMIF("+fundRange+",\""+criterion+"\","+ir+"$J$10:$J$"+last+")";
                    ws.Cells[row,6].Formula="SUMIF("+fundRange+",\""+criterion+"\","+ir+"$L$10:$L$"+last+")+SUMIF("+fundRange+",\""+criterion+"\","+ir+"$N$10:$N$"+last+")";
                    ws.Cells[row,7].Formula="SUMIF("+fundRange+",\""+criterion+"\","+ir+"$P$10:$P$"+last+")";
                    ws.Cells[row,8].Formula="SUMIF("+fundRange+",\""+criterion+"\","+ir+"$R$10:$R$"+last+")+SUMIF("+fundRange+",\""+criterion+"\","+ir+"$T$10:$T$"+last+")";
                    ws.Cells[row,9].Formula="SUMIF("+fundRange+",\""+criterion+"\","+ir+"$U$10:$U$"+last+")+SUMIF("+fundRange+",\""+criterion+"\","+ir+"$V$10:$V$"+last+")";ws.Cells[row,10].Formula="SUM(E"+row+":I"+row+")";
                }
            }
            var overallTable=ws.Cells[7,2,requiredLast,10];overallTable.Style.Border.Top.Style=ExcelBorderStyle.Thin;overallTable.Style.Border.Bottom.Style=ExcelBorderStyle.Thin;overallTable.Style.Border.Left.Style=ExcelBorderStyle.Thin;overallTable.Style.Border.Right.Style=ExcelBorderStyle.Thin;ws.Cells[8,5,requiredLast,10].Style.Numberformat.Format="#,##0;[Red]-#,##0";overallTable.Style.HorizontalAlignment=ExcelHorizontalAlignment.Center;overallTable.Style.VerticalAlignment=ExcelVerticalAlignment.Center;
        }

        static void WriteNewSummarySheet(ExcelWorksheet ws,ExcelWorksheet individual,int individualLast,List<ResultRow> rows,BillingPeriod period)
        {
            ws.Cells[1,1].Value=period.Year+"년 "+period.Month+"월 고지분 사대사회보험 검증 결과";
            string sites=String.Join(", ",rows.Select(x=>x.WorkplaceNumber).Where(x=>!IsMissingWorkplace(x)).Distinct().OrderBy(x=>x));
            string displayedSites=sites.Length>0?String.Join(", ",sites.Split(',').Select(x=>FormattedSiteNumber(x.Trim()))):"미확인";ws.Cells[6,7].Style.Numberformat.Format="@";ws.Cells[6,7].Formula="=\""+displayedSites+"\"";
            string[] insurance={"건강보험","국민연금","고용보험","산재보험"};
            string ir=SheetReference(individual);int firstIndividual=10,lastIndividual=Math.Max(10,individualLast);
            string healthEmployer="SUM("+ir+"H"+firstIndividual+":H"+lastIndividual+")+SUM("+ir+"J"+firstIndividual+":J"+lastIndividual+")+SUM("+ir+"L"+firstIndividual+":L"+lastIndividual+")+SUM("+ir+"N"+firstIndividual+":N"+lastIndividual+")";
            string healthPersonal="SUM("+ir+"G"+firstIndividual+":G"+lastIndividual+")+SUM("+ir+"I"+firstIndividual+":I"+lastIndividual+")+SUM("+ir+"K"+firstIndividual+":K"+lastIndividual+")+SUM("+ir+"M"+firstIndividual+":M"+lastIndividual+")";
            string[] employerFormula={healthEmployer,"SUM("+ir+"P"+firstIndividual+":P"+lastIndividual+")","SUM("+ir+"R"+firstIndividual+":R"+lastIndividual+")+SUM("+ir+"T"+firstIndividual+":T"+lastIndividual+")","SUM("+ir+"U"+firstIndividual+":U"+lastIndividual+")+SUM("+ir+"V"+firstIndividual+":V"+lastIndividual+")"};
            string[] personalFormula={healthPersonal,"SUM("+ir+"O"+firstIndividual+":O"+lastIndividual+")","SUM("+ir+"Q"+firstIndividual+":Q"+lastIndividual+")+SUM("+ir+"S"+firstIndividual+":S"+lastIndividual+")","0"};
            ws.Cells[6,9].Value="보험구분";ws.Cells[6,10].Value="재원";ws.Cells[6,11].Value="감면적용";ws.Cells[6,12].Value="감면금액";
            for(int clearRow=7;clearRow<=11;clearRow++)for(int clearCol=9;clearCol<=12;clearCol++){ws.Cells[clearRow,clearCol].Value=null;ws.Cells[clearRow,clearCol].Formula="";}
            var discountTable=ws.Cells[6,9,10,12];discountTable.Style.Border.Top.Style=ExcelBorderStyle.Thin;discountTable.Style.Border.Bottom.Style=ExcelBorderStyle.Thin;discountTable.Style.Border.Left.Style=ExcelBorderStyle.Thin;discountTable.Style.Border.Right.Style=ExcelBorderStyle.Thin;
            var obsoleteDiscountRow=ws.Cells[11,9,11,12];obsoleteDiscountRow.Style.Border.Top.Style=ExcelBorderStyle.None;obsoleteDiscountRow.Style.Border.Bottom.Style=ExcelBorderStyle.None;obsoleteDiscountRow.Style.Border.Left.Style=ExcelBorderStyle.None;obsoleteDiscountRow.Style.Border.Right.Style=ExcelBorderStyle.None;
            for(int i=0;i<insurance.Length;i++)
            {
                int c=4+i,discountRow=7+i;
                ws.Cells[8,c].Value=rows.Where(x=>x.Insurance==insurance[i]).Sum(x=>x.Employer+x.Charge);
                ws.Cells[9,c].Formula=personalFormula[i];
                ws.Cells[10,c].Formula=employerFormula[i];
                ws.Cells[discountRow,9].Value=insurance[i];ws.Cells[discountRow,10].Value="";ws.Cells[discountRow,11].Value="";ws.Cells[discountRow,12].Value=0;
            }
            var summaryStatusAddresses=new HashSet<string>(new[]{"D8","E8","F8","G8"},StringComparer.OrdinalIgnoreCase);
            foreach(var oldRule in ws.ConditionalFormatting.Where(x=>x.Address!=null&&summaryStatusAddresses.Contains(x.Address.Address)).ToList())ws.ConditionalFormatting.Remove(oldRule);
            for(int c=4;c<=7;c++){int discountRow=c+3;string notice=ws.Cells[8,c].Address,personal=ws.Cells[9,c].Address,employer=ws.Cells[10,c].Address;var ok=ws.ConditionalFormatting.AddExpression(ws.Cells[8,c]);ok.Formula="ABS("+notice+"-("+personal+"+"+employer+"))=ABS($L$"+discountRow+")";ok.Style.Fill.PatternType=ExcelFillStyle.Solid;ok.Style.Fill.BackgroundColor.Color=Color.FromArgb(198,239,206);ok.Style.Font.Color.Color=Color.FromArgb(0,97,0);ok.Style.Font.Bold=true;}
            string[] funds={"공무원","계약제교원","교특","학회"};
            for(int i=0;i<funds.Length;i++)
            {
                int r=15+i;string criterion=funds[i];ws.Cells[r,2].Value=criterion=="교특"?"교육공무직(교특)":criterion=="학회"?"교육공무직(학회)":criterion;
                string fc=(criterion=="학회"||criterion=="교특")?criterion+"*":criterion,range=ir+"$C$"+firstIndividual+":$C$"+lastIndividual;
                ws.Cells[r,3].Formula="SUMIF("+range+",\""+fc+"\","+ir+"$H$"+firstIndividual+":$H$"+lastIndividual+")+SUMIF("+range+",\""+fc+"\","+ir+"$J$"+firstIndividual+":$J$"+lastIndividual+")+SUMIF("+range+",\""+fc+"\","+ir+"$L$"+firstIndividual+":$L$"+lastIndividual+")+SUMIF("+range+",\""+fc+"\","+ir+"$N$"+firstIndividual+":$N$"+lastIndividual+")";
                ws.Cells[r,4].Formula="SUMIF("+range+",\""+fc+"\","+ir+"$P$"+firstIndividual+":$P$"+lastIndividual+")";
                ws.Cells[r,5].Formula="SUMIF("+range+",\""+fc+"\","+ir+"$R$"+firstIndividual+":$R$"+lastIndividual+")+SUMIF("+range+",\""+fc+"\","+ir+"$T$"+firstIndividual+":$T$"+lastIndividual+")";
                ws.Cells[r,6].Formula="SUMIF("+range+",\""+fc+"\","+ir+"$U$"+firstIndividual+":$U$"+lastIndividual+")+SUMIF("+range+",\""+fc+"\","+ir+"$V$"+firstIndividual+":$V$"+lastIndividual+")";
                ws.Cells[r,7].Formula="SUM(C"+r+":F"+r+")";
            }
            foreach(string merged in ws.MergedCells.Where(x=>!String.IsNullOrWhiteSpace(x)).Where(x=>{var a=new ExcelAddress(x);return a.Start.Row<=22&&a.End.Row>=21&&a.Start.Column<=3&&a.End.Column>=2;}).ToList())ws.Cells[merged].Merge=false;ws.Cells[21,2].Value="개인별";ws.Cells[22,2].Value="전체";ws.Cells[21,3].Value=null;ws.Cells[22,3].Value=null;var selectorArea=ws.Cells[21,2,22,3];selectorArea.Style.Border.Top.Style=ExcelBorderStyle.None;selectorArea.Style.Border.Bottom.Style=ExcelBorderStyle.None;selectorArea.Style.Border.Left.Style=ExcelBorderStyle.None;selectorArea.Style.Border.Right.Style=ExcelBorderStyle.None;ws.Cells[21,2].Style.Border.Top.Style=ExcelBorderStyle.Thick;ws.Cells[21,2].Style.Border.Left.Style=ExcelBorderStyle.Thick;ws.Cells[21,2].Style.Border.Right.Style=ExcelBorderStyle.Thick;ws.Cells[21,2].Style.Border.Bottom.Style=ExcelBorderStyle.Thin;ws.Cells[22,2].Style.Border.Top.Style=ExcelBorderStyle.Thin;ws.Cells[22,2].Style.Border.Bottom.Style=ExcelBorderStyle.Thick;ws.Cells[22,2].Style.Border.Left.Style=ExcelBorderStyle.Thick;ws.Cells[22,2].Style.Border.Right.Style=ExcelBorderStyle.Thick;ws.Cells[21,2,22,2].Style.HorizontalAlignment=ExcelHorizontalAlignment.Center;ws.Cells[21,2,22,2].Style.VerticalAlignment=ExcelVerticalAlignment.Center;ws.Cells[22,4,22,12].Clear();ws.Cells[24,11,24,23].Clear();
            var groups=rows.GroupBy(ResultIdentityKey).OrderBy(g=>g.First().Fund).ThenBy(g=>g.First().Name).ToList();int start=25,lastData=ResizeTemplateDataRows(ws,start,groups.Count,62,13);
            for(int i=0;i<groups.Count;i++)
            {
                int r=start+i,sourceRow=firstIndividual+i;
                ws.Cells[r,2].Formula="IF(C"+r+"=\"\",\"\",SUBTOTAL(103,$C$"+start+":C"+r+"))";ws.Cells[r,3].Formula=ir+"E"+sourceRow;ws.Cells[r,4].Formula=ir+"F"+sourceRow;
                ws.Cells[r,5].Formula=ir+"H"+sourceRow+"+"+ir+"J"+sourceRow;ws.Cells[r,6].Formula=ir+"L"+sourceRow+"+"+ir+"N"+sourceRow;
                ws.Cells[r,7].Formula=ir+"P"+sourceRow;ws.Cells[r,8].Formula=ir+"R"+sourceRow+"+"+ir+"T"+sourceRow;ws.Cells[r,9].Formula=ir+"U"+sourceRow+"+"+ir+"V"+sourceRow;ws.Cells[r,10].Formula="SUM(E"+r+":I"+r+")";ws.Cells[r,13].Formula=ir+"C"+sourceRow;
            }
            for(int c=5;c<=10;c++){string col=ExcelCellAddress.GetColumnLetter(c);ws.Cells[24,c].Formula=groups.Count==0?"0":"SUBTOTAL(109,"+col+start+":"+col+lastData+")";}
            ws.Column(13).Hidden=true;ws.Cells[8,4,10,7].Style.Numberformat.Format="#,##0;[Red]-#,##0";ws.Cells[15,3,18,7].Style.Numberformat.Format="#,##0;[Red]-#,##0";if(groups.Count>0)ws.Cells[start,5,lastData,10].Style.Numberformat.Format="#,##0;[Red]-#,##0";ws.Cells[24,5,24,10].Style.Numberformat.Format="#,##0;[Red]-#,##0";ws.Cells[1,1,lastData,12].Style.HorizontalAlignment=ExcelHorizontalAlignment.Center;ws.Cells[1,1,lastData,12].Style.VerticalAlignment=ExcelVerticalAlignment.Center;
        }

        static void WriteNewSummarySheetLegacy(ExcelWorksheet ws,ExcelWorksheet individual,int individualLast,List<ResultRow> rows,BillingPeriod period)
        {
            ws.Cells[1,1].Value=period.Year+"년 "+period.Month+"월 고지분 사대사회보험 검증 결과";
            string sites=String.Join(", ",rows.Select(x=>x.WorkplaceNumber).Where(x=>!IsMissingWorkplace(x)).Distinct().OrderBy(x=>x));ws.Cells[6,6].Value=sites.Length>0?sites:"사업장관리번호 미확인";
            string[] insurance={"건강보험","국민연금","고용보험","산재보험"};string ir=SheetReference(individual);int firstIndividual=10,lastIndividual=Math.Max(10,individualLast);
            string healthEmployer="SUM("+ir+"G"+firstIndividual+":G"+lastIndividual+")+SUM("+ir+"I"+firstIndividual+":I"+lastIndividual+")+SUM("+ir+"K"+firstIndividual+":K"+lastIndividual+")+SUM("+ir+"M"+firstIndividual+":M"+lastIndividual+")";
            string healthPersonal="SUM("+ir+"F"+firstIndividual+":F"+lastIndividual+")+SUM("+ir+"H"+firstIndividual+":H"+lastIndividual+")+SUM("+ir+"J"+firstIndividual+":J"+lastIndividual+")+SUM("+ir+"L"+firstIndividual+":L"+lastIndividual+")";
            string[] employerFormula={healthEmployer,"SUM("+ir+"O"+firstIndividual+":O"+lastIndividual+")","SUM("+ir+"Q"+firstIndividual+":Q"+lastIndividual+")+SUM("+ir+"S"+firstIndividual+":S"+lastIndividual+")","SUM("+ir+"T"+firstIndividual+":T"+lastIndividual+")+SUM("+ir+"U"+firstIndividual+":U"+lastIndividual+")"};
            string[] personalFormula={healthPersonal,"SUM("+ir+"N"+firstIndividual+":N"+lastIndividual+")","SUM("+ir+"P"+firstIndividual+":P"+lastIndividual+")+SUM("+ir+"R"+firstIndividual+":R"+lastIndividual+")","0"};
            // 감면 적용은 "적용 재원 → 적용 대상자 → 감면금액" 순서로 선택한다.
            // I열은 감면 사유 등을 기록하는 참고 칸으로 남긴다.
            ws.Cells[6,9].Value="감면 재원";ws.Cells[6,10].Value="적용 재원";ws.Cells[6,11].Value="적용 대상자";ws.Cells[6,12].Value="감면금액";
            for(int i=0;i<insurance.Length;i++)
            {
                int c=3+i,discountRow=7+i;ws.Cells[8,c].Formula=employerFormula[i]+"-IF(AND(J"+discountRow+"<>\"\",K"+discountRow+"<>\"\"),L"+discountRow+",0)";ws.Cells[9,c].Formula=personalFormula[i];ws.Cells[10,c].Value=rows.Where(x=>x.Insurance==insurance[i]).Sum(x=>x.Employer+x.Charge);ws.Cells[11,c].Formula=ws.Cells[8,c].Address+"+"+ws.Cells[9,c].Address;
                ws.Cells[discountRow,8].Value=insurance[i];ws.Cells[discountRow,9].Value="";ws.Cells[discountRow,10].Value="";ws.Cells[discountRow,11].Value="";ws.Cells[discountRow,12].Value=0;
            }
            ws.Row(11).Height=ws.Row(10).Height;for(int c=2;c<=6;c++)CopyBasicStyle(ws.Cells[10,c],ws.Cells[11,c]);ws.Cells[10,2].Value="고지금액";ws.Cells[11,2].Value="납부금액";
            for(int c=3;c<=6;c++){var ok=ws.ConditionalFormatting.AddExpression(ws.Cells[11,c]);ok.Formula=ws.Cells[11,c].Address+"="+ws.Cells[10,c].Address;ok.Style.Fill.PatternType=ExcelFillStyle.Solid;ok.Style.Fill.BackgroundColor.Color=Color.FromArgb(198,239,206);ok.Style.Font.Color.Color=Color.FromArgb(0,97,0);ok.Style.Font.Bold=true;}
            string[] funds={"공무원","계약제교원","교특","학회"};
            for(int i=0;i<funds.Length;i++)
            {
                int r=15+i;string criterion=funds[i];ws.Cells[r,2].Value=criterion=="교특"?"교육공무직(교특)":criterion=="학회"?"교육공무직(학회)":criterion;
                string fc=criterion=="학회"?"학회*":criterion,range=ir+"$V$"+firstIndividual+":$V$"+lastIndividual,discount="SUMIFS($L$7:$L$10,$J$7:$J$10,\""+criterion+"\",$H$7:$H$10,";
                ws.Cells[r,3].Formula="SUMIF("+range+",\""+fc+"\","+ir+"$F$"+firstIndividual+":$F$"+lastIndividual+")+SUMIF("+range+",\""+fc+"\","+ir+"$G$"+firstIndividual+":$G$"+lastIndividual+")+SUMIF("+range+",\""+fc+"\","+ir+"$H$"+firstIndividual+":$H$"+lastIndividual+")+SUMIF("+range+",\""+fc+"\","+ir+"$I$"+firstIndividual+":$I$"+lastIndividual+")+SUMIF("+range+",\""+fc+"\","+ir+"$J$"+firstIndividual+":$J$"+lastIndividual+")+SUMIF("+range+",\""+fc+"\","+ir+"$K$"+firstIndividual+":$K$"+lastIndividual+")+SUMIF("+range+",\""+fc+"\","+ir+"$L$"+firstIndividual+":$L$"+lastIndividual+")+SUMIF("+range+",\""+fc+"\","+ir+"$M$"+firstIndividual+":$M$"+lastIndividual+")-"+discount+"\"건강보험\")";
                ws.Cells[r,4].Formula="SUMIF("+range+",\""+fc+"\","+ir+"$N$"+firstIndividual+":$N$"+lastIndividual+")+SUMIF("+range+",\""+fc+"\","+ir+"$O$"+firstIndividual+":$O$"+lastIndividual+")-"+discount+"\"국민연금\")";
                ws.Cells[r,5].Formula="SUMIF("+range+",\""+fc+"\","+ir+"$P$"+firstIndividual+":$P$"+lastIndividual+")+SUMIF("+range+",\""+fc+"\","+ir+"$Q$"+firstIndividual+":$Q$"+lastIndividual+")+SUMIF("+range+",\""+fc+"\","+ir+"$R$"+firstIndividual+":$R$"+lastIndividual+")+SUMIF("+range+",\""+fc+"\","+ir+"$S$"+firstIndividual+":$S$"+lastIndividual+")-"+discount+"\"고용보험\")";
                ws.Cells[r,6].Formula="SUMIF("+range+",\""+fc+"\","+ir+"$T$"+firstIndividual+":$T$"+lastIndividual+")+SUMIF("+range+",\""+fc+"\","+ir+"$U$"+firstIndividual+":$U$"+lastIndividual+")-"+discount+"\"산재보험\")";
                ws.Cells[r,7].Formula="SUM(C"+r+":F"+r+")";
            }
            ws.Cells[22,2].Value="전체";ws.Cells[22,3].Value=null;
            var groups=rows.GroupBy(ResultIdentityKey).OrderBy(g=>g.First().Fund).ThenBy(g=>g.First().Name).ToList();int start=25;
            PrepareTemplateRows(ws,start,Math.Max(1,groups.Count),8);
            for(int i=0;i<groups.Count;i++)
            {
                int r=start+i,sourceRow=firstIndividual+i;string name="C"+r,personalDiscount="SUMIFS($L$7:$L$10,$K$7:$K$10,"+name+",$H$7:$H$10,",totalHealth="("+ir+"G"+sourceRow+"+"+ir+"I"+sourceRow+"+"+ir+"K"+sourceRow+"+"+ir+"M"+sourceRow+"-"+personalDiscount+"\"건강보험\"))";ws.Cells[r,2].Formula="IF(C"+r+"=\"\",\"\",SUBTOTAL(103,$C$"+start+":C"+r+"))";ws.Cells[r,3].Formula=ir+"D"+sourceRow;ws.Cells[r,4].Formula=ir+"E"+sourceRow;ws.Cells[r,5].Formula="IF("+totalHealth+"=0,0,ROUNDDOWN(ABS("+totalHealth+")/1.1314/10,0)*10*SIGN("+totalHealth+"))";ws.Cells[r,6].Formula=totalHealth+"-E"+r;ws.Cells[r,7].Formula=ir+"Q"+sourceRow+"+"+ir+"S"+sourceRow+"-"+personalDiscount+"\"고용보험\")";ws.Cells[r,8].Formula=ir+"T"+sourceRow+"+"+ir+"U"+sourceRow+"-"+personalDiscount+"\"산재보험\")";ws.Cells[r,9].Formula=ir+"V"+sourceRow;ws.Cells[r,10].Formula="SUM(E"+r+":H"+r+")";
            }
            ws.Cells[24,10].Value="보험료 계";ws.Column(9).Hidden=true;ws.Cells[8,3,11,6].Style.Numberformat.Format="#,##0;[Red]-#,##0";ws.Cells[15,3,18,7].Style.Numberformat.Format="#,##0;[Red]-#,##0";if(groups.Count>0)ws.Cells[start,5,start+groups.Count-1,10].Style.Numberformat.Format="#,##0;[Red]-#,##0";ws.Cells[1,1,Math.Max(24,start+groups.Count-1),12].Style.HorizontalAlignment=ExcelHorizontalAlignment.Center;ws.Cells[1,1,Math.Max(24,start+groups.Count-1),12].Style.VerticalAlignment=ExcelVerticalAlignment.Center;
        }

        static string SheetReference(ExcelWorksheet ws){return "'"+ws.Name.Replace("'","''")+"'!";}

        static int WriteNewIndividualSheet(ExcelWorksheet ws,List<ResultRow> rows,ExcelWorksheet review,Dictionary<string,int> reviewRows,ExcelWorksheet summary)
        {
            string sites=String.Join(", ",rows.Select(x=>x.WorkplaceNumber).Where(x=>!IsMissingWorkplace(x)).Distinct().OrderBy(x=>x)),displayedSites=sites.Length>0?String.Join(", ",sites.Split(',').Select(x=>FormattedSiteNumber(x.Trim()))):"미확인";ws.Cells[6,5].Style.Numberformat.Format="@";ws.Cells[6,5].Formula="=\""+displayedSites+"\"";ws.Cells[5,2].Value="전체";
            var groups=rows.GroupBy(ResultIdentityKey).OrderBy(g=>g.First().Fund).ThenBy(g=>g.First().Name).ToList();int start=10,last=ResizeTemplateDataRows(ws,start,groups.Count,47,23);
            for(int i=0;i<groups.Count;i++)
            {
                var g=groups[i];ResultRow first=g.First();int r=start+i,reviewRow;ws.Row(r).Hidden=false;bool linked=reviewRows.TryGetValue(ResultIdentityKey(first),out reviewRow);ws.Cells[r,2].Formula="IF(E"+r+"=\"\",\"\",SUBTOTAL(103,$E$"+start+":E"+r+"))";ws.Cells[r,4].Value=first.Job;ws.Cells[r,5].Value=DisplayName(g,rows);ws.Cells[r,6].Value=first.Birth;
                ResultRow h=CombineInsurance(g,"건강보험"),p=CombineInsurance(g,"국민연금"),e=CombineInsurance(g,"고용보험"),ind=CombineInsurance(g,"산재보험");decimal hc=h.ChargeHealth,lc=h.ChargeLongTerm,hs=h.SettlementPersonalHealth,ls=h.SettlementPersonalLongTerm,hce=h.EmployerHealth,lce=h.EmployerLongTerm,hse=h.SettlementEmployerHealth,lse=h.SettlementEmployerLongTerm;
                object[] values={hc,hce,hs,hse,lc,lce,ls,lse,p.Charge,p.Employer,e.Charge-e.SettlementPersonal,e.Employer-e.SettlementEmployer,e.SettlementPersonal,e.SettlementEmployer,ind.Employer-ind.SettlementEmployer,ind.SettlementEmployer};for(int c=0;c<values.Length;c++)ws.Cells[r,7+c].Value=values[c];
                ws.Cells[r,3].Value=first.Fund;
                ws.Cells[r,23].Formula="=$C"+r;
                string sr=SheetReference(summary),person="$D"+r;
                person="$E"+r;
                ws.Cells[r,8].Formula="$AB"+r+"-SUMIFS("+sr+"$L$7:$L$10,"+sr+"$K$7:$K$10,"+person+","+sr+"$I$7:$I$10,\"건강보험\")";
                ws.Cells[r,16].Formula="$AC"+r+"-SUMIFS("+sr+"$L$7:$L$10,"+sr+"$K$7:$K$10,"+person+","+sr+"$I$7:$I$10,\"국민연금\")";
                ws.Cells[r,18].Formula="$AD"+r+"-SUMIFS("+sr+"$L$7:$L$10,"+sr+"$K$7:$K$10,"+person+","+sr+"$I$7:$I$10,\"고용보험\")";
                ws.Cells[r,21].Formula="$AE"+r+"-SUMIFS("+sr+"$L$7:$L$10,"+sr+"$K$7:$K$10,"+person+","+sr+"$I$7:$I$10,\"산재보험\")";
                ws.Cells[r,28].Value=hce;ws.Cells[r,29].Value=p.Employer;ws.Cells[r,30].Value=e.Employer-e.SettlementEmployer;ws.Cells[r,31].Value=ind.Employer-ind.SettlementEmployer;
            }
            if(groups.Count==0)ws.Cells[start,5].Value="해당 사업장 대상자가 없습니다.";ws.Column(23).Hidden=true;for(int c=28;c<=31;c++)ws.Column(c).Hidden=true;if(groups.Count>0)ws.Cells[start,7,start+groups.Count-1,22].Style.Numberformat.Format="#,##0;[Red]-#,##0";ApplyAlternatingFill(ws,start,last,2,22,Color.FromArgb(189,215,238));ws.Cells[1,1,last,22].Style.HorizontalAlignment=ExcelHorizontalAlignment.Center;ws.Cells[1,1,last,22].Style.VerticalAlignment=ExcelVerticalAlignment.Center;return last;
        }

        static ReviewSheetResult WriteNewReviewSheet(ExcelWorksheet ws,List<ResultRow> rows)
        {
            var result=new ReviewSheetResult();
            string sites=String.Join(", ",rows.Select(x=>x.WorkplaceNumber).Where(x=>!IsMissingWorkplace(x)).Distinct().OrderBy(x=>x));for(int staleCol=1;staleCol<=18;staleCol++){ws.Cells[3,staleCol].Value=null;ws.Cells[3,staleCol].Formula="";}foreach(int headerCol in new[]{2,3,4,5,6}){foreach(string merged in ws.MergedCells.Where(x=>!String.IsNullOrWhiteSpace(x)).Where(x=>{var a=new ExcelAddress(x);return a.Start.Column<=headerCol&&a.End.Column>=headerCol&&a.Start.Row<=7&&a.End.Row>=5;}).ToList())ws.Cells[merged].Merge=false;ws.Cells[5,headerCol,6,headerCol].Merge=true;}ws.Cells[4,2].Value="사업장관리번호: "+(sites.Length>0?String.Join(", ",sites.Split(',').Select(x=>FormattedSiteNumber(x.Trim()))):"미확인");ws.Column(2).Width=5.664;for(int reviewCol=3;reviewCol<=18;reviewCol++)ws.Column(reviewCol).Width=11.554;ws.Cells[15,2,16,18].Clear();
            var groups=rows.GroupBy(ResultIdentityKey).Where(g=>g.First().Fund=="분류필요"||g.Any(x=>x.Status!="정상"&&x.Status!="부과확인")||g.Any(x=>IsMissingWorkplace(x.WorkplaceNumber))||g.Any(x=>Math.Abs(x.SettlementPersonal-x.DeductionSettlement)>0.5m)).OrderBy(g=>g.First().Fund).ThenBy(g=>g.First().Name).ToList();int start=7,last=ResizeTemplateDataRows(ws,start,groups.Count,35,18);
            for(int i=0;i<groups.Count;i++)
            {
                var g=groups[i];ResultRow first=g.First();int r=start+i;ws.Row(r).Hidden=false;ws.Row(r).Height=18;result.Rows[ResultIdentityKey(first)]=r;ws.Cells[r,2].Value=i+1;ws.Cells[r,3].Value=first.Fund;ws.Cells[r,4].Value=DisplayName(g,rows);ws.Cells[r,5].Value=first.Birth;
                ResultRow h=CombineInsurance(g,"건강보험"),p=CombineInsurance(g,"국민연금"),e=CombineInsurance(g,"고용보험");
                decimal payrollH=h.DeductionHealth+h.DeductionSettlementHealth,payrollL=h.DeductionLongTerm+h.DeductionSettlementLongTerm,chargeH=h.ChargeHealth+h.SettlementPersonalHealth,chargeL=h.ChargeLongTerm+h.SettlementPersonalLongTerm;
                decimal payrollPensionTotal=g.Where(x=>x.Insurance=="국민연금").Sum(x=>x.Deduction),noticePensionTotal=p.Charge;
                decimal payrollEmploymentTotal=g.Where(x=>x.Insurance=="고용보험").Sum(x=>x.Deduction),noticeEmploymentTotal=e.Charge;
                object[] values={payrollH,chargeH,chargeH-payrollH,payrollL,chargeL,chargeL-payrollL,payrollPensionTotal,noticePensionTotal,noticePensionTotal-payrollPensionTotal,payrollEmploymentTotal,noticeEmploymentTotal,noticeEmploymentTotal-payrollEmploymentTotal};for(int c=0;c<values.Length;c++)ws.Cells[r,7+c].Value=values[c];
                string[] diffCells={"I","L","O","R"};var notices=new List<string>();if(first.Fund=="분류필요")notices.Add("분류 필요");decimal[] diffs={chargeH-payrollH,chargeL-payrollL,noticePensionTotal-payrollPensionTotal,noticeEmploymentTotal-payrollEmploymentTotal};string[] labels={"건강보험료","장기요양보험료","국민연금보험료","고용보험료"};for(int d=0;d<diffs.Length;d++)if(Math.Abs(diffs[d])>0.5m)notices.Add(labels[d]+" "+Math.Abs(diffs[d]).ToString("#,##0")+"원 "+(diffs[d]>0?"추징":"환급"));foreach(string insuranceName in new[]{"건강보험","국민연금","고용보험","산재보험"}){decimal settlementDiff=g.Where(x=>x.Insurance==insuranceName).Sum(x=>x.SettlementPersonal-x.DeductionSettlement);if(Math.Abs(settlementDiff)>0.5m)notices.Add(insuranceName+" 정산보험료 "+Math.Abs(settlementDiff).ToString("#,##0")+"원 "+(settlementDiff>0?"추징":"환급"));}ws.Cells[r,6].Value=String.Join(", ",notices.Distinct());
                foreach(string col in diffCells)AddDifferenceDirectionRules(ws,ws.Cells[col+r]);
            }
            if(groups.Count==0){ws.Cells[start,3].Value="확인 필요한 대상자가 없습니다.";}
            if(groups.Count>0)ws.Cells[start,7,start+groups.Count-1,18].Style.Numberformat.Format="#,##0;[Red]-#,##0";ApplyAlternatingFill(ws,start,last,2,18,Color.FromArgb(252,228,214));var reviewTable=ws.Cells[5,2,last,18];reviewTable.Style.Border.Top.Style=ExcelBorderStyle.Thin;reviewTable.Style.Border.Bottom.Style=ExcelBorderStyle.Thin;reviewTable.Style.Border.Left.Style=ExcelBorderStyle.Thin;reviewTable.Style.Border.Right.Style=ExcelBorderStyle.Thin;ws.Cells[5,2,5,18].Style.Border.Top.Style=ExcelBorderStyle.Thick;ws.Cells[last,2,last,18].Style.Border.Bottom.Style=ExcelBorderStyle.Thick;ws.Cells[5,2,last,2].Style.Border.Left.Style=ExcelBorderStyle.Thick;ws.Cells[5,18,last,18].Style.Border.Right.Style=ExcelBorderStyle.Thick;ws.Cells[last,6].Style.Border.Bottom.Style=ExcelBorderStyle.Double;ws.Cells[1,1,last,18].Style.HorizontalAlignment=ExcelHorizontalAlignment.Center;ws.Cells[1,1,last,18].Style.VerticalAlignment=ExcelVerticalAlignment.Center;result.Last=last;return result;
        }

        static void AddDifferenceHighlight(ExcelWorksheet ws,ExcelRange range,string formula){var rule=ws.ConditionalFormatting.AddExpression(range);rule.Formula=formula;rule.Style.Fill.PatternType=ExcelFillStyle.Solid;rule.Style.Fill.BackgroundColor.Color=Color.FromArgb(255,199,206);rule.Style.Font.Color.Color=Color.FromArgb(156,0,6);rule.Style.Font.Bold=true;}

        static void AddDifferenceDirectionRules(ExcelWorksheet ws,ExcelRange cell)
        {
            var red=ws.ConditionalFormatting.AddLessThan(cell);red.Formula="0";red.Style.Fill.PatternType=ExcelFillStyle.Solid;red.Style.Fill.BackgroundColor.Color=Color.FromArgb(255,199,206);red.Style.Font.Color.Color=Color.FromArgb(156,0,6);red.Style.Font.Bold=true;
            var blue=ws.ConditionalFormatting.AddGreaterThan(cell);blue.Formula="0";blue.Style.Fill.PatternType=ExcelFillStyle.Solid;blue.Style.Fill.BackgroundColor.Color=Color.FromArgb(189,215,238);blue.Style.Font.Color.Color=Color.FromArgb(31,78,121);blue.Style.Font.Bold=true;
        }

        static void ApplyAlternatingFill(ExcelWorksheet ws,int firstRow,int lastRow,int firstCol,int lastCol,Color color)
        {
            for(int r=firstRow;r<=lastRow;r++){var range=ws.Cells[r,firstCol,r,lastCol];range.Style.Fill.PatternType=ExcelFillStyle.Solid;range.Style.Fill.BackgroundColor.SetColor(((r-firstRow)%2)==1?color:Color.White);range.Style.Fill.BackgroundColor.Tint=0;}
        }

        static ResultRow CombineInsurance(IEnumerable<ResultRow> rows,string insurance)
        {
            var selected=rows.Where(x=>x.Insurance==insurance).ToList();return new ResultRow{Insurance=insurance,Charge=selected.Sum(x=>x.Charge==0&&x.SettlementPersonal!=0?x.SettlementPersonal:x.Charge),Employer=selected.Sum(x=>x.Employer),SettlementPersonal=selected.Sum(x=>x.SettlementPersonal),SettlementEmployer=selected.Sum(x=>x.SettlementEmployer),Deduction=selected.Sum(x=>x.Deduction),DeductionSettlement=selected.Sum(x=>x.DeductionSettlement),DeductionHealth=selected.Sum(x=>x.DeductionHealth),DeductionLongTerm=selected.Sum(x=>x.DeductionLongTerm),DeductionSettlementHealth=selected.Sum(x=>x.DeductionSettlementHealth),DeductionSettlementLongTerm=selected.Sum(x=>x.DeductionSettlementLongTerm),ChargeHealth=selected.Sum(x=>x.ChargeHealth),ChargeLongTerm=selected.Sum(x=>x.ChargeLongTerm),SettlementPersonalHealth=selected.Sum(x=>x.SettlementPersonalHealth),SettlementPersonalLongTerm=selected.Sum(x=>x.SettlementPersonalLongTerm),EmployerHealth=selected.Sum(x=>x.EmployerHealth),EmployerLongTerm=selected.Sum(x=>x.EmployerLongTerm),SettlementEmployerHealth=selected.Sum(x=>x.SettlementEmployerHealth),SettlementEmployerLongTerm=selected.Sum(x=>x.SettlementEmployerLongTerm),HasHealthComponents=selected.Any(x=>x.HasHealthComponents)};
        }
        static void SplitCombinedHealth(decimal total,out decimal health,out decimal longTerm){if(total==0){health=0;longTerm=0;return;}decimal sign=total<0?-1:1,abs=Math.Abs(total);health=Math.Floor((abs/1.1314m)/10m)*10m*sign;longTerm=total-health;}
        static string DisplayName(IEnumerable<ResultRow> group,List<ResultRow> allRows)
        {
            ResultRow first=group.First();bool daily=first.Fund=="교특(일용)"||first.Fund=="학회(일용근로)"||(!String.IsNullOrWhiteSpace(first.Source)&&first.Source.Contains("단기기간제 근로자"));bool overlap=allRows.GroupBy(ResultIdentityKey).Count(g=>Key(g.First().Name,g.First().Birth)==Key(first.Name,first.Birth))>1;return daily&&overlap?first.Name+"(일용)":first.Name;
        }
        static int ResizeSummaryPersonRows(ExcelWorksheet ws,int count)
        {
            const int start=25,templateCount=10,totalTemplateRow=35;
            if(count>templateCount)ws.InsertRow(totalTemplateRow,count-templateCount,totalTemplateRow-1);
            else if(count<templateCount)ws.DeleteRow(start+count,templateCount-count);
            int totalRow=start+count;
            for(int r=start;r<totalRow;r++)for(int c=2;c<=13;c++)ws.Cells[r,c].Value=null;
            return totalRow;
        }
        static int ResizeTemplateDataRows(ExcelWorksheet ws,int start,int requestedCount,int templateEnd,int visibleColumns)
        {
            int count=Math.Max(1,requestedCount),templateCount=templateEnd-start+1;
            if(count>templateCount)ws.InsertRow(templateEnd+1,count-templateCount,start);
            else if(count<templateCount)ws.DeleteRow(start+count,templateCount-count);
            for(int r=start;r<start+count;r++)for(int c=2;c<=Math.Max(visibleColumns,27);c++)ws.Cells[r,c].Value=null;
            return start+count-1;
        }
        static void PrepareTemplateRows(ExcelWorksheet ws,int start,int count,int visibleColumns)
        {
            int required=start+count-1;if(ws.Dimension!=null&&required>ws.Dimension.End.Row)for(int r=ws.Dimension.End.Row+1;r<=required;r++){ws.Row(r).Height=ws.Row(start).Height;for(int c=1;c<=visibleColumns;c++)CopyBasicStyle(ws.Cells[start,c],ws.Cells[r,c]);}
            int clearTo=Math.Max(required,ws.Dimension==null?required:ws.Dimension.End.Row);for(int r=start;r<=clearTo;r++)for(int c=2;c<=Math.Max(visibleColumns,22);c++)ws.Cells[r,c].Value=null;
        }

        static void AddTemplateInteractivity(ExcelPackage package,List<SiteSheetSet> sheetSets)
        {
            foreach(var staleName in package.Workbook.Names.Where(x=>x.Name.StartsWith("DiscountChoice_",StringComparison.OrdinalIgnoreCase)||x.Name.StartsWith("DiscountList_",StringComparison.OrdinalIgnoreCase)||x.Name.StartsWith("DiscountNames_",StringComparison.OrdinalIgnoreCase)).ToList())package.Workbook.Names.Remove(staleName.Name);
            ExcelWorksheet lists=package.Workbook.Worksheets.Add("선택목록");string[] funds=SummarySelectionCategories;for(int i=0;i<funds.Length;i++)lists.Cells[i+1,1].Value=funds[i];for(int i=0;i<FundSelectionCategories.Length;i++)lists.Cells[i+1,34].Value=FundSelectionCategories[i];
            var names=package.Workbook.Worksheets["근무자별 부담금"].Dimension==null?new List<string>():Enumerable.Range(2,package.Workbook.Worksheets["근무자별 부담금"].Dimension.End.Row-1).Select(r=>Convert.ToString(package.Workbook.Worksheets["근무자별 부담금"].Cells[r,2].Value)).Where(x=>!String.IsNullOrWhiteSpace(x)).Distinct().OrderBy(x=>x).ToList();lists.Cells[1,2].Value="전체";for(int i=0;i<names.Count;i++)lists.Cells[i+2,2].Value=names[i];string[] discountFunds={"공무원","계약제교원","교특","학회"};for(int i=0;i<discountFunds.Length;i++)lists.Cells[i+1,3].Value=discountFunds[i];lists.Hidden=eWorkSheetHidden.Hidden;
            if(package.Workbook.VbaProject==null)package.Workbook.CreateVBAProject();package.Workbook.VbaProject.Name="InsuranceValidator180";
            for(int setIndex=0;setIndex<sheetSets.Count;setIndex++)
            {
                RemoveExtendedDataValidations(sheetSets[setIndex].Summary);RemoveExtendedDataValidations(sheetSets[setIndex].Individual);RemoveExtendedDataValidations(sheetSets[setIndex].Review);sheetSets[setIndex].Summary.DataValidations.Clear();sheetSets[setIndex].Individual.DataValidations.Clear();sheetSets[setIndex].Review.DataValidations.Clear();
                SiteSheetSet set=sheetSets[setIndex];AddListValidation(set.Summary,set.Summary.Cells[7,10,10,10],"'선택목록'!$C$1:$C$"+discountFunds.Length);AddListValidation(set.Summary,set.Summary.Cells[22,2],"'선택목록'!$A$1:$A$"+funds.Length);AddListValidation(set.Individual,set.Individual.Cells[5,2],"'선택목록'!$A$1:$A$"+funds.Length);AddListValidation(set.Individual,set.Individual.Cells[10,3,set.IndividualLast,3],"'선택목록'!$AH$1:$AH$"+FundSelectionCategories.Length);AddListValidation(set.Review,set.Review.Cells[7,3,set.ReviewLast,3],"'선택목록'!$AH$1:$AH$"+FundSelectionCategories.Length);PruneDataValidations(set.Summary,"J7:J10","B22");PruneDataValidations(set.Individual,"B5",set.Individual.Cells[10,3,set.IndividualLast,3].Address);PruneDataValidations(set.Review,set.Review.Cells[7,3,set.ReviewLast,3].Address);
                int summaryLast=Math.Max(25,24+Math.Max(1,set.IndividualLast-9)),listBase=4+setIndex*4;string individualName=set.Individual.Name.Replace("\"","\"\"");
                set.Summary.CodeModule.Code="Option Explicit\r\nPrivate Sub Worksheet_Change(ByVal Target As Range)\r\n On Error GoTo SafeExit\r\n Application.EnableEvents = False\r\n If Not Intersect(Target,Me.Range(\"B22\")) Is Nothing Then FilterFund Me,25,"+summaryLast+",13,CStr(Me.Range(\"B22\").Value)\r\n Dim changed As Range\r\n If Not Intersect(Target,Me.Range(\"J7:J10\")) Is Nothing Then\r\n  For Each changed In Intersect(Target,Me.Range(\"J7:J10\"))\r\n   UpdateDiscountNameList Me,changed.Row,\""+individualName+"\",10,"+set.IndividualLast+","+listBase+"+(changed.Row-7)\r\n  Next changed\r\n End If\r\n If Not Intersect(Target,Me.Range(\"K7:K10\")) Is Nothing Then\r\n  For Each changed In Intersect(Target,Me.Range(\"K7:K10\"))\r\n   If Trim$(CStr(Me.Cells(changed.Row,10).Value)) = \"\" Then Me.Cells(changed.Row,11).ClearContents\r\n  Next changed\r\n End If\r\n If Not Intersect(Target,Me.Range(\"L7:L10\")) Is Nothing Then\r\n  For Each changed In Intersect(Target,Me.Range(\"L7:L10\"))\r\n   If Trim$(CStr(Me.Cells(changed.Row,10).Value)) = \"\" Or Trim$(CStr(Me.Cells(changed.Row,11).Value)) = \"\" Then\r\n    MsgBox \"재원과 감면 적용 대상자를 먼저 선택해 주세요.\", vbExclamation, \"감면 적용\"\r\n    Me.Cells(changed.Row,12).Value = 0\r\n   End If\r\n  Next changed\r\n End If\r\n Application.CalculateFull\r\nSafeExit:\r\n Application.EnableEvents = True\r\nEnd Sub\r\n";
                string reviewName=set.Review.Name.Replace("\"","\"\"");
                set.Individual.CodeModule.Code="Option Explicit\r\nPrivate Sub Worksheet_Change(ByVal Target As Range)\r\n On Error GoTo SafeExit\r\n Application.EnableEvents = False\r\n If Not Intersect(Target,Me.Range(\"B5\")) Is Nothing Then FilterFund Me,10,"+set.IndividualLast+",3,CStr(Me.Range(\"B5\").Value)\r\n If Not Intersect(Target,Me.Range(\"C10:C"+set.IndividualLast+"\")) Is Nothing Then SyncFundFromIndividual Me,ThisWorkbook.Worksheets(\""+reviewName+"\"),Intersect(Target,Me.Range(\"C10:C"+set.IndividualLast+"\")),7,"+set.ReviewLast+"\r\n Application.CalculateFull\r\nSafeExit:\r\n Application.EnableEvents = True\r\nEnd Sub\r\n";
                set.Review.CodeModule.Code="Option Explicit\r\nPrivate Sub Worksheet_Change(ByVal Target As Range)\r\n On Error GoTo SafeExit\r\n If Intersect(Target,Me.Range(\"C7:C"+set.ReviewLast+"\")) Is Nothing Then Exit Sub\r\n Application.EnableEvents = False\r\n SyncFundFromReview Me,ThisWorkbook.Worksheets(\""+individualName+"\"),Intersect(Target,Me.Range(\"C7:C"+set.ReviewLast+"\")),10,"+set.IndividualLast+"\r\n Application.CalculateFull\r\nSafeExit:\r\n Application.EnableEvents = True\r\nEnd Sub\r\n";
            }
            ExcelVBAModule module=package.Workbook.VbaProject.Modules.FirstOrDefault(x=>x.Name=="FundFilter");if(module==null)module=package.Workbook.VbaProject.Modules.AddModule("FundFilter");module.Code=@"Option Explicit
Public Sub FilterFund(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, ByVal fundCol As Long, ByVal selectedFund As String)
 Dim r As Long, value As String
 Application.ScreenUpdating = False
 For r = firstRow To lastRow
  value = CStr(ws.Cells(r, fundCol).Value)
   ws.Rows(r).Hidden = Not (selectedFund = ""전체"" Or value = selectedFund Or ((selectedFund = ""학회"" Or selectedFund = ""교특"") And Left$(value, 2) = selectedFund))
 Next r
 Application.ScreenUpdating = True
End Sub

Public Sub UpdateDiscountNameList(ByVal ws As Worksheet, ByVal discountRow As Long, ByVal personSheetName As String, ByVal firstRow As Long, ByVal lastRow As Long, ByVal listColumn As Long)
 Dim src As Worksheet, listWs As Worksheet, r As Long, outRow As Long, selectedFund As String, value As String, formulaText As String, rangeName As String
 Set src = ThisWorkbook.Worksheets(personSheetName)
 Set listWs = ThisWorkbook.Worksheets(""선택목록"")
 selectedFund = CStr(ws.Cells(discountRow, 10).Value)
 ws.Cells(discountRow, 11).ClearContents
 ws.Cells(discountRow, 12).Value = 0
 On Error Resume Next
 ws.Cells(discountRow, 11).Validation.Delete
 On Error GoTo 0
 listWs.Columns(listColumn).ClearContents
 If selectedFund = """" Then Exit Sub
 outRow = 1
 For r = firstRow To lastRow
  value = CStr(src.Cells(r, 3).Value)
   If value = selectedFund Or ((selectedFund = ""학회"" Or selectedFund = ""교특"") And Left$(value, 2) = selectedFund) Then
    listWs.Cells(outRow, listColumn).Value = src.Cells(r, 5).Value
   outRow = outRow + 1
  End If
 Next r
 If outRow > 1 Then
  formulaText = ""='선택목록'!$"" & Split(listWs.Cells(1, listColumn).Address, ""$"")(1) & ""$1:$"" & Split(listWs.Cells(1, listColumn).Address, ""$"")(1) & ""$"" & CStr(outRow - 1)
  rangeName = ""DiscountNames_"" & CStr(listColumn)
  On Error Resume Next
  ThisWorkbook.Names(rangeName).Delete
  On Error GoTo 0
  ThisWorkbook.Names.Add Name:=rangeName, RefersTo:=formulaText
  ws.Cells(discountRow, 11).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=""="" & rangeName
  ws.Cells(discountRow, 11).Validation.InCellDropdown = True
 End If
End Sub

Public Sub SyncFundFromReview(ByVal reviewWs As Worksheet, ByVal individualWs As Worksheet, ByVal changedCells As Range, ByVal individualFirstRow As Long, ByVal individualLastRow As Long)
 Dim changed As Range, r As Long, personName As String, birthValue As String
 For Each changed In changedCells.Cells
  personName = CStr(reviewWs.Cells(changed.Row, 4).Value)
  birthValue = CStr(reviewWs.Cells(changed.Row, 5).Value)
  For r = individualFirstRow To individualLastRow
   If CStr(individualWs.Cells(r, 5).Value) = personName And CStr(individualWs.Cells(r, 6).Value) = birthValue Then individualWs.Cells(r, 3).Value = changed.Value
  Next r
 Next changed
End Sub

Public Sub SyncFundFromIndividual(ByVal individualWs As Worksheet, ByVal reviewWs As Worksheet, ByVal changedCells As Range, ByVal reviewFirstRow As Long, ByVal reviewLastRow As Long)
 Dim changed As Range, r As Long, personName As String, birthValue As String
 For Each changed In changedCells.Cells
  personName = CStr(individualWs.Cells(changed.Row, 5).Value)
  birthValue = CStr(individualWs.Cells(changed.Row, 6).Value)
  For r = reviewFirstRow To reviewLastRow
   If CStr(reviewWs.Cells(r, 4).Value) = personName And CStr(reviewWs.Cells(r, 5).Value) = birthValue Then reviewWs.Cells(r, 3).Value = changed.Value
  Next r
 Next changed
End Sub";
        }
        static void AddListValidation(ExcelWorksheet ws,ExcelRange range,string formula){var validation=ws.DataValidations.AddListValidation(range.Address);validation.Formula.ExcelFormula=formula;validation.ShowErrorMessage=true;validation.Error="목록에서 값을 선택해 주세요.";}

        static void RemoveExtendedDataValidations(ExcelWorksheet ws)
        {
            XmlDocument xml=ws.WorksheetXml;XmlNodeList found=xml.SelectNodes("//*[local-name()='extLst']//*[local-name()='dataValidations']");if(found==null||found.Count==0)return;
            foreach(XmlNode node in found.Cast<XmlNode>().ToList())
            {
                XmlNode ext=node.ParentNode;ext.RemoveChild(node);
                if(!ext.ChildNodes.Cast<XmlNode>().Any(x=>x.NodeType==XmlNodeType.Element))
                {
                    XmlNode extList=ext.ParentNode;extList.RemoveChild(ext);
                    if(!extList.ChildNodes.Cast<XmlNode>().Any(x=>x.NodeType==XmlNodeType.Element))extList.ParentNode.RemoveChild(extList);
                }
            }
        }

        static void PruneDataValidations(ExcelWorksheet ws,params string[] allowedRanges)
        {
            HashSet<string> allowed=new HashSet<string>(allowedRanges,StringComparer.OrdinalIgnoreCase);XmlDocument xml=ws.WorksheetXml;XmlNodeList found=xml.SelectNodes("//*[local-name()='dataValidation']");if(found==null)return;
            foreach(XmlNode node in found.Cast<XmlNode>().ToList())
            {
                string range=node.Attributes["sqref"]==null?"":node.Attributes["sqref"].Value;if(range.Length==0){XmlNode sqref=node.SelectSingleNode(".//*[local-name()='sqref']");if(sqref!=null)range=sqref.InnerText;}
                if(!allowed.Contains(range))node.ParentNode.RemoveChild(node);
            }
            foreach(XmlNode container in xml.SelectNodes("//*[local-name()='dataValidations']").Cast<XmlNode>().ToList())
            {
                int count=container.ChildNodes.Cast<XmlNode>().Count(x=>x.NodeType==XmlNodeType.Element&&x.LocalName=="dataValidation");
                if(count==0)container.ParentNode.RemoveChild(container);else if(container.Attributes["count"]!=null)container.Attributes["count"].Value=count.ToString(CultureInfo.InvariantCulture);
            }
        }

        static void WriteValidationDetailSheet(ExcelWorksheet ws,List<ResultRow> rows,Dictionary<string,int> missingRows)
        {
            string[] baseHeaders={"재원구분","성명","생년월일","급여자료","직종"};
            for(int c=1;c<=baseHeaders.Length;c++){ws.Cells[1,c,2,c].Merge=true;ws.Cells[1,c].Value=baseHeaders[c-1];}
            string[] insurance={"건강보험","국민연금","고용보험","산재보험"};
            string[] subHeaders={"기관부담","개인부과","급여공제","차액"};
            for(int i=0;i<insurance.Length;i++)
            {
                int start=6+i*4;ws.Cells[1,start,1,start+3].Merge=true;ws.Cells[1,start].Value=insurance[i];
                for(int c=0;c<4;c++)ws.Cells[2,start+c].Value=subHeaders[c];
            }
            ws.Cells[1,22,2,22].Merge=true;ws.Cells[1,22].Value="종합판정";
            ws.Cells[1,23,2,23].Merge=true;ws.Cells[1,23].Value="비고";
            for(int i=0;i<insurance.Length;i++){int c=24+i;ws.Cells[1,c,2,c].Merge=true;ws.Cells[1,c].Value=insurance[i]+" 사업장번호";}
            var header=ws.Cells[1,1,2,27];header.Style.Font.Bold=true;header.Style.Font.Color.SetColor(Color.White);header.Style.Fill.PatternType=ExcelFillStyle.Solid;header.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(68,114,196));header.Style.HorizontalAlignment=ExcelHorizontalAlignment.Center;header.Style.VerticalAlignment=ExcelVerticalAlignment.Center;header.Style.Border.Top.Style=ExcelBorderStyle.Thin;header.Style.Border.Bottom.Style=ExcelBorderStyle.Thin;header.Style.Border.Left.Style=ExcelBorderStyle.Thin;header.Style.Border.Right.Style=ExcelBorderStyle.Thin;
            int r=3;
            foreach(var g in rows.GroupBy(x=>Key(x.Name,x.Birth)).OrderBy(x=>x.First().Fund).ThenBy(x=>x.First().Name))
            {
                ResultRow f=g.First();int missingRow;bool linked=missingRows.TryGetValue(Key(f.Name,f.Birth),out missingRow);
                if(linked)ws.Cells[r,1].Formula="'누락자'!A"+missingRow;else ws.Cells[r,1].Value=f.Fund;
                ws.Cells[r,2].Value=f.Name;ws.Cells[r,3].Value=f.Birth;ws.Cells[r,4].Value=String.Join(", ",g.Select(x=>x.Source).Where(x=>!String.IsNullOrWhiteSpace(x)).Distinct());ws.Cells[r,5].Value=f.Job;
                for(int i=0;i<insurance.Length;i++)
                {
                    int start=6+i*4;var items=g.Where(x=>x.Insurance==insurance[i]);
                    decimal rawDeduction=items.Sum(x=>x.Deduction);
                    if(linked)
                    {
                        int missingInstitution=6+i*2,missingPersonal=missingInstitution+1;
                        string fundCell="$A"+r,personalCell=ExcelCellAddress.GetColumnLetter(start+1)+r;
                        ws.Cells[r,start].Formula="IF("+fundCell+"=\"휴직\",0,'누락자'!"+ExcelCellAddress.GetColumnLetter(missingInstitution)+missingRow+")";
                        ws.Cells[r,start+1].Formula="IF("+fundCell+"=\"휴직\",0,'누락자'!"+ExcelCellAddress.GetColumnLetter(missingPersonal)+missingRow+")";
                        string raw=rawDeduction.ToString(CultureInfo.InvariantCulture);
                        if(f.Fund=="분류필요")ws.Cells[r,start+2].Formula="IF("+fundCell+"=\"휴직\",0,IF("+fundCell+"=\"분류필요\","+raw+","+personalCell+"))";
                        else ws.Cells[r,start+2].Formula="IF("+fundCell+"=\"휴직\",0,"+raw+")";
                        ws.Cells[r,start+3].Formula=personalCell+"-"+ExcelCellAddress.GetColumnLetter(start+2)+r;
                    }
                    else
                    {
                        ws.Cells[r,start].Value=items.Sum(x=>x.Employer);ws.Cells[r,start+1].Value=items.Sum(x=>x.Charge);ws.Cells[r,start+2].Value=rawDeduction;ws.Cells[r,start+3].Value=items.Sum(x=>x.Difference);
                    }
                }
                if(linked)
                {
                    string d1="I"+r,d2="M"+r,d3="Q"+r,d4="U"+r;
                    ws.Cells[r,22].Formula="IF($A"+r+"=\"휴직\",\"휴직 제외\",IF($A"+r+"=\"분류필요\",\"분류필요\",IF(ABS("+d1+")+ABS("+d2+")+ABS("+d3+")+ABS("+d4+")=0,\"정상\",IF(AND("+d1+">=0,"+d2+">=0,"+d3+">=0,"+d4+">=0),\"추납 확인\",IF(AND("+d1+"<=0,"+d2+"<=0,"+d3+"<=0,"+d4+"<=0),\"환급 확인\",\"차액 확인\")))))";
                }
                else
                {
                    string[] states=g.Select(x=>x.Status).Where(x=>!String.IsNullOrWhiteSpace(x)&&x!="정상"&&x!="부과확인").Distinct().ToArray();
                    string state=states.Length==0?"정상":String.Join(", ",states);ws.Cells[r,22].Value=state;ColorStatus(ws.Cells[r,22],state);
                }
                ws.Cells[r,23].Value=String.Join(", ",g.Select(x=>x.Note).Where(x=>!String.IsNullOrWhiteSpace(x)).Distinct());for(int i=0;i<insurance.Length;i++)ws.Cells[r,24+i].Value=String.Join(", ",g.Where(x=>x.Insurance==insurance[i]).Select(x=>x.WorkplaceNumber).Where(x=>!String.IsNullOrWhiteSpace(x)).Distinct());r++;
            }
            int last=Math.Max(3,r-1);ws.View.FreezePanes(3,1);ws.Cells[3,6,last,21].Style.Numberformat.Format="#,##0;[Red]-#,##0";
            ws.Cells[1,1,last,27].Style.VerticalAlignment=ExcelVerticalAlignment.Center;ws.Cells[1,1,last,27].Style.Border.Top.Style=ExcelBorderStyle.Thin;ws.Cells[1,1,last,27].Style.Border.Bottom.Style=ExcelBorderStyle.Thin;ws.Cells[1,1,last,27].Style.Border.Left.Style=ExcelBorderStyle.Thin;ws.Cells[1,1,last,27].Style.Border.Right.Style=ExcelBorderStyle.Thin;
            for(int c=1;c<=27;c++){ws.Column(c).AutoFit();if(ws.Column(c).Width>24)ws.Column(c).Width=24;}ws.Column(4).Width=28;ws.Column(23).Width=38;
        }

        static void ApplyWorksheetTabColors(ExcelPackage package)
        {
            Color blue=Color.FromArgb(68,114,196),orange=Color.FromArgb(237,125,49),gray=Color.FromArgb(165,165,165),yellow=Color.FromArgb(255,217,102),lightGreen=Color.FromArgb(198,224,180);
            foreach(ExcelWorksheet ws in package.Workbook.Worksheets)
            {
                if(ws.Name.StartsWith("원본_",StringComparison.OrdinalIgnoreCase))ws.TabColor=gray;
                else if(ws.Name=="작업내역")ws.TabColor=yellow;
                else if(ws.Name=="자료인식")ws.TabColor=yellow;
                else if(ws.Name.StartsWith("검증결과(",StringComparison.OrdinalIgnoreCase))ws.TabColor=lightGreen;
                else if(ws.Name.StartsWith("개인별내역(",StringComparison.OrdinalIgnoreCase))ws.TabColor=blue;
                else if(ws.Name.StartsWith("확인명단(",StringComparison.OrdinalIgnoreCase))ws.TabColor=orange;
                else if(ws.Name=="_MissingData")ws.TabColor=orange;
                else if(ws.Name=="근무자별 부담금"||ws.Name=="재원별총괄")ws.TabColor=Color.Empty;
                else ws.TabColor=blue;
            }
        }

        static void ApplyShrinkToFit(ExcelPackage package)
        {
            foreach(ExcelWorksheet ws in package.Workbook.Worksheets)if(ws.Dimension!=null){ExcelRange used=ws.Cells[ws.Dimension.Address];used.Style.WrapText=false;used.Style.ShrinkToFit=true;}
        }

        static void AddDifferenceDrilldown(ExcelPackage package,ExcelWorksheet dashboard)
        {
            package.Workbook.CreateVBAProject();package.Workbook.VbaProject.Name="InsuranceDifferenceViewer";
            dashboard.CodeModule.Code="Option Explicit\r\nPrivate Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)\r\n    If Intersect(Target, Me.Range(\"E5:E9\")) Is Nothing Then Exit Sub\r\n    Cancel = True\r\n    ShowDifferenceList CStr(Me.Cells(Target.Row, 2).Value)\r\nEnd Sub\r\n";
            ExcelVBAModule module=package.Workbook.VbaProject.Modules.AddModule("DifferenceViewer");
            module.Code=@"Option Explicit
Public Sub ShowDifferenceList(ByVal insuranceName As String)
    Dim src As Worksheet, dst As Worksheet, lastRow As Long, r As Long, outRow As Long
    Dim chargeCol As Long, deductionCol As Long, diff As Double, includeRow As Boolean
    On Error GoTo HandleError
    Application.ScreenUpdating = False
    Application.CalculateFull
    Set src = ThisWorkbook.Worksheets(""근무자별 부담금"")
    On Error Resume Next
    Set dst = ThisWorkbook.Worksheets(""개인차액명단"")
    On Error GoTo HandleError
    If dst Is Nothing Then
        Set dst = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        dst.Name = ""개인차액명단""
    Else
        dst.Cells.Clear
    End If
    dst.Tab.Color = RGB(68, 114, 196)
    Select Case insuranceName
        Case ""건강보험(공무원)"", ""건강보험(그 외)"": chargeCol = 11: deductionCol = 25
        Case ""국민연금"": chargeCol = 12: deductionCol = 26
        Case ""고용보험"": chargeCol = 13: deductionCol = 27
        Case ""산재보험"": chargeCol = 14: deductionCol = 28
        Case Else: GoTo CleanExit
    End Select
    dst.Range(""A1:J1"").Value = Array(""보험구분"", ""재원구분"", ""성명"", ""생년월일"", ""급여자료"", ""직종"", ""급여공제액"", ""개인부과액"", ""개인차액"", ""확인"")
    dst.Range(""A1:J1"").Font.Bold = True
    dst.Range(""A1:J1"").Interior.Color = RGB(68, 114, 196)
    dst.Range(""A1:J1"").Font.Color = RGB(255, 255, 255)
    lastRow = src.Cells(src.Rows.Count, 2).End(xlUp).Row
    outRow = 2
    For r = 2 To lastRow
        includeRow = (CStr(src.Cells(r, 1).Value) <> ""휴직"")
        If includeRow And Left$(insuranceName, 4) = ""건강보험"" Then includeRow = (CStr(src.Cells(r, 29).Value) = insuranceName)
        If includeRow Then
            diff = CDbl(Val(src.Cells(r, chargeCol).Value)) - CDbl(Val(src.Cells(r, deductionCol).Value))
            If Abs(diff) > 0.0001 Then
                dst.Cells(outRow, 1).Value = insuranceName
                dst.Cells(outRow, 2).Value = src.Cells(r, 1).Value
                dst.Cells(outRow, 3).Value = src.Cells(r, 2).Value
                dst.Cells(outRow, 4).Value = src.Cells(r, 3).Value
                dst.Cells(outRow, 5).Value = src.Cells(r, 4).Value
                dst.Cells(outRow, 6).Value = src.Cells(r, 5).Value
                dst.Cells(outRow, 7).Value = src.Cells(r, deductionCol).Value
                dst.Cells(outRow, 8).Value = src.Cells(r, chargeCol).Value
                dst.Cells(outRow, 9).Value = diff
                If diff > 0 Then dst.Cells(outRow, 10).Value = ""추납 확인"" Else dst.Cells(outRow, 10).Value = ""환급 확인""
                outRow = outRow + 1
            End If
        End If
    Next r
    If outRow = 2 Then dst.Cells(2, 1).Value = ""현재 남아 있는 개인차액 대상자가 없습니다.""
    dst.Columns(""A:J"").AutoFit
    dst.Columns(""G:I"").NumberFormat = ""#,##0;[Red]-#,##0""
    dst.Range(""A1:J1"").AutoFilter
    dst.Activate
    dst.Range(""A1"").Select
CleanExit:
    Application.ScreenUpdating = True
    Exit Sub
HandleError:
    Application.ScreenUpdating = True
    MsgBox ""개인차액 명단 생성 중 오류: "" & Err.Description, vbExclamation
End Sub
";
        }

        static void SetForceFullCalculation(string path)
        {
            using(FileStream fs=new FileStream(path,FileMode.Open,FileAccess.ReadWrite,FileShare.None))
            using(ZipArchive zip=new ZipArchive(fs,ZipArchiveMode.Update))
            {
                ZipArchiveEntry entry=zip.GetEntry("xl/workbook.xml");if(entry==null)return;string xml;
                using(StreamReader sr=new StreamReader(entry.Open(),Encoding.UTF8)){xml=sr.ReadToEnd();}
                if(Regex.IsMatch(xml,@"<calcPr\b[^>]*/>"))xml=Regex.Replace(xml,@"<calcPr\b[^>]*/>","<calcPr calcMode=\"auto\" fullCalcOnLoad=\"1\" forceFullCalc=\"1\"/>");
                else xml=xml.Replace("</workbook>","<calcPr calcMode=\"auto\" fullCalcOnLoad=\"1\" forceFullCalc=\"1\"/></workbook>");
                entry.Delete();ZipArchiveEntry replacement=zip.CreateEntry("xl/workbook.xml",System.IO.Compression.CompressionLevel.Optimal);using(StreamWriter sw=new StreamWriter(replacement.Open(),new UTF8Encoding(false))){sw.Write(xml);}
            }
        }

        static readonly string[] FundCategories={"공무원","계약제교원","교특(교육공무직)","교특(일용)","학회(교육공무직)","학회(강사)","학회(일용)","휴직","분류없음","분류필요"};
        static readonly string[] SummaryFundCategories={"공무원","계약제교원","교특","학회","기타급여","분류필요"};
        static readonly string[] FundSelectionCategories={"공무원","계약제교원","교특(교육공무직)","교특(일용)","학회(교육공무직)","학회(강사)","학회(일용)","휴직","분류없음"};
        static readonly string[] SummarySelectionCategories={"전체","공무원","계약제교원","교특","학회"};
        static void CopyInputSheets(ExcelPackage target,List<InputFile> items,List<Recognition> log)
        {
            foreach(var item in items)
            {
                if(String.IsNullOrWhiteSpace(item.Path)||!File.Exists(item.Path))continue;
                try
                {
                    using(ExcelPackage source=new ExcelPackage(new FileInfo(item.Path)))
                    {
                        int index=0;
                        foreach(ExcelWorksheet src in source.Workbook.Worksheets)
                        {
                            index++;string baseName="원본_"+item.Kind+(source.Workbook.Worksheets.Count>1?"_"+SafeSheetName(src.Name):"");string name=UniqueSheetName(target,baseName,index);
                            ExcelWorksheet dst=target.Workbook.Worksheets.Add(name);CopyWorksheetSnapshot(src,dst);
                        }
                    }
                }
                catch(Exception ex){log.Add(new Recognition{Kind="원본시트 복사: "+item.Kind,File=Path.GetFileName(item.Path),State="확인필요",Detail=ex.Message});}
            }
        }
        internal static string SafeSheetName(string name){string s=Regex.Replace(name??"",@"[\\/:*?\[\]]","_");return s.Length>18?s.Substring(0,18):s;}
        internal static string UniqueSheetName(ExcelPackage p,string baseName,int index){string n=baseName.Length>31?baseName.Substring(0,31):baseName;int k=index;while(p.Workbook.Worksheets.Any(x=>x.Name==n)){string suffix="_"+(k++);n=baseName.Substring(0,Math.Min(baseName.Length,31-suffix.Length))+suffix;}return n;}
        internal static void CopyWorksheetSnapshot(ExcelWorksheet src,ExcelWorksheet dst)
        {
            if(src.Dimension==null)return;int scanRows=Math.Min(src.Dimension.End.Row,10000),scanCols=Math.Min(src.Dimension.End.Column,250),maxRow=1,maxCol=1,blankRun=0;
            for(int r=1;r<=scanRows;r++)
            {
                bool hasValue=false;
                for(int c=1;c<=scanCols;c++){ExcelRange cell=src.Cells[r,c];if(cell.Value!=null||!String.IsNullOrWhiteSpace(cell.Formula)){hasValue=true;if(r>maxRow)maxRow=r;if(c>maxCol)maxCol=c;}}
                if(hasValue)blankRun=0;else if(maxRow>1&&++blankRun>=200)break;
            }
            for(int c=1;c<=maxCol;c++){dst.Column(c).Width=src.Column(c).Width;dst.Column(c).Hidden=src.Column(c).Hidden;}
            for(int r=1;r<=maxRow;r++)
            {
                dst.Row(r).Height=src.Row(r).Height;dst.Row(r).Hidden=src.Row(r).Hidden;
                for(int c=1;c<=maxCol;c++)
                {
                    ExcelRange from=src.Cells[r,c],to=dst.Cells[r,c];to.Value=from.Value;CopyBasicStyle(from,to);
                }
            }
            foreach(string merged in src.MergedCells)try{dst.Cells[merged].Merge=true;}catch{}
            dst.View.ShowGridLines=src.View.ShowGridLines;
        }
        static void CopyBasicStyle(ExcelRange from,ExcelRange to)
        {
            try{to.Style.Numberformat.Format=from.Style.Numberformat.Format;to.Style.HorizontalAlignment=from.Style.HorizontalAlignment;to.Style.VerticalAlignment=from.Style.VerticalAlignment;to.Style.WrapText=from.Style.WrapText;to.Style.ShrinkToFit=from.Style.ShrinkToFit;to.Style.Indent=from.Style.Indent;to.Style.TextRotation=from.Style.TextRotation;
                to.Style.Font.Name=from.Style.Font.Name;to.Style.Font.Size=from.Style.Font.Size;to.Style.Font.Bold=from.Style.Font.Bold;to.Style.Font.Italic=from.Style.Font.Italic;to.Style.Font.UnderLine=from.Style.Font.UnderLine;
                to.Style.Fill.PatternType=from.Style.Fill.PatternType;CopyColor(from.Style.Fill.BackgroundColor,to.Style.Fill.BackgroundColor);CopyColor(from.Style.Font.Color,to.Style.Font.Color);
                CopyBorder(from.Style.Border.Top,to.Style.Border.Top);CopyBorder(from.Style.Border.Bottom,to.Style.Border.Bottom);CopyBorder(from.Style.Border.Left,to.Style.Border.Left);CopyBorder(from.Style.Border.Right,to.Style.Border.Right);
            }catch{}
        }
        static void CopyBorder(ExcelBorderItem from,ExcelBorderItem to){try{to.Style=from.Style;CopyColor(from.Color,to.Color);}catch{}}
        static void CopyColor(ExcelColor from,ExcelColor to){try{if(!String.IsNullOrWhiteSpace(from.Rgb)){string s=from.Rgb;long n=Int64.Parse(s,NumberStyles.HexNumber);to.SetColor(Color.FromArgb((int)n));}}catch{}}
        static Dictionary<string,int> WriteMissingSheet(ExcelWorksheet ws,List<ResultRow> rows)
        {
            string[] single={"재원구분/제외","성명","생년월일","급여자료","직종"};for(int c=1;c<=single.Length;c++){ws.Cells[1,c,2,c].Merge=true;ws.Cells[1,c].Value=single[c-1];}
            string[] insurance={"건강보험","국민연금","고용보험","산재보험"};for(int i=0;i<insurance.Length;i++){int c=6+i*2;ws.Cells[1,c,1,c+1].Merge=true;ws.Cells[1,c].Value=insurance[i];ws.Cells[2,c].Value="기관";ws.Cells[2,c+1].Value="개인";}
            ws.Cells[1,14,2,14].Merge=true;ws.Cells[1,14].Value="반영상태";ws.Cells[1,15,2,15].Merge=true;ws.Cells[1,15].Value="확인내용";
            ws.Cells[1,16,2,16].Merge=true;ws.Cells[1,16].Value="사업장번호(보험별)";
            var header=ws.Cells[1,1,2,16];header.Style.Font.Bold=true;header.Style.Font.Color.SetColor(Color.White);header.Style.Fill.PatternType=ExcelFillStyle.Solid;header.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(68,114,196));header.Style.HorizontalAlignment=ExcelHorizontalAlignment.Center;header.Style.VerticalAlignment=ExcelVerticalAlignment.Center;header.Style.Border.Top.Style=ExcelBorderStyle.Thin;header.Style.Border.Bottom.Style=ExcelBorderStyle.Thin;header.Style.Border.Left.Style=ExcelBorderStyle.Thin;header.Style.Border.Right.Style=ExcelBorderStyle.Thin;
            var map=new Dictionary<string,int>();int r=3;
            var candidates=rows.GroupBy(ResultIdentityKey).Where(g=>g.First().Fund=="분류필요"||g.Any(x=>x.Status.Contains("누락"))).OrderBy(g=>g.First().Name);
            foreach(var g in candidates)
            {
                ResultRow f=g.First();string key=ResultIdentityKey(f);map[key]=r;
                ws.Cells[r,1].Value=f.Fund;ws.Cells[r,2].Value=f.Name;ws.Cells[r,3].Value=f.Birth;ws.Cells[r,4].Value=String.Join(", ",g.Select(x=>x.Source).Where(x=>!String.IsNullOrWhiteSpace(x)).Distinct());ws.Cells[r,5].Value=f.Job;
                ws.Cells[r,6].Value=g.Where(x=>x.Insurance=="건강보험").Sum(x=>x.Employer);ws.Cells[r,7].Value=g.Where(x=>x.Insurance=="건강보험").Sum(x=>x.Charge);
                ws.Cells[r,8].Value=g.Where(x=>x.Insurance=="국민연금").Sum(x=>x.Employer);ws.Cells[r,9].Value=g.Where(x=>x.Insurance=="국민연금").Sum(x=>x.Charge);
                ws.Cells[r,10].Value=g.Where(x=>x.Insurance=="고용보험").Sum(x=>x.Employer);ws.Cells[r,11].Value=g.Where(x=>x.Insurance=="고용보험").Sum(x=>x.Charge);
                ws.Cells[r,12].Value=g.Where(x=>x.Insurance=="산재보험").Sum(x=>x.Employer);ws.Cells[r,13].Value=g.Where(x=>x.Insurance=="산재보험").Sum(x=>x.Charge);
                ws.Cells[r,14].Formula="IF(A"+r+"=\"휴직\",\"총괄 제외\",\"총괄 반영\")";ws.Cells[r,15].Value=String.Join(", ",g.Select(x=>x.Status).Where(x=>x!="정상"&&x!="부과확인").Distinct());ws.Cells[r,16].Value=String.Join(", ",g.GroupBy(x=>x.Insurance).Select(x=>x.Key+": "+String.Join("/",x.Select(y=>y.WorkplaceNumber).Where(y=>!String.IsNullOrWhiteSpace(y)).Distinct())));r++;
            }
            int last=Math.Max(3,r-1);var fund=ws.DataValidations.AddListValidation("A3:A"+last);foreach(string c in FundCategories)fund.Formula.Values.Add(c);
            if(r==3){ws.Cells[3,1].Value="분류필요";ws.Cells[3,14].Formula="IF(A3=\"휴직\",\"총괄 제외\",\"총괄 반영\")";}
            ws.View.FreezePanes(3,1);ws.Cells[3,6,last,13].Style.Numberformat.Format="#,##0;[Red]-#,##0";
            ws.Cells[3,1,last,1].Style.Fill.PatternType=ExcelFillStyle.Solid;ws.Cells[3,1,last,1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255,242,204));ws.Cells[3,6,last,13].Style.Fill.PatternType=ExcelFillStyle.Solid;ws.Cells[3,6,last,13].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255,242,204));
            for(int c=1;c<=16;c++){ws.Column(c).AutoFit();if(ws.Column(c).Width>28)ws.Column(c).Width=28;}ws.Column(4).Width=28;ws.Column(15).Width=35;ws.Column(16).Width=35;ws.Cells.Style.VerticalAlignment=ExcelVerticalAlignment.Center;return map;
        }

        static int WritePersonSheet(ExcelWorksheet ws,List<ResultRow> rows,Dictionary<string,int> missingRows)
        {
            string[] h={"재원구분","성명","생년월일","급여자료","직종","건강보험 기관부담","국민연금 기관부담","고용보험 기관부담","산재보험 기관부담","기관부담 합계","건강보험 개인부과","국민연금 개인부과","고용보험 개인부과","산재보험 개인부과","개인부과 합계","총 납입예정액","확인상태","건강 정산 개인(포함액)","건강 정산 기관(포함액)","고용 정산 개인(포함액)","고용 정산 기관(포함액)","산재 정산 개인(포함액)","산재 정산 기관(포함액)"};WriteHeader(ws,h);
            string[] siteHeaders={"건강 사업장번호","국민연금 사업장번호","고용 사업장번호","산재 사업장번호"};for(int i=0;i<siteHeaders.Length;i++)ws.Cells[1,31+i].Value=siteHeaders[i];StyleBlockHeader(ws.Cells[1,31,1,34]);
            ws.Cells[1,35].Value="국민 정산 개인(포함액)";ws.Cells[1,36].Value="국민 정산 기관(포함액)";StyleBlockHeader(ws.Cells[1,35,1,36]);
            int r=2;
            foreach(var g in rows.GroupBy(ResultIdentityKey).OrderBy(x=>x.First().Fund).ThenBy(x=>x.First().Name))
            {
                ResultRow f=g.First();int mr;bool linked=missingRows.TryGetValue(ResultIdentityKey(f),out mr);
                if(linked)ws.Cells[r,1].Formula="_MissingData!A"+mr;else ws.Cells[r,1].Value=f.Fund;
                ws.Cells[r,2].Value=f.Name;ws.Cells[r,3].Value=f.Birth;ws.Cells[r,4].Value=f.Source;ws.Cells[r,5].Value=f.Job;
                if(linked)
                {
                    string prefix="IF(_MissingData!$A$"+mr+"=\"휴직\",0,";int[] sourceCols={6,8,10,12,7,9,11,13};
                    for(int c=0;c<4;c++)ws.Cells[r,6+c].Formula=prefix+"_MissingData!"+ExcelCellAddress.GetColumnLetter(sourceCols[c])+mr+")";
                    for(int c=0;c<4;c++)ws.Cells[r,11+c].Formula=prefix+"_MissingData!"+ExcelCellAddress.GetColumnLetter(sourceCols[c+4])+mr+")";
                    string[] insurance={"건강보험","국민연금","고용보험","산재보험"};
                    for(int c=0;c<4;c++)
                    {
                        string raw=g.Where(x=>x.Insurance==insurance[c]).Sum(x=>x.Deduction).ToString(CultureInfo.InvariantCulture);
                        if(f.Fund=="분류필요")ws.Cells[r,25+c].Formula="IF(_MissingData!$A$"+mr+"=\"휴직\",0,IF(_MissingData!$A$"+mr+"=\"분류필요\","+raw+","+ExcelCellAddress.GetColumnLetter(11+c)+r+"))";
                        else ws.Cells[r,25+c].Formula="IF(_MissingData!$A$"+mr+"=\"휴직\",0,"+raw+")";
                    }
                    string d1="K"+r+"-Y"+r,d2="L"+r+"-Z"+r,d3="M"+r+"-AA"+r,d4="N"+r+"-AB"+r;
                    ws.Cells[r,17].Formula="IF(A"+r+"=\"휴직\",\"휴직 제외\",IF(A"+r+"=\"분류필요\",\"분류필요\",IF(ABS("+d1+")+ABS("+d2+")+ABS("+d3+")+ABS("+d4+")=0,\"정상\",IF(AND("+d1+">=0,"+d2+">=0,"+d3+">=0,"+d4+">=0),\"추납 확인\",IF(AND("+d1+"<=0,"+d2+"<=0,"+d3+"<=0,"+d4+"<=0),\"환급 확인\",\"차액 확인\")))))";
                }
                else
                {
                    ws.Cells[r,6].Value=g.Where(x=>x.Insurance=="건강보험").Sum(x=>x.Employer);ws.Cells[r,7].Value=g.Where(x=>x.Insurance=="국민연금").Sum(x=>x.Employer);ws.Cells[r,8].Value=g.Where(x=>x.Insurance=="고용보험").Sum(x=>x.Employer);ws.Cells[r,9].Value=g.Where(x=>x.Insurance=="산재보험").Sum(x=>x.Employer);
                    ws.Cells[r,11].Value=g.Where(x=>x.Insurance=="건강보험").Sum(x=>x.Charge);ws.Cells[r,12].Value=g.Where(x=>x.Insurance=="국민연금").Sum(x=>x.Charge);ws.Cells[r,13].Value=g.Where(x=>x.Insurance=="고용보험").Sum(x=>x.Charge);ws.Cells[r,14].Value=g.Where(x=>x.Insurance=="산재보험").Sum(x=>x.Charge);
                    ws.Cells[r,25].Value=g.Where(x=>x.Insurance=="건강보험").Sum(x=>x.Deduction);ws.Cells[r,26].Value=g.Where(x=>x.Insurance=="국민연금").Sum(x=>x.Deduction);ws.Cells[r,27].Value=g.Where(x=>x.Insurance=="고용보험").Sum(x=>x.Deduction);ws.Cells[r,28].Value=g.Where(x=>x.Insurance=="산재보험").Sum(x=>x.Deduction);
                    string[] bad=g.Select(x=>x.Status).Where(x=>x!="정상"&&x!="부과확인").Distinct().ToArray();ws.Cells[r,17].Value=bad.Length==0?"정상":String.Join(", ",bad);ColorStatus(ws.Cells[r,17],bad.Length==0?"정상":bad[0]);
                }
                string settlePrefix=linked?"IF(_MissingData!$A$"+mr+"=\"휴직\",0,":"";decimal[] settlements={g.Where(x=>x.Insurance=="건강보험").Sum(x=>x.SettlementPersonal),g.Where(x=>x.Insurance=="건강보험").Sum(x=>x.SettlementEmployer),g.Where(x=>x.Insurance=="고용보험").Sum(x=>x.SettlementPersonal),g.Where(x=>x.Insurance=="고용보험").Sum(x=>x.SettlementEmployer),g.Where(x=>x.Insurance=="산재보험").Sum(x=>x.SettlementPersonal),g.Where(x=>x.Insurance=="산재보험").Sum(x=>x.SettlementEmployer)};
                for(int c=0;c<settlements.Length;c++){if(linked)ws.Cells[r,18+c].Formula=settlePrefix+settlements[c].ToString(CultureInfo.InvariantCulture)+")";else ws.Cells[r,18+c].Value=settlements[c];}
                decimal pensionSettlementPersonal=g.Where(x=>x.Insurance=="국민연금").Sum(x=>x.SettlementPersonal),pensionSettlementEmployer=g.Where(x=>x.Insurance=="국민연금").Sum(x=>x.SettlementEmployer);
                if(linked){ws.Cells[r,35].Formula=settlePrefix+pensionSettlementPersonal.ToString(CultureInfo.InvariantCulture)+")";ws.Cells[r,36].Formula=settlePrefix+pensionSettlementEmployer.ToString(CultureInfo.InvariantCulture)+")";}else{ws.Cells[r,35].Value=pensionSettlementPersonal;ws.Cells[r,36].Value=pensionSettlementEmployer;}ws.Cells[r,37].Value=f.Reason;
                ResultRow health=g.FirstOrDefault(x=>x.Insurance=="건강보험");string healthGroup=health!=null&&health.ChargeSource=="건강보험(공무원)"?"건강보험(공무원)":health!=null&&health.ChargeSource=="건강보험(비공무원)"?"건강보험(그 외)":f.Fund=="공무원"?"건강보험(공무원)":"건강보험(그 외)";
                ws.Cells[r,29].Value=healthGroup;ws.Cells[r,10].Formula="SUM(F"+r+":I"+r+")";ws.Cells[r,15].Formula="SUM(K"+r+":N"+r+")";ws.Cells[r,16].Formula="J"+r+"+O"+r;ws.Cells[r,24].Formula="A"+r+"&COUNTIF($A$2:A"+r+",A"+r+")";ws.Cells[r,30].Formula="IF(LEFT(A"+r+",2)=\"학회\",\"학회\"&COUNTIF($A$2:A"+r+",\"학회*\"),\"\")";string[] siteInsurance={"건강보험","국민연금","고용보험","산재보험"};for(int i=0;i<siteInsurance.Length;i++)ws.Cells[r,31+i].Value=String.Join(", ",g.Where(x=>x.Insurance==siteInsurance[i]).Select(x=>x.WorkplaceNumber).Where(x=>!String.IsNullOrWhiteSpace(x)).Distinct());r++;
            }
            int last=Math.Max(2,r-1);var dv=ws.DataValidations.AddListValidation("A2:A"+last);foreach(string c in FundCategories)dv.Formula.Values.Add(c);
            ws.Cells[1,37].Value="기간제 근무 사유(원근로자 현황)";StyleBlockHeader(ws.Cells[1,37]);Finish(ws,37,r-1);ws.Cells[2,6,last,16].Style.Numberformat.Format="#,##0;[Red]-#,##0";ws.Cells[2,18,last,23].Style.Numberformat.Format="#,##0;[Red]-#,##0";ws.Cells[2,35,last,36].Style.Numberformat.Format="#,##0;[Red]-#,##0";ws.Column(1).Style.Fill.PatternType=ExcelFillStyle.Solid;ws.Column(1).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255,242,204));ws.Cells[1,1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(68,114,196));for(int c=24;c<=30;c++)ws.Column(c).Hidden=true;for(int c=31;c<=36;c++){ws.Column(c).AutoFit();if(ws.Column(c).Width<16)ws.Column(c).Width=16;}ws.Column(37).Hidden=true;return r-1;
        }

        static void WriteUiSummarySheet(ExcelWorksheet ws,List<ResultRow> rows,BillingPeriod period)
        {
            string[] headers={"사업장관리번호","재원","인원","건강개인","건강기관","장기요양개인","장기요양기관","국민개인","국민기관","고용개인","고용기관","산재개인","산재기관","기관부담계","확인필요","대체근로자","연도","월","건강차액","장기요양차액","국민차액","고용차액","산재차액"};WriteHeader(ws,headers);int row=2;var sites=rows.Where(x=>!IsMissingWorkplace(x.WorkplaceNumber)).GroupBy(x=>x.WorkplaceNumber).OrderBy(x=>x.Key);
            foreach(var site in sites)
            {
                var people=site.GroupBy(ResultIdentityKey).ToList();foreach(var fundGroup in people.GroupBy(g=>UiFundName(g.First().Fund)).OrderBy(g=>UiFundOrder(g.Key)))
                {
                    decimal hp=0,he=0,lp=0,le=0,pp=0,pe=0,ep=0,ee=0,ip=0,ie=0,hd=0,ld=0,pd=0,ed=0,id=0;int review=0,shortTerm=0;
                    foreach(var person in fundGroup)
                    {
                        ResultRow h=CombineInsurance(person,"건강보험"),p=CombineInsurance(person,"국민연금"),e=CombineInsurance(person,"고용보험"),ind=CombineInsurance(person,"산재보험");decimal personHp=h.ChargeHealth+h.SettlementPersonalHealth,personLp=h.ChargeLongTerm+h.SettlementPersonalLongTerm,personHe=h.EmployerHealth+h.SettlementEmployerHealth,personLe=h.EmployerLongTerm+h.SettlementEmployerLongTerm,payrollHp=h.DeductionHealth+h.DeductionSettlementHealth,payrollLp=h.DeductionLongTerm+h.DeductionSettlementLongTerm;if(personHp==0&&personLp==0&&h.Charge+h.SettlementPersonal!=0){personHp=h.Charge+h.SettlementPersonal;payrollHp=h.Deduction+h.DeductionSettlement;}if(personHe==0&&personLe==0&&h.Employer+h.SettlementEmployer!=0)personHe=h.Employer+h.SettlementEmployer;hp+=personHp;lp+=personLp;he+=personHe;le+=personLe;pp+=p.Charge+p.SettlementPersonal;pe+=p.Employer+p.SettlementEmployer;ep+=e.Charge;ee+=e.Employer;ip+=ind.Charge;ie+=ind.Employer;hd+=personHp-payrollHp;ld+=personLp-payrollLp;pd+=p.Charge+p.SettlementPersonal-p.Deduction-p.DeductionSettlement;ed+=e.Charge-e.Deduction-e.DeductionSettlement;id+=ind.Charge-ind.Deduction-ind.DeductionSettlement;
                        ResultRow first=person.First();if(first.Fund=="분류필요"||person.Any(x=>x.Status!="정상"&&x.Status!="부과확인")||person.Any(x=>Math.Abs(x.SettlementPersonal-x.DeductionSettlement)>.5m))review++;if(first.Fund.Contains("일용")||(!String.IsNullOrWhiteSpace(first.Source)&&first.Source.Contains("단기기간제 근로자"))||(!String.IsNullOrWhiteSpace(first.Reason)&&Regex.IsMatch(first.Reason,"대체|단기")))shortTerm++;
                    }
                    object[] values={site.Key,fundGroup.Key,fundGroup.Count(),hp,he,lp,le,pp,pe,ep,ee,ip,ie,he+le+pe+ee+ie,review,shortTerm,period.Year,period.Month,hd,ld,pd,ed,id};for(int c=0;c<values.Length;c++)ws.Cells[row,c+1].Value=values[c];row++;
                }
            }
            Finish(ws,headers.Length,Math.Max(1,row-1));if(row>2)ws.Cells[2,4,row-1,23].Style.Numberformat.Format="#,##0;[Red]-#,##0";
        }
        static void WriteUiIndividualSheet(ExcelWorksheet ws,List<ResultRow> rows,BillingPeriod period)
        {
            string[] headers={"사업장관리번호","재원","이름","생년월일","직종명","대사결과","건강고지","건강급여","건강차액","국민고지","국민급여","국민차액","고용고지","고용급여","고용차액","산재고지","산재급여","산재차액","연도","월","확인사유","건강개인","건강기관","장기요양개인","장기요양기관","국민개인","국민기관","고용개인","고용기관","산재개인","산재기관","건강차액분리","장기요양차액분리","대체근로자","요약기여유효"};WriteHeader(ws,headers);int row=2;
            foreach(var site in rows.Where(x=>!IsMissingWorkplace(x.WorkplaceNumber)).GroupBy(x=>x.WorkplaceNumber).OrderBy(x=>x.Key))foreach(var person in site.GroupBy(ResultIdentityKey).OrderBy(x=>UiFundOrder(UiFundName(x.First().Fund))).ThenBy(x=>x.First().Name))
            {
                ResultRow first=person.First(),h=CombineInsurance(person,"건강보험"),p=CombineInsurance(person,"국민연금"),e=CombineInsurance(person,"고용보험"),ind=CombineInsurance(person,"산재보험");decimal personHp=h.ChargeHealth+h.SettlementPersonalHealth,personLp=h.ChargeLongTerm+h.SettlementPersonalLongTerm,personHe=h.EmployerHealth+h.SettlementEmployerHealth,personLe=h.EmployerLongTerm+h.SettlementEmployerLongTerm,payrollHp=h.DeductionHealth+h.DeductionSettlementHealth,payrollLp=h.DeductionLongTerm+h.DeductionSettlementLongTerm;if(personHp==0&&personLp==0&&h.Charge+h.SettlementPersonal!=0){personHp=h.Charge+h.SettlementPersonal;payrollHp=h.Deduction+h.DeductionSettlement;}if(personHe==0&&personLe==0&&h.Employer+h.SettlementEmployer!=0)personHe=h.Employer+h.SettlementEmployer;decimal healthNotice=personHp+personLp,healthPayroll=payrollHp+payrollLp,pensionNotice=p.Charge+p.SettlementPersonal,pensionPayroll=p.Deduction+p.DeductionSettlement,employmentNotice=e.Charge,employmentPayroll=e.Deduction+e.DeductionSettlement,industrialNotice=ind.Employer+ind.SettlementEmployer,industrialPayroll=industrialNotice;decimal[] differences={healthNotice-healthPayroll,pensionNotice-pensionPayroll,employmentNotice-employmentPayroll};bool positive=differences.Any(x=>x>.5m),negative=differences.Any(x=>x<-.5m),classified=first.Fund!="분류필요",unusual=person.Any(x=>x.Status!="정상"&&x.Status!="부과확인"&&x.Status!="추납"&&x.Status!="환급"),shortTerm=first.Fund.Contains("일용")||(!String.IsNullOrWhiteSpace(first.Source)&&first.Source.Contains("단기기간제 근로자"))||(!String.IsNullOrWhiteSpace(first.Reason)&&Regex.IsMatch(first.Reason,"대체|단기"));string status=!classified||unusual||positive&&negative?"확인 필요":positive?"추징 필요":negative?"환급 필요":"정상";string reason=!classified?"재원 분류 필요":positive&&negative?"보험별 추징·환급 혼재":String.Join(", ",person.Select(x=>x.Status).Where(x=>x!="정상"&&x!="부과확인").Distinct());object[] values={site.Key,UiIndividualFundName(first.Fund),first.Name,first.Birth,first.Job,status,healthNotice,healthPayroll,differences[0],pensionNotice,pensionPayroll,differences[1],employmentNotice,employmentPayroll,differences[2],industrialNotice,industrialPayroll,0,period.Year,period.Month,reason,personHp,personHe,personLp,personLe,pensionNotice,p.Employer+p.SettlementEmployer,employmentNotice,e.Employer,ind.Charge,ind.Employer,personHp-payrollHp,personLp-payrollLp,shortTerm?1:0,1};for(int c=0;c<values.Length;c++)ws.Cells[row,c+1].Value=values[c];row++;
            }
            Finish(ws,headers.Length,Math.Max(1,row-1));if(row>2)ws.Cells[2,7,row-1,33].Style.Numberformat.Format="#,##0;[Red]-#,##0";
        }
        static string UiIndividualFundName(string fund){return fund=="분류필요"?"분류필요":UiFundName(fund);}
        static string UiFundName(string fund){if((fund??"").StartsWith("공무원"))return "공무원";if((fund??"").StartsWith("계약제교원"))return "계약제교원";if((fund??"").StartsWith("교특"))return "교특회계";if((fund??"").StartsWith("학회"))return "학교회계";return "기타";}
        static int UiFundOrder(string fund){switch(fund){case "공무원":return 0;case "계약제교원":return 1;case "교특회계":return 2;case "학교회계":return 3;default:return 4;}}

        static void WriteFundSummarySheet(ExcelWorksheet ws,int personLast)
        {
            string[] h={"재원구분","보험종류","대상자수","개인부과액","기관부담금","납입예정액"};WriteHeader(ws,h);string[] ins={"건강보험","국민연금","고용보험","산재보험"};int r=2;
            for(int f=0;f<SummaryFundCategories.Length;f++)for(int i=0;i<ins.Length;i++)
            {
                string fund=SummaryFundCategories[f],criterion=fund=="학회"?"학회*":fund;ws.Cells[r,1].Value=fund;ws.Cells[r,2].Value=ins[i];string inst=ExcelCellAddress.GetColumnLetter(6+i),personal=ExcelCellAddress.GetColumnLetter(11+i);
                ws.Cells[r,3].Formula="COUNTIFS('근무자별 부담금'!$A$2:$A$"+personLast+",\""+criterion+"\",'근무자별 부담금'!$"+inst+"$2:$"+inst+"$"+personLast+",\">0\")";
                ws.Cells[r,4].Formula="SUMIF('근무자별 부담금'!$A$2:$A$"+personLast+",\""+criterion+"\",'근무자별 부담금'!$"+personal+"$2:$"+personal+"$"+personLast+")";
                ws.Cells[r,5].Formula="SUMIF('근무자별 부담금'!$A$2:$A$"+personLast+",\""+criterion+"\",'근무자별 부담금'!$"+inst+"$2:$"+inst+"$"+personLast+")";ws.Cells[r,6].Formula="D"+r+"+E"+r;r++;
            }
            ws.Cells[r,1].Value="총계";for(int c=3;c<=6;c++)ws.Cells[r,c].Formula="SUM("+ws.Cells[2,c].Address+":"+ws.Cells[r-1,c].Address+")";StyleTotal(ws.Cells[r,1,r,6]);Finish(ws,h.Length,r);ws.Cells[2,4,r,6].Style.Numberformat.Format="#,##0;[Red]-#,##0";
        }
        static void WriteWorkplaceSummarySheet(ExcelWorksheet ws,List<ResultRow> rows)
        {
            ws.Cells[1,1,1,9].Merge=true;ws.Cells[1,1].Value="사업장번호별 4대보험 고지·납입 현황";ws.Cells[1,1].Style.Font.Size=18;ws.Cells[1,1].Style.Font.Bold=true;ws.Cells[1,1].Style.HorizontalAlignment=ExcelHorizontalAlignment.Center;
            ws.Cells[2,1,2,9].Merge=true;ws.Cells[2,1].Value="EDI 엑셀 내부의 사업장관리번호를 우선 사용하며, 표기가 없으면 파일명 속 사업장번호를 사용합니다.";ws.Cells[2,1].Style.Font.Color.SetColor(Color.DimGray);
            int r=4;string[] headers={"사업장번호","보험구분","대상자수","급여공제액","개인부과액","개인차액","기관부담금","고지금액 총액","확인"};
            foreach(var site in rows.GroupBy(x=>String.IsNullOrWhiteSpace(x.WorkplaceNumber)?"미확인":x.WorkplaceNumber).OrderBy(x=>x.Key=="미확인"?"ZZZZZZZZZZZZZZZ":x.Key))
            {
                ws.Cells[r,1,r,9].Merge=true;ws.Cells[r,1].Value="사업장번호: "+site.Key;StyleSectionTitle(ws.Cells[r,1,r,9]);r++;
                for(int c=0;c<headers.Length;c++)ws.Cells[r,c+1].Value=headers[c];StyleBlockHeader(ws.Cells[r,1,r,9]);r++;
                int firstData=r;
                var groups=site.GroupBy(x=>x.Insurance=="건강보험"?(x.ChargeSource=="건강보험(공무원)"?"건강보험(공무원)":"건강보험(그 외)"):x.Insurance).OrderBy(x=>x.Key);
                foreach(var g in groups)
                {
                    ws.Cells[r,1].Value=site.Key;ws.Cells[r,2].Value=g.Key;ws.Cells[r,3].Value=g.Select(x=>Key(x.Name,x.Birth)).Distinct().Count();ws.Cells[r,4].Value=g.Sum(x=>x.Deduction);ws.Cells[r,5].Value=g.Sum(x=>x.Charge);ws.Cells[r,6].Formula="E"+r+"-D"+r;ws.Cells[r,7].Value=g.Sum(x=>x.Employer);ws.Cells[r,8].Formula="E"+r+"+G"+r;ws.Cells[r,9].Formula=g.Key=="산재보험"?"\"부과확인\"":"IF(F"+r+"=0,\"일치\",\"차액확인\")";r++;
                }
                ws.Cells[r,1].Value=site.Key;ws.Cells[r,2].Value="사업장 합계";for(int c=3;c<=8;c++)ws.Cells[r,c].Formula="SUM("+ws.Cells[firstData,c].Address+":"+ws.Cells[r-1,c].Address+")";ws.Cells[r,9].Formula="IF(F"+r+"=0,\"일치\",\"차액확인\")";StyleTotal(ws.Cells[r,1,r,9]);r+=2;
            }
            if(r==4){ws.Cells[4,1].Value="인식된 보험 고지자료가 없습니다.";r=5;}
            ws.View.FreezePanes(4,1);ws.Cells[1,1,Math.Max(1,r-1),9].Style.VerticalAlignment=ExcelVerticalAlignment.Center;ws.Cells[5,4,Math.Max(5,r-1),8].Style.Numberformat.Format="#,##0;[Red]-#,##0";for(int c=1;c<=9;c++){ws.Column(c).AutoFit();if(ws.Column(c).Width<12)ws.Column(c).Width=12;if(ws.Column(c).Width>26)ws.Column(c).Width=26;}ws.Column(1).Width=18;ws.Column(2).Width=22;
        }

        static void WriteDashboard(ExcelWorksheet ws,List<ResultRow> rows,int personLast)
        {
            ws.Cells[2,2,2,14].Merge=true;ws.Cells[2,2].Value="4대보험 고지·납입 총괄표";ws.Cells[2,2].Style.Font.Size=18;ws.Cells[2,2].Style.Font.Bold=true;ws.Cells[2,2].Style.HorizontalAlignment=ExcelHorizontalAlignment.Center;
            string[] ih={"보험구분","급여공제액","개인부과액","개인차액","기관부담 고지액","고지금액 총액","분류완료 합계","미분류 금액","기관부담 대조차액","확인","할인·조정액","할인 반영 재원","실제 납입금액"};for(int c=0;c<ih.Length;c++)ws.Cells[4,c+2].Value=ih[c];StyleBlockHeader(ws.Cells[4,2,4,14]);
            string[] displayIns={"건강보험(공무원)","건강보험(그 외)","국민연금","고용보험","산재보험"};int[] insuranceIndex={0,0,1,2,3};
            for(int i=0;i<displayIns.Length;i++)
            {
                int r=5+i,index=insuranceIndex[i];bool health=i<2;string instCol=ExcelCellAddress.GetColumnLetter(6+index),personalCol=ExcelCellAddress.GetColumnLetter(11+index),deductionCol=ExcelCellAddress.GetColumnLetter(25+index);ws.Cells[r,2].Value=displayIns[i];
                string healthCriteria="'근무자별 부담금'!$AC$2:$AC$"+personLast+",B"+r+",";
                ws.Cells[r,3].Formula=health?"SUMIF("+healthCriteria+"'근무자별 부담금'!$"+deductionCol+"$2:$"+deductionCol+"$"+personLast+")":"SUM('근무자별 부담금'!$"+deductionCol+"$2:$"+deductionCol+"$"+personLast+")";
                ws.Cells[r,4].Formula=health?"SUMIF("+healthCriteria+"'근무자별 부담금'!$"+personalCol+"$2:$"+personalCol+"$"+personLast+")":"SUM('근무자별 부담금'!$"+personalCol+"$2:$"+personalCol+"$"+personLast+")";
                ws.Cells[r,5].Formula="D"+r+"-C"+r;ws.Cells[r,6].Formula=health?"SUMIF("+healthCriteria+"'근무자별 부담금'!$"+instCol+"$2:$"+instCol+"$"+personLast+")":"SUM('근무자별 부담금'!$"+instCol+"$2:$"+instCol+"$"+personLast+")";ws.Cells[r,7].Formula="D"+r+"+F"+r;
                string classified=health?"SUMIFS('근무자별 부담금'!$"+instCol+"$2:$"+instCol+"$"+personLast+",'근무자별 부담금'!$A$2:$A$"+personLast+",\"<>분류필요\",'근무자별 부담금'!$A$2:$A$"+personLast+",\"<>휴직\",'근무자별 부담금'!$AC$2:$AC$"+personLast+",B"+r+")":"SUMIFS('근무자별 부담금'!$"+instCol+"$2:$"+instCol+"$"+personLast+",'근무자별 부담금'!$A$2:$A$"+personLast+",\"<>분류필요\",'근무자별 부담금'!$A$2:$A$"+personLast+",\"<>휴직\")";
                string unclassified=health?"SUMIFS('근무자별 부담금'!$"+instCol+"$2:$"+instCol+"$"+personLast+",'근무자별 부담금'!$A$2:$A$"+personLast+",\"분류필요\",'근무자별 부담금'!$AC$2:$AC$"+personLast+",B"+r+")":"SUMIF('근무자별 부담금'!$A$2:$A$"+personLast+",\"분류필요\",'근무자별 부담금'!$"+instCol+"$2:$"+instCol+"$"+personLast+")";
                ws.Cells[r,8].Formula=classified+"-IF(AND(M"+r+"<>\"분류필요\",M"+r+"<>\"휴직\"),L"+r+",0)";ws.Cells[r,9].Formula=unclassified+"-IF(M"+r+"=\"분류필요\",L"+r+",0)";ws.Cells[r,10].Formula="F"+r+"-L"+r+"-H"+r+"-I"+r;ws.Cells[r,11].Formula="IF(J"+r+"=0,\"일치\",\"불일치\")";
                ws.Cells[r,12].Value=0;ws.Cells[r,13].Value=i==0?"공무원":i==1?"교특":"학회(교육공무직)";var discountFund=ws.DataValidations.AddListValidation(ws.Cells[r,13].Address);foreach(string fund in FundCategories.Where(x=>x!="휴직"))discountFund.Formula.Values.Add(fund);ws.Cells[r,14].Formula="G"+r+"-L"+r;
            }
            int topTotal=10;ws.Cells[5,12,9,13].Style.Fill.PatternType=ExcelFillStyle.Solid;ws.Cells[5,12,9,13].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255,242,204));ws.Cells[topTotal,2].Value="총계";for(int c=3;c<=10;c++)ws.Cells[topTotal,c].Formula="SUM("+ws.Cells[5,c].Address+":"+ws.Cells[9,c].Address+")";ws.Cells[topTotal,11].Formula="IF(J"+topTotal+"=0,\"일치\",\"불일치\")";ws.Cells[topTotal,12].Formula="SUM(L5:L9)";ws.Cells[topTotal,14].Formula="SUM(N5:N9)";StyleTotal(ws.Cells[topTotal,2,topTotal,14]);ws.Cells[5,3,topTotal,12].Style.Numberformat.Format="#,##0;[Red]-#,##0";ws.Cells[5,14,topTotal,14].Style.Numberformat.Format="#,##0;[Red]-#,##0";
            int sectionRow=12,headerRow=13,dataRow=14;ws.Cells[sectionRow,2,sectionRow,10].Merge=true;ws.Cells[sectionRow,2].Value="재원별 보험료 및 납입예정액";StyleSectionTitle(ws.Cells[sectionRow,2,sectionRow,10]);
            string[] fh={"재원구분","건강보험","국민연금","고용보험","산재보험","기관부담 총액","개인부과 총액","납입예정 총액","인원수"};for(int c=0;c<fh.Length;c++)ws.Cells[headerRow,c+2].Value=fh[c];StyleBlockHeader(ws.Cells[headerRow,2,headerRow,10]);string[] reportFunds=SummaryFundCategories;
            for(int i=0;i<reportFunds.Length;i++)
            {
                int r=dataRow+i;string fundCriterion=reportFunds[i]=="학회"?"학회*":reportFunds[i];ws.Cells[r,2].Value=reportFunds[i];for(int c=3;c<=6;c++){string pc=ExcelCellAddress.GetColumnLetter(c+3),insuranceCriterion=c==3?"건강보험*":c==4?"국민연금":c==5?"고용보험":"산재보험";ws.Cells[r,c].Formula="SUMIF('근무자별 부담금'!$A$2:$A$"+personLast+",\""+fundCriterion+"\",'근무자별 부담금'!$"+pc+"$2:$"+pc+"$"+personLast+")-SUMIFS($L$5:$L$9,$M$5:$M$9,\""+fundCriterion+"\",$B$5:$B$9,\""+insuranceCriterion+"\")";}ws.Cells[r,7].Formula="SUM(C"+r+":F"+r+")";ws.Cells[r,8].Formula="SUMIF('근무자별 부담금'!$A$2:$A$"+personLast+",\""+fundCriterion+"\",'근무자별 부담금'!$O$2:$O$"+personLast+")";ws.Cells[r,9].Formula="G"+r+"+H"+r;ws.Cells[r,10].Formula="COUNTIFS('근무자별 부담금'!$A$2:$A$"+personLast+",\""+fundCriterion+"\",'근무자별 부담금'!$P$2:$P$"+personLast+",\">0\")";
            }
            int ft=dataRow+reportFunds.Length;ws.Cells[ft,2].Value="총계";for(int c=3;c<=10;c++)ws.Cells[ft,c].Formula="SUM("+ws.Cells[dataRow,c].Address+":"+ws.Cells[ft-1,c].Address+")";StyleTotal(ws.Cells[ft,2,ft,10]);ws.Cells[dataRow,3,ft,9].Style.Numberformat.Format="#,##0;[Red]-#,##0";
            int titleRow=ft+2,selectRow=ft+4,headRow=ft+6,start=headRow+1,displayRows=Math.Max(40,personLast-1);
            ws.Cells[titleRow,2,titleRow,12].Merge=true;ws.Cells[titleRow,2].Value="선택 재원 근무자별 부담금";StyleSectionTitle(ws.Cells[titleRow,2,titleRow,12]);ws.Cells[selectRow,2].Value="재원 선택";ws.Cells[selectRow,2].Style.Font.Bold=true;ws.Cells[selectRow,3].Value="공무원";ws.Cells[selectRow,3].Style.Fill.PatternType=ExcelFillStyle.Solid;ws.Cells[selectRow,3].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255,242,204));var choice=ws.DataValidations.AddListValidation(ws.Cells[selectRow,3].Address);foreach(string c in FundSelectionCategories)choice.Formula.Values.Add(c);
            string[] lh={"성명","생년월일","급여자료","직종","건강보험","국민연금","고용보험","산재보험","기관부담 합계","개인부과 합계","확인상태"};for(int c=0;c<lh.Length;c++)ws.Cells[headRow,c+2].Value=lh[c];StyleBlockHeader(ws.Cells[headRow,2,headRow,12]);
            int[] srcCols={2,3,4,5,6,7,8,9,10,15,17};
            for(int r=start;r<start+displayRows;r++)
            {
                ws.Cells[r,1].Formula="IF($C$"+selectRow+"=\"학회\",IFERROR(MATCH(\"학회\"&(ROW()-"+(start-1)+"),'근무자별 부담금'!$AD$2:$AD$"+personLast+",0),\"\"),IFERROR(MATCH($C$"+selectRow+"&(ROW()-"+(start-1)+"),'근무자별 부담금'!$X$2:$X$"+personLast+",0),\"\"))";
                for(int c=0;c<srcCols.Length;c++){string sourceCol=ExcelCellAddress.GetColumnLetter(srcCols[c]);ws.Cells[r,c+2].Formula="IF($A"+r+"=\"\",\"\",INDEX('근무자별 부담금'!$"+sourceCol+"$2:$"+sourceCol+"$"+personLast+",$A"+r+"))";}
            }
            ws.Column(1).Hidden=true;ws.Cells[start,6,start+displayRows-1,11].Style.Numberformat.Format="#,##0;[Red]-#,##0";ws.View.FreezePanes(5,2);for(int c=2;c<=12;c++){ws.Column(c).AutoFit();if(ws.Column(c).Width<12)ws.Column(c).Width=12;if(ws.Column(c).Width>24)ws.Column(c).Width=24;}ws.Column(5).Width=22;
        }
        static void StyleBlockHeader(ExcelRange x){x.Style.Font.Bold=true;x.Style.Font.Color.SetColor(Color.White);x.Style.Fill.PatternType=ExcelFillStyle.Solid;x.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(68,114,196));x.Style.HorizontalAlignment=ExcelHorizontalAlignment.Center;x.Style.Border.BorderAround(ExcelBorderStyle.Thin,Color.Black);}
        static void StyleSectionTitle(ExcelRange x){x.Style.Font.Bold=true;x.Style.Font.Size=12;x.Style.Fill.PatternType=ExcelFillStyle.Solid;x.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(221,235,247));x.Style.HorizontalAlignment=ExcelHorizontalAlignment.Center;}
        static void StyleTotal(ExcelRange x){x.Style.Font.Bold=true;x.Style.Fill.PatternType=ExcelFillStyle.Solid;x.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(226,239,218));x.Style.Border.Top.Style=ExcelBorderStyle.Thin;x.Style.Border.Bottom.Style=ExcelBorderStyle.Thin;}
        static void WriteHeader(ExcelWorksheet ws,string[] h){for(int i=0;i<h.Length;i++)ws.Cells[1,i+1].Value=h[i];var x=ws.Cells[1,1,1,h.Length];x.Style.Font.Bold=true;x.Style.Font.Color.SetColor(Color.White);x.Style.Fill.PatternType=ExcelFillStyle.Solid;x.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(68,114,196));x.Style.HorizontalAlignment=ExcelHorizontalAlignment.Center;}
        static void Finish(ExcelWorksheet ws,int cols,int last){ws.View.FreezePanes(2,1);if(last>=1)ws.Cells[1,1,Math.Max(1,last),cols].AutoFilter=true;for(int c=1;c<=cols;c++){ws.Column(c).AutoFit();if(ws.Column(c).Width>28)ws.Column(c).Width=28;}ws.Cells.Style.VerticalAlignment=ExcelVerticalAlignment.Center;}
        static void ColorStatus(ExcelRange cell,string status){cell.Style.Font.Bold=true;cell.Style.Fill.PatternType=ExcelFillStyle.Solid;if(status=="정상")cell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(198,239,206));else if(status.Contains("누락"))cell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255,199,206));else cell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255,235,156));}

        class SheetInfo{public ExcelWorksheet Sheet;public int HeaderRow;public Dictionary<int,string> Headers;}
        static SheetInfo FindSheet(ExcelPackage p,string[] required1,string[] required2)
        {
            SheetInfo best=null;int bestScore=-1;
            foreach(var ws in p.Workbook.Worksheets){if(ws.Dimension==null)continue;int maxR=Math.Min(ws.Dimension.End.Row,60),maxC=Math.Min(ws.Dimension.End.Column,200);for(int r=1;r<=maxR;r++){Dictionary<int,string> h=new Dictionary<int,string>(),cur=new Dictionary<int,string>();for(int c=1;c<=maxC;c++){string current=Text(ws.Cells[r,c].Value),previous=r>1?Text(ws.Cells[r-1,c].Value):"";if(current.Length>0)cur[c]=current;string s=current;if(previous.Length>0&&current.Length>0&&Norm(previous)!=Norm(current))s=previous+" "+current;if(s.Length>0)h[c]=s;}int score=(Has(cur,required1)?10:0)+(Has(cur,required2)?10:0)+Math.Min(cur.Count,20);if(score>=bestScore){bestScore=score;best=new SheetInfo{Sheet=ws,HeaderRow=r,Headers=h};}}}
            if(best==null||!Has(best.Headers,required1)||!Has(best.Headers,required2))return null;return best;
        }
        static bool Has(Dictionary<int,string> h,string[] a){return h.Values.Any(v=>Match(v,a));}
        static int FindCol(SheetInfo s,string[] aliases){foreach(var x in s.Headers)if(Match(x.Value,aliases))return x.Key;return 0;}
        static string DetectWorkplaceNumber(ExcelPackage package,SheetInfo info,string path)
        {
            string[] aliases={"사업장관리번호","사업장번호","사업장기호","단위사업장번호","사업장관리기호"};int col=FindCol(info,aliases);
            if(col>0&&info.Sheet.Dimension!=null)for(int r=info.HeaderRow+1;r<=Math.Min(info.Sheet.Dimension.End.Row,info.HeaderRow+30);r++){string found=NormalizeWorkplaceNumber(info.Sheet.Cells[r,col].Value);if(found.Length>0)return found;}
            foreach(ExcelWorksheet ws in package.Workbook.Worksheets)
            {
                if(ws.Dimension==null)continue;int mr=Math.Min(ws.Dimension.End.Row,40),mc=Math.Min(ws.Dimension.End.Column,50);
                for(int r=1;r<=mr;r++)for(int c=1;c<=mc;c++)
                {
                    string label=Norm(Text(ws.Cells[r,c].Value));if(!aliases.Any(x=>label.Contains(Norm(x))))continue;
                    string embedded=NormalizeWorkplaceNumber(ws.Cells[r,c].Value);if(embedded.Length>0&&Regex.Replace(Text(ws.Cells[r,c].Value),"[^0-9]","").Length>=8)return embedded;
                    foreach(int[] offset in new[]{new[]{0,1},new[]{1,0},new[]{0,2},new[]{1,1}}){int rr=r+offset[0],cc=c+offset[1];if(rr<=ws.Dimension.End.Row&&cc<=ws.Dimension.End.Column){string found=NormalizeWorkplaceNumber(ws.Cells[rr,cc].Value);if(found.Length>0)return found;}}
                }
            }
            Match fileMatch=Regex.Match(Path.GetFileNameWithoutExtension(path),@"(?<!\d)(\d{11})(?!\d)");if(fileMatch.Success)return fileMatch.Groups[1].Value;
            fileMatch=Regex.Match(Path.GetFileNameWithoutExtension(path),@"(?<!\d)(\d{8,15})(?!\d)");return fileMatch.Success?fileMatch.Groups[1].Value:"미확인";
        }
        static string NormalizeWorkplaceNumber(object value)
        {
            string raw=Text(value),digits=Regex.Replace(raw,"[^0-9]","");if(digits.Length>=8&&digits.Length<=15)return digits;return "";
        }
        static bool Match(string value,string[] aliases){string v=Norm(value);foreach(string a in aliases){string n=Norm(a);if(v==n||(n.Length>=2&&v.Contains(n)))return true;}return false;}
        static string Norm(string s){return Regex.Replace((s??"").ToLowerInvariant(),"[^0-9a-z가-힣]","");}
        static string Text(object o){return o==null?"":Convert.ToString(o,CultureInfo.InvariantCulture).Trim();}
        static string CleanName(object o){return Regex.Replace(Text(o),@"\s+","").Trim();}
        static string Birth6(object o){if(o==null)return "";if(o is DateTime)return ((DateTime)o).ToString("yyMMdd");double d;if(Double.TryParse(Text(o),out d)&&d>20000&&d<70000){try{return DateTime.FromOADate(d).ToString("yyMMdd");}catch{}}string s=Regex.Replace(Text(o),"[^0-9]","");if(s.Length>=8)return s.Substring(s.Length==8?2:0,6);return s.Length>=6?s.Substring(0,6):s;}
        static string Key(string name,string birth){return CleanName(name)+"|"+birth;}
        static decimal Num(ExcelWorksheet ws,int r,int c){if(c<=0)return 0;object v=ws.Cells[r,c].Value;if(v==null)return 0;decimal x;if(Decimal.TryParse(Regex.Replace(Text(v),"[^0-9.-]",""),NumberStyles.Any,CultureInfo.InvariantCulture,out x))return x;return 0;}
        static string FundOf(string defaultFund,string job){string j=Norm(job);if(j.Contains("기간제")||j.Contains("계약제")||j.Contains("시간강사"))return "계약제교원";if(j.Contains("교특"))return "교특";return String.IsNullOrWhiteSpace(defaultFund)?"분류필요":defaultFund;}
        static string Cols(params int[] c){return String.Join("+",c.Where(x=>x>0).Select(x=>ExcelCellAddress.GetColumnLetter(x)).ToArray());}
        static Recognition Ok(string kind,string path,SheetInfo s,int rows,string detail){return new Recognition{Kind=kind,File=Path.GetFileName(path),Sheet=s.Sheet.Name,HeaderRow=s.HeaderRow.ToString(),Rows=rows.ToString(),State="정상",Detail=detail};}
        static Recognition Fail(string kind,string path,string detail){return new Recognition{Kind=kind,File=Path.GetFileName(path),State="확인필요",Detail=detail};}
    }

    static class SubmissionGenerator
    {
        class SubmitPerson
        {
            public string Fund, Name, Birth, Job, Reason; public decimal Health, HealthOnly, LongTerm, Pension, Employment, Industrial, HealthSettlement, PensionSettlement, EmploymentSettlement, IndustrialSettlement, HealthBase, PensionBase, EmploymentBase, IndustrialBase; public bool HasHealthParts,ShortTerm;
        }
        class UiSubmissionIdentity{public string Site,Fund,Name,Birth,Job,Reason;public bool ShortTerm,HasSummaryBreakdown;public decimal HealthEmployer,LongTermEmployer,PensionEmployer,EmploymentEmployer,IndustrialEmployer;}
        class HealthParts { public decimal Health, LongTerm; }

        public static string Create(string resultPath,string outputFolder,bool teacher,SubmissionInfo submissionInfo=null)
        {
            if(submissionInfo==null)submissionInfo=new SubmissionInfo();
            using(ExcelPackage source=new ExcelPackage(new FileInfo(resultPath)))
            {
                try{source.Workbook.Calculate();}catch{}
                ExcelWorksheet people=source.Workbook.Worksheets["근무자별 부담금"];
                if(people==null)throw new InvalidOperationException("선택한 파일에서 '근무자별 부담금' 시트를 찾지 못했습니다. 계산(검증) 탭에서 만든 결과 파일을 선택해 주세요.");
                int year=DateTime.Now.Year,month=DateTime.Now.Month;ExcelWorksheet info=source.Workbook.Worksheets["제출정보"];
                if(info!=null){year=ToInt(info.Cells[1,2].Value,year);month=ToInt(info.Cells[2,2].Value,month);}
                string targetFund=teacher?"계약제교원":"교특 + 학회(일용근로)";List<SubmitPerson> targets=new List<SubmitPerson>();
                Dictionary<string,HealthParts> healthParts=ReadHealthParts(source);Dictionary<string,string> reviewFundOverrides=ReadReviewFundOverrides(source);Dictionary<string,UiSubmissionIdentity> uiOverrides=ReadUiSubmissionIdentities(source);
                Dictionary<string,decimal> healthBases=ReadWageBases(source,"건강보험",new[]{"보수월액"}),pensionBases=ReadWageBases(source,"국민연금",new[]{"기준소득월액"}),employmentBases=ReadWageBases(source,"고용보험",new[]{"월평균보수금액","월평균보수액"}),industrialBases=ReadWageBases(source,"산재보험",new[]{"월평균보수액","월평균보수금액"});
                int last=people.Dimension==null?1:people.Dimension.End.Row;
                for(int r=2;r<=last;r++)
                {
                    string name=CellText(source,people,r,2);if(String.IsNullOrWhiteSpace(name))continue;string birth=NormalizeBirth(CellText(source,people,r,3)),key=PersonKey(name,birth),fund=CellText(source,people,r,1),job=CellText(source,people,r,5),reviewFund;bool sourceShortTerm=Regex.IsMatch(NormalizeHeader(fund)+NormalizeHeader(job),"일용|단기|대체",RegexOptions.IgnoreCase);UiSubmissionIdentity ui;if(reviewFundOverrides.TryGetValue(key,out reviewFund)&&!sourceShortTerm)fund=reviewFund;if(uiOverrides.TryGetValue(key,out ui)){if(!sourceShortTerm)fund=ui.Fund;if(!String.IsNullOrWhiteSpace(ui.Job))job=ui.Job;if(!String.IsNullOrWhiteSpace(submissionInfo.Site)&&!String.Equals(ui.Site,submissionInfo.Site,StringComparison.OrdinalIgnoreCase))continue;}bool selected=teacher?IsTeacherSubmissionFund(fund):IsWorkerSubmissionFund(fund);if(!selected)continue;SubmitPerson item=new SubmitPerson{Fund=fund,Name=name,Birth=birth,Job=job,Reason=CellText(source,people,r,37),ShortTerm=sourceShortTerm||ui!=null&&ui.ShortTerm,Health=CellNumber(source,people,r,6),Pension=CellNumber(source,people,r,7),Employment=CellNumber(source,people,r,8),Industrial=CellNumber(source,people,r,9),HealthSettlement=CellNumber(source,people,r,19),EmploymentSettlement=CellNumber(source,people,r,21),IndustrialSettlement=CellNumber(source,people,r,23),PensionSettlement=CellNumber(source,people,r,36)};if(!sourceShortTerm&&ui!=null&&ui.HasSummaryBreakdown){item.Health=ui.HealthEmployer+ui.LongTermEmployer;item.Pension=ui.PensionEmployer;item.Employment=ui.EmploymentEmployer;item.Industrial=ui.IndustrialEmployer;}
                    HealthParts parts;if(!sourceShortTerm&&ui!=null&&ui.HasSummaryBreakdown){item.HealthOnly=ui.HealthEmployer;item.LongTerm=ui.LongTermEmployer;item.HasHealthParts=true;}else if(healthParts.TryGetValue(key,out parts)){decimal delta=item.Health-parts.Health-parts.LongTerm;item.HealthOnly=parts.Health+delta;item.LongTerm=parts.LongTerm;item.HasHealthParts=true;}healthBases.TryGetValue(key,out item.HealthBase);pensionBases.TryGetValue(key,out item.PensionBase);employmentBases.TryGetValue(key,out item.EmploymentBase);industrialBases.TryGetValue(key,out item.IndustrialBase);targets.Add(item);
                }
                if(!teacher)foreach(KeyValuePair<string,UiSubmissionIdentity> pair in uiOverrides){UiSubmissionIdentity ui=pair.Value;if(!ui.ShortTerm||targets.Any(x=>PersonKey(x.Name,x.Birth)==pair.Key))continue;if(!String.IsNullOrWhiteSpace(submissionInfo.Site)&&!String.Equals(ui.Site,submissionInfo.Site,StringComparison.OrdinalIgnoreCase))continue;targets.Add(new SubmitPerson{Fund=ui.Fund,Name=ui.Name,Birth=ui.Birth,Job=ui.Job,Reason=String.IsNullOrWhiteSpace(ui.Reason)?"1개월 미만 대체근로자":ui.Reason,ShortTerm=true,Health=ui.HealthEmployer+ui.LongTermEmployer,HealthOnly=ui.HealthEmployer,LongTerm=ui.LongTermEmployer,HasHealthParts=ui.HasSummaryBreakdown,Pension=ui.PensionEmployer,Employment=ui.EmploymentEmployer,Industrial=ui.IndustrialEmployer});}
                if(targets.Count==0){string found=String.Join(", ",Enumerable.Range(2,Math.Max(0,last-1)).Select(r=>CellText(source,people,r,1)).Where(x=>!String.IsNullOrWhiteSpace(x)).Distinct().Take(10));throw new InvalidOperationException(targetFund+"으로 분류된 제출 대상자가 없습니다. 검증 결과의 '근무자별 부담금' A열 분류를 확인해 주세요."+(found.Length>0?"\r\n현재 확인된 분류: "+found:""));}
                if(teacher)return CreateTeacherWithExcel(targets,year,month,submissionInfo,outputFolder);
                return CreateWorkerWithExcel(targets,year,month,submissionInfo,outputFolder);
            }
        }

        static bool IsWorkerSubmissionFund(string value)
        {
            string n=NormalizeHeader(value);return n.StartsWith("교특")||n.Contains("교육공무직교특")||n.Contains("교특교육공무직")||n.Contains("단기기간제")||n.Contains("일용");
        }
        static bool IsTeacherSubmissionFund(string value){string n=NormalizeHeader(value);return n=="계약제교원"||n.Contains("계약제교원");}
        static bool IsFixedTermWorker(SubmitPerson person){return person!=null&&(person.ShortTerm||Regex.IsMatch(NormalizeHeader(person.Fund)+NormalizeHeader(person.Job)+NormalizeHeader(person.Reason),"기간제|대체|일용|단기|계약직",RegexOptions.IgnoreCase));}

        static Dictionary<string,UiSubmissionIdentity> ReadUiSubmissionIdentities(ExcelPackage source)
        {
            var result=new Dictionary<string,UiSubmissionIdentity>();ExcelWorksheet ws=source.Workbook.Worksheets["UI개인별데이터"];if(ws==null||ws.Dimension==null)return result;for(int r=2;r<=ws.Dimension.End.Row;r++){string name=CellText(source,ws,r,3),birth=NormalizeBirth(CellText(source,ws,r,4));if(String.IsNullOrWhiteSpace(name))continue;result[PersonKey(name,birth)]=new UiSubmissionIdentity{Site=CellText(source,ws,r,1),Fund=CellText(source,ws,r,2),Name=name,Birth=birth,Job=CellText(source,ws,r,5),Reason=CellText(source,ws,r,21),HealthEmployer=CellNumber(source,ws,r,23),LongTermEmployer=CellNumber(source,ws,r,25),PensionEmployer=CellNumber(source,ws,r,27),EmploymentEmployer=CellNumber(source,ws,r,29),IndustrialEmployer=CellNumber(source,ws,r,31),ShortTerm=ToInt(ws.Cells[r,34].Value,0)>0,HasSummaryBreakdown=ToInt(ws.Cells[r,35].Value,0)>0};}return result;
        }

        static Dictionary<string,string> ReadReviewFundOverrides(ExcelPackage source)
        {
            Dictionary<string,string> result=new Dictionary<string,string>();
            foreach(ExcelWorksheet ws in source.Workbook.Worksheets.Where(x=>x.Name.StartsWith("확인명단(")||x.Name.StartsWith("확인필요(")))
            {
                if(ws.Dimension==null)continue;
                for(int r=1;r<=ws.Dimension.End.Row;r++)
                {
                    string fund=CellText(source,ws,r,3),name=CellText(source,ws,r,4),birth=NormalizeBirth(CellText(source,ws,r,5)),normalized=NormalizeHeader(fund);
                    if(String.IsNullOrWhiteSpace(name)||String.IsNullOrWhiteSpace(birth)||String.IsNullOrWhiteSpace(fund)||normalized=="분류필요"||normalized=="전체")continue;
                    result[PersonKey(name,birth)]=fund;
                }
            }
            return result;
        }

        static string CreateTeacherWithExcel(List<SubmitPerson> people,int year,int month,SubmissionInfo info,string outputFolder)
        {
            string round=RoundText(info.Round);
            string code=SafeFilePart(info.RecipientCode,"수신자기호"),institution=SafeFilePart(info.InstitutionName,"기관명");
            string outputPath=UniquePath(Path.Combine(outputFolder,code+"_"+institution+"_"+year+"년 "+month+"월("+round+") 계약제교원 인건비 신청.xlsx"));
            string tempFolder=Path.Combine(Path.GetTempPath(),"TeacherSubmission_"+Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempFolder);string templatePath=Path.Combine(tempFolder,"계약제교원_제출서식.xls");
            using(Stream resource=Assembly.GetExecutingAssembly().GetManifestResourceStream("InsurancePayrollValidator.TeacherTemplate.xls"))
            {
                if(resource==null)throw new InvalidOperationException("내장된 계약제교원 원본 서식을 불러오지 못했습니다.");
                using(FileStream file=new FileStream(templatePath,FileMode.Create,FileAccess.Write))resource.CopyTo(file);
            }

            object excelObject=null,booksObject=null,bookObject=null,appSheetObject=null,detailSheetObject=null;
            try
            {
                Type excelType=Type.GetTypeFromProgID("Excel.Application");
                if(excelType==null)throw new InvalidOperationException("Microsoft Excel이 설치되어 있지 않아 제출 서식을 만들 수 없습니다.");
                excelObject=Activator.CreateInstance(excelType);dynamic excel=excelObject;
                excel.Visible=false;excel.DisplayAlerts=false;excel.ScreenUpdating=false;excel.EnableEvents=false;
                booksObject=excel.Workbooks;dynamic books=booksObject;bookObject=books.Open(templatePath);dynamic book=bookObject;
                appSheetObject=book.Worksheets[1];detailSheetObject=book.Worksheets[5];dynamic app=appSheetObject,detail=detailSheetObject;
                ResizeTeacherApplicationTable(app,people.Count);
                ResizeTeacherInsuranceTable(detail,people.Count);

                app.Range["A1"].Value2=year+"년 "+month+"월("+round+") 계약제교원 인건비 신청서";
                app.Range["A2"].Value2="담당자 : "+(info.ManagerName??"")+", 전화번호("+(info.Phone??"")+")";
                app.Range["O2"].Value2="세외계좌번호: ("+(info.BankName??"")+") "+(info.AccountNumber??"");
                app.Range["A5"].Value2=info.RecipientCode??"";app.Range["B5"].Value2=info.InstitutionName??"";
                detail.Range["A1"].Value2=year+"년 "+month+"월("+round+") 계약제교원 4대보험료 산출 서식";
                detail.Range["A4"].Value2=info.InstitutionName??"";

                int appLast=4+people.Count,detailLast=5+people.Count;
                app.Range["C5:P"+appLast].ClearContents();detail.Range["A6:I"+detailLast].ClearContents();
                for(int i=0;i<people.Count;i++)
                {
                    SubmitPerson x=people[i];int ar=5+i,dr=6+i;decimal health,longTerm;GetHealthParts(x,out health,out longTerm);
                    app.Cells[ar,3].Value2=x.Name;app.Cells[ar,7].Value2=x.Pension;app.Cells[ar,8].Value2=health;app.Cells[ar,9].Value2=longTerm;app.Cells[ar,10].Value2=x.Industrial;app.Cells[ar,11].Value2=x.Employment;
                    app.Cells[ar,6].Formula="=SUM(D"+ar+":E"+ar+")";app.Cells[ar,12].Formula="=SUM(G"+ar+":K"+ar+")";app.Cells[ar,13].Formula="=F"+ar+"+L"+ar;app.Cells[ar,15].Value2="4대보험료 산출내역(시트5번) 별첨";
                    detail.Cells[dr,1].Value2=i+1;detail.Cells[dr,2].Value2=x.Name;detail.Cells[dr,3].Value2=x.HealthBase;detail.Cells[dr,4].Value2=x.Pension;detail.Cells[dr,5].Value2=health;detail.Cells[dr,6].Value2=longTerm;detail.Cells[dr,7].Value2=x.Industrial;detail.Cells[dr,8].Value2=x.Employment;
                }
                int appTotal=appLast+1,detailTotal=detailLast+1;
                app.Cells[appTotal,1].Value2="합계";for(int c=4;c<=13;c++)app.Cells[appTotal,c].Formula="=SUM("+ColumnLetter(c)+"5:"+ColumnLetter(c)+appLast+")";
                detail.Cells[detailTotal,1].Value2="신청금액 합계";for(int c=4;c<=8;c++)detail.Cells[detailTotal,c].Formula="=SUM("+ColumnLetter(c)+"6:"+ColumnLetter(c)+detailLast+")";detail.Cells[detailTotal,9].Formula="=SUM(D"+detailTotal+":H"+detailTotal+")";
                excel.CutCopyMode=false;book.SaveAs(outputPath,51);
                return outputPath;
            }
            catch(Exception ex)
            {
                throw new InvalidOperationException("계약제교원 원본 서식 변환 중 오류가 발생했습니다. Excel을 모두 닫은 뒤 다시 실행해 주세요.\r\n"+ex.Message,ex);
            }
            finally
            {
                try{if(bookObject!=null)((dynamic)bookObject).Close(false);}catch{}
                try{if(excelObject!=null)((dynamic)excelObject).Quit();}catch{}
                ReleaseCom(detailSheetObject);ReleaseCom(appSheetObject);ReleaseCom(bookObject);ReleaseCom(booksObject);ReleaseCom(excelObject);
                try{if(Directory.Exists(tempFolder))Directory.Delete(tempFolder,true);}catch{}
                GC.Collect();GC.WaitForPendingFinalizers();
            }
        }

        static string CreateWorkerWithExcel(List<SubmitPerson> people,int year,int month,SubmissionInfo info,string outputFolder)
        {
            if(people.Count>95)throw new InvalidOperationException("교육공무직 제출 대상이 95명을 초과해 내장 서식 범위를 넘었습니다.");
            string round=RoundText(info.Round);
            string code=SafeFilePart(info.RecipientCode,"수신자기호"),institution=SafeFilePart(info.InstitutionName,"기관명");
            string outputPath=UniquePath(Path.Combine(outputFolder,code+"_"+institution+"_("+round+")"+year+". "+month+"월 4대보험 기관부담금 교육공무직원 인건비 신청.xlsx"));
            string tempFolder=Path.Combine(Path.GetTempPath(),"WorkerSubmission_"+Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempFolder);string templatePath=Path.Combine(tempFolder,"worker_template.xlsx");
            using(Stream resource=Assembly.GetExecutingAssembly().GetManifestResourceStream("InsurancePayrollValidator.WorkerTemplate.xlsx"))
            {
                if(resource==null)throw new InvalidOperationException("내장된 교육공무직 원본 서식을 불러오지 못했습니다.");
                using(FileStream file=new FileStream(templatePath,FileMode.Create,FileAccess.Write))resource.CopyTo(file);
            }

            object excelObject=null,booksObject=null,bookObject=null,sheetObject=null;
            try
            {
                Type excelType=Type.GetTypeFromProgID("Excel.Application");
                if(excelType==null)throw new InvalidOperationException("Microsoft Excel이 설치되어 있지 않아 제출 서식을 만들 수 없습니다.");
                excelObject=Activator.CreateInstance(excelType);dynamic excel=excelObject;
                excel.Visible=false;excel.DisplayAlerts=false;excel.ScreenUpdating=false;excel.EnableEvents=false;
                booksObject=excel.Workbooks;dynamic books=booksObject;bookObject=books.Open(templatePath);dynamic book=bookObject;
                sheetObject=book.Worksheets[1];dynamic ws=sheetObject;
                ws.Range["A1"].Value2=year+". "+month+"월 4대보험 기관부담금 - 인건비 신청("+round+")";
                decimal industrialRate=ParseRate(info.IndustrialRate);
                for(int i=0;i<people.Count;i++)
                {
                    SubmitPerson x=people[i];int r=7+i;decimal health,longTerm;GetHealthParts(x,out health,out longTerm);
                    bool shortTerm=IsFixedTermWorker(x);
                    ws.Cells[r,1].Value2=i+1;ws.Cells[r,2].Value2=info.RecipientCode??"";ws.Cells[r,3].Value2=info.InstitutionName??"";ws.Cells[r,4].Value2=info.ManagerName??"";
                    ws.Cells[r,5].Value2=x.Job;ws.Cells[r,6].Value2=x.Name;ws.Cells[r,7].Value2=shortTerm?"기간제":"무기";if(shortTerm)ws.Cells[r,8].Value2=x.Reason??"";ws.Cells[r,9].Value2="N";
                    if(shortTerm)
                    {
                        if(health!=0)ws.Cells[r,15].Value2=health;
                        if(longTerm!=0)ws.Cells[r,19].Value2=longTerm;
                        if(x.Pension!=0)ws.Cells[r,23].Value2=x.Pension;
                        ws.Cells[r,27].Value2=x.Employment;
                        ws.Cells[r,31].Value2=x.Industrial;
                    }
                    else
                    {
                        ws.Cells[r,10].Value2=x.HealthBase;
                        ws.Cells[r,11].Value2=x.PensionBase;
                        ws.Cells[r,12].Value2=x.EmploymentBase!=0?x.EmploymentBase:x.IndustrialBase;
                    }
                    ws.Cells[r,29].Value2=industrialRate;
                }
                excel.CalculateFullRebuild();excel.CutCopyMode=false;book.SaveAs(outputPath,51);return outputPath;
            }
            catch(Exception ex)
            {
                throw new InvalidOperationException("교육공무직 원본 서식 작성 중 오류가 발생했습니다. Excel을 모두 닫은 뒤 다시 실행해 주세요.\r\n"+ex.Message,ex);
            }
            finally
            {
                try{if(bookObject!=null)((dynamic)bookObject).Close(false);}catch{}
                try{if(excelObject!=null)((dynamic)excelObject).Quit();}catch{}
                ReleaseCom(sheetObject);ReleaseCom(bookObject);ReleaseCom(booksObject);ReleaseCom(excelObject);
                try{if(Directory.Exists(tempFolder))Directory.Delete(tempFolder,true);}catch{}
                GC.Collect();GC.WaitForPendingFinalizers();
            }
        }

        static void ResizeTeacherApplicationTable(dynamic sheet,int count)
        {
            sheet.Range["A5:A12"].UnMerge();sheet.Range["B5:B12"].UnMerge();int totalRow=13;
            if(count<8){sheet.Rows[(5+count)+":12"].Delete();totalRow=5+count;}
            else if(count>8){for(int i=0;i<count-8;i++){sheet.Rows[totalRow].Insert();sheet.Rows[totalRow-1].Copy(sheet.Rows[totalRow]);sheet.Rows[totalRow].RowHeight=sheet.Rows[totalRow-1].RowHeight;totalRow++;}}
            sheet.Range["A5:A"+(4+count)].Merge();sheet.Range["B5:B"+(4+count)].Merge();
        }

        static void ResizeTeacherInsuranceTable(dynamic sheet,int count)
        {
            sheet.Rows[14].Delete();int totalRow=14;
            if(count<8){sheet.Rows[(6+count)+":13"].Delete();totalRow=6+count;}
            else if(count>8){for(int i=0;i<count-8;i++){sheet.Rows[totalRow].Insert();sheet.Rows[totalRow-1].Copy(sheet.Rows[totalRow]);sheet.Rows[totalRow].RowHeight=sheet.Rows[totalRow-1].RowHeight;totalRow++;}}
        }

        static string ColumnLetter(int column){string result="";while(column>0){column--;result=(char)('A'+column%26)+result;column/=26;}return result;}
        static void ReleaseCom(object value){if(value!=null&&Marshal.IsComObject(value)){try{Marshal.FinalReleaseComObject(value);}catch{}}}

        static void FillTeacher(ExcelPackage p,List<SubmitPerson> people,int year,int month,SubmissionInfo info)
        {
            ExcelWorksheet app=p.Workbook.Worksheets.FirstOrDefault(x=>x.Name.StartsWith("1.")&&x.Name.Contains("인건비 신청서"));
            ExcelWorksheet detail=p.Workbook.Worksheets.FirstOrDefault(x=>x.Name.StartsWith("4.")&&x.Name.Contains("4대보험료"));
            if(app==null||detail==null)throw new InvalidOperationException("새 계약제교원 제출 서식의 신청서 또는 4대보험 산출내역 탭을 찾지 못했습니다.");
            if(people.Count>8)throw new InvalidOperationException("새 계약제교원 제출 서식은 표 크기를 유지하기 위해 최대 8명까지 작성할 수 있습니다. 계약제교원 대상이 "+people.Count+"명입니다.");

            string round=RoundText(info.Round);
            app.Cells[1,1].Value=year+"년 "+month+"월("+round+") 계약제교원 인건비 신청서";
            detail.Cells[1,1].Value=year+"년 "+month+"월("+round+") 계약제교원 4대보험료 산출 서식";
            app.Cells[2,1].Value="담당자 : "+(info.ManagerName??"")+", 전화번호("+(info.Phone??"")+")";
            app.Cells[2,15].Value="세외계좌번호: ("+(info.BankName??"")+") "+(info.AccountNumber??"");
            app.Cells[5,1].Value=info.RecipientCode??"";
            app.Cells[5,2].Value=info.InstitutionName??"";
            detail.Cells[4,1].Value=info.InstitutionName??"";

            // 고정된 제출용 표의 색상, 테두리, 행 높이, 열 너비는 유지하고 내용만 교체한다.
            for(int r=5;r<=12;r++)
            {
                for(int c=3;c<=16;c++)app.Cells[r,c].Value=null;
                app.Cells[r,6].Formula="SUM(D"+r+":E"+r+")";
                app.Cells[r,12].Formula="SUM(G"+r+":K"+r+")";
                app.Cells[r,13].Formula="F"+r+"+L"+r;
            }
            for(int r=6;r<=14;r++)for(int c=1;c<=9;c++)detail.Cells[r,c].Value=null;

            for(int i=0;i<people.Count;i++)
            {
                SubmitPerson x=people[i];int ar=5+i,dr=6+i;decimal health,longTerm;GetHealthParts(x,out health,out longTerm);
                app.Cells[ar,3].Value=x.Name;
                app.Cells[ar,7].Value=x.Pension;
                app.Cells[ar,8].Value=health;
                app.Cells[ar,9].Value=longTerm;
                app.Cells[ar,10].Value=x.Industrial;
                app.Cells[ar,11].Value=x.Employment;
                app.Cells[ar,15].Value="4대보험료 산출내역(시트5번) 별첨";

                detail.Cells[dr,1].Value=i+1;
                detail.Cells[dr,2].Value=x.Name;
                detail.Cells[dr,3].Value=x.HealthBase;
                detail.Cells[dr,4].Value=x.Pension;
                detail.Cells[dr,5].Value=health;
                detail.Cells[dr,6].Value=longTerm;
                detail.Cells[dr,7].Value=x.Industrial;
                detail.Cells[dr,8].Value=x.Employment;
            }

            app.Cells[13,1].Value="합계";
            for(int c=4;c<=13;c++)app.Cells[13,c].Formula="SUM("+app.Cells[5,c].Address+":"+app.Cells[12,c].Address+")";
            detail.Cells[15,1].Value="신청금액 합계";
            for(int c=4;c<=8;c++)detail.Cells[15,c].Formula="SUM("+detail.Cells[6,c].Address+":"+detail.Cells[14,c].Address+")";
            detail.Cells[15,9].Formula="SUM(D15:H15)";
        }

        static void FillWorker(ExcelPackage p,List<SubmitPerson> people,int year,int month,SubmissionInfo info)
        {
            ExcelWorksheet ws=p.Workbook.Worksheets[1];PrepareWorkerTemplateValidation(ws);ws.Cells[1,1].Value=year+". "+month+"월 4대보험 기관부담금 - 인건비 신청("+RoundText(info.Round)+")";
            if(people.Count>95)throw new InvalidOperationException("교육공무직 제출 대상이 95명을 초과해 내장 서식 범위를 넘었습니다.");
            decimal industrialRate=ParseRate(info.IndustrialRate);
            // 잠금된 원본 서식의 빈 행, 공유수식, 셀 스타일은 지우지 않는다.
            // 원본에는 보험별 계(P/T/X/AB/AF)와 전체 계(AH)가 7~101행 전체에 이미 들어 있다.
            for(int i=0;i<people.Count;i++)
            {
                SubmitPerson x=people[i];int r=7+i;decimal health,longTerm;GetHealthParts(x,out health,out longTerm);
                bool shortTerm=IsFixedTermWorker(x);ws.Cells[r,1].Value=i+1;ws.Cells[r,2].Value=info.RecipientCode??"";ws.Cells[r,3].Value=info.InstitutionName??"";ws.Cells[r,4].Value=info.ManagerName??"";ws.Cells[r,5].Value=x.Job;ws.Cells[r,6].Value=x.Name;ws.Cells[r,7].Value=shortTerm?"기간제":"무기";ws.Cells[r,8].Value=shortTerm?(object)x.Reason:null;ws.Cells[r,9].Value="N";
                if(shortTerm)
                {
                    // 대체근로자는 월보수액과 당월 보험료를 건드리지 않고 보험별 정산보험료 칸에만 실제 금액을 기록한다.
                    ws.Cells[r,15].Value=health;
                    ws.Cells[r,19].Value=longTerm;
                    ws.Cells[r,23].Value=x.Pension;
                    ws.Cells[r,27].Value=x.Employment;
                    ws.Cells[r,31].Value=x.Industrial;
                }
                else
                {
                    // 일반 교육공무직은 월보수액만 채우고 보험료 계산은 잠금된 기본 서식의 수식에 맡긴다.
                    ws.Cells[r,10].Value=x.HealthBase;
                    ws.Cells[r,11].Value=x.PensionBase;
                    ws.Cells[r,12].Value=x.EmploymentBase!=0?x.EmploymentBase:x.IndustrialBase;
                    ws.Cells[r,29].Value=industrialRate;
                }
            }
        }

        static void PrepareWorkerTemplateValidation(ExcelWorksheet ws)
        {
            // 원본 직종 드롭다운은 255자를 넘는 직접 목록이라 구버전 EPPlus 저장 시 오류가 난다.
            // 목록 내용은 그대로 AZ 숨김열로 옮기고 유효성 검사의 참조만 짧은 셀 범위로 바꾼다.
            const string mainNs="http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            XmlNodeList validations=ws.WorksheetXml.GetElementsByTagName("dataValidation",mainNs);
            foreach(XmlNode validation in validations)
            {
                XmlAttribute range=validation.Attributes["sqref"];if(range==null||!range.Value.Split(' ').Contains("E7:E101"))continue;
                XmlNode formula=null;foreach(XmlNode child in validation.ChildNodes)if(child.LocalName=="formula1"){formula=child;break;}
                if(formula==null)continue;string raw=formula.InnerText.Trim();if(raw.Length<=255||!raw.StartsWith("\"")||!raw.EndsWith("\""))continue;
                raw=raw.Substring(1,raw.Length-2).Replace("\"\"","\"");string[] values=raw.Split(new[]{','},StringSplitOptions.RemoveEmptyEntries).Select(x=>x.Trim()).Where(x=>x.Length>0).ToArray();
                for(int i=0;i<values.Length;i++)ws.Cells[i+1,52].Value=values[i];
                ws.Column(52).Hidden=true;formula.InnerText="$AZ$1:$AZ$"+values.Length;break;
            }
        }

        static void SplitHealth(decimal total,out decimal health,out decimal longTerm)
        {
            if(total==0){health=0;longTerm=0;return;}decimal sign=total<0?-1:1,abs=Math.Abs(total);health=Math.Floor((abs/1.1314m)/10m)*10m*sign;longTerm=total-health;
        }
        static void GetHealthParts(SubmitPerson person,out decimal health,out decimal longTerm){if(person.HasHealthParts){health=person.HealthOnly;longTerm=person.LongTerm;}else SplitHealth(person.Health,out health,out longTerm);}
        static Dictionary<string,HealthParts> ReadHealthParts(ExcelPackage package)
        {
            var result=new Dictionary<string,HealthParts>();
            foreach(ExcelWorksheet ws in package.Workbook.Worksheets.Where(x=>x.Name.Contains("건강보험")||x.Name.Contains("건강비공무원")))
            {
                if(ws.Dimension==null)continue;int header=0,nameCol=0,birthCol=0,healthCol=0,healthYearCol=0,healthInterestCol=0,longCol=0,longYearCol=0,longInterestCol=0;
                for(int r=1;r<=Math.Min(10,ws.Dimension.End.Row)&&header==0;r++)for(int c=1;c<=Math.Min(80,ws.Dimension.End.Column);c++)if(NormalizeHeader(Convert.ToString(ws.Cells[r,c].Value))=="성명"){header=r;nameCol=c;break;}
                if(header==0)continue;
                for(int c=1;c<=Math.Min(80,ws.Dimension.End.Column);c++)
                {
                    string h=NormalizeHeader(Convert.ToString(ws.Cells[header,c].Value));if(h=="주민등록번호"||h=="생년월일")birthCol=c;else if(h=="고지금액")healthCol=c;else if(h=="연말정산")healthYearCol=c;else if(h=="건강환급금이자")healthInterestCol=c;else if(h=="요양고지보험료")longCol=c;else if(h=="요양연말정산보험료")longYearCol=c;else if(h=="요양환급금이자")longInterestCol=c;
                }
                if(nameCol==0||healthCol==0||longCol==0)continue;
                for(int r=header+1;r<=ws.Dimension.End.Row;r++)
                {
                    string name=NormalizeName(Convert.ToString(ws.Cells[r,nameCol].Value));if(name.Length==0)continue;string birth=birthCol>0?NormalizeBirth(Convert.ToString(ws.Cells[r,birthCol].Value)):"";string key=PersonKey(name,birth);HealthParts part;if(!result.TryGetValue(key,out part)){part=new HealthParts();result[key]=part;}part.Health+=CellDecimal(ws,r,healthCol)+CellDecimal(ws,r,healthYearCol)+CellDecimal(ws,r,healthInterestCol);part.LongTerm+=CellDecimal(ws,r,longCol)+CellDecimal(ws,r,longYearCol)+CellDecimal(ws,r,longInterestCol);
                }
            }
            return result;
        }
        static Dictionary<string,decimal> ReadWageBases(ExcelPackage package,string sheetNamePart,string[] baseAliases)
        {
            var result=new Dictionary<string,decimal>();
            foreach(ExcelWorksheet ws in package.Workbook.Worksheets.Where(x=>x.Name.Contains(sheetNamePart)))
            {
                if(ws.Dimension==null)continue;int header=0,nameCol=0,birthCol=0,baseCol=0;
                for(int r=1;r<=Math.Min(12,ws.Dimension.End.Row);r++)
                {
                    int nc=0,bc=0,vc=0;for(int c=1;c<=Math.Min(100,ws.Dimension.End.Column);c++){string h=NormalizeHeader(Convert.ToString(ws.Cells[r,c].Value));if(h=="성명"||h=="근로자명"||h=="가입자명")nc=c;if(h=="주민등록번호"||h=="생년월일")bc=c;if(baseAliases.Any(a=>h.Contains(NormalizeHeader(a))))vc=c;}
                    if(nc>0&&vc>0){header=r;nameCol=nc;birthCol=bc;baseCol=vc;break;}
                }
                if(header==0)continue;
                for(int r=header+1;r<=ws.Dimension.End.Row;r++){string name=NormalizeName(Convert.ToString(ws.Cells[r,nameCol].Value));if(name.Length==0)continue;string birth=birthCol>0?NormalizeBirth(Convert.ToString(ws.Cells[r,birthCol].Value)):"";decimal amount=CellDecimal(ws,r,baseCol);string key=PersonKey(name,birth);if(!result.ContainsKey(key)||result[key]==0)result[key]=amount;}
            }
            return result;
        }
        static void SanitizeWorkerOutput(string path)
        {
            using(FileStream fs=new FileStream(path,FileMode.Open,FileAccess.ReadWrite,FileShare.None))using(ZipArchive zip=new ZipArchive(fs,ZipArchiveMode.Update))
            {
                string[] names=zip.Entries.Where(e=>e.FullName.StartsWith("xl/comments",StringComparison.OrdinalIgnoreCase)&&e.FullName.EndsWith(".xml",StringComparison.OrdinalIgnoreCase)).Select(e=>e.FullName).ToArray();
                foreach(string name in names){ZipArchiveEntry entry=zip.GetEntry(name);string xml;using(StreamReader sr=new StreamReader(entry.Open(),Encoding.UTF8)){xml=sr.ReadToEnd();}xml=Regex.Replace(xml,"\\s+shapeId=\"[^\"]*\"","");entry.Delete();ZipArchiveEntry replacement=zip.CreateEntry(name,System.IO.Compression.CompressionLevel.Optimal);using(StreamWriter sw=new StreamWriter(replacement.Open(),new UTF8Encoding(false))){sw.Write(xml);}}
                ZipArchiveEntry workbook=zip.GetEntry("xl/workbook.xml");if(workbook!=null){string xml;using(StreamReader sr=new StreamReader(workbook.Open(),Encoding.UTF8)){xml=sr.ReadToEnd();}xml=Regex.Replace(xml,"<definedName\\b(?![^>]*name=\"_xlnm\\.)[^>]*>.*?</definedName>","",RegexOptions.Singleline);workbook.Delete();ZipArchiveEntry replacement=zip.CreateEntry("xl/workbook.xml",System.IO.Compression.CompressionLevel.Optimal);using(StreamWriter sw=new StreamWriter(replacement.Open(),new UTF8Encoding(false))){sw.Write(xml);}}
            }
        }
        static decimal ReadDecimal(object value){decimal n;Decimal.TryParse(Regex.Replace(Convert.ToString(value,CultureInfo.InvariantCulture)??"","[^0-9.-]",""),NumberStyles.Any,CultureInfo.InvariantCulture,out n);return n;}
        static decimal CellDecimal(ExcelWorksheet ws,int row,int col){return col>0?ReadDecimal(ws.Cells[row,col].Value):0;}
        static string NormalizeHeader(string value){return Regex.Replace((value??"").ToLowerInvariant(),"[^0-9a-z가-힣]","");}
        static string NormalizeName(string value){return Regex.Replace(value??"",@"\s+","");}
        static string NormalizeBirth(string value){string s=Regex.Replace(value??"","[^0-9]","");if(s.Length>=8)return s.Substring(s.Length==8?2:0,6);return s.Length>=6?s.Substring(0,6):s;}
        static string PersonKey(string name,string birth){return NormalizeName(name)+"|"+birth;}
        static int ToInt(object value,int fallback){int n;return Int32.TryParse(Convert.ToString(value,CultureInfo.InvariantCulture),out n)&&n>0?n:fallback;}
        static string CellText(ExcelPackage p,ExcelWorksheet ws,int row,int col)
        {
            string text=Convert.ToString(ws.Cells[row,col].Value,CultureInfo.InvariantCulture);if(!String.IsNullOrWhiteSpace(text))return text.Trim();
            string f=ws.Cells[row,col].Formula;if(String.IsNullOrWhiteSpace(f))return "";Match m=Regex.Match(f,@"'?([^']+)'?!\$?([A-Z]+)\$?(\d+)");if(!m.Success)return "";ExcelWorksheet refWs=p.Workbook.Worksheets[m.Groups[1].Value];if(refWs==null)return "";return Convert.ToString(refWs.Cells[Int32.Parse(m.Groups[3].Value),ColumnNumber(m.Groups[2].Value)].Value,CultureInfo.InvariantCulture).Trim();
        }
        static decimal CellNumber(ExcelPackage p,ExcelWorksheet ws,int row,int col)
        {
            decimal n;if(Decimal.TryParse(Convert.ToString(ws.Cells[row,col].Value,CultureInfo.InvariantCulture),NumberStyles.Any,CultureInfo.InvariantCulture,out n))return n;
            string f=ws.Cells[row,col].Formula;if(String.IsNullOrWhiteSpace(f))return 0;Match m=Regex.Match(f,@"'?([^']+)'?!\$?([A-Z]+)\$?(\d+)");if(!m.Success)return 0;ExcelWorksheet refWs=p.Workbook.Worksheets[m.Groups[1].Value];if(refWs==null)return 0;Decimal.TryParse(Convert.ToString(refWs.Cells[Int32.Parse(m.Groups[3].Value),ColumnNumber(m.Groups[2].Value)].Value,CultureInfo.InvariantCulture),NumberStyles.Any,CultureInfo.InvariantCulture,out n);return n;
        }
        static int ColumnNumber(string letters){int n=0;foreach(char c in letters.ToUpperInvariant())n=n*26+(c-'A'+1);return n;}
        static decimal ParseRate(string value){string s=(value??"").Trim();bool percent=s.Contains("%");s=s.Replace("%","").Replace(",","");decimal rate;if(!Decimal.TryParse(s,NumberStyles.Any,CultureInfo.InvariantCulture,out rate))rate=0.008m;if(percent)rate/=100m;return rate;}
        static string SafeFilePart(string value,string fallback){string s=String.IsNullOrWhiteSpace(value)?fallback:value.Trim();foreach(char c in Path.GetInvalidFileNameChars())s=s.Replace(c,'_');return s;}
        static string RoundText(string value){string s=(value??"").Trim();if(s.Length==0)return "n차";string digits=Regex.Replace(s,"[^0-9]","");if(digits.Length>0)return digits+"차";s=Regex.Replace(s,"차수?$","");return s.Length==0?"n차":s+"차";}
        static string UniquePath(string path){if(!File.Exists(path))return path;string dir=Path.GetDirectoryName(path),name=Path.GetFileNameWithoutExtension(path),ext=Path.GetExtension(path);int i=2;do{path=Path.Combine(dir,name+" ("+i+")"+ext);i++;}while(File.Exists(path));return path;}
    }
}
