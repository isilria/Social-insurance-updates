using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using OfficeOpenXml;

namespace InsurancePayrollValidator
{
    // Test-channel state is deliberately kept apart from the installed stable version.
    static class TestStore202
    {
        public static string Root { get { string p=Environment.GetEnvironmentVariable("SOCIAL_INSURANCE_TEST_HOME"); if(String.IsNullOrWhiteSpace(p))p=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"살구아빠","사회보험_2.0.2_테스트");Directory.CreateDirectory(p);return p; } }
        public static string FilePath(string name){return Path.Combine(Root,name);}
        public static string Encode(string text){return Convert.ToBase64String(Encoding.UTF8.GetBytes(text??""));}
        public static string Decode(string text){return Encoding.UTF8.GetString(Convert.FromBase64String(text));}
        public static List<string[]> Read(string name){var result=new List<string[]>();string path=FilePath(name);if(!File.Exists(path))return result;foreach(string line in File.ReadAllLines(path,Encoding.UTF8))try{result.Add(line.Split('\t').Select(Decode).ToArray());}catch(FormatException){}return result;}
        public static void Write(string name,IEnumerable<string[]> rows){string path=FilePath(name),temp=path+".tmp";File.WriteAllLines(temp,rows.Select(r=>String.Join("\t",r.Select(Encode))),Encoding.UTF8);if(File.Exists(path))File.Replace(temp,path,null);else File.Move(temp,path);}
    }

    partial class MainForm
    {
        readonly Dictionary<string,string> importOrigins202=new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase);
        readonly List<string[]> importErrors202=new List<string[]>();
        readonly List<string> extractionFolders202=new List<string>();
        readonly Dictionary<string,string> reviewNotes202=new Dictionary<string,string>();
        bool externalResult202;
        Label toolsStatus202;
        string lastSavedHash202;
        readonly HashSet<string> retainedReviewKeys202=new HashSet<string>(StringComparer.Ordinal);

        bool IsListedReview202(IndividualRowData row){return row!=null&&(retainedReviewKeys202.Contains(ReviewKey(row))||IsReviewRow(row));}
        void RememberReviewRows202(){if(individualDashboard!=null)foreach(var row in individualDashboard.Rows.Where(IsReviewRow))retainedReviewKeys202.Add(ReviewKey(row));}
        void LoadReviewState(ExcelPackage p)
        {
            LoadReviewStateBase(p);retainedReviewKeys202.Clear();var ws=p.Workbook.Worksheets["UI확인목록"];
            if(ws!=null&&ws.Dimension!=null)for(int r=2;r<=ws.Dimension.End.Row;r++)if(!String.IsNullOrWhiteSpace(ws.Cells[r,1].Text))retainedReviewKeys202.Add(ws.Cells[r,1].Text);
        }
        string[] ReviewFundsForSite202(string site)
        {
            string[] allowed={"공무원","계약제교원","교특회계","학교회계"};
            var present=new HashSet<string>((individualDashboard==null?Enumerable.Empty<IndividualRowData>():individualDashboard.Rows).Where(x=>x.Site==site).Select(x=>x.Fund));
            var choices=allowed.Where(present.Contains).ToArray();return choices.Length>0?choices:allowed;
        }
        void UpdateReviewFundOptions202()
        {
            if(reviewApplyFundSelector==null)return;string site=reviewSiteSelector!=null&&reviewSiteSelector.SelectedIndex>=0&&reviewSiteSelector.SelectedIndex<reviewSiteKeys.Count?reviewSiteKeys[reviewSiteSelector.SelectedIndex]:"";
            bool was=reviewFilterLoading;reviewFilterLoading=true;reviewApplyFundSelector.SetItems(new[]{"적용 재원 선택"}.Concat(ReviewFundsForSite202(site)));reviewApplyFundSelector.SelectedIndex=0;reviewFilterLoading=was;
        }
        void WriteRetainedReview202(ExcelPackage p)
        {
            RememberReviewRows202();var ws=p.Workbook.Worksheets["UI확인목록"]??p.Workbook.Worksheets.Add("UI확인목록");ws.Cells.Clear();ws.Cells[1,1].Value="대상키";int r=2;foreach(string key in retainedReviewKeys202.OrderBy(x=>x))ws.Cells[r++,1].Value=key;ws.Hidden=eWorkSheetHidden.Hidden;
        }

        void Initialize202()
        {
            Text="사회보험 재원별 대사 보조 도우미 Ver. 2.0.2";sidebarVersionLabel.Text="Ver 2.0.2";
            foreach(Control c in Descendants202(pages["설정"])){
                Label l=c as Label;if(l!=null&&l.Text.Contains("현재 버전"))l.Text="현재 버전  Ver. 2.0.2";
                // Official release keeps the user's existing update preference.
            }
            var button=OutputButton("인식 내역","search",854,9,180,34,UiBlue,false);button.Tag="ThemeAccentAction";button.Click+=(s,e)=>Safe202(ShowRecognition202);pages["파일 등록"].Controls.Add(button);
            // Closing is immediate. Explicit save actions remain available.
            FormClosed+=(s,e)=>{foreach(string folder in extractionFolders202)try{Directory.Delete(folder,true);}catch{}};
        }
        IEnumerable<Control> Descendants202(Control parent){foreach(Control c in parent.Controls){yield return c;foreach(Control child in Descendants202(c))yield return child;}}
        void Safe202(Action action){try{action();}catch(Exception ex){string path=TestStore202.FilePath("최근오류.txt");try{File.WriteAllText(path,DateTime.Now.ToString("s")+Environment.NewLine+ex,Encoding.UTF8);}catch{}MessageBox.Show(this,ex.Message+"\r\n\r\n오류 기록: "+path,"작업 확인",MessageBoxButtons.OK,MessageBoxIcon.Error);}}
        void BuildTools202(Control page)
        {
            page.Controls.Add(TitleLabel("작업 관리",8,10,20F));
            page.Controls.Add(new Label{Text="Ver. 2.0.2 테스트 · 저장한 결과를 다시 열고, 인식 내역과 처리 이력을 확인합니다.",Location=new Point(10,52),AutoSize=true,ForeColor=UiMuted});
            string[] labels={"인식 내역 / 미인식 파일","현재 작업 저장","저장한 작업 열기","최근 작업","확인완료 이력 / 되돌리기","재원 분류 규칙","지난 자료와 금액 비교"};
            string[] notes={"파일별 종류·사업장·인원, ZIP 내부 파일 및 오류 확인","보정·감면·확인 이력을 포함한 XLSM 저장","저장된 XLSM의 사본을 열어 이어서 작업","최근 저장·불러오기 목록에서 선택","확인 사유·처리일을 조회하고 미확인으로 복원","사업장별 성명·직종 규칙을 저장하고 적용 전 확인","다른 월의 저장 결과와 보험별 기관부담 증감 비교"};
            Action[] actions={ShowRecognition202,SaveWorkspace202,OpenWorkspace202,ShowRecent202,ShowAudit202,EditRules202,Compare202};
            for(int i=0;i<labels.Length;i++){int index=i;int y=91+i*65;var card=Card(8,y,1030,57,UiTheme.Card);var b=OutputButton(labels[i],"file",14,9,250,38,UiBlue,false);b.Tag="ThemeAccentAction";b.Click+=(s,e)=>Safe202(actions[index]);card.Controls.Add(b);card.Controls.Add(new Label{Text=notes[i],Location=new Point(288,21),AutoSize=true,ForeColor=UiMuted});page.Controls.Add(card);}
            toolsStatus202=new Label{Text="기관정보·계좌번호·산재요율은 제출서 생성 화면에서 입력하면 다음 실행에도 유지됩니다.",Location=new Point(12,558),Size=new Size(1000,48),ForeColor=UiMuted};page.Controls.Add(toolsStatus202);
        }
        void RequireResult202(){if(individualDashboard==null||validationResult==null||!File.Exists(validationResult.Text))throw new InvalidOperationException("먼저 대사를 실행하거나 저장한 작업을 열어 주세요.");}
        void AnalyzeRegisteredFiles(IEnumerable<string> paths)
        {
            var flat=new List<string>();foreach(string path in paths.Distinct(StringComparer.OrdinalIgnoreCase)){
                if(!File.Exists(path)){importErrors202.Add(new[]{Path.GetFileName(path),"파일 없음",path});continue;}
                if(!Path.GetExtension(path).Equals(".zip",StringComparison.OrdinalIgnoreCase)){flat.Add(path);importOrigins202[path]=Path.GetFileName(path);continue;}
                string dir=Path.Combine(Path.GetTempPath(),"Insurance202_"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(dir);extractionFolders202.Add(dir);
                try{using(ZipArchive zip=ZipFile.OpenRead(path)){int count=0;foreach(ZipArchiveEntry entry in zip.Entries){if(String.IsNullOrEmpty(entry.Name))continue;string ext=Path.GetExtension(entry.Name).ToLowerInvariant();if(!new[]{".xlsx",".xlsm",".xls"}.Contains(ext)){importErrors202.Add(new[]{Path.GetFileName(path)+" / "+entry.FullName,"지원하지 않는 파일",ext});continue;}
                    if(entry.Length>150L*1024*1024||count>=500)throw new InvalidDataException("ZIP의 파일 크기 또는 개수가 처리 범위를 초과했습니다.");
                    string itemDir=Path.Combine(dir,(count++).ToString());Directory.CreateDirectory(itemDir);string dest=Path.Combine(itemDir,entry.Name);entry.ExtractToFile(dest);flat.Add(dest);importOrigins202[dest]=Path.GetFileName(path)+" / "+entry.FullName;}
                    if(count==0)importErrors202.Add(new[]{Path.GetFileName(path),"Excel 파일 없음","ZIP 내용을 확인해 주세요."});}}
                catch(Exception ex){importErrors202.Add(new[]{Path.GetFileName(path),"압축 읽기 오류",ex.Message});}
            }
            if(flat.Count>0)try{AnalyzePreparedFiles(flat);}finally{UseWaitCursor=false;}
            if(importErrors202.Count>0)fileAnalysisStatus.Text="인식 내역에서 확인할 파일 "+importErrors202.Count+"건";
        }
        List<string[]> RecognitionRows202()
        {
            var rows=new List<string[]>();foreach(string path in registeredFileSites.Keys){var kinds=registeredFiles.Where(x=>x.Value.Contains(path,StringComparer.OrdinalIgnoreCase)).Select(x=>x.Key).ToArray();HashSet<string> people;registeredFilePeople.TryGetValue(path,out people);string origin;importOrigins202.TryGetValue(path,out origin);rows.Add(new[]{origin??Path.GetFileName(path),String.Join(", ",kinds),String.Join(", ",registeredFileSites[path].Select(FormatSite)),people==null?"":people.Count.ToString(),kinds.Length==0?"미인식":registeredFileSites[path].Count==0?"사업장 연결 확인":"인식됨"});}
            rows.AddRange(importErrors202.Select(x=>new[]{x[0],"","","",x[1]+": "+x[2]}));
            if(validationResult!=null&&File.Exists(validationResult.Text))using(var p=new ExcelPackage(new FileInfo(validationResult.Text))){var ws=p.Workbook.Worksheets["자료인식"];if(ws!=null&&ws.Dimension!=null)for(int r=2;r<=ws.Dimension.End.Row;r++)rows.Add(new[]{ws.Cells[r,2].Text,ws.Cells[r,1].Text,"(대사 처리 기록)",ws.Cells[r,5].Text,String.Join(" / ",new[]{ws.Cells[r,3].Text,ws.Cells[r,6].Text,ws.Cells[r,7].Text}.Where(x=>x.Length>0))});}
            return rows;
        }
        void ShowRecognition202(){ShowTable202("파일 인식 내역 · 인원은 파일 내 식별값 기준",new[]{"파일 / ZIP 내부 경로","자료 종류","사업장관리번호","인원","확인 사항"},RecognitionRows202());}
        Form TableForm202(string title,string[] headers,IEnumerable<string[]> rows,out DataGridView grid)
        {
            var form=new Form{Text=title,StartPosition=FormStartPosition.CenterParent,Size=new Size(1080,600),MinimumSize=new Size(700,420),Font=new Font("맑은 고딕",9F),BackColor=UiTheme.Page};
            grid=new DataGridView{Dock=DockStyle.Fill,ReadOnly=true,AllowUserToAddRows=false,AllowUserToDeleteRows=false,RowHeadersVisible=false,AutoSizeColumnsMode=DataGridViewAutoSizeColumnsMode.Fill,BackgroundColor=UiTheme.Card,SelectionMode=DataGridViewSelectionMode.FullRowSelect,MultiSelect=false};
            grid.DefaultCellStyle.BackColor=UiTheme.Card;grid.DefaultCellStyle.ForeColor=UiTheme.Text;grid.DefaultCellStyle.Alignment=DataGridViewContentAlignment.MiddleCenter;grid.AutoSizeRowsMode=DataGridViewAutoSizeRowsMode.AllCells;grid.DefaultCellStyle.WrapMode=DataGridViewTriState.True;
            for(int i=0;i<headers.Length;i++)grid.Columns.Add("c"+i,headers[i]);foreach(string[] row in rows)grid.Rows.Add(row.Cast<object>().ToArray());form.Controls.Add(grid);return form;
        }
        void ShowTable202(string title,string[] headers,IEnumerable<string[]> rows){DataGridView g;using(Form form=TableForm202(title,headers,rows,out g))form.ShowDialog(this);}
        void EnsureEditable202()
        {
            if(!externalResult202)return;string original=validationResult.Text;string target=NewTemporaryResultPath();File.Copy(original,target);temporaryResultPath=target;validationResult.Text=target;externalResult202=false;
        }
        void SaveWorkspace202()
        {
            RequireResult202();if(discountDrafts.Count>0)PersistDiscountChangesCore(validationResult.Text);PersistReviewChangesCore(validationResult.Text);
            using(var d=new SaveFileDialog{Filter="대사 작업 (*.xlsm)|*.xlsm",FileName="사회보험_대사_"+individualDashboard.Year+"년_"+individualDashboard.Month+"월_"+DateTime.Now.ToString("yyyyMMdd_HHmmss")+".xlsm",OverwritePrompt=true})if(d.ShowDialog(this)==DialogResult.OK){if(!String.Equals(Path.GetFullPath(validationResult.Text),Path.GetFullPath(d.FileName),StringComparison.OrdinalIgnoreCase))File.Copy(validationResult.Text,d.FileName,true);Remember202(d.FileName);lastSavedHash202=Hash202(validationResult.Text);if(toolsStatus202!=null)toolsStatus202.Text="저장 완료: "+d.FileName;}
        }
        void OpenWorkspace202(){using(var d=new OpenFileDialog{Filter="저장한 대사 결과 (*.xlsm;*.xlsx)|*.xlsm;*.xlsx"})if(d.ShowDialog(this)==DialogResult.OK)OpenWorkspacePath202(d.FileName);}
        void OpenWorkspacePath202(string path)
        {
            using(var p=new ExcelPackage(new FileInfo(path)))if(p.Workbook.Worksheets["UI개인별데이터"]==null)throw new InvalidDataException("2.0 계열에서 저장한 대사 결과 파일을 선택해 주세요.");
            string target=NewTemporaryResultPath();File.Copy(path,target);temporaryResultPath=target;lastSavedHash202=Hash202(target);validationResult.Text=target;externalResult202=false;reviewNotes202.Clear();reviewFundDrafts.Clear();discountDrafts.Clear();LoadResultIntoUi(target);Remember202(path);ShowPage("총괄표");
        }
        void Remember202(string path){var rows=TestStore202.Read("recent.tsv").Where(r=>r.Length>=2&&!String.Equals(r[1],path,StringComparison.OrdinalIgnoreCase)).ToList();rows.Insert(0,new[]{DateTime.Now.ToString("yyyy-MM-dd HH:mm"),path});TestStore202.Write("recent.tsv",rows.Take(20));}
        void ShowRecent202(){var rows=TestStore202.Read("recent.tsv").Where(x=>x.Length>=2).Select(x=>new[]{x[0],x[1],File.Exists(x[1])?"열기 가능":"파일 없음"}).ToList();DataGridView g;using(Form f=TableForm202("최근 작업 · 더블클릭하여 열기",new[]{"최근 작업일","파일","상태"},rows,out g)){g.CellDoubleClick+=(s,e)=>{if(e.RowIndex<0)return;string path=Convert.ToString(g.Rows[e.RowIndex].Cells[1].Value);Safe202(()=>{OpenWorkspacePath202(path);f.Close();});};f.ShowDialog(this);}}
        string AskReason202()
        {
            using(var f=new Form{Text="확인완료 사유",Size=new Size(480,210),StartPosition=FormStartPosition.CenterParent,FormBorderStyle=FormBorderStyle.FixedDialog,MaximizeBox=false,MinimizeBox=false}){
                var box=new ComboBox{Location=new Point(20,35),Width=420,DropDownStyle=ComboBoxStyle.DropDown};box.Items.AddRange(new object[]{"급여 반영 예정","정산분 확인","휴직 관련 확인","대체근로자 확인","원자료 대조 완료"});f.Controls.Add(new Label{Text="확인한 사유를 선택하거나 입력해 주세요.",Location=new Point(20,12),AutoSize=true});f.Controls.Add(box);var ok=new Button{Text="확인완료 저장",Location=new Point(190,105),Width=140};ok.Click+=(s,e)=>{if(!String.IsNullOrWhiteSpace(box.Text))f.DialogResult=DialogResult.OK;};var cancel=new Button{Text="취소",Location=new Point(340,105),Width=100,DialogResult=DialogResult.Cancel};f.Controls.Add(ok);f.Controls.Add(cancel);f.AcceptButton=ok;f.CancelButton=cancel;return f.ShowDialog(this)==DialogResult.OK?box.Text.Trim():null;}
        }
        void MarkSelectedReviewsChecked()
        {
            var selected=FilteredReviewRows().Where(x=>reviewSelections.Contains(ReviewKey(x))).ToList();if(selected.Count==0){reviewSelectionLabel.Text="대상을 체크해 주세요.";return;}
            Safe202(()=>{bool undo=selected.All(IsReviewCompleted);foreach(var row in selected){if(undo)reviewCheckedKeys.Remove(ReviewKey(row));else reviewCheckedKeys.Add(ReviewKey(row));}PersistReviewChangesCore(validationResult.Text);reviewSelectionLabel.Text="저장 완료";reviewAmountLabel.Text=selected.Count+"명 "+(undo?"미확인으로 변경":"확인완료 반영");});
        }
        void PersistReviewState(string path)
        {
            var changes=new List<string[]>();using(var p=new ExcelPackage(new FileInfo(path))){var states=p.Workbook.Worksheets["UI확인상태"];var old=new HashSet<string>();if(states!=null&&states.Dimension!=null)for(int r=2;r<=states.Dimension.End.Row;r++)old.Add(states.Cells[r,1].Text);
                var data=p.Workbook.Worksheets["UI개인별데이터"];if(data!=null&&data.Dimension!=null)for(int r=2;r<=data.Dimension.End.Row;r++){string key=String.Join("|",new[]{data.Cells[r,1].Text,data.Cells[r,3].Text,Regex.Replace(data.Cells[r,4].Text,"[^0-9]","")});var row=individualDashboard.Rows.FirstOrDefault(x=>StablePersonKey(x)==key);if(row==null)continue;bool before=old.Contains(key),after=reviewCheckedKeys.Contains(key);string note;reviewNotes202.TryGetValue(key,out note);if(before!=after||data.Cells[r,2].Text!=row.Fund)changes.Add(new[]{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),row.Site,row.Name,before!=after?(after?"확인완료":"미확인 복원"):"재원 변경",data.Cells[r,2].Text,row.Fund,note??"재원 분류 변경",Environment.UserName,key});}}
            string temp=path+"."+Guid.NewGuid().ToString("N")+".xlsm";try{File.Copy(path,temp);PersistReviewStateBase(temp);using(var p=new ExcelPackage(new FileInfo(temp))){var log=p.Workbook.Worksheets["처리이력"]??p.Workbook.Worksheets.Add("처리이력");string[] h={"처리일","사업장","성명","처리","변경 전 재원","변경 후 재원","사유","처리자","대상키"};for(int c=0;c<h.Length;c++)log.Cells[1,c+1].Value=h[c];int index=(log.Dimension==null?1:log.Dimension.End.Row)+1;foreach(var row in changes){for(int c=0;c<row.Length;c++)log.Cells[index,c+1].Value=row[c];index++;}log.Column(9).Hidden=true;log.Cells[1,1,Math.Max(1,index-1),8].Style.HorizontalAlignment=OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;log.Cells[1,1,1,8].Style.Font.Bold=true;for(int c=1;c<=8;c++)log.Column(c).Width=c==7?35:22;WriteRetainedReview202(p);p.Save();}File.Replace(temp,path,null);reviewNotes202.Clear();}finally{if(File.Exists(temp))File.Delete(temp);}
        }
        void ShowAudit202()
        {
            RequireResult202();var rows=individualDashboard.Rows.Where(IsReviewCompleted).Select(x=>new[]{x.Site,x.Name,x.Fund,ReviewKey(x)}).ToList();DataGridView g;using(Form f=TableForm202("확인완료 명단 · 더블클릭하면 미확인으로 복원",new[]{"사업장","성명","재원","대상키"},rows,out g)){
                g.Columns[3].Visible=false;var history=new Button{Text="전체 처리 이력 보기",Dock=DockStyle.Bottom,Height=42};f.Controls.Add(history);history.Click+=(s,e)=>Safe202(()=>{using(var p=new ExcelPackage(new FileInfo(validationResult.Text))){var ws=p.Workbook.Worksheets["처리이력"];var entries=new List<string[]>();if(ws!=null&&ws.Dimension!=null)for(int r=2;r<=ws.Dimension.End.Row;r++)entries.Add(Enumerable.Range(1,8).Select(c=>ws.Cells[r,c].Text).ToArray());ShowTable202("처리 이력",new[]{"처리일","사업장","성명","처리","변경 전","변경 후","사유","처리자"},entries);}});
                g.CellDoubleClick+=(s,e)=>{if(e.RowIndex<0)return;if(MessageBox.Show(f,"선택한 대상을 다시 확인 필요 상태로 돌릴까요?","미확인 복원",MessageBoxButtons.YesNo)!=DialogResult.Yes)return;Safe202(()=>{string key=Convert.ToString(g.Rows[e.RowIndex].Cells[3].Value);reviewCheckedKeys.Remove(key);reviewNotes202[key]="사용자가 미확인으로 복원";PersistReviewChangesCore(validationResult.Text);g.Rows.RemoveAt(e.RowIndex);});};f.ShowDialog(this);}
        }
        void EditRules202()
        {
            DataGridView grid;using(Form f=TableForm202("재원 규칙 · 위쪽 규칙 우선 / 빈 사업장은 모든 사업장",new[]{"사업장번호(선택)","성명 또는 직종","일치 문자열","적용 재원"},new List<string[]>(),out grid)){
                grid.ReadOnly=false;grid.AllowUserToAddRows=true;grid.AllowUserToDeleteRows=true;
                grid.Columns.RemoveAt(3);grid.Columns.Add(new DataGridViewComboBoxColumn{Name="fund",HeaderText="적용 재원",DataSource=new[]{"공무원","계약제교원","교특회계","학교회계"}});
                grid.Columns.RemoveAt(1);grid.Columns.Insert(1,new DataGridViewComboBoxColumn{Name="field",HeaderText="성명 또는 직종",DataSource=new[]{"성명","직종"}});
                grid.Rows.Clear();foreach(var rule in TestStore202.Read("rules.tsv").Where(x=>x.Length==4))grid.Rows.Add(rule.Cast<object>().ToArray());
                var save=new Button{Text="규칙 저장 후 적용 대상 확인",Dock=DockStyle.Bottom,Height=42};f.Controls.Add(save);save.Click+=(s,e)=>Safe202(()=>{grid.EndEdit();var rules=grid.Rows.Cast<DataGridViewRow>().Where(r=>!r.IsNewRow).Select(r=>r.Cells.Cast<DataGridViewCell>().Select(c=>Convert.ToString(c.Value).Trim()).ToArray()).ToList();if(rules.Any(r=>!new[]{"성명","직종"}.Contains(r[1])||r[2].Length==0||!new[]{"공무원","계약제교원","교특회계","학교회계"}.Contains(r[3])))throw new InvalidOperationException("각 규칙의 기준·일치 문자열·재원을 입력해 주세요.");TestStore202.Write("rules.tsv",rules);if(individualDashboard==null){MessageBox.Show("규칙을 저장했습니다. 대사 후 같은 화면에서 적용 대상을 확인하세요.");return;}ApplyRules202(rules);});f.ShowDialog(this);}
        }
        static bool RuleMatches202(string[] r,IndividualRowData p){return r.Length==4&&!p.ShortTerm&&(String.IsNullOrWhiteSpace(r[0])||Regex.Replace(r[0],"[^0-9]","")==Regex.Replace(p.Site??"","[^0-9]",""))&&(r[1]=="성명"?String.Equals(p.Name,r[2],StringComparison.Ordinal):r[1]=="직종"&&(p.Job??"").IndexOf(r[2],StringComparison.OrdinalIgnoreCase)>=0);}
        void ApplyRules202(List<string[]> rules)
        {
            var changes=new List<Tuple<IndividualRowData,string>>();foreach(var row in individualDashboard.Rows){var rule=rules.FirstOrDefault(r=>RuleMatches202(r,row));if(rule!=null&&row.Fund!=rule[3])changes.Add(Tuple.Create(row,rule[3]));}
            if(changes.Count==0){MessageBox.Show("변경할 대상이 없습니다. 대체근로자는 별도 분류를 유지합니다.");return;}DataGridView g;using(Form f=TableForm202("재원 규칙 적용 미리보기",new[]{"사업장","성명","직종","현재 재원","변경 재원"},changes.Select(x=>new[]{x.Item1.Site,x.Item1.Name,x.Item1.Job,x.Item1.Fund,x.Item2}),out g)){var apply=new Button{Text=changes.Count+"명에게 적용하고 저장",Dock=DockStyle.Bottom,Height=42};f.Controls.Add(apply);apply.Click+=(s,e)=>Safe202(()=>{foreach(var x in changes)reviewFundDrafts[ReviewKey(x.Item1)]=x.Item2;PersistReviewChangesCore(validationResult.Text);f.Close();});f.ShowDialog(this);}
        }
        void Compare202()
        {
            RequireResult202();using(var d=new OpenFileDialog{Title="비교할 지난 대사 결과 선택",Filter="대사 결과 (*.xlsm;*.xlsx)|*.xlsm;*.xlsx"})if(d.ShowDialog(this)==DialogResult.OK)using(var p=new ExcelPackage(new FileInfo(d.FileName))){var ws=p.Workbook.Worksheets["UI개인별데이터"];if(ws==null||ws.Dimension==null)throw new InvalidDataException("개인별 대사 데이터가 없는 파일입니다.");var previous=new Dictionary<string,decimal[]>();for(int r=2;r<=ws.Dimension.End.Row;r++){string key=ws.Cells[r,1].Text;decimal[] v;if(!previous.TryGetValue(key,out v))previous[key]=v=new decimal[6];v[0]++;int[] cols={23,25,27,29,31};for(int c=0;c<cols.Length;c++)v[c+1]+=UiDecimal(ws.Cells[r,cols[c]].Value);}
                var current=individualDashboard.Rows.GroupBy(x=>x.Site).ToDictionary(g=>g.Key,g=>new[]{(decimal)g.Count(),g.Sum(x=>x.SummaryHealthEmployer),g.Sum(x=>x.SummaryLongTermEmployer),g.Sum(x=>x.SummaryPensionEmployer),g.Sum(x=>x.SummaryEmploymentEmployer),g.Sum(x=>x.SummaryIndustrialEmployer)});var result=new List<string[]>();foreach(string site in previous.Keys.Union(current.Keys).OrderBy(x=>x)){decimal[] a,b;if(!previous.TryGetValue(site,out a))a=new decimal[6];if(!current.TryGetValue(site,out b))b=new decimal[6];for(int i=0;i<6;i++)result.Add(new[]{FormatSite(site),new[]{"인원","건강 기관","장기요양 기관","국민 기관","고용 기관","산재 기관"}[i],a[i].ToString("#,##0"),b[i].ToString("#,##0"),(b[i]-a[i]).ToString("+#,##0;-#,##0;0")});}ShowTable202("이전: "+Path.GetFileName(d.FileName)+" / 현재: "+individualDashboard.Year+"년 "+individualDashboard.Month+"월",new[]{"사업장","항목","이전","현재","증감"},result);}
        }
        bool Preflight202(bool teacher,SubmissionInfo info)
        {
            RequireResult202();var missing=new List<string>();if(String.IsNullOrWhiteSpace(info.RecipientCode))missing.Add("수신자기호");if(String.IsNullOrWhiteSpace(info.InstitutionName))missing.Add("기관명");if(String.IsNullOrWhiteSpace(info.ManagerName))missing.Add("담당자명");if(String.IsNullOrWhiteSpace(info.Phone))missing.Add("전화번호");if(teacher&&(String.IsNullOrWhiteSpace(info.BankName)||String.IsNullOrWhiteSpace(info.AccountNumber)))missing.Add("은행명 / 계좌번호");if(missing.Count>0){MessageBox.Show("입력할 항목: "+String.Join(", ",missing),"제출 전 점검");return false;}
            var rows=individualDashboard.Rows.Where(x=>x.Site==info.Site&&(teacher?x.Fund=="계약제교원":x.Fund=="교특회계"||x.ShortTerm)).ToList();if(rows.Count==0){MessageBox.Show("선택 사업장에 제출 대상자가 없습니다.");return false;}
            bool pending=reviewFundDrafts.Count>0||discountDrafts.Count>0;if(pending){MessageBox.Show("재원 또는 감면 수정사항을 먼저 저장해 주세요.","제출 전 점검");return false;}
            // The legacy submission engine uses a person-only key. Reject cross-site collisions instead of mixing them.
            bool ambiguous=individualDashboard.Rows.GroupBy(x=>x.Name+"|"+Regex.Replace(x.Birth??"","[^0-9]","")).Any(g=>g.Select(x=>x.Site).Distinct().Count()>1&&g.Any(x=>x.Site==info.Site));
            if(ambiguous){MessageBox.Show("동일인이 여러 사업장에 있습니다. 이 테스트판에서는 제출 대상 혼합을 막기 위해 생성을 중단합니다. 사업장별로 자료를 분리하여 대사해 주세요.","사업장 중복 확인");return false;}
            return true; // The screen already presents totals; proceed directly to file generation.
        }
        public void SelfTest202(string fixture,string report)
        {
            var log=new List<string>();string work=report+".xlsm";File.Copy(fixture,work,true);validationResult.Text=work;LoadResultIntoUi(work);RequireResult202();
            var target=individualDashboard.Rows.First(x=>!IsReviewCompleted(x));string key=ReviewKey(target);decimal[] before={target.HealthDifference,target.PensionDifference,target.EmploymentDifference,target.IndustrialDifference};
            reviewCheckedKeys.Add(key);reviewNotes202[key]="자동검증: 원자료 대조 완료";PersistReviewChangesCore(work);LoadResultIntoUi(work);target=individualDashboard.Rows.Single(x=>ReviewKey(x)==key);
            if(!IsReviewCompleted(target)||FilteredReviewRows().Any(x=>ReviewKey(x)==key))throw new Exception("확인완료 / 미확인 목록 연동 실패");
            if(!before.SequenceEqual(new[]{target.HealthDifference,target.PensionDifference,target.EmploymentDifference,target.IndustrialDifference}))throw new Exception("확인완료 중 차액 변경");log.Add("PASS: 확인완료 저장·재로딩 및 반환/추징 차액 유지");
            using(var p=new ExcelPackage(new FileInfo(work))){var audit=p.Workbook.Worksheets["처리이력"];if(audit==null||!Enumerable.Range(2,audit.Dimension.End.Row-1).Any(r=>audit.Cells[r,9].Text==key&&audit.Cells[r,7].Text=="자동검증: 원자료 대조 완료"))throw new Exception("사유 기록 실패");}log.Add("PASS: 처리일·사유·처리자 이력 기록");
            reviewCheckedKeys.Remove(key);reviewNotes202[key]="자동검증: 미확인 복원";PersistReviewChangesCore(work);LoadResultIntoUi(work);target=individualDashboard.Rows.Single(x=>ReviewKey(x)==key);if(IsReviewCompleted(target))throw new Exception("복원 실패");log.Add("PASS: 확인완료 되돌리기·재로딩");
            string[] rule={target.Site,"성명",target.Name,"계약제교원"};if(!RuleMatches202(rule,target))throw new Exception("규칙 누락");string savedName=target.Name;target.Name=savedName+"다른사람";if(RuleMatches202(rule,target))throw new Exception("성명 부분 일치 오류");target.Name=savedName;
            string site=target.Site;target.Site="99999999999";if(RuleMatches202(rule,target))throw new Exception("다른 사업장 규칙 오적용");target.Site=site;bool shortTerm=target.ShortTerm;target.ShortTerm=true;if(RuleMatches202(rule,target))throw new Exception("대체근로자 규칙 오적용");target.ShortTerm=shortTerm;
            log.Add("PASS: 재원 규칙 정확한 성명·사업장 제한·대체근로자 제외");
            var eligible=individualDashboard.Rows.First(x=>!x.ShortTerm);string personKey=ReviewKey(eligible),fund=eligible.Fund;reviewFundDrafts[personKey]=fund=="공무원"?"교특회계":"공무원";string expected=reviewFundDrafts[personKey];PersistReviewChangesCore(work);LoadResultIntoUi(work);if(individualDashboard.Rows.Single(x=>ReviewKey(x)==personKey).Fund!=expected)throw new Exception("재원 저장 실패");log.Add("PASS: 공무원 포함 재원 변경 저장 및 전체 화면 갱신");
            string empty=report+".unknown.xlsx";using(var p=new ExcelPackage()){p.Workbook.Worksheets.Add("메모").Cells[1,1].Value="일반 메모";p.SaveAs(new FileInfo(empty));}if(ClassifyInput(empty).Length!=0)throw new Exception("미인식 자료가 급여로 분류됨");log.Add("PASS: 미인식 파일은 급여대장으로 강제 분류하지 않음");
            Remember202(work);if(!TestStore202.Read("recent.tsv").Any(r=>r.Length>1&&r[1]==work))throw new Exception("최근 작업 저장 실패");log.Add("PASS: 최근 작업 저장");
            string beforeHash=Hash202(work);OpenWorkspacePath202(work);if(validationResult.Text==work||Hash202(work)!=beforeHash)throw new Exception("다시 열기 원본 보존 실패");log.Add("PASS: 저장한 작업을 사본으로 열어 원본 보존");
            File.WriteAllLines(report,log,Encoding.UTF8);
        }
        static string Hash202(string path){using(var sha=System.Security.Cryptography.SHA256.Create())using(var s=File.OpenRead(path))return Convert.ToBase64String(sha.ComputeHash(s));}
        public void ReviewFixTest202(string fixture,string report)
        {
            var log=new List<string>();string path=report+".xlsm";File.Copy(fixture,path,true);validationResult.Text=path;LoadResultIntoUi(path);
            if(pages.ContainsKey("작업 관리")||navigationButtons.ContainsKey("작업 관리"))throw new Exception("작업 관리 메뉴 잔존");log.Add("PASS: 작업 관리 메뉴 제거");
            var target=FilteredReviewRows().First(x=>!IsReviewCompleted(x));string key=ReviewKey(target);decimal[] raw={target.HealthNotice,target.HealthPayroll,target.PensionNotice,target.PensionPayroll,target.HealthDifference,target.PensionDifference,target.EmploymentDifference,target.SummaryHealthPersonal,target.SummaryHealthEmployer};
            decimal total=summaryDashboard.Sites.Values.Sum(s=>s.Rows.Sum(x=>x.InsuranceTotal));reviewSelections.Add(key);MarkSelectedReviewsChecked();
            if(!IsReviewCompleted(target)||!FilteredReviewRows().Any(x=>ReviewKey(x)==key))throw new Exception("확인완료가 명단에서 사라짐");if(HasCollectionDirection(target)||HasRefundDirection(target)||target.Status!="정상")throw new Exception("확인완료 경고 판정 오류");
            if(total!=summaryDashboard.Sites.Values.Sum(s=>s.Rows.Sum(x=>x.InsuranceTotal)))throw new Exception("실제 부과금액 변경");
            if(!raw.SequenceEqual(new[]{target.HealthNotice,target.HealthPayroll,target.PensionNotice,target.PensionPayroll,target.HealthDifference,target.PensionDifference,target.EmploymentDifference,target.SummaryHealthPersonal,target.SummaryHealthEmployer}))throw new Exception("원금액 변경");
            foreach(var group in individualDashboard.Rows.GroupBy(x=>new{x.Site,Fund=x.Fund=="분류필요"?"기타":x.Fund})){var summary=summaryDashboard.Sites[group.Key.Site].Rows.Single(x=>x.Fund==group.Key.Fund);if(summary.HealthDifference!=group.Where(x=>!IsReviewCompleted(x)).Sum(x=>x.SummaryHealthDifference)||summary.EmploymentDifference!=group.Where(x=>!IsReviewCompleted(x)).Sum(x=>x.EmploymentDifference))throw new Exception("총괄표 확인완료 제외 오류");}
            log.Add("PASS: 선택창 없이 확인완료 저장 / 명단 유지 / 총괄 경고 제외 / 원금액 보존");
            LoadResultIntoUi(path);target=individualDashboard.Rows.Single(x=>ReviewKey(x)==key);if(!IsReviewCompleted(target)||!individualDashboard.Rows.Where(IsListedReview202).Any(x=>ReviewKey(x)==key))throw new Exception("재실행 상태 유지 실패");log.Add("PASS: 다시 열어도 확인완료 상태와 명단 유지");
            int idx=reviewSiteKeys.IndexOf(target.Site);reviewSiteSelector.SelectedIndex=idx;reviewSelections.Add(key);MarkSelectedReviewsChecked();if(IsReviewCompleted(target)||!FilteredReviewRows().Any(x=>ReviewKey(x)==key))throw new Exception("일괄 미확인 복원 오류");log.Add("PASS: 일괄 변경으로 미확인 복원");
            reviewFundDrafts[key]="공무원";SaveReviewChanges();LoadResultIntoUi(path);target=individualDashboard.Rows.Single(x=>ReviewKey(x)==key);if(target.Fund!="공무원"||!IsListedReview202(target))throw new Exception("공무원 선택 / 명단 유지 오류");log.Add("PASS: 저장만 눌러 공무원 재원 반영 / 명단 유지");
            var model=individualDashboard;individualDashboard=new IndividualDashboardModel();individualDashboard.Rows.Add(new IndividualRowData{Site="PUBLIC",Fund="공무원"});individualDashboard.Rows.Add(new IndividualRowData{Site="OTHER",Fund="교특회계"});individualDashboard.Rows.Add(new IndividualRowData{Site="OTHER",Fund="학교회계"});
            if(!ReviewFundsForSite202("PUBLIC").SequenceEqual(new[]{"공무원"})||!ReviewFundsForSite202("OTHER").SequenceEqual(new[]{"교특회계","학교회계"}))throw new Exception("사업장별 재원 선택 오류");individualDashboard=model;log.Add("PASS: 공무원 사업장 / 교특·학회 사업장 재원 목록 분리");
            var noPayroll=new IndividualRowData{Site=target.Site,Name="검증용_급여없음",Birth="990101",Fund="공무원",HealthNotice=10000m,HealthPayroll=0m,SummaryHealthPersonal=10000m,SummaryHealthEmployer=0m,SummaryHealthDifference=10000m,HasSummaryBreakdown=true};
            individualDashboard.Rows.Add(noPayroll);NormalizeIndividualStatuses();RebuildSummaryDashboardFromIndividuals();if(!HasCollectionDirection(noPayroll))throw new Exception("급여 없는 검증대상 준비 실패");
            reviewCheckedKeys.Add(ReviewKey(noPayroll));NormalizeIndividualStatuses();RebuildSummaryDashboardFromIndividuals();if(noPayroll.Status!="정상"||HasCollectionDirection(noPayroll)||noPayroll.HealthDifference!=10000m)throw new Exception("급여 없는 확인완료 판정 오류");
            log.Add("PASS: 급여 0원·고지 10,000원 확인완료 시 정상 처리 및 실제 차액 보존");
            individualDashboard.Rows.Remove(noPayroll);reviewCheckedKeys.Remove(ReviewKey(noPayroll));NormalizeIndividualStatuses();RebuildSummaryDashboardFromIndividuals();
            File.WriteAllLines(report,log,Encoding.UTF8);
        }
    }
}
