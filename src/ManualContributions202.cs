using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using OfficeOpenXml;

namespace InsurancePayrollValidator
{
    partial class MainForm
    {
        public void SubmissionFlowTest202(string fixture,string output)
        {
            Directory.CreateDirectory(output);validationResult.Text=fixture;LoadResultIntoUi(fixture);
            foreach(bool teacher in new[]{false,true}){
                var person=individualDashboard.Rows.First(x=>teacher?x.Fund=="계약제교원":x.Fund=="교특회계");
                var info=new SubmissionInfo{RecipientCode="테스트기관",InstitutionName="검증학교",ManagerName="테스트",Phone="000-0000",BankName="테스트은행",AccountNumber="000",Round="1차",IndustrialRate="0.008",Site=person.Site};
                if(!Preflight202(teacher,info))throw new Exception("사전 검증 통과 실패");
                string path=SubmissionGenerator.Create(fixture,output,teacher,info);
                if(!File.Exists(path)||new FileInfo(path).Length==0||Path.GetExtension(path)!=".xlsx")throw new Exception("엑셀 파일 생성 실패");
                using(var p=new ExcelPackage(new FileInfo(path))){if(p.Workbook.Worksheets.Count<1)throw new Exception("시트 없음");}
                File.AppendAllText(Path.Combine(output,"검증기록.txt"),"PASS: "+(teacher?"계약제교원":"교육공무직")+" 사전검증 팝업 없이 통과 / 실제 xlsx 생성 및 재열기: "+Path.GetFileName(path)+Environment.NewLine,Encoding.UTF8);
            }
        }
        bool HasActualRefund202(IndividualRowData row){return row!=null&&row.Fund!="분류필요"&&(row.HealthDifference<-.5m||row.PensionDifference<-.5m||row.EmploymentDifference<-.5m||row.IndustrialDifference<-.5m);}
        bool HasActualCollection202(IndividualRowData row){return row!=null&&row.Fund!="분류필요"&&(row.HealthDifference>.5m||row.PensionDifference>.5m||row.EmploymentDifference>.5m||row.IndustrialDifference>.5m);}
        static decimal[] ManualValues202(IndividualRowData r)
        {
            decimal care=r.SummaryLongTermPersonal-r.SummaryLongTermDifference;
            return new[]{r.HealthPayroll-care,care,r.PensionPayroll,r.EmploymentPayroll,r.SummaryHealthEmployer,r.SummaryLongTermEmployer,r.SummaryPensionEmployer,r.SummaryEmploymentEmployer};
        }
        void ShowContributionEditor202(IndividualRowData row,Point anchor)
        {
            Safe202(delegate {
                RequireResult202();
                if(!row.HasSummaryBreakdown)throw new InvalidOperationException("보험별 상세 데이터가 없는 이전 결과입니다. 먼저 새로 대사해 주세요.");
                string path=validationResult.Text,key=StablePersonKey(row),hash=Hash202(path);
                if(reviewBubble!=null&&!reviewBubble.IsDisposed)reviewBubble.Close();
                var popup=new ReviewEditBubble202(row,ReviewReasonText(row),IsReviewCompleted(row),values=>{
                    if(validationResult.Text!=path||Hash202(path)!=hash)throw new InvalidOperationException("결과 자료가 변경되었습니다. 창을 닫고 다시 열어 주세요.");
                    SaveManual202(path,key,values);
                });
                reviewBubble=popup;
                {
                    Rectangle area=Screen.FromPoint(anchor).WorkingArea;
                    popup.Location=new Point(Math.Max(area.Left,Math.Min(Right-8,area.Right-popup.Width)),Math.Max(area.Top,Math.Min(anchor.Y-170,area.Bottom-popup.Height)));
                    popup.Show(this);
                }
            });
        }
        void SaveManual202(string path,string key,decimal[] values)
        {
            if(values==null||values.Length!=8||values.Any(v=>v!=Decimal.Truncate(v)||Math.Abs(v)>1000000000000m))throw new ArgumentException("금액은 원 단위 정수로 입력해 주세요.");
            var matches=individualDashboard.Rows.Where(x=>StablePersonKey(x)==key).ToList();
            if(matches.Count!=1)throw new InvalidOperationException("동일 사업장의 성명·생년월일이 중복되어 자동 보정할 수 없습니다.");
            var row=matches[0];decimal[] before=ManualValues202(row);
            if(before.SequenceEqual(values))return;
            string temp=path+"."+Guid.NewGuid().ToString("N")+".xlsm";
            try {
                File.Copy(path,temp);
                row.HealthPayroll=values[0]+values[1];row.SummaryLongTermDifference=row.SummaryLongTermPersonal-values[1];
                row.PensionPayroll=values[2];row.EmploymentPayroll=values[3];
                row.SummaryHealthEmployer=values[4];row.SummaryLongTermEmployer=values[5];row.SummaryPensionEmployer=values[6];row.SummaryEmploymentEmployer=values[7];
                reviewCheckedKeys.Remove(key);retainedReviewKeys202.Add(key);row.ReviewReason="부담금 수기 보정";
                NormalizeIndividualStatuses();RebuildSummaryDashboardFromIndividuals();
                PersistReviewStateBase(temp);
                using(var p=new ExcelPackage(new FileInfo(temp))) {
                    var ws=p.Workbook.Worksheets["UI개인별데이터"];
                    int[] rows=Enumerable.Range(2,ws.Dimension.End.Row-1).Where(r=>String.Join("|",new[]{ws.Cells[r,1].Text,ws.Cells[r,3].Text,System.Text.RegularExpressions.Regex.Replace(ws.Cells[r,4].Text,"[^0-9]","")})==key).ToArray();
                    if(rows.Length!=1)throw new InvalidOperationException("저장 대상의 고유 정보를 확인할 수 없습니다.");
                    int n=rows[0];int[] cols={8,11,14,23,25,27,29,32,33};
                    decimal[] saved={row.HealthPayroll,row.PensionPayroll,row.EmploymentPayroll,values[4],values[5],values[6],values[7],row.SummaryHealthDifference,row.SummaryLongTermDifference};
                    for(int c=0;c<cols.Length;c++)ws.Cells[n,cols[c]].Value=saved[c];
                    var log=p.Workbook.Worksheets["수기보정이력"]??p.Workbook.Worksheets.Add("수기보정이력");
                    string[] headers={"변경일시","사업장","성명","항목","변경 전","변경 후","대상키"};
                    string[] names={"급여 건강보험","급여 장기요양","급여 국민연금","급여 고용보험","기관 건강보험","기관 장기요양","기관 국민연금","기관 고용보험"};
                    for(int c=0;c<headers.Length;c++)log.Cells[1,c+1].Value=headers[c];
                    int nr=log.Dimension.End.Row+1;
                    for(int c=0;c<8;c++)if(before[c]!=values[c]){
                        object[] data={DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),row.Site,row.Name,names[c],before[c],values[c],key};
                        for(int j=0;j<data.Length;j++)log.Cells[nr,j+1].Value=data[j];nr++;
                    }
                    log.Column(7).Hidden=true;log.Column(1).Width=23;log.Column(2).Width=22;log.Column(3).Width=16;log.Column(4).Width=23;log.Column(5).Width=18;log.Column(6).Width=18;
                    log.Cells[2,5,Math.Max(2,nr-1),6].Style.Numberformat.Format="#,##0;[Red]-#,##0";
                    WriteRetainedReview202(p);p.Save();
                }
                File.Replace(temp,path,null);
            } catch {LoadResultIntoUi(path);throw;}
            finally {if(File.Exists(temp))File.Delete(temp);}
            LoadResultIntoUi(path);InitializeReviewFilters();RefreshLinkedResultViews();
        }
        public void ManualTest202(string fixture,string report)
        {
            string path=report+".xlsm";File.Copy(fixture,path,true);validationResult.Text=path;LoadResultIntoUi(path);
            var row=individualDashboard.Rows.First(x=>x.HasSummaryBreakdown&&x.Fund!="분류필요"&&x.HealthNotice>0);
            string key=StablePersonKey(row);decimal[] original=ManualValues202(row),v=ManualValues202(row);
            v[0]=row.SummaryHealthPersonal;v[1]=row.SummaryLongTermPersonal;v[2]=row.PensionNotice;v[3]=row.EmploymentNotice;v[4]+=123;
            string other=String.Join(";",individualDashboard.Rows.Where(x=>StablePersonKey(x)!=key).Select(x=>String.Join(",",ManualValues202(x))));
            SaveManual202(path,key,v);row=individualDashboard.Rows.Single(x=>StablePersonKey(x)==key);
            if(!ManualValues202(row).SequenceEqual(v)||row.HealthDifference!=0||row.SummaryLongTermDifference!=0||row.PensionDifference!=0||row.EmploymentDifference!=0)throw new Exception("보정 저장/재로딩 실패");
            if(other!=String.Join(";",individualDashboard.Rows.Where(x=>StablePersonKey(x)!=key).Select(x=>String.Join(",",ManualValues202(x)))))throw new Exception("다른 대상자 변경");
            if(!IsListedReview202(row))throw new Exception("목록 유지 실패");
            var summary=summaryDashboard.Sites[row.Site].Rows.First(x=>x.Fund==row.Fund);
            if(summary.HealthEmployer!=individualDashboard.Rows.Where(x=>x.Site==row.Site&&x.Fund==row.Fund).Sum(x=>x.SummaryHealthEmployer))throw new Exception("총괄 기관부담금 반영 실패");
            PersistReviewChangesCore(path);LoadResultIntoUi(path);row=individualDashboard.Rows.Single(x=>StablePersonKey(x)==key);
            if(!ManualValues202(row).SequenceEqual(v))throw new Exception("일반 저장 시 보정값 손실");
            v[0]-=100;SaveManual202(path,key,v);row=individualDashboard.Rows.Single(x=>StablePersonKey(x)==key);
            if(row.HealthDifference!=100||row.SummaryHealthDifference!=100||!HasCollectionDirection(row))throw new Exception("차액 재계산 실패");
            string hash=Hash202(path);try{SaveManual202(path,key,new decimal[1]);throw new Exception("invalid accepted");}catch(ArgumentException){}
            if(Hash202(path)!=hash)throw new Exception("잘못된 입력 저장");
            using(var source=new ExcelPackage(new FileInfo(fixture)))using(var result=new ExcelPackage(new FileInfo(path))){
                foreach(var sheet in source.Workbook.Worksheets.Where(x=>x.Name.StartsWith("원본"))) {
                    var saved=result.Workbook.Worksheets[sheet.Name];
                    if(saved==null||sheet.WorksheetXml.OuterXml!=saved.WorksheetXml.OuterXml)throw new Exception("원본 시트 변경: "+sheet.Name);
                }
                if(result.Workbook.Worksheets["수기보정이력"]==null)throw new Exception("이력 누락");
            }
            decimal[] applied=null;
            using(var form=new ReviewEditBubble202(row,"급여대장 공제액 확인",false,x=>applied=x)){
                form.Show();Application.DoEvents();
                ((Button)form.Controls.Find("ApplyNotice",true).Single()).PerformClick();
                if(applied!=null||Hash202(path)!=hash)throw new Exception("고지금액 적용 시 자동 저장 발생");
                using(var bitmap=new Bitmap(form.Width,form.Height)){form.DrawToBitmap(bitmap,new Rectangle(Point.Empty,bitmap.Size));bitmap.Save(report+".png");}
                ((Button)form.Controls.Find("SaveCorrection",true).Single()).PerformClick();
            }
            if(applied==null||!applied.Take(4).SequenceEqual(new[]{row.SummaryHealthPersonal,row.SummaryLongTermPersonal,row.PensionNotice,row.EmploymentNotice})||!applied.Skip(4).SequenceEqual(ManualValues202(row).Skip(4)))throw new Exception("고지금액 적용 값 오류");
            SaveManual202(path,key,applied);
            SaveManual202(path,key,original);
            foreach(string mode in new[]{"반환","추징"}){
                row=individualDashboard.Rows.Single(x=>StablePersonKey(x)==key);
                var amounts=ManualValues202(row);amounts[0]=row.SummaryHealthPersonal+(mode=="반환"?100:-100);amounts[1]=row.SummaryLongTermPersonal;amounts[2]=row.PensionNotice;amounts[3]=row.EmploymentNotice;
                SaveManual202(path,key,amounts);reviewCheckedKeys.Add(key);PersistReviewChangesCore(path);LoadResultIntoUi(path);
                row=individualDashboard.Rows.Single(x=>StablePersonKey(x)==key);
                if(!IsReviewCompleted(row)||row.Status!="정상")throw new Exception("확인완료 상태 손실");
                adjustmentSiteSelector.SelectedIndex=adjustmentSiteKeys.IndexOf(row.Site);
                adjustmentFundSelector.SelectedIndex=0;SelectAdjustmentMode(mode);
                if(!FilteredAdjustmentRows().Any(x=>StablePersonKey(x)==key))throw new Exception(mode+" 확인완료 대상 목록 누락");
                var exportRows=RowsForAdjustmentExport();
                if(!exportRows.Any(x=>StablePersonKey(x)==key))throw new Exception(mode+" 출력 대상 누락");
                string exported=report+"."+mode+".xlsx";
                AdjustmentReportGenerator.CreateExcel(exported,exportRows,mode,individualDashboard.Year,individualDashboard.Month,FormatSite(row.Site));
                using(var p=new ExcelPackage(new FileInfo(exported))){if(!p.Workbook.Worksheets.Any(ws=>ws.Dimension!=null&&ws.Cells[ws.Dimension.Address].Any(c=>c.Text==row.Name)))throw new Exception(mode+" 출력 엑셀 명단 누락");}
            }
            SaveManual202(path,key,original);
            string beforeClose=Hash202(path);Show();Application.DoEvents();Close();if(Hash202(path)!=beforeClose)throw new Exception("종료 시 엑셀 변경");
            File.WriteAllText(report,"PASS: 확인완료 후 반환/추징 양쪽 목록 유지, 재로딩 후 유지, 실제 Excel 출력 명단 포함, 확인완료 정상 상태 유지. 고지금액 적용, 저장/재로딩, 원본 보존, 종료 시 엑셀 변경 없음.",Encoding.UTF8);
        }
    }
    sealed class ReviewEditBubble202:ReviewDetailBubble
    {
        public ReviewEditBubble202(IndividualRowData row,string reason,bool done,Action<decimal[]> save):base(row,row.Fund,reason,done)
        {
            Height=675;TopMost=false;StartPosition=FormStartPosition.Manual;
            var panel=new Panel{Location=new Point(22,306),Size=new Size(306,354),BackColor=UiTheme.Card};
            Controls.Add(panel);
            panel.Controls.Add(new Label{Text="급여대장 금액 보정",Location=new Point(8,0),Size=new Size(164,25),ForeColor=UiTheme.Accent,Font=new Font("맑은 고딕",8.3F,FontStyle.Bold)});
            var applyNotice=new NoticeApplyButton202{Name="ApplyNotice",Text="고지금액 적용",Location=new Point(176,-1),Size=new Size(124,28)};
            panel.Controls.Add(applyNotice);
            var grid=new TableLayoutPanel{Location=new Point(8,29),Size=new Size(292,180),ColumnCount=4,RowCount=6,CellBorderStyle=TableLayoutPanelCellBorderStyle.Single};
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute,56));for(int i=0;i<3;i++)grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,33.333F));
            for(int i=0;i<6;i++)grid.RowStyles.Add(new RowStyle(SizeType.Percent,16.667F));
            string[] headers={"보험","급여대장","고지금액","차액"};
            for(int i=0;i<4;i++)grid.Controls.Add(new Label{Text=headers[i],Dock=DockStyle.Fill,TextAlign=ContentAlignment.MiddleCenter,ForeColor=UiTheme.Muted,BackColor=UiTheme.Header},i,0);
            decimal care=row.SummaryLongTermPersonal-row.SummaryLongTermDifference;
            decimal[] initial={row.HealthPayroll-care,care,row.PensionPayroll,row.EmploymentPayroll,row.SummaryHealthEmployer,row.SummaryLongTermEmployer,row.SummaryPensionEmployer,row.SummaryEmploymentEmployer};
            decimal[] notices={row.SummaryHealthPersonal,row.SummaryLongTermPersonal,row.PensionNotice,row.EmploymentNotice};
            var inputs=new TextBox[4];var diffs=new Label[4];var totals=new Label[3];
            string[] names={"건강","요양","국민","고용","합계"};
            for(int r=0;r<5;r++)grid.Controls.Add(new Label{Text=names[r],Dock=DockStyle.Fill,TextAlign=ContentAlignment.MiddleCenter,ForeColor=UiTheme.Text},0,r+1);
            for(int c=0;c<3;c++){totals[c]=new Label{Dock=DockStyle.Fill,TextAlign=ContentAlignment.MiddleCenter,ForeColor=UiTheme.Text,BackColor=UiTheme.Surface};grid.Controls.Add(totals[c],c+1,5);}
            Action refresh=()=>{
                decimal sum=0;bool valid=true;
                for(int i=0;i<4;i++){decimal amount;if(ContributionEditor202.TryAmount(inputs[i].Text,out amount)){sum+=amount;decimal d=notices[i]-amount;diffs[i].Text=d.ToString("#,##0");diffs[i].ForeColor=d==0?Color.ForestGreen:d>0?Color.Firebrick:Color.RoyalBlue;}else{valid=false;diffs[i].Text="확인";}}
                totals[0].Text=valid?sum.ToString("#,##0"):"확인";totals[1].Text=notices.Sum().ToString("#,##0");totals[2].Text=valid?(notices.Sum()-sum).ToString("#,##0"):"확인";
            };
            for(int i=0;i<4;i++){
                inputs[i]=new TextBox{Text=initial[i].ToString("#,##0"),Dock=DockStyle.Fill,Margin=new Padding(2,4,2,2),TextAlign=HorizontalAlignment.Right,Font=new Font("맑은 고딕",7.5F),BackColor=UiTheme.Surface,ForeColor=UiTheme.Text};
                diffs[i]=new Label{Dock=DockStyle.Fill,TextAlign=ContentAlignment.MiddleCenter};
                grid.Controls.Add(inputs[i],1,i+1);grid.Controls.Add(new Label{Text=notices[i].ToString("#,##0"),Dock=DockStyle.Fill,TextAlign=ContentAlignment.MiddleCenter,ForeColor=UiTheme.Text},2,i+1);grid.Controls.Add(diffs[i],3,i+1);
                inputs[i].TextChanged+=(s,e)=>refresh();
            }
            applyNotice.Click+=(s,e)=>{for(int i=0;i<4;i++)inputs[i].Text=notices[i].ToString("#,##0");refresh();};
            refresh();panel.Controls.Add(grid);
            panel.Controls.Add(new Label{Text="기관부담금 합계  "+(initial.Skip(4).Sum()+row.SummaryIndustrialEmployer).ToString("#,##0")+"원\r\n산재 기관부담금  "+row.SummaryIndustrialEmployer.ToString("#,##0")+"원",Location=new Point(8,217),Size=new Size(292,39),ForeColor=UiTheme.Text});
            panel.Controls.Add(new Label{Text="급여대장 칸을 수정한 후 저장·반영하세요.\r\n고지금액·기관부담금과 원본 자료는 유지됩니다.",Location=new Point(8,261),Size=new Size(292,40),ForeColor=UiTheme.Muted,Font=new Font("맑은 고딕",7.5F)});
            var ok=new Button{Name="SaveCorrection",Text="저장 · 반영",Location=new Point(104,309),Size=new Size(116,32),BackColor=UiTheme.Accent,ForeColor=Color.White,FlatStyle=FlatStyle.Flat};
            var cancel=new Button{Text="닫기",Location=new Point(228,309),Size=new Size(72,32),BackColor=UiTheme.Surface,ForeColor=UiTheme.Text,FlatStyle=FlatStyle.Flat};
            cancel.Click+=(s,e)=>Close();
            ok.Click+=(s,e)=>{
                decimal[] next=(decimal[])initial.Clone();for(int i=0;i<4;i++)if(!ContributionEditor202.TryAmount(inputs[i].Text,out next[i])){MessageBox.Show(this,"급여대장 금액을 원 단위 정수로 입력해 주세요.");inputs[i].Focus();return;}
                try{if(save!=null)save(next);Close();}catch(Exception ex){MessageBox.Show(this,"저장하지 못했습니다.\r\n"+ex.Message,"부담금 보정",MessageBoxButtons.OK,MessageBoxIcon.Warning);}
            };
            panel.Controls.Add(ok);panel.Controls.Add(cancel);AcceptButton=ok;CancelButton=cancel;
        }
    }
    sealed class NoticeApplyButton202:Button
    {
        bool hovered;
        public NoticeApplyButton202(){FlatStyle=FlatStyle.Flat;FlatAppearance.BorderSize=0;Cursor=Cursors.Hand;Font=new Font("맑은 고딕",8F,FontStyle.Bold);SetStyle(ControlStyles.UserPaint|ControlStyles.OptimizedDoubleBuffer|ControlStyles.AllPaintingInWmPaint,true);}
        protected override void OnMouseEnter(EventArgs e){hovered=true;Invalidate();base.OnMouseEnter(e);}
        protected override void OnMouseLeave(EventArgs e){hovered=false;Invalidate();base.OnMouseLeave(e);}
        protected override void OnPaint(PaintEventArgs e){
            e.Graphics.SmoothingMode=System.Drawing.Drawing2D.SmoothingMode.AntiAlias;e.Graphics.Clear(Parent==null?UiTheme.Card:Parent.BackColor);
            using(var shadow=UiDrawing.Rounded(new RectangleF(2,3,Width-4,Height-4),8))using(var brush=new SolidBrush(Color.FromArgb(45,UiTheme.Accent)))e.Graphics.FillPath(brush,shadow);
            using(var path=UiDrawing.Rounded(new RectangleF(1,1,Width-4,Height-5),8))using(var brush=new SolidBrush(UiTheme.Dark?UiTheme.Surface:hovered?Color.FromArgb(216,229,255):Color.FromArgb(236,242,255)))using(var pen=new Pen(UiTheme.Accent,1.1F)){e.Graphics.FillPath(brush,path);e.Graphics.DrawPath(pen,path);}
            UiDrawing.Text(e.Graphics,Text,Font,UiTheme.Accent,new Rectangle(1,0,Width-4,Height-4),ContentAlignment.MiddleCenter);
            if(Focused)ControlPaint.DrawFocusRectangle(e.Graphics,new Rectangle(5,5,Width-12,Height-12));
        }
    }
    sealed class ContributionEditor202:Form
    {
        public ContributionEditor202(IndividualRowData row,Action<decimal[]> save)
        {
            AutoScaleMode=AutoScaleMode.None;Font=new Font("맑은 고딕",10);ClientSize=new Size(760,390);FormBorderStyle=FormBorderStyle.FixedToolWindow;
            Text="부담금 수기 보정";StartPosition=FormStartPosition.Manual;BackColor=Color.White;ShowInTaskbar=false;
            Controls.Add(new Label{Text=row.Name+" · 부담금 보정",Location=new Point(24,18),Size=new Size(710,30),Font=new Font(Font,FontStyle.Bold)});
            Controls.Add(new Label{Text="사업장 "+row.Site+"  /  "+row.Fund+"     (단위: 원)",Location=new Point(24,50),Size=new Size(710,25),ForeColor=Color.DimGray});
            string[] headers={"","건강보험","장기요양","국민연금","고용보험"};
            var grid=new TableLayoutPanel{Location=new Point(24,90),Size=new Size(710,185),ColumnCount=5,RowCount=5,CellBorderStyle=TableLayoutPanelCellBorderStyle.Single};
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute,145));for(int i=0;i<4;i++)grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,25));
            for(int i=0;i<5;i++)grid.RowStyles.Add(new RowStyle(SizeType.Percent,20));
            for(int i=0;i<5;i++)grid.Controls.Add(new Label{Text=headers[i],Dock=DockStyle.Fill,TextAlign=ContentAlignment.MiddleCenter,BackColor=Color.FromArgb(232,240,253)},i,0);
            string[] labels={"급여대장 공제액","개인 고지금액","차액 (고지 - 급여)","기관부담금"};
            for(int i=0;i<4;i++)grid.Controls.Add(new Label{Text=labels[i],Dock=DockStyle.Fill,TextAlign=ContentAlignment.MiddleCenter},0,i+1);
            decimal care=row.SummaryLongTermPersonal-row.SummaryLongTermDifference;
            decimal[] values={row.HealthPayroll-care,care,row.PensionPayroll,row.EmploymentPayroll,row.SummaryHealthEmployer,row.SummaryLongTermEmployer,row.SummaryPensionEmployer,row.SummaryEmploymentEmployer};
            decimal[] notice={row.SummaryHealthPersonal,row.SummaryLongTermPersonal,row.PensionNotice,row.EmploymentNotice};
            var boxes=new TextBox[8];
            for(int c=0;c<4;c++){
                int index=c;
                boxes[c]=new TextBox{Text=values[c].ToString("#,##0"),Dock=DockStyle.Fill,TextAlign=HorizontalAlignment.Right,Margin=new Padding(6)};
                boxes[c+4]=new TextBox{Text=values[c+4].ToString("#,##0"),Dock=DockStyle.Fill,TextAlign=HorizontalAlignment.Right,Margin=new Padding(6)};
                var diff=new Label{Dock=DockStyle.Fill,TextAlign=ContentAlignment.MiddleRight,Padding=new Padding(0,0,6,0)};
                Action refresh=()=>{decimal v;if(TryAmount(boxes[index].Text,out v)){decimal d=notice[index]-v;diff.Text=d.ToString("#,##0");diff.ForeColor=d==0?Color.ForestGreen:Color.Firebrick;}else diff.Text="입력 확인";};
                boxes[c].TextChanged+=(s,e)=>refresh();refresh();
                grid.Controls.Add(boxes[c],c+1,1);grid.Controls.Add(new Label{Text=notice[c].ToString("#,##0"),Dock=DockStyle.Fill,TextAlign=ContentAlignment.MiddleRight,Padding=new Padding(0,0,6,0)},c+1,2);
                grid.Controls.Add(diff,c+1,3);grid.Controls.Add(boxes[c+4],c+1,4);
            }
            Controls.Add(grid);
            Controls.Add(new Label{Text="원본 급여대장은 변경하지 않습니다. 고지금액은 비교용입니다.\r\n기관부담금은 현재 적용액(감면 반영 후)입니다. 보정 후 차액을 다시 판정합니다.",Location=new Point(24,286),Size=new Size(710,45),ForeColor=Color.DimGray,Font=new Font("맑은 고딕",9)});
            var ok=new Button{Text="저장 · 반영",Location=new Point(484,342),Size=new Size(130,33),BackColor=Color.FromArgb(43,111,220),ForeColor=Color.White,FlatStyle=FlatStyle.Flat};
            var cancel=new Button{Text="취소",Location=new Point(624,342),Size=new Size(110,33),DialogResult=DialogResult.Cancel};
            ok.Click+=(s,e)=>{var next=new decimal[8];for(int i=0;i<8;i++)if(!TryAmount(boxes[i].Text,out next[i])){MessageBox.Show(this,"금액을 원 단위 정수로 입력해 주세요. (0 또는 음수도 가능)");boxes[i].Focus();return;}try{if(save!=null)save(next);DialogResult=DialogResult.OK;Close();}catch(Exception ex){MessageBox.Show(this,"저장하지 못했습니다.\r\n"+ex.Message,"부담금 보정",MessageBoxButtons.OK,MessageBoxIcon.Warning);}};
            Controls.Add(ok);Controls.Add(cancel);AcceptButton=ok;CancelButton=cancel;
        }
        internal static bool TryAmount(string text,out decimal value){return Decimal.TryParse(text,NumberStyles.AllowLeadingSign|NumberStyles.AllowThousands|NumberStyles.AllowLeadingWhite|NumberStyles.AllowTrailingWhite,CultureInfo.InvariantCulture,out value)&&value==Decimal.Truncate(value)&&Math.Abs(value)<=1000000000000m;}
    }
}
