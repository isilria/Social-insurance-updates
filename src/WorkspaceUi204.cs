using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using OfficeOpenXml;

namespace InsurancePayrollValidator
{
    partial class MainForm
    {
        SubmissionInfo WorkbookInfo204(){return new SubmissionInfo{RecipientCode=recipientCode.Text,InstitutionName=institutionName.Text,Round=submissionRound.Text,IndustrialRate=industrialRate.Text,BankName=bankName.Text,AccountNumber=accountNumber.Text,ManagerName=managerName.Text,Phone=phone.Text};}
        void ExportWorkspace204(string path){WorkspaceWorkbook204.Export(validationResult.Text,path,WorkbookInfo204());}
        void OpenWorkbook204(string path)
        {
            string target=NewTemporaryResultPath();bool loaded=false;
            try{
                SubmissionInfo info=WorkspaceWorkbook204.Import(path,target);
                // Parse and validate first. The currently open work is not changed by a failed import.
                temporaryResultPath=target;validationResult.Text=target;externalResult202=false;reviewNotes202.Clear();reviewFundDrafts.Clear();discountDrafts.Clear();
                LoadResultIntoUi(target);NormalizeIndividualStatuses();RebuildSummaryDashboardFromIndividuals();PersistDiscountWorkbook(target);
                lastSavedHash202=Hash202(target);recipientCode.Text=info.RecipientCode;institutionName.Text=info.InstitutionName;submissionRound.Text=info.Round;industrialRate.Text=info.IndustrialRate;bankName.Text=info.BankName;accountNumber.Text=info.AccountNumber;managerName.Text=info.ManagerName;phone.Text=info.Phone;
                Remember202(path);ShowPage("총괄표");loaded=true;
            }finally{if(!loaded&&File.Exists(target))File.Delete(target);}
        }
        public void ExportWorkspaceForTest204(string native,string destination){validationResult.Text=native;LoadResultIntoUi(native);ExportWorkspace204(destination);}
        static void Check204(bool ok,string message,List<string> log){if(!ok)throw new Exception(message);log.Add("PASS: "+message);}
        public void WorkbookTest204(string native,string directory)
        {
            Directory.CreateDirectory(directory);var log=new List<string>();string originalHash=Hash202(native);
            string input=Path.Combine(directory,"original.xlsm");File.Copy(native,input,true);validationResult.Text=input;LoadResultIntoUi(input);
            var before=individualDashboard.Rows.Select(x=>new[]{x.Site,x.Fund,x.Name,x.Birth,x.Job,x.ShortTerm.ToString()}).ToArray();
            decimal[] totals=WorkbookTotals204();
            string exported=Path.Combine(directory,"작업보관.xlsx");ExportWorkspace204(exported);
            Check204(WorkspaceWorkbook204.IsArchive(exported),"독립 엑셀 작업 파일 생성",log);
            using(var p=new ExcelPackage(new FileInfo(exported))){
                Check204(p.Workbook.Worksheets.Count>=15,"총괄표·메뉴별 탭과 복원 데이터 보관",log);
                Check204(p.Workbook.Worksheets.First().Name=="총괄표","총괄표를 첫 탭으로 배치",log);
                Check204(p.Workbook.Worksheets["대상자정보"].Cells[8,2].Text!="P001","시안 대상자ID를 실제 작업 키로 교체",log);
                Check204(WorkspaceWorkbook204.Number(p.Workbook.Worksheets["총괄표"].Cells[20,3].Value)==totals[0],"엑셀 개인부담 총액 = 화면 총액",log);
                Check204(WorkspaceWorkbook204.Number(p.Workbook.Worksheets["총괄표"].Cells[53,7].Value)==totals[0]+totals[1],"엑셀 감면 후 납입 총액 = 화면 총액",log);
                foreach(var ws in p.Workbook.Worksheets.Where(w=>w.Dimension!=null&&!w.Name.StartsWith("_SI")))foreach(var cell in ws.Cells[ws.Dimension.Address])if(cell.Value is ExcelErrorValue)throw new Exception("수식 오류: "+ws.Name+"!"+cell.Address+" "+cell.Value);
                log.Add("PASS: 모든 표시 탭 수식 오류 검사");
            }
            OpenWorkbook204(exported);Check204(totals.SequenceEqual(WorkbookTotals204()),"저장→불러오기 총액 유지",log);
            Check204(before.Length==individualDashboard.Rows.Count&&before.All(a=>individualDashboard.Rows.Any(x=>x.Site==a[0]&&x.Fund==a[1]&&x.Name==a[2]&&x.Birth==a[3]&&x.Job==a[4]&&x.ShortTerm.ToString()==a[5])),"동명이인·사업장·재원·직종·대체근로자 유지",log);
            string edited=Path.Combine(directory,"엑셀수정.xlsx");File.Copy(exported,edited,true);
            string targetName,targetSite,targetBirth,originalFund;decimal originalEmploymentPayroll;
            using(var p=new ExcelPackage(new FileInfo(edited))){
                var people=p.Workbook.Worksheets["대상자정보"];string id=people.Cells[8,2].Text;targetSite=people.Cells[8,3].Text;targetName=people.Cells[8,4].Text;targetBirth=people.Cells[8,5].Text;originalFund=people.Cells[8,6].Text;
                people.Cells[8,8].Value="복원 검증 직종";people.Cells[8,6].Value=originalFund=="학교회계"?"교특회계":"학교회계";
                var review=p.Workbook.Worksheets["확인필요"];review.Cells[8,7].Value="확인완료";review.Cells[8,8].Value="원자료 대조 메모";
                var payroll=p.Workbook.Worksheets["원본급여"];int row=Enumerable.Range(8,payroll.Dimension.End.Row-7).First(r=>payroll.Cells[r,3].Text==id&&payroll.Cells[r,4].Text=="고용보험");originalEmploymentPayroll=WorkspaceWorkbook204.Number(payroll.Cells[row,9].Value)+WorkspaceWorkbook204.Number(payroll.Cells[row,10].Value);payroll.Cells[row,9].Value=WorkspaceWorkbook204.Number(payroll.Cells[row,9].Value)+100;
                var disc=p.Workbook.Worksheets["감면적용"];disc.Cells[8,6].Value=WorkspaceWorkbook204.Number(disc.Cells[8,6].Value)+30;disc.Cells[8,15].Value="감면 검증 메모";
                p.Workbook.Worksheets["반환·추징"].Cells[8,12].Value="처리 내역 메모";p.Workbook.Worksheets["작업설정"].Cells[13,3].Value="10월 마감 작업";
                p.Save();
            }
            WorkspaceWorkbook204.RecalculateExcel(edited);
            OpenWorkbook204(edited);var changed=individualDashboard.Rows.Single(x=>x.Site==targetSite&&x.Name==targetName&&x.Birth==targetBirth);
            Check204(changed.Job=="복원 검증 직종"&&changed.Fund!=originalFund,"엑셀 직종·재원 수정 복원",log);
            Check204(changed.EmploymentPayroll==originalEmploymentPayroll+100,"엑셀 보험료 수정 복원",log);
            Check204(changed.ReviewReason=="원자료 대조 메모"&&IsReviewCompleted(changed),"확인 상태·메모 복원",log);
            Check204(WorkbookTotals204()[1]==totals[1]-30,"감면 중복 차감 없이 복원",log);
            string second=Path.Combine(directory,"재저장.xlsx");ExportWorkspace204(second);using(var p=new ExcelPackage(new FileInfo(second))){Check204(p.Workbook.Worksheets["작업설정"].Cells[13,3].Text=="10월 마감 작업","작업 메모 재저장",log);Check204(p.Workbook.Worksheets["감면적용"].Cells[8,15].Text=="감면 검증 메모","감면 메모 재저장",log);Check204(p.Workbook.Worksheets["반환·추징"].Cells[8,12].Text=="처리 내역 메모","반환·추징 메모 재저장",log);}
            decimal[] changedTotals=WorkbookTotals204();OpenWorkbook204(second);Check204(changedTotals.SequenceEqual(WorkbookTotals204()),"반복 저장·불러오기 총액과 감면 유지",log);
            foreach(string problem in new[]{"duplicate","missing","identity","invalidamount","corrupt"}){
                string broken=Path.Combine(directory,problem+".xlsx");File.Copy(exported,broken,true);using(var p=new ExcelPackage(new FileInfo(broken))){var people=p.Workbook.Worksheets["대상자정보"];if(problem=="duplicate")people.Cells[9,2].Value=people.Cells[8,2].Value;else if(problem=="missing")people.Cells[8,2].Value="";else if(problem=="identity")people.Cells[8,4].Value="변경한 이름";else if(problem=="invalidamount")p.Workbook.Worksheets["원본급여"].Cells[8,9].Value="문자 금액";else p.Workbook.Worksheets["_SI작업복원204"].Cells[1,2].Value="손상";p.Save();}
                bool rejected=false;try{WorkspaceWorkbook204.Import(broken,Path.Combine(directory,problem+".xlsm"));}catch{rejected=true;}Check204(rejected,"잘못된 파일 차단: "+problem,log);
            }
            Check204(originalHash==Hash202(native),"기존 입력 파일 보존",log);File.WriteAllLines(Path.Combine(directory,"검증결과.txt"),log,Encoding.UTF8);
        }
        decimal[] WorkbookTotals204(){return new[]{individualDashboard.Rows.Sum(x=>x.SummaryHealthPersonal+x.SummaryLongTermPersonal+x.SummaryPensionPersonal+x.SummaryEmploymentPersonal+x.SummaryIndustrialPersonal),individualDashboard.Rows.Sum(x=>x.SummaryHealthEmployer+x.SummaryLongTermEmployer+x.SummaryPensionEmployer+x.SummaryEmploymentEmployer+x.SummaryIndustrialEmployer)};}
    }
}
