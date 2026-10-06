using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace InsurancePayrollValidator
{
    // The original result stays byte-for-byte in the archive. Display sheets are editable projections.
    static class WorkspaceWorkbook204
    {
        const string Archive="_SI작업복원204", Map="_SI대상키204";
        static readonly string[] Insurance={"건강보험","장기요양보험","국민연금","고용보험","산재보험"};
        public static decimal Number(object x){decimal d;if(x==null||Convert.ToString(x)=="")return 0;if(!Decimal.TryParse(Convert.ToString(x,CultureInfo.InvariantCulture),NumberStyles.Number,CultureInfo.InvariantCulture,out d))throw new InvalidDataException("금액에 숫자가 아닌 값이 있습니다: "+x);return d;}
        static string Hash(byte[] b){using(var sha=SHA256.Create())return BitConverter.ToString(sha.ComputeHash(b)).Replace("-","");}
        static void Set(ExcelWorksheet w,int r,int c,object v){w.Cells[r,c].Formula="";w.Cells[r,c].Value=v;}
        static string Id(int r){return "P"+r.ToString("D6");}
        static string Shift(string f,int delta){return Regex.Replace(f,@"(?<![A-Za-z0-9_])(\$?[A-Z]{1,3})(\$?)([0-9]+)",m=>m.Groups[1].Value+m.Groups[2].Value+(Int32.Parse(m.Groups[3].Value)+(m.Groups[2].Value=="$"?0:delta)).ToString());}
        static void Expand(ExcelPackage p,int capacity)
        {
            int personEnd=capacity+7,detailEnd=capacity*5+7;
            string[] one={"대상자정보","개인별_개인부담","개인별_기관부담","감면적용","확인필요","내부결재_미리보기"};
            string[] five={"원본급여","원본고지","계산내역","반환·추징"};
            foreach(var ws in p.Workbook.Worksheets){
                bool person=one.Contains(ws.Name),detail=five.Contains(ws.Name);
                int end=person?personEnd:detail?detailEnd:0;
                if(end==0)continue;
                int oldEnd=person?47:207,cols=ws.Dimension.End.Column;
                // Capture formulas before clearing prototype data; each group has five insurance rows.
                var formula=new Dictionary<int,string[]>();
                for(int g=0;g<(ws.Name=="계산내역"?5:1);g++)formula[g]=Enumerable.Range(1,cols).Select(c=>ws.Cells[8+g,c].Formula).ToArray();
                if(end>oldEnd){foreach(string merged in ws.MergedCells.ToList())if(new ExcelAddress(merged).Start.Row>oldEnd&&new ExcelAddress(merged).Start.Row<=end)ws.Cells[merged].Merge=false;ws.Cells[oldEnd+1,1,end,cols].Clear();}
                for(int r=8;r<=end;r++){
                    int g=ws.Name=="계산내역"?(r-8)%5:0;
                    for(int c=1;c<=cols;c++){
                        if(r>oldEnd)ws.Cells[r,c].StyleID=ws.Cells[8+g,c].StyleID;
                        Set(ws,r,c,null);
                        string f=formula[g][c-1];
                        if(!String.IsNullOrEmpty(f))ws.Cells[r,c].Formula=Shift(f,r-(8+g));
                    }
                    if(ws.Name=="계산내역"){
                        int pr=8+(r-8)/5;
                        ws.Cells[r,3].Formula="IF('대상자정보'!B"+pr+"=\"\",\"\",'대상자정보'!B"+pr+")";
                        ws.Cells[r,8].Formula="IF(C"+r+"=\"\",\"\",\""+Insurance[g]+"\")";
                    }
                    ws.Row(r).Height=ws.Row(8).Height;
                }
                foreach(var table in ws.Tables){int lastCol=table.Address.End.Column;string address=ws.Cells[7,2,end,lastCol].Address;table.TableXml.DocumentElement.SetAttribute("ref",address);foreach(System.Xml.XmlNode node in table.TableXml.DocumentElement.ChildNodes)if(node.LocalName=="autoFilter")((System.Xml.XmlElement)node).SetAttribute("ref",address);}
                ws.PrinterSettings.PrintArea=ws.Cells[2,2,end,cols];
            }
            foreach(var ws in p.Workbook.Worksheets)if(ws.Dimension!=null){
                foreach(var cell in ws.Cells[ws.Dimension.Address]){
                    if(!String.IsNullOrEmpty(cell.Formula))cell.Formula=cell.Formula.Replace("$207","$"+detailEnd).Replace("$47","$"+personEnd);
                    else if(cell.Value is string && (cell.Text.Contains("가상")||cell.Text.Contains("시안")))cell.Value=cell.Text.Replace("가상 자료 시안","저장된 작업 자료").Replace("가상 자료","저장된 작업 자료").Replace("심플 시안","대사 작업").Replace("시안용 가상 값","저장된 값").Replace("가상하늘학교","기관").Replace("가상 원본 입력","저장된 원본 금액").Replace("설계시안","대사작업");
                }
                ws.HeaderFooter.OddFooter.LeftAlignedText="사회보험 대사 작업";
            }
            // The generated IDs differ from the demonstration IDs. Keep input validations live.
            p.Workbook.Names.Add("WorkbookPeople204",p.Workbook.Worksheets["대상자정보"].Cells[8,2,personEnd,2]);
            foreach(string name in new[]{"감면적용","확인필요","반환·추징"}){
                var ws=p.Workbook.Worksheets[name];foreach(var v in ws.DataValidations.ToList())if(v.Address.Start.Column==2)ws.DataValidations.Remove(v);
                var validation=ws.DataValidations.AddListValidation("B8:B"+(name=="반환·추징"?detailEnd:personEnd));validation.Formula.ExcelFormula="WorkbookPeople204";
            }
            var dash=p.Workbook.Worksheets["총괄표"];foreach(var v in dash.DataValidations.ToList())if(v.Address.Address=="D6")dash.DataValidations.Remove(v);
        }
        static ExcelWorksheet Required(ExcelPackage p,string name){var w=p.Workbook.Worksheets[name];if(w==null)throw new InvalidDataException("작업 파일의 필수 탭이 없습니다: "+name);return w;}
        public static bool IsArchive(string path){using(var p=new ExcelPackage(new FileInfo(path)))return p.Workbook.Worksheets[Archive]!=null;}
        public static void Export(string native,string destination,SubmissionInfo info)
        {
            if(Path.GetFullPath(native).Equals(Path.GetFullPath(destination),StringComparison.OrdinalIgnoreCase))throw new InvalidOperationException("원본 작업 파일과 다른 이름으로 저장해 주세요.");
            using(var source=new ExcelPackage(new FileInfo(native)))
            using(var resource=Assembly.GetExecutingAssembly().GetManifestResourceStream("InsurancePayrollValidator.WorkspaceTemplate.xlsx"))
            using(var p=new ExcelPackage(resource)){
                var data=Required(source,"UI개인별데이터");
                var rows=Enumerable.Range(2,data.Dimension.End.Row-1).Where(r=>data.Cells[r,1].Text!=""&&data.Cells[r,3].Text!="").OrderBy(r=>Number(data.Cells[r,34].Value)>0?1:0).ThenBy(r=>data.Cells[r,3].Text).ToList();
                if(rows.Count==0)throw new InvalidDataException("보관할 개인별 자료가 없습니다.");
                int capacity=Math.Max(40,rows.Count);Expand(p,capacity);
                var map=p.Workbook.Worksheets.Add(Map);map.Cells[1,1].Value="ID";map.Cells[1,2].Value="원본행";map.Hidden=eWorkSheetHidden.VeryHidden;
                var people=p.Workbook.Worksheets["대상자정보"];var pay=p.Workbook.Worksheets["원본급여"];var notice=p.Workbook.Worksheets["원본고지"];var disc=p.Workbook.Worksheets["감면적용"];var rev=p.Workbook.Worksheets["확인필요"];var adj=p.Workbook.Worksheets["반환·추징"];
                var discounts=source.Workbook.Worksheets["UI감면상태"];var checkedKeys=source.Workbook.Worksheets["UI확인상태"];
                string period=Number(data.Cells[rows[0],19].Value).ToString("0")+"-"+Number(data.Cells[rows[0],20].Value).ToString("00");
                for(int i=0;i<rows.Count;i++){
                    int nr=rows[i],r=8+i;string id=Id(nr),site=data.Cells[nr,1].Text,name=data.Cells[nr,3].Text,birth=data.Cells[nr,4].Text,fund=data.Cells[nr,2].Text;bool shortTerm=Number(data.Cells[nr,34].Value)>0;
                    string key=site+"|"+name+"|"+Regex.Replace(birth,"[^0-9]","");
                    map.Cells[i+2,1].Value=id;map.Cells[i+2,2].Value=nr;map.Cells[i+2,3].Value=key;
                    object[] identity={id,site,name,birth,fund,shortTerm?"1개월 미만 대체근로자":fund=="공무원"?"공무원":fund=="계약제교원"?"계약제교원":"교육공무직",data.Cells[nr,5].Text,"",data.Cells[nr,41].Text};
                    for(int c=0;c<identity.Length;c++)Set(people,r,2+c,identity[c]);
                    decimal hp=Number(data.Cells[nr,22].Value),lp=Number(data.Cells[nr,24].Value);
                    if(Number(data.Cells[nr,35].Value)==0){hp=Number(data.Cells[nr,7].Value);lp=0;}
                    decimal[] charges={hp,lp,Number(data.Cells[nr,26].Value),Number(data.Cells[nr,28].Value),Number(data.Cells[nr,30].Value)};
                    decimal payrollTotal=Number(data.Cells[nr,8].Value),longPayroll=lp-Number(data.Cells[nr,33].Value);
                    if(Number(data.Cells[nr,35].Value)==0)longPayroll=0;
                    decimal[] payroll={payrollTotal-longPayroll,longPayroll,Number(data.Cells[nr,11].Value),Number(data.Cells[nr,14].Value),0};
                    decimal[] employer={Number(data.Cells[nr,23].Value),Number(data.Cells[nr,25].Value),Number(data.Cells[nr,27].Value),Number(data.Cells[nr,29].Value),Number(data.Cells[nr,31].Value)};
                    decimal[] discount=new decimal[4];bool autoE=false,autoI=false;
                    if(discounts!=null&&discounts.Dimension!=null)for(int dr=2;dr<=discounts.Dimension.End.Row;dr++)if(discounts.Cells[dr,1].Text==key+"|"+fund+"|"+(shortTerm?"1":"0")||discounts.Cells[dr,1].Text==key){autoE=Number(discounts.Cells[dr,2].Value)>0;autoI=Number(discounts.Cells[dr,3].Value)>0;for(int k=0;k<4;k++)discount[k]=Number(discounts.Cells[dr,4+k].Value);break;}
                    employer[0]+=discount[0];employer[2]+=discount[1];employer[3]+=discount[2]+(autoE?250:0);employer[4]+=discount[3]+(autoI?250:0);
                    Set(disc,r,2,id);for(int k=0;k<4;k++)Set(disc,r,6+k,discount[k]);Set(disc,r,10,autoE?"예":"아니오");Set(disc,r,11,autoI?"예":"아니오");Set(disc,r,12,250);Set(disc,r,13,250);
                    Set(rev,r,2,id);bool complete=checkedKeys!=null&&checkedKeys.Dimension!=null&&Enumerable.Range(2,Math.Max(0,checkedKeys.Dimension.End.Row-1)).Any(cr=>checkedKeys.Cells[cr,1].Text==key);Set(rev,r,7,complete?"확인완료":"미확인");Set(rev,r,8,data.Cells[nr,21].Text);
                    for(int k=0;k<5;k++){
                        int rr=8+i*5+k;
                        object[] pv={"PAY-"+id+"-"+k,id,Insurance[k],period,Path.GetFileName(native),"UI개인별데이터",nr,payroll[k],0,"프로그램 저장 공제 합계"};
                        for(int c=0;c<pv.Length;c++)Set(pay,rr,2+c,pv[c]);
                        decimal settlement=Number(data.Cells[nr,36+k].Value);
                        object[] nv={"NOT-"+id+"-"+k,id,Insurance[k],period,Path.GetFileName(native),"UI개인별데이터",nr,charges[k],0,null,employer[k]-settlement,settlement,null,"저장 고지 합계. 세부 원자료는 작업 복원 데이터에 보관",charges[k],employer[k]};
                        for(int c=0;c<nv.Length;c++)if(c!=9&&c!=12)Set(notice,rr,2+c,nv[c]);
                        Set(adj,rr,2,id);Set(adj,rr,5,Insurance[k]);Set(adj,rr,10,"미처리");Set(adj,rr,11,0);
                    }
                }
                // Keep management notes with their IDs when the workbook is exported again.
                foreach(string name in new[]{"감면적용","반환·추징"}){
                    var old=source.Workbook.Worksheets["WB204_"+name];var target=p.Workbook.Worksheets[name];if(old==null||old.Dimension==null)continue;
                    var notes=new Dictionary<string,int>();for(int r=8;r<=old.Dimension.End.Row;r++)if(old.Cells[r,2].Text!="")notes[old.Cells[r,2].Text+(name=="반환·추징"?"|"+old.Cells[r,5].Text:"")]=r;
                    for(int r=8;r<=target.Dimension.End.Row;r++){int nr;string key=target.Cells[r,2].Text+(name=="반환·추징"?"|"+target.Cells[r,5].Text:"");if(!notes.TryGetValue(key,out nr))continue;foreach(int c in name=="반환·추징"?new[]{10,11,12}:new[]{15})Set(target,r,c,old.Cells[nr,c].Value);}
                }
                var cfg=p.Workbook.Worksheets["작업설정"];Set(cfg,8,3,"SI-WORKBOOK-204");Set(cfg,9,3,"1");Set(cfg,10,3,Guid.NewGuid().ToString("N"));Set(cfg,11,3,period);Set(cfg,12,3,info.InstitutionName??"");Set(cfg,13,3,ReadNote(source));Set(cfg,15,3,info.RecipientCode??"");Set(cfg,16,3,info.Round??"");Set(cfg,17,3,info.IndustrialRate??"");Set(cfg,18,3,info.BankName??"");Set(cfg,19,3,info.AccountNumber??"");Set(cfg,22,3,"프로그램에서 저장한 작업");
                Set(cfg,38,2,"이 파일은 테스트판의 총괄표 > 저장한 작업 열기로 다시 불러올 수 있습니다.");
                var dash=p.Workbook.Worksheets["총괄표"];Set(dash,3,2,period+" · "+(info.InstitutionName??"")+" · 대사 작업 보관");Set(dash,6,4,"전체");
                var sites=p.Workbook.Worksheets.Add("_SI사업장204");sites.Cells[1,1].Value="전체";int sr=2;foreach(string site in rows.Select(n=>data.Cells[n,1].Text).Distinct())sites.Cells[sr++,1].Value=site;sites.Hidden=eWorkSheetHidden.VeryHidden;p.Workbook.Names.Add("WorkbookSites204",sites.Cells[1,1,sr-1,1]);var dv=dash.DataValidations.AddListValidation("D6");dv.Formula.ExcelFormula="WorkbookSites204";
                var files=p.Workbook.Worksheets["파일등록"];files.Cells[8,2,Math.Max(18,files.Dimension.End.Row),10].Clear();var recognition=source.Workbook.Worksheets["자료인식"];if(recognition!=null&&recognition.Dimension!=null)for(int n=2;n<=recognition.Dimension.End.Row;n++){int rr=n+6;Set(files,rr,2,"FILE"+n);Set(files,rr,3,recognition.Cells[n,1].Text);Set(files,rr,4,recognition.Cells[n,2].Text);Set(files,rr,5,recognition.Cells[n,3].Text);Set(files,rr,8,recognition.Cells[n,5].Text);}
                var help=p.Workbook.Worksheets["사용안내"];Set(help,17,3,"총괄표의 총괄 엑셀 저장/저장한 작업 열기 버튼으로 XLSX를 저장하고 다시 열 수 있습니다. 원본 파일을 이동해도 작업 복원 데이터가 함께 보관됩니다.");Set(help,21,2,"노란 입력칸의 재원·직종·공제·고지·감면·메모를 수정할 수 있습니다. 대상자 ID·사업장·이름·생년월일은 변경하지 마세요.");
                // Embed the source only after projections have been created. No original input is modified.
                byte[] bytes=File.ReadAllBytes(native);var archive=p.Workbook.Worksheets.Add(Archive);archive.Cells[1,1].Value="SI-WORKBOOK-204";archive.Cells[1,2].Value=Hash(bytes);archive.Cells[1,3].Value=info.ManagerName??"";archive.Cells[1,4].Value=info.Phone??"";string b64=Convert.ToBase64String(bytes);for(int offset=0,index=2;offset<b64.Length;offset+=30000,index++)archive.Cells[index,1].Value=b64.Substring(offset,Math.Min(30000,b64.Length-offset));archive.Hidden=eWorkSheetHidden.VeryHidden;
                p.Workbook.CalcMode=ExcelCalcMode.Automatic;
                string tmp=destination+"."+Guid.NewGuid().ToString("N")+".xlsx";try{p.SaveAs(new FileInfo(tmp));RecalculateExcel(tmp);if(File.Exists(destination))File.Replace(tmp,destination,null);else File.Move(tmp,destination);}finally{if(File.Exists(tmp))File.Delete(tmp);}
            }
        }
        public static void RecalculateExcel(string path)
        {
            Type type=Type.GetTypeFromProgID("Excel.Application");if(type==null)throw new InvalidOperationException("엑셀 작업 보관 파일의 수식 계산에는 Microsoft Excel 설치가 필요합니다.");
            object app=null,book=null,books=null;
            try{
                app=Activator.CreateInstance(type);Com(app,"Visible",BindingFlags.SetProperty,false);Com(app,"DisplayAlerts",BindingFlags.SetProperty,false);Com(app,"EnableEvents",BindingFlags.SetProperty,false);Com(app,"AutomationSecurity",BindingFlags.SetProperty,3);
                books=Com(app,"Workbooks",BindingFlags.GetProperty);book=Com(books,"Open",BindingFlags.InvokeMethod,Path.GetFullPath(path),0,false);Com(app,"CalculateFullRebuild",BindingFlags.InvokeMethod);Com(book,"Save",BindingFlags.InvokeMethod);
                // Native Excel validates formulas with the same engine used by the user.
                using(var p=new ExcelPackage(new FileInfo(path)))foreach(var ws in p.Workbook.Worksheets.Where(w=>w.Dimension!=null&&!w.Name.StartsWith("_SI")))foreach(var cell in ws.Cells[ws.Dimension.Address])if(cell.Value is ExcelErrorValue)throw new InvalidDataException("엑셀 수식 오류: "+ws.Name+"!"+cell.Address+" "+cell.Value);
            }finally{if(book!=null){try{Com(book,"Close",BindingFlags.InvokeMethod,false);}finally{Marshal.FinalReleaseComObject(book);}}if(books!=null)Marshal.FinalReleaseComObject(books);if(app!=null){try{Com(app,"Quit",BindingFlags.InvokeMethod);}finally{Marshal.FinalReleaseComObject(app);}}}
        }
        static object Com(object target,string member,BindingFlags flags,params object[] args){return target.GetType().InvokeMember(member,flags,null,target,args,CultureInfo.GetCultureInfo("en-US"));}
        static string ReadNote(ExcelPackage source){var ws=source.Workbook.Worksheets["WB204_설정"];return ws==null?"":ws.Cells[1,1].Text;}
        static Dictionary<string,int> Index(ExcelWorksheet ws,HashSet<string> ids,bool required,bool insurance){
            var result=new Dictionary<string,int>();int end=ws.Tables.Count>0?ws.Tables.First().Address.End.Row:ws.Dimension.End.Row;for(int r=8;r<=end;r++){string id=ws.Cells[r,2].Text;if(id=="")continue;if(!ids.Contains(id))throw new InvalidDataException(ws.Name+": 알 수 없는 대상자ID "+id);string key=id+(insurance?"|"+ws.Cells[r,5].Text:"");if(result.ContainsKey(key))throw new InvalidDataException(ws.Name+": 대상자ID가 중복되었습니다: "+key);result.Add(key,r);}
            if(required&&!ids.SetEquals(result.Keys))throw new InvalidDataException(ws.Name+": 대상자 명단이 변경되거나 누락되었습니다.");return result;
        }
        static decimal Money(ExcelWorksheet ws,int r,int c){decimal d=Number(ws.Cells[r,c].Value);if(d!=Decimal.Truncate(d)||Math.Abs(d)>1000000000000m)throw new InvalidDataException(ws.Name+": 금액은 원 단위 정수로 입력해 주세요.");return d;}
        public static SubmissionInfo Import(string file,string target)
        {
            using(var p=new ExcelPackage(new FileInfo(file))){
                var archive=Required(p,Archive);if(archive.Cells[1,1].Text!="SI-WORKBOOK-204")throw new InvalidDataException("지원하지 않는 작업 파일 버전입니다.");
                var b=new StringBuilder();for(int r=2;r<=archive.Dimension.End.Row;r++)b.Append(archive.Cells[r,1].Text);byte[] native=Convert.FromBase64String(b.ToString());if(Hash(native)!=archive.Cells[1,2].Text)throw new InvalidDataException("작업 복원 데이터가 손상되었습니다.");
                using(var source=new ExcelPackage(new MemoryStream(native))){
                    var cfg=Required(p,"작업설정");if(cfg.Cells[8,3].Text!="SI-WORKBOOK-204"||cfg.Cells[9,3].Text!="1")throw new InvalidDataException("작업 설정의 저장형식/스키마 버전을 변경하지 마세요.");
                    var data=Required(source,"UI개인별데이터");var map=Required(p,Map);var ids=new HashSet<string>();var nativeRows=new Dictionary<string,int>();
                    for(int r=2;r<=map.Dimension.End.Row;r++){string id=map.Cells[r,1].Text;int nr=(int)Number(map.Cells[r,2].Value);if(nr<2||nr>data.Dimension.End.Row||id!=Id(nr)||!ids.Add(id))throw new InvalidDataException("대상키 복원 데이터가 올바르지 않습니다.");nativeRows.Add(id,nr);}
                    int nativeCount=Enumerable.Range(2,data.Dimension.End.Row-1).Count(r=>data.Cells[r,1].Text!=""&&data.Cells[r,3].Text!="");if(nativeCount!=ids.Count)throw new InvalidDataException("대상키 복원 명단이 누락되었습니다.");
                    var people=Required(p,"대상자정보");var pi=Index(people,ids,true,false);var rev=Required(p,"확인필요");var ri=Index(rev,ids,true,false);var discount=Required(p,"감면적용");var di=Index(discount,ids,true,false);
                    var pay=Required(p,"원본급여");var notice=Required(p,"원본고지");var payIndex=new Dictionary<string,int>();var noticeIndex=new Dictionary<string,int>();
                    foreach(var pair in new[]{Tuple.Create(pay,payIndex),Tuple.Create(notice,noticeIndex)})for(int r=8;r<=pair.Item1.Dimension.End.Row;r++){string id=pair.Item1.Cells[r,3].Text;if(id=="")continue;string kind=pair.Item1.Cells[r,4].Text,key=id+"|"+kind;if(!ids.Contains(id)||!Insurance.Contains(kind)||pair.Item2.ContainsKey(key))throw new InvalidDataException(pair.Item1.Name+": 원본 대상자/보험키가 중복되거나 잘못되었습니다.");if(pair.Item1.Cells[r,5].Text!=cfg.Cells[11,3].Text)throw new InvalidDataException("대상년월이 일치하지 않습니다.");pair.Item2.Add(key,r);}
                    if(payIndex.Count!=ids.Count*5||noticeIndex.Count!=ids.Count*5)throw new InvalidDataException("저장된 공제/고지 행이 누락되었습니다.");
                    var checkedKeys=source.Workbook.Worksheets["UI확인상태"]??source.Workbook.Worksheets.Add("UI확인상태");checkedKeys.Cells.Clear();checkedKeys.Cells[1,1].Value="대상키";checkedKeys.Cells[1,2].Value="조회상태";int checkedRow=2;
                    var ds=source.Workbook.Worksheets["UI감면상태"]??source.Workbook.Worksheets.Add("UI감면상태");ds.Cells.Clear();string[] headers={"대상키","고용자동","산재자동","건강기타","국민기타","고용기타","산재기타","적용기준"};for(int c=0;c<headers.Length;c++)ds.Cells[1,c+1].Value=headers[c];int dsr=2;
                    foreach(string id in ids){
                        int nr=nativeRows[id],r=pi[id],dr=di[id],rr=ri[id];string site=data.Cells[nr,1].Text,name=data.Cells[nr,3].Text,birth=data.Cells[nr,4].Text;
                        if(people.Cells[r,3].Text!=site||people.Cells[r,4].Text!=name||people.Cells[r,5].Text!=birth)throw new InvalidDataException("대상자ID·사업장·이름·생년월일은 변경하지 마세요: "+name);
                        string fund=people.Cells[r,6].Text;if(!new[]{"공무원","계약제교원","교특회계","학교회계","기타급여","분류필요","휴직"}.Contains(fund))throw new InvalidDataException("지원하지 않는 재원입니다: "+fund);
                        bool shortTerm=Number(data.Cells[nr,34].Value)>0;if((people.Cells[r,7].Text=="1개월 미만 대체근로자")!=shortTerm)throw new InvalidDataException("대체근로자 유형은 변경하지 마세요: "+name);
                        string job=people.Cells[r,8].Text;if(job.Length>100||job.Contains("\n")||job.Contains("\r"))throw new InvalidDataException("직종은 100자 이내 한 줄로 입력해 주세요.");
                        decimal[] charges=new decimal[5],payroll=new decimal[5],employer=new decimal[5];
                        for(int k=0;k<5;k++){string key=id+"|"+Insurance[k];int pr=payIndex[key],nn=noticeIndex[key];payroll[k]=Money(pay,pr,9)+Money(pay,pr,10);charges[k]=Money(notice,nn,9)+Money(notice,nn,10);employer[k]=Money(notice,nn,12)+Money(notice,nn,13);Set(data,nr,36+k,Money(notice,nn,13));}
                        decimal[] reductions=new decimal[4];for(int k=0;k<4;k++){reductions[k]=Money(discount,dr,6+k);if(reductions[k]<0)throw new InvalidDataException("감면 금액은 0 이상이어야 합니다.");}
                        string ae=discount.Cells[dr,10].Text,ai=discount.Cells[dr,11].Text;if(!new[]{"예","아니오"}.Contains(ae)||!new[]{"예","아니오"}.Contains(ai))throw new InvalidDataException("자동감면은 예/아니오로 입력해 주세요.");
                        if(Money(discount,dr,12)!=250||Money(discount,dr,13)!=250)throw new InvalidDataException("자동감면 단가는 기존 프로그램과 동일한 250원입니다.");
                        employer[0]-=reductions[0];employer[2]-=reductions[1];employer[3]-=reductions[2]+(ae=="예"?250:0);employer[4]-=reductions[3]+(ai=="예"?250:0);
                        Set(data,nr,2,fund);Set(data,nr,5,job);Set(data,nr,7,charges[0]+charges[1]);Set(data,nr,8,payroll[0]+payroll[1]);Set(data,nr,9,charges[0]+charges[1]-payroll[0]-payroll[1]);
                        int[] noticeCols={10,13},payCols={11,14},diffCols={12,15};for(int k=2;k<4;k++){Set(data,nr,noticeCols[k-2],charges[k]);Set(data,nr,payCols[k-2],payroll[k]);Set(data,nr,diffCols[k-2],charges[k]-payroll[k]);}
                        // Industrial reconciliation is an employer comparison in the existing app.
                        decimal industrialNotice=Number(data.Cells[nr,16].Value)+employer[4]-Number(data.Cells[nr,31].Value);Set(data,nr,16,industrialNotice);Set(data,nr,18,industrialNotice-Number(data.Cells[nr,17].Value));
                        for(int k=0;k<5;k++){Set(data,nr,22+k*2,charges[k]);Set(data,nr,23+k*2,employer[k]);}Set(data,nr,32,charges[0]-payroll[0]);Set(data,nr,33,charges[1]-payroll[1]);
                        string status=rev.Cells[rr,7].Text;if(!new[]{"확인완료","확인 완료","미확인","보류","확인중","해당없음"}.Contains(status))throw new InvalidDataException("확인 상태는 확인완료/미확인/확인중/해당없음/보류로 입력해 주세요.");
                        Set(data,nr,21,rev.Cells[rr,8].Text);string personKey=site+"|"+name+"|"+Regex.Replace(birth,"[^0-9]","");
                        if(status=="확인완료"||status=="확인 완료"){checkedKeys.Cells[checkedRow,1].Value=personKey;checkedKeys.Cells[checkedRow++,2].Value="확인 완료";}
                        object[] dsv={personKey+"|"+fund+"|"+(shortTerm?"1":"0"),ae=="예"?1:0,ai=="예"?1:0,reductions[0],reductions[1],reductions[2],reductions[3],"기관부담차감"};for(int c=0;c<dsv.Length;c++)ds.Cells[dsr,c+1].Value=dsv[c];dsr++;
                    }
                    foreach(string name in new[]{"감면적용","반환·추징"}){var input=Required(p,name);if(name=="반환·추징")Index(input,ids,false,true);string saved="WB204_"+name;if(source.Workbook.Worksheets[saved]!=null)source.Workbook.Worksheets.Delete(saved);var notes=source.Workbook.Worksheets.Add(saved);int end=input.Tables.First().Address.End.Row;for(int r=8;r<=end;r++)if(input.Cells[r,2].Text!="")foreach(int c in name=="반환·추징"?new[]{2,5,10,11,12}:new[]{2,15})notes.Cells[r,c].Value=input.Cells[r,c].Value;notes.Hidden=eWorkSheetHidden.Hidden;}
                    var settings=source.Workbook.Worksheets["WB204_설정"]??source.Workbook.Worksheets.Add("WB204_설정");settings.Cells[1,1].Value=cfg.Cells[13,3].Text;settings.Hidden=eWorkSheetHidden.Hidden;checkedKeys.Hidden=eWorkSheetHidden.Hidden;ds.Hidden=eWorkSheetHidden.Hidden;
                    source.SaveAs(new FileInfo(target));
                    return new SubmissionInfo{InstitutionName=cfg.Cells[12,3].Text,RecipientCode=cfg.Cells[15,3].Text,Round=cfg.Cells[16,3].Text,IndustrialRate=cfg.Cells[17,3].Text,BankName=cfg.Cells[18,3].Text,AccountNumber=cfg.Cells[19,3].Text,ManagerName=archive.Cells[1,3].Text,Phone=archive.Cells[1,4].Text};
                }
            }
        }
    }
}
