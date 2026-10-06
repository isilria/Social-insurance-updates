using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace InsurancePayrollValidator
{
    static partial class AdjustmentReportGenerator
    {
        static readonly Color[] GroupColors={Color.FromArgb(242,170,132),Color.FromArgb(115,159,180),Color.FromArgb(149,202,130)};
        static readonly Color[] SubColors={Color.FromArgb(246,198,173),Color.FromArgb(161,191,205),Color.FromArgb(184,220,172)};
        const string ReportNote="※ 본 자료는 반환 또는 추징 대상자 확인용 보조자료입니다. 원자료와 최종 금액을 반드시 확인해 주세요.";
        static string ReportMode(string mode){return mode=="전체"?"반환·추징":mode;}
        static string TotalCaption(string mode){return mode=="전체"?"조정총액":mode+"총액";}
        static decimal[] PersonalValues(IndividualRowData d){return new[]{d.HealthNotice,d.HealthPayroll,d.HealthDifference,d.PensionNotice,d.PensionPayroll,d.PensionDifference,d.EmploymentNotice,d.EmploymentPayroll,d.EmploymentDifference};}

        public static void CreateExcel(string path,List<IndividualRowData> rows,string mode,int year,int month,string site)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path)));
            using(var package=new ExcelPackage()){
                var ws=package.Workbook.Worksheets.Add(mode=="전체"?"반환추징내역":mode+"내역");
                int last=Math.Max(9,8+rows.Count),noteRow=last+5;
                ws.Cells[1,2,noteRow,18].Style.Font.Name="맑은 고딕";ws.Cells[1,2,noteRow,18].Style.Font.Size=11;
                ws.Cells[1,2,1,18].Merge=true;ws.Row(1).Height=16;
                ws.Cells[2,2,2,18].Merge=true;ws.Cells[2,2].Value="사회보험료 "+ReportMode(mode)+" 대상 내역";
                ws.Cells[2,2].Style.Font.Size=18;ws.Cells[2,2].Style.Font.Bold=true;ws.Cells[2,2].Style.Font.Color.SetColor(Navy);ws.Row(2).Height=31;
                ws.Cells[4,2].Value="고지년월";ws.Cells[4,3,4,6].Merge=true;ws.Cells[4,3].Value=year+"년 "+month+"월";
                ws.Cells[5,2].Value="사업장";ws.Cells[5,3,5,8].Merge=true;ws.Cells[5,3].Value=site;
                string[] fixedHeads={"No.","구분","사업장 관리번호","재원","이름","주민번호","직종명"};
                for(int c=2;c<=8;c++){ws.Cells[7,c,8,c].Merge=true;ws.Cells[7,c].Value=fixedHeads[c-2];}
                ws.Cells[7,18,8,18].Merge=true;ws.Cells[7,18].Value=TotalCaption(mode);
                using(var header=ws.Cells[7,2,8,18]){header.Style.Font.Bold=true;header.Style.Font.Color.SetColor(Color.White);header.Style.Fill.PatternType=ExcelFillStyle.Solid;header.Style.Fill.BackgroundColor.SetColor(Navy);header.Style.WrapText=true;}
                for(int g=0;g<3;g++){
                    int c=9+g*3;ws.Cells[7,c,7,c+2].Merge=true;ws.Cells[7,c].Value=new[]{"건강","국민","고용"}[g];
                    ws.Cells[7,c,7,c+2].Style.Fill.BackgroundColor.SetColor(GroupColors[g]);ws.Cells[8,c,8,c+2].Style.Fill.BackgroundColor.SetColor(SubColors[g]);
                    ws.Cells[7,c,8,c+2].Style.Font.Color.SetColor(Color.Black);
                    for(int k=0;k<3;k++)ws.Cells[8,c+k].Value=new[]{"고지금액","급여대장","차액"}[k];
                }
                ws.Row(7).Height=27;ws.Row(8).Height=27;
                int r=9,index=1;
                foreach(var d in rows){
                    object[] identity={index++,Kind(d,mode),d.Site,d.Fund,d.Name,MaskBirth(d.Birth),d.Job};
                    for(int c=0;c<identity.Length;c++)ws.Cells[r,c+2].Value=identity[c];
                    decimal[] values=PersonalValues(d);for(int c=0;c<values.Length;c++)ws.Cells[r,9+c].Value=(double)values[c];
                    // Export the completed reconciliation snapshot, as in the supplied examples and existing reports.
                    ws.Cells[r,18].Value=Amount(d,mode);
                    ws.Cells[r,9,r,18].Style.Numberformat.Format="#,##0;[Red]-#,##0;0";
                    for(int g=0;g<3;g++)ws.Cells[r,11+g*3].Style.Font.Color.SetColor(DiffColor(values[g*3+2]));
                    ws.Cells[r,18].Style.Font.Bold=true;ws.Cells[r,18].Style.Font.Color.SetColor(mode=="반환"?Blue:mode=="추징"?Red:Navy);
                    ws.Cells[r,8].Style.WrapText=true;ws.Row(r).Height=30;r++;
                }
                using(var body=ws.Cells[7,2,last,18]){body.Style.HorizontalAlignment=ExcelHorizontalAlignment.Center;body.Style.VerticalAlignment=ExcelVerticalAlignment.Center;body.Style.Border.Top.Style=body.Style.Border.Bottom.Style=body.Style.Border.Left.Style=body.Style.Border.Right.Style=ExcelBorderStyle.Thin;body.Style.Border.Top.Color.SetColor(Grid);body.Style.Border.Bottom.Color.SetColor(Grid);body.Style.Border.Left.Color.SetColor(Grid);body.Style.Border.Right.Color.SetColor(Grid);}
                ws.Cells[last,2,last,18].Style.Border.Bottom.Style=ExcelBorderStyle.Medium;
                ws.Cells[noteRow,2,noteRow,18].Merge=true;ws.Cells[noteRow,2].Value=ReportNote;ws.Cells[noteRow,2].Style.Font.Color.SetColor(Color.FromArgb(126,32,32));ws.Row(noteRow).Height=22;
                ws.Column(1).Width=3;double[] widths={10.28,12,20,14,12,17,24};for(int c=2;c<=8;c++)ws.Column(c).Width=widths[c-2];for(int c=9;c<=18;c++)ws.Column(c).Width=13;
                ws.View.FreezePanes(9,9);ws.PrinterSettings.PaperSize=ePaperSize.A4;ws.PrinterSettings.Orientation=eOrientation.Landscape;ws.PrinterSettings.FitToPage=true;ws.PrinterSettings.FitToWidth=1;ws.PrinterSettings.FitToHeight=0;ws.PrinterSettings.RepeatRows=new ExcelAddress("7:8");ws.PrinterSettings.PrintArea=ws.Cells[1,2,noteRow,18];
                ws.PrinterSettings.LeftMargin=.25M;ws.PrinterSettings.RightMargin=.25M;ws.PrinterSettings.TopMargin=.35M;ws.PrinterSettings.BottomMargin=.35M;
                package.SaveAs(new FileInfo(path));
            }
        }
        static void DrawPdfPage(Graphics g,List<IndividualRowData> rows,string mode,int year,int month,string site,int page,int pages,int offset,int totalCount)
        {
            using(var title=new Font("맑은 고딕",27F,FontStyle.Bold))using(var meta=new Font("맑은 고딕",12F))using(var head=new Font("맑은 고딕",10F,FontStyle.Bold))using(var cell=new Font("맑은 고딕",10F))using(var small=new Font("맑은 고딕",9F)){
                Draw(g,"사회보험료 "+ReportMode(mode)+" 대상 내역",title,Navy,new RectangleF(48,35,1658,55),StringAlignment.Near);
                Draw(g,"고지년월   "+year+"년 "+month+"월",meta,Navy,new RectangleF(48,115,1000,32),StringAlignment.Near);
                Draw(g,"사업장      "+site,meta,Navy,new RectangleF(48,155,1658,32),StringAlignment.Near);
                int[] widths={45,65,132,92,90,108,137,97,97,97,97,97,97,97,97,97,116};
                int x=48,y=220;const int half=38,rowH=45;
                string[] fixedHeads={"No.","구분","사업장\n관리번호","재원","이름","주민번호","직종명"};
                using(var pen=new Pen(Grid)){
                    for(int c=0;c<17;c++){
                        if(c<7||c==16){using(var fill=new SolidBrush(Navy))g.FillRectangle(fill,x,y,widths[c],half*2);DrawReportCell(g,c==16?TotalCaption(mode):fixedHeads[c],head,Color.White,new RectangleF(x,y,widths[c],half*2));g.DrawRectangle(pen,x,y,widths[c],half*2);}
                        else{int group=(c-7)/3,k=(c-7)%3;if(k==0){using(var fill=new SolidBrush(GroupColors[group]))g.FillRectangle(fill,x,y,widths[c]*3,half);DrawReportCell(g,new[]{"건강","국민","고용"}[group],head,Color.Black,new RectangleF(x,y,widths[c]*3,half));g.DrawRectangle(pen,x,y,widths[c]*3,half);}using(var fill=new SolidBrush(SubColors[group]))g.FillRectangle(fill,x,y+half,widths[c],half);DrawReportCell(g,new[]{"고지금액","급여대장","차액"}[k],head,Color.Black,new RectangleF(x,y+half,widths[c],half));g.DrawRectangle(pen,x,y+half,widths[c],half);}
                        x+=widths[c];
                    }
                    y+=half*2;
                    for(int r=0;r<rows.Count;r++){
                        var d=rows[r];var values=PersonalValues(d);var texts=new List<string>{(offset+r+1).ToString(),Kind(d,mode),d.Site,d.Fund,d.Name,MaskBirth(d.Birth),d.Job};texts.AddRange(values.Select(UiDrawing.Money));texts.Add(UiDrawing.Money(Amount(d,mode)));x=48;
                        for(int c=0;c<17;c++){Color ink=c>=7&&c<16&&(c-7)%3==2?DiffColor(values[c-7]):c==16?(mode=="반환"?Blue:mode=="추징"?Red:Navy):Navy;DrawReportCell(g,texts[c],c==2||c==6?small:cell,ink,new RectangleF(x+2,y,widths[c]-4,rowH));g.DrawRectangle(pen,x,y,widths[c],rowH);x+=widths[c];}y+=rowH;
                    }
                }
                Draw(g,ReportNote,small,Color.FromArgb(126,32,32),new RectangleF(48,1120,1658,40),StringAlignment.Near);
                Draw(g,page+" / "+pages,small,Color.Gray,new RectangleF(48,1180,1658,28),StringAlignment.Far);
            }
        }
        static void DrawReportCell(Graphics g,string value,Font font,Color color,RectangleF rect)
        {
            using(var brush=new SolidBrush(color))using(var format=new StringFormat{Alignment=StringAlignment.Center,LineAlignment=StringAlignment.Center})g.DrawString(value??"",font,brush,rect,format);
        }
    }
}
