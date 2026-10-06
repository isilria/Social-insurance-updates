using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace InsurancePayrollValidator
{
    partial class MainForm
    {
        bool submissionSettingsReady;
        Label submissionSaveStatus,footerWarning;
        Panel footerWarningPanel;
        const string FooterWarningText="본 프로그램은 개인 제작 프로그램입니다. 사용 전 현행 지침과 규정을 확인하시기 바랍니다.";

        void InitializeSubmissionPersistence()
        {
            submissionSaveStatus=new Label{Text="기본정보는 입력 시 자동 저장됩니다",Location=new Point(500,12),Size=new Size(512,18),TextAlign=ContentAlignment.MiddleRight,Font=new Font("맑은 고딕",7.5F),ForeColor=UiMuted,Tag="ThemeMuted",Name="SubmissionSaveStatus"};
            recipientCode.Parent.Parent.Controls.Add(submissionSaveStatus);
            submissionSettingsReady=true;
            foreach(TextBox field in new[]{recipientCode,institutionName,managerName,phone,bankName,accountNumber,submissionRound,industrialRate,submitOutput})
                field.TextChanged+=(s,e)=>SaveSubmissionInfo();
        }

        void InitializeFooterWarning()
        {
            footerWarningPanel=new Panel{BackColor=UiTheme.Page,Name="FooterWarningPanel"};
            footerWarning=new Label{Text=FooterWarningText,Dock=DockStyle.Fill,Padding=new Padding(12,0,12,0),Font=new Font("맑은 고딕",8F),TextAlign=ContentAlignment.MiddleLeft,Name="FooterWarning"};
            footerWarningPanel.Controls.Add(footerWarning);Controls.Add(footerWarningPanel);
            sidebar.Dock=DockStyle.None;sidebar.Anchor=AnchorStyles.Top|AnchorStyles.Bottom|AnchorStyles.Left;
            ClientSizeChanged+=(s,e)=>LayoutFooterWarning();
            LayoutFooterWarning();UpdateFooterWarningTheme();
        }
        void LayoutFooterWarning()
        {
            if(footerWarningPanel==null)return;
            int height=Math.Max(30,TextRenderer.MeasureText(FooterWarningText,footerWarning.Font,new Size(Math.Max(100,ClientSize.Width-24),0),TextFormatFlags.WordBreak).Height+12);
            footerWarningPanel.SetBounds(0,Math.Max(0,ClientSize.Height-height),ClientSize.Width,height);
            sidebar.Height=Math.Max(1,ClientSize.Height-height);
            contentHost.SetBounds(sidebar.Width,0,Math.Max(1,ClientSize.Width-sidebar.Width),Math.Max(1,ClientSize.Height-height));
            footerWarningPanel.BringToFront();
        }
        void UpdateFooterWarningTheme()
        {
            if(footerWarning==null)return;
            footerWarningPanel.BackColor=UiTheme.Page;footerWarning.BackColor=UiTheme.Page;
            footerWarning.ForeColor=UiTheme.Dark?Color.FromArgb(160,191,255):Color.FromArgb(35,83,181);
        }
    }
}
