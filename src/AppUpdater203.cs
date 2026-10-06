using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace InsurancePayrollValidator
{
    static class AppUpdater
    {
        public const string ManifestUrl="https://raw.githubusercontent.com/isilria/Social-insurance-updates/main/latest.ini";
        public static Version CurrentVersion{get{return Assembly.GetExecutingAssembly().GetName().Version;}}
        public static Version NormalizeVersion(Version v){return new Version(v.Major,v.Minor,Math.Max(0,v.Build),Math.Max(0,v.Revision));}
        public static bool IsNewer(Version latest,Version current){return NormalizeVersion(latest)>NormalizeVersion(current);}
        public sealed class Release
        {
            public Version Version;public Uri Url;public string Sha256,Notes;
        }
        public static Release ParseManifest(string manifest)
        {
            var values=new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase);
            foreach(string raw in manifest.Replace("\r","").Split('\n')){string line=raw.Trim().TrimStart('\uFEFF');int p=line.IndexOf('=');if(p<=0||line.StartsWith("#"))continue;values[line.Substring(0,p).Trim()]=line.Substring(p+1).Trim();}
            string version,url,sha,notes;Version latest;Uri uri;
            if(!values.TryGetValue("version",out version)||!Version.TryParse(version,out latest)||!values.TryGetValue("url",out url)||!Uri.TryCreate(url,UriKind.Absolute,out uri)||uri.Scheme!=Uri.UriSchemeHttps||!values.TryGetValue("sha256",out sha)||!Regex.IsMatch(sha,"\\A[0-9a-fA-F]{64}\\z"))throw new InvalidDataException("업데이트 정보 형식이 올바르지 않습니다.");
            string file=Uri.UnescapeDataString(Path.GetFileName(uri.LocalPath));
            if(!file.EndsWith(".exe",StringComparison.OrdinalIgnoreCase)||file.IndexOfAny(Path.GetInvalidFileNameChars())>=0)throw new InvalidDataException("업데이트 실행파일 이름이 올바르지 않습니다.");
            values.TryGetValue("notes",out notes);return new Release{Version=latest,Url=uri,Sha256=sha,Notes=notes};
        }
        // The production manifest accepts HTTPS only; a local file URI also lets regression tests exercise the exact download/hash code offline.
        public static string DownloadVerified(Uri url,string sha,string folder)
        {
            Directory.CreateDirectory(folder);string path=Path.Combine(folder,"SocialInsurance_download_"+Guid.NewGuid().ToString("N")+".exe");
            try{
                ServicePointManager.SecurityProtocol|=SecurityProtocolType.Tls12;
                using(var client=new WebClient()){client.Headers[HttpRequestHeader.UserAgent]="SocialInsuranceUpdater/"+CurrentVersion;client.DownloadFile(url,path);}
                string actual;using(var hash=SHA256.Create())using(var stream=File.OpenRead(path))actual=BitConverter.ToString(hash.ComputeHash(stream)).Replace("-","");
                if(!String.Equals(actual,sha,StringComparison.OrdinalIgnoreCase))throw new InvalidDataException("내려받은 업데이트 파일의 무결성 확인에 실패했습니다.");
                using(var stream=File.OpenRead(path)){if(stream.Length<64||stream.ReadByte()!=77||stream.ReadByte()!=90)throw new InvalidDataException("내려받은 파일이 Windows 실행파일이 아닙니다.");}
                return path;
            }catch{if(File.Exists(path))File.Delete(path);throw;}
        }
        public static string StageVerified(string downloaded,string fileName,Version version,string currentExe,string fallbackFolder)
        {
            string name=Path.GetFileName(fileName),folder=Path.GetDirectoryName(Path.GetFullPath(currentExe));
            if(String.Equals(Path.Combine(folder,name),Path.GetFullPath(currentExe),StringComparison.OrdinalIgnoreCase))name=Path.GetFileNameWithoutExtension(name)+"_"+NormalizeVersion(version).ToString()+".exe";
            string target=Path.Combine(folder,name);
            try{File.Copy(downloaded,target,true);return target;}
            catch(IOException){}catch(UnauthorizedAccessException){}
            Directory.CreateDirectory(fallbackFolder);
            target=Path.Combine(fallbackFolder,Path.GetFileNameWithoutExtension(name)+"_"+Guid.NewGuid().ToString("N")+".exe");
            File.Copy(downloaded,target,false);return target;
        }
        public static ProcessStartInfo LaunchInfo(string target){return new ProcessStartInfo(Path.GetFullPath(target)){WorkingDirectory=Path.GetDirectoryName(Path.GetFullPath(target)),UseShellExecute=true};}
        public static string CheckAndInstall(IWin32Window owner,Version current,bool interactive)
        {
            string downloaded=null;
            try{
                ServicePointManager.SecurityProtocol|=SecurityProtocolType.Tls12;
                string manifest;using(var client=new WebClient()){client.Headers[HttpRequestHeader.UserAgent]="SocialInsuranceUpdater/"+current;client.Encoding=Encoding.UTF8;manifest=client.DownloadString(ManifestUrl);}
                Release release=ParseManifest(manifest);
                if(!IsNewer(release.Version,current))return "최신 여부  최신 버전입니다.";
                string available="최신 여부  Ver. "+release.Version+" 업데이트 가능";
                if(!interactive)return available;
                if(MessageBox.Show(owner,"새 버전 Ver. "+release.Version+"이 있습니다."+(String.IsNullOrWhiteSpace(release.Notes)?"":"\r\n\r\n"+release.Notes)+"\r\n\r\n지금 내려받아 실행할까요?","업데이트 확인",MessageBoxButtons.YesNo,MessageBoxIcon.Information)!=DialogResult.Yes)return available;
                downloaded=DownloadVerified(release.Url,release.Sha256,Path.GetTempPath());
                string target=StageVerified(downloaded,Uri.UnescapeDataString(Path.GetFileName(release.Url.LocalPath)),release.Version,Application.ExecutablePath,Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),"Downloads"));
                using(Process started=Process.Start(LaunchInfo(target))){if(started==null)throw new IOException("새 버전 실행을 시작하지 못했습니다. 다시 시도해 주세요.");}
                Application.Exit();return "업데이트 실행 중";
            }catch(Exception ex){if(interactive)MessageBox.Show(owner,"업데이트를 완료하지 못했습니다.\r\n\r\n"+ex.Message,"업데이트 확인",MessageBoxButtons.OK,MessageBoxIcon.Warning);return "최신 여부  확인 실패";}
            finally{if(downloaded!=null&&File.Exists(downloaded))try{File.Delete(downloaded);}catch{}}
        }
    }
}
