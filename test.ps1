param([string]$Variant='src',[string]$HealthFixture='',[string]$PensionFixture='')
$ErrorActionPreference='Stop'
$root=Join-Path $PSScriptRoot $Variant
$out=Join-Path $PSScriptRoot ('qa-'+$Variant)
New-Item -ItemType Directory -Path $out -Force | Out-Null
$ep=Join-Path $root 'dependencies/EPPlus.dll'
Copy-Item -LiteralPath $ep -Destination $out -Force
$csc='C:/Windows/Microsoft.NET/Framework64/v4.0.30319/csc.exe'
$argsList=@('/nologo','/target:exe','/main:Regression',('/out:'+(Join-Path $out 'Regression.exe')),'/r:System.dll','/r:System.Core.dll','/r:System.Drawing.dll','/r:System.Windows.Forms.dll','/r:System.IO.Compression.dll','/r:System.IO.Compression.FileSystem.dll','/r:Microsoft.CSharp.dll',('/r:'+$ep))
$resources=@{'templates/teacher_template.xls'='TeacherTemplate.xls';'templates/worker_template.xlsx'='WorkerTemplate.xlsx';'templates/validation_template_distribution.xlsx'='ValidationTemplate.xlsx';'assets/ui_reference_app_icon.ico'='AppIcon.ico';'assets/ui_reference_icon_transparent.png'='ReferenceIcon.png'}
foreach($key in $resources.Keys){$argsList+=('/resource:'+(Join-Path $root $key)+',InsurancePayrollValidator.'+$resources[$key])}
foreach($file in @('InsurancePayrollValidator_Ver2.0.cs','TestFeatures202.cs','ManualContributions202.cs','PrintQuality202.cs')){$argsList+=(Join-Path $root $file)}
$argsList+=(Join-Path $PSScriptRoot 'Regression.cs')
& $csc @argsList
if($LASTEXITCODE -ne 0){throw 'test compile failed'}
$runArgs=@($out)
if($HealthFixture){$runArgs+=$HealthFixture}
if($PensionFixture){if(!$HealthFixture){throw 'HealthFixture required with PensionFixture'};$runArgs+=$PensionFixture}
if($Variant -eq 'baseline'){$runArgs+='--baseline'}
& (Join-Path $out 'Regression.exe') @runArgs
if($LASTEXITCODE -ne 0){throw 'regression failed'}
