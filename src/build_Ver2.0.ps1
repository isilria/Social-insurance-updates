$ErrorActionPreference = 'Stop'
$sourceRoot = $PSScriptRoot
$source = Join-Path $sourceRoot 'InsurancePayrollValidator_Ver2.0.cs'
$epplus = Join-Path $sourceRoot 'dependencies\EPPlus.dll'
$teacher = Join-Path $sourceRoot 'templates\teacher_template.xls'
$worker = Join-Path $sourceRoot 'templates\worker_template.xlsx'
$validation = Join-Path $sourceRoot 'templates\validation_template_distribution.xlsx'
$icon = Join-Path $sourceRoot 'assets\ui_reference_app_icon.ico'
$referenceIcon = Join-Path $sourceRoot 'assets\ui_reference_icon_transparent.png'
$buildFolder = Join-Path $sourceRoot 'build'
$outputName = '사회보험_재원별_대사_보조_도우미_Ver2.0.2.exe'
$output = Join-Path $buildFolder $outputName
$compiler = 'C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe'

New-Item -ItemType Directory -Path $buildFolder -Force | Out-Null

& $compiler /nologo /target:winexe /platform:anycpu /optimize+ /win32icon:$icon /out:$output `
    /r:System.dll /r:System.Core.dll /r:System.Drawing.dll /r:System.Windows.Forms.dll `
    /r:System.IO.Compression.dll /r:System.IO.Compression.FileSystem.dll /r:Microsoft.CSharp.dll `
    /r:$epplus `
    /resource:"$epplus,InsurancePayrollValidator.EPPlus.dll" `
    /resource:"$teacher,InsurancePayrollValidator.TeacherTemplate.xls" `
    /resource:"$worker,InsurancePayrollValidator.WorkerTemplate.xlsx" `
    /resource:"$validation,InsurancePayrollValidator.ValidationTemplate.xlsx" `
    /resource:"$icon,InsurancePayrollValidator.AppIcon.ico" `
    /resource:"$referenceIcon,InsurancePayrollValidator.ReferenceIcon.png" `
    $source (Join-Path $sourceRoot 'TestFeatures202.cs') (Join-Path $sourceRoot 'ManualContributions202.cs') (Join-Path $sourceRoot 'PrintQuality202.cs')

if ($LASTEXITCODE -ne 0) { throw "Build failed: $LASTEXITCODE" }
Write-Host "Completed: $output"
