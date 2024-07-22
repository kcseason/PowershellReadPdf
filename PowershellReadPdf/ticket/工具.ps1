#获取当前路径
$scriptPath = Split-Path $MyInvocation.MyCommand.Path -Parent
#解压缩文件夹路径
$extractFolder =Join-Path $scriptPath "1.原始发票和行程单"

#读取PDF方法
function convert-PDFToText
{
    param([Parameter(Mandatory=$true)][string]$file);
           
    logOutPut("开始加载DLL")
    Add-Type -Path "C:\Users\30908\Desktop\发票\BouncyCastle.Crypto.dll"
    Add-Type -Path "C:\Users\30908\Desktop\发票\itextsharp.dll"
    logOutPut("开始读取PDF")
    $pdf = New-Object iTextSharp.text.pdf.pdfreader -ArgumentList $file
    $text = ""
    for($page=1;$page -le $pdf.NumberOfPages;$page++)
    {
        $text = $text+[iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($pdf,$page)
    }
    $pdf.Close()
    logOutPut("读取PDF完毕")
    return $text
}

# 记录日志方法
function logOutPut()
{
    param([Parameter(Mandatory=$false)][string]$logMsg)

    $now = Get-Date
    $logInfo = $now.ToString('yyyy-MM-dd HH:mm:ss.fff')+' '
    Write-Host $logInfo$logMsg
}

# 获取行程日期
function getTrafficDate()
{   
    param([Parameter(Mandatory=$false)][string]$pdfText)

    $arrayPattern = New-Object System.Collections.ArrayList
    # 行程时间
    $arrayPattern.Add("(?<=\u884c\u7a0b\u65f6\u95f4：)\d{4}-\d{2}-\d{2}")
    $arrayPattern.Add("(?<=\u884c\u7a0b\u8d77\u6b62\u65e5\u671f：)\d{4}-\d{2}-\d{2}")

    foreach($pattern in $arrayPattern)
    {
        logOutPut("开始读取日期")
        $match = [regex]::Match($pdfText,$pattern)
        logOutPut($match.Value)
        if($match.Value -ne $null -and $match.Value -ne "")
        {
            return $match.Value.ToString()
        }
    }
    logOutPut("无法识别行程日期")
    return "无法识别行程日期"
}

# 开始处理
logOutPut("开始处理")
# 获取当前目录中的所有ZIP文件
logOutPut("开始解压缩文件......")
$zipFiles = Get-ChildItem -Path $scriptPath -Filter *.zip 
# 遍历每个ZIP文件
foreach ($zipFile in $zipFiles) {
    # 解压ZIP文件到同名文件夹
    logOutPut("正在解压缩"+$zipFile)
    Expand-Archive -LiteralPath $zipFile.FullName -DestinationPath $extractFolder -Force
    logOutPut("解压缩完成")
}

#新建处理完文件夹
$finishedFolder = Join-Path $scriptPath "2.已整理的发票和行程单"
$undefinedFolder = Join-Path $scriptPath "3.无法识别行程日期"
logOutPut("新建整理好文件夹："+$finishedFolder)
New-Item -Path $finishedFolder -ItemType Directory -Force
#寻找同一单发票和行程单
$excludePrefix = "行程单"
# 获取当前目录中的所有PDF文件
$pdfFiles = Get-ChildItem -Path $extractFolder -File | Where-Object {$_.Name -notlike "$excludePrefix*" }
$pdfFiles2 = Get-ChildItem -Path $extractFolder -Filter "$excludePrefix*.pdf"
# 遍历每个PDF文件
foreach($pdfFile in $pdfFiles){
    $copy = $false
    foreach($pdfFile2 in $pdfFiles2){
        if($excludePrefix+"-"+$pdfFile.BaseName -eq $pdfFile2.BaseName)
        {
            logOutPut("读取"+$pdfFile2.BaseName+"行程时间")
            # 读取PDF文件内容
            $pdfText = convert-PDFToText $pdfFile2.FullName

            $folderDate = getTrafficDate $pdfText
            $folderDate = $folderDate[2]

            $finishedFolderSub = Join-Path $finishedFolder $folderDate
            $index =2 
            $tempPath = $finishedFolderSub
            while(Test-Path -Path $tempPath -PathType Container)
            {
                $tempPath = $finishedFolderSub+"("+$index+")" 
                $index++   
            }
            if($index -ge 2)
            {
                $finishedFolderSub = $tempPath
            }

            logOutPut("新建行程日期文件夹"+$folderDate)
            New-Item -Path $finishedFolderSub -ItemType Directory -Force
            logOutPut("复制发票："+$pdfFile.BaseName)
            Copy-Item -Path $pdfFile.FullName -Destination $finishedFolderSub\$pdfFile -Force
            logOutPut("复制行程单："+$pdfFile.BaseName)
            Copy-Item -Path $pdfFile2.FullName -Destination $finishedFolderSub\$pdfFile2 -Force
            logOutPut("复制完毕")
            $copy = $true
            break 
        }
    }
    # 没有具体行程日期
    if($copy -eq $false)
    {
        if((Test-Path -Path $undefinedFolder) -eq $false)
        {
            New-Item -Path $undefinedFolder -ItemType Directory -Force
        }
        logOutPut("复制发票："+$pdfFile.BaseName)
        Copy-Item -Path $pdfFile.FullName -Destination $undefinedFolder\$pdfFile -Force
        $copy = $true
        logOutPut("复制完毕")
    }
}

Start-Sleep 5000