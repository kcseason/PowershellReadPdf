# 获取当前目录中的所有ZIP文件
$zipFiles = Get-ChildItem -Path . -Filter *.zip
 
# 遍历每个ZIP文件
foreach ($zipFile in $zipFiles) {
    # 解压ZIP文件到同名文件夹
    Write-Host ‘正在解压缩’+$zipFile
    Expand-Archive -LiteralPath $zipFile.FullName -DestinationPath '原始发票和行程单' -Force
    Write-Host ‘解压缩完成’
}

#新建文件夹
$newFolderPath = ".\已整理的发票和行程单\" 
New-Item -Path $newFolderPath  -ItemType Directory -Force
#寻找同一单发票和行程单
$folderPath = ".\原始发票和行程单" 
$excludePrefix = "行程单"
# 获取当前目录中的所有PDF文件
$pdfFiles = Get-ChildItem -Path $folderPath -File | Where-Object {$_.Name -notlike "$excludePrefix*" }
$pdfFiles2 = Get-ChildItem -Path $folderPath -Filter "$excludePrefix*.pdf"
# 遍历每个PDF文件
foreach($pdfFile in $pdfFiles){
    foreach($pdfFile2 in $pdfFiles2){
        if("行程单-"+$pdfFile.BaseName -eq $pdfFile2.BaseName)
        {
            # 读取PDF文件内容
            #$pdfText = Get-PDFText -Path $pdfFile2.BaseName
            # 输出PDF文件内容
            #Write-Host $pdfText

            $pattern = "\d{4}年\d{2}月\d{2}日"
            $matches = [regex]::Matches($pdfFile.BaseName, $pattern)
            $newFolderPath2 = $newFolderPath+$matches[0].Value
            $index = 2
            $tempPath = $newFolderPath2
            while (Test-Path -Path $tempPath -PathType Container) {
                $tempPath = $newFolderPath2+"("+$index+")"
                $index++;
            } 
            if($index -gt 2)
            {
                $newFolderPath2 = $tempPath
            }
            New-Item -Path $newFolderPath2 -ItemType Directory -Force
            Copy-Item -Path $pdfFile.FullName -Destination $newFolderPath2\$pdfFile -Force
            Copy-Item -Path $pdfFile2.FullName -Destination $newFolderPath2\$pdfFile2 -Force
            break 
        }
    }
}

Start-Sleep 5000