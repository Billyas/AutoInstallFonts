# 批量安装TTF字体的PowerShell脚本（需管理员权限）
$FontsFolder = (Get-Location).Path  # 使用当前目录


$SystemFontsPath = "$env:Windir\Fonts"

$Shell = New-Object -ComObject Shell.Application
$FontsNamespace = $Shell.Namespace(0x14)  # 0x14对应Fonts文件夹

Get-ChildItem -Path $FontsFolder -Filter *.ttf -Recurse | ForEach-Object {
    try {
        $FontFile = $_.FullName
       
        Copy-Item -Path $FontFile -Destination $SystemFontsPath -Force
      
        $FontsNamespace.CopyHere($FontFile, 0x14)
        Write-Host "已安装字体: $($_.Name)" -ForegroundColor Green
    } catch {
        Write-Host "安装失败: $($_.Name) ($($_.Exception.Message))" -ForegroundColor Red
    }
}
