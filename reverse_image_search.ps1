#requires -Version 5.1
# 以图识图工具 — 反向图片搜索桌面应用 (纯 PowerShell，零安装依赖)
# 用户上传图片，并行向多家反向图片搜索引擎提交，把来源链接列出并可在电脑浏览器中打开。

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls
Add-Type -AssemblyName System.Windows.Forms

[System.Windows.Forms.Application]::EnableVisualStyles()

$script:UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
$script:SearchResults = [System.Collections.Generic.List[PSCustomObject]]::new()
$script:ImagePath = ""

$Engines = @(
    [PSCustomObject]@{ Name="TinEye";         Url="https://www.tineye.com/search";                 Field="upload";                  Max=10; Exclude=@("tineye.com") }
    [PSCustomObject]@{ Name="Yandex";         Url="https://yandex.com/images/search";              Field="upfile";                  Max=10; Exclude=@("yandex.com","yandex.ru") }
    [PSCustomObject]@{ Name="Bing";           Url="https://www.bing.com/images/searchbyimage/upload";Field="file";                   Max=10; Exclude=@("bing.com","microsoft.com","msn.com") }
    [PSCustomObject]@{ Name="Google";         Url="https://www.google.com/searchbyimage/upload";    Field="encoded_image";           Max=10; Exclude=@("google.com","googleusercontent.com","gstatic.com","googleapis.com") }
    [PSCustomObject]@{ Name="Baidu";          Url="https://image.baidu.com/";                      Field="file";                    Max=8;  Exclude=@("baidu.com") }
    [PSCustomObject]@{ Name="Sogou";          Url="https://pic.sogou.com/ris";                     Field="file";                    Max=8;  Exclude=@("sogou.com") }
    [PSCustomObject]@{ Name="So";             Url="https://st.so.com/";                            Field="file";                    Max=8;  Exclude=@("so.com","qihoo.com") }
)

function Get-ImageContentType {
    param($Path)
    $ext = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
    switch ($ext) {
        ".jpg"   { return "image/jpeg" }
        ".jpeg"  { return "image/jpeg" }
        ".png"   { return "image/png" }
        ".gif"   { return "image/gif" }
        ".webp"  { return "image/webp" }
        ".bmp"   { return "image/bmp" }
        default  { return "application/octet-stream" }
    }
}

function Build-Multipart {
    param($TextFields, $FileField, $FilePath, $FileName)
    $boundary = [System.Guid]::NewGuid().ToString("N")
    $ct = "multipart/form-data; boundary=$boundary"
    $enc = [System.Text.Encoding]::UTF8
    $ms = [System.IO.MemoryStream]::new()
    $writer = New-Object System.IO.StreamWriter($ms, $enc)
    try {
        foreach ($key in $TextFields.Keys) {
            $writer.Write("--$boundary`r`n")
            $writer.Write("Content-Disposition: form-data; name=`"$key`"`r`n`r`n")
            $writer.Write("$($TextFields[$key])`r`n")
        }
        $writer.Write("--$boundary`r`n")
        $writer.Write("Content-Disposition: form-data; name=`"$FileField`"; filename=`"$FileName`"`r`n")
        $writer.Write("Content-Type: $(Get-ImageContentType $FilePath)`r`n`r`n")
        $writer.Flush()
        $fs = [System.IO.File]::OpenRead($FilePath)
        try { $fs.CopyTo($ms) } finally { $fs.Close() }
        $nl = $enc.GetBytes("`r`n")
        $ms.Write($nl, 0, $nl.Count)
        $writer.Write("--$boundary--`r`n")
        $writer.Flush()
    } finally {
        $writer.Dispose()
    }
    return @{ Bytes=$ms.ToArray(); ContentType=$ct }
}

function Resolve-Url {
    param($Url, $Base)
    if ($Url -match '^https?://') { return $Url }
    if ($Url -match '^//') {
        $u = [uri]$Base
        return "$($u.Scheme):$Url"
    }
    if ($Url.StartsWith("/")) {
        $u = [uri]$Base
        return "$($u.Scheme)://$($u.Host)$Url"
    }
    $base = $Base
    if ($base -match '/[^/]*$') { $base = $base -replace '/[^/]*$', '' }
    return "$base/$Url"
}

function Strip-Tags {
    param($s)
    if ([string]::IsNullOrEmpty($s)) { return "" }
    $s = $s -replace '<[^>]+>', ' '
    $s = $s -replace '&nbsp;', ' '
    $s = $s -replace '&amp;', '&'
    $s = $s -replace '&lt;', '<'
    $s = $s -replace '&gt;', '>'
    $s = $s -replace '&quot;', '"'
    $s = $s -replace '\s+', ' '
    $s = $s -replace '^\s+|\s+$', ''
    return $s
}

function Get-EngineLinks {
    param($Engine, $ImagePath)
    $out = [System.Collections.Generic.List[PSCustomObject]]::new()
    try {
        $fileName = [System.IO.Path]::GetFileName($ImagePath)
        $mp = Build-Multipart @{} $Engine.Field $ImagePath $fileName
        $req = @{
            Uri            = $Engine.Url
            Method         = "POST"
            Body           = $mp.Bytes
            ContentType    = $mp.ContentType
            UserAgent      = $script:UA
            MaximumRedirection = 15
            TimeoutSec     = 28
            UseBasicParsing  = $true
            ErrorAction      = "SilentlyContinue"
        }
        $resp = Invoke-WebRequest @req
        $html = if ($resp) { $resp.Content } else { "" }
        if ($html.Length -gt 2000000) { $html = $html.Substring(0, 2000000) }
        $baseUrl = if ($resp -and $resp.ResponseUri) { $resp.ResponseUri.AbsoluteUri } else { $Engine.Url }

        $matches = [regex]::Matches($html, "(?is)<a\s+href\s*=\s*[`"`'](?<url>[^`"`']+)[`"`']")
        $seen = New-Object System.Collections.Generic.HashSet[string]
        $count = 0
        foreach ($m in $matches) {
            $url = $m.Groups['url'].Value
            $text = Strip-Tags $m.Value
            if ([string]::IsNullOrWhiteSpace($url)) { continue }
            $abs = Resolve-Url $url $baseUrl
            if ($abs -notmatch '^https?://') { continue }
            try { $host = ([uri]$abs).Host } catch { continue }
            $skip = $false
            foreach ($d in $Engine.Exclude) {
                if ($host -eq $d -or $host.EndsWith(".$d")) { $skip = $true; break }
            }
            if ($skip) { continue }
            if ($seen.Contains($abs)) { continue }
            $seen.Add($abs) | Out-Null
            if ($abs -match '\.(css|js|png|jpe?g|gif|svg|ico|woff2?|ttf|eot)(\?|$)') { continue }
            $title = $abs
            if ($title.Length -gt 140) { $title = $title.Substring(0, 140) + "..." }
            $out.Add([PSCustomObject]@{ Engine=$Engine.Name; Title=$title; Url=$abs })
            $count++
            if ($count -ge $Engine.Max) { break }
        }
    } catch {
        $errMsg = $_.Exception.Message
        $out.Add([PSCustomObject]@{ Engine=$Engine.Name; Title="错误：$errMsg"; Url="" })
    }
    return $out
}

# ---------------- UI ----------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "以图识图工具 — 反向图片搜索"
$form.Size = New-Object System.Drawing.Size(1050, 720)
$form.MinimumSize = New-Object System.Drawing.Size(820, 540)
$form.StartPosition = "CenterScreen"

$pnlTop = New-Object System.Windows.Forms.Panel
$pnlTop.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlTop.Height = 64
$pnlTop.Padding = New-Object System.Windows.Forms.Padding(10,8,10,8)

$btnChoose = New-Object System.Windows.Forms.Button
$btnChoose.Text = "选择图片"
$btnChoose.Location = New-Object System.Drawing.Point(0,14)
$btnChoose.Size = New-Object System.Drawing.Size(115,36)
$btnChoose.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left

$lblFile = New-Object System.Windows.Forms.Label
$lblFile.Location = New-Object System.Drawing.Point(130,20)
$lblFile.Size = New-Object System.Drawing.Size(460,24)
$lblFile.Text = "尚未选择图片"
$lblFile.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

$pnlTop.Controls.AddRange(@($btnChoose,$lblFile))

$pic = New-Object System.Windows.Forms.PictureBox
$pic.Location = New-Object System.Drawing.Point(12,82)
$pic.Size = New-Object System.Drawing.Size(280,280)
$pic.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
$pic.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$pic.BackColor = [System.Drawing.Color]::White

$lblHint = New-Object System.Windows.Forms.Label
$lblHint.Location = New-Object System.Drawing.Point(12,370)
$lblHint.Size = New-Object System.Drawing.Size(280,40)
$lblHint.Text = "图片预览区"
$lblHint.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter

$pnlMid = New-Object System.Windows.Forms.Panel
$pnlMid.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlMid.Height = 60
$pnlMid.Padding = New-Object System.Windows.Forms.Padding(10,6,10,6)

$btnSearch = New-Object System.Windows.Forms.Button
$btnSearch.Text = "开始搜索（多引擎并行）"
$btnSearch.Location = New-Object System.Drawing.Point(0,12)
$btnSearch.Size = New-Object System.Drawing.Size(170,38)
$btnSearch.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left

$btnOpenSel = New-Object System.Windows.Forms.Button
$btnOpenSel.Text = "浏览器打开选中项"
$btnOpenSel.Location = New-Object System.Drawing.Point(182,12)
$btnOpenSel.Size = New-Object System.Drawing.Size(150,38)
$btnOpenSel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left

$btnOpenAll = New-Object System.Windows.Forms.Button
$btnOpenAll.Text = "全部在浏览器打开"
$btnOpenAll.Location = New-Object System.Drawing.Point(344,12)
$btnOpenAll.Size = New-Object System.Drawing.Size(140,38)
$btnOpenAll.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left

$btnCopy = New-Object System.Windows.Forms.Button
$btnCopy.Text = "复制选中链接"
$btnCopy.Location = New-Object System.Drawing.Point(496,12)
$btnCopy.Size = New-Object System.Drawing.Size(130,38)
$btnCopy.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left

$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text = "清空结果"
$btnClear.Location = New-Object System.Drawing.Point(638,12)
$btnClear.Size = New-Object System.Drawing.Size(100,38)
$btnClear.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Location = New-Object System.Drawing.Point(750,18)
$lblStatus.Size = New-Object System.Drawing.Size(280,24)
$lblStatus.Text = "就绪"
$lblStatus.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

$pnlMid.Controls.AddRange(@($btnSearch,$btnOpenSel,$btnOpenAll,$btnCopy,$btnClear,$lblStatus))

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Dock = [System.Windows.Forms.DockStyle]::Top
$progress.Height = 14
$progress.Visible = $false

$lst = New-Object System.Windows.Forms.ListView
$lst.Dock = [System.Windows.Forms.DockStyle]::Fill
$lst.View = [System.Windows.Forms.View]::Details
$lst.FullRowSelect = $true
$lst.GridLines = $true
$lst.Columns.AddRange(@(
    (New-Object System.Windows.Forms.ColumnHeader -Property @{ Text="引擎"; Width=100 }),
    (New-Object System.Windows.Forms.ColumnHeader -Property @{ Text="标题"; Width=340 }),
    (New-Object System.Windows.Forms.ColumnHeader -Property @{ Text="链接"; Width=560 })
))

$log = New-Object System.Windows.Forms.RichTextBox
$log.Dock = [System.Windows.Forms.DockStyle]::Fill
$log.ReadOnly = $true
$log.Font = New-Object System.Drawing.Font("Consolas",9.0,[System.Drawing.FontStyle]::Regular)
$log.BackColor = [System.Drawing.Color]::FromArgb(248,248,248)

function Write-Log {
    param($s)
    $log.AppendText("[$(Get-Date -Format 'HH:mm:ss')] $s`r`n")
    $log.ScrollToCaret()
}

$split = New-Object System.Windows.Forms.SplitContainer
$split.Dock = [System.Windows.Forms.DockStyle]::Fill
$split.Orientation = [System.Windows.Forms.Orientation]::Vertical
$split.SplitterDistance = 400
$split.Panel1.Controls.Add($lst)
$split.Panel2.Controls.Add($log)

$form.Controls.AddRange(@($pnlTop,$pic,$lblHint,$pnlMid,$progress,$split))

# ---------------- Events ----------------
$btnChoose.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "图片文件|*.jpg;*.jpeg;*.png;*.gif;*.bmp;*.webp|所有文件|*.*"
    $dlg.Title = "选择要识图的图片"
    $dlg.Multiselect = $false
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $script:ImagePath = $dlg.FileName
        $lblFile.Text = [System.IO.Path]::GetFileName($script:ImagePath)
        try { $pic.Image = [System.Drawing.Image]::FromFile($script:ImagePath) } catch {}
        $lblHint.Text = "已加载：$([System.IO.Path]::GetFileName($script:ImagePath))"
        Write-Log ("已选择图片：{0}" -f $script:ImagePath)
    }
})

$openUrl = {
    param($url)
    if ([string]::IsNullOrWhiteSpace($url)) { return }
    try { [System.Diagnostics.Process]::Start($url) | Out-Null } catch {
        [System.Windows.Forms.MessageBox]::Show("无法在浏览器中打开链接：`n`n$url`n`n$($_.Exception.Message)", "打开失败", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
}

$btnSearch.Add_Click({
    if ([string]::IsNullOrWhiteSpace($script:ImagePath)) {
        [System.Windows.Forms.MessageBox]::Show("请先选择一张图片。", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    $btnSearch.Enabled = $false
    $lst.Items.Clear()
    $script:SearchResults.Clear()
    $progress.Visible = $true
    $progress.Value = 0
    $lblStatus.Text = "正在搜索..."
    Write-Log "开始多引擎反向图片搜索（共 $($Engines.Count) 家）"
    $bw.RunWorkerAsync()
})

$btnOpenSel.Add_Click({
    $items = $lst.SelectedItems
    if ($items.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("请先在结果列表中选中一条。", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    foreach ($it in $items) { &$openUrl $it.Tag }
})

$btnOpenAll.Add_Click({
    if ($lst.Items.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("暂无结果可打开。", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    $confirm = [System.Windows.Forms.MessageBox]::Show("将在浏览器中打开全部 $($lst.Items.Count) 条链接，继续？", "确认", [System.Windows.Forms.MessageBoxButtons]::OKCancel, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirm -eq [System.Windows.Forms.DialogResult]::OK) {
        foreach ($it in $lst.Items) { &$openUrl $it.Tag }
    }
})

$btnCopy.Add_Click({
    $it = $lst.SelectedItems[0]
    if (-not $it) {
        [System.Windows.Forms.MessageBox]::Show("请先选中一条结果。", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    try {
        [System.Windows.Forms.Clipboard]::SetText($it.Tag)
        $lblStatus.Text = "链接已复制到剪贴板"
        Write-Log ("已复制：{0}" -f $it.Tag)
    } catch {
        [System.Windows.Forms.MessageBox]::Show("复制失败：$($_.Exception.Message)", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

$btnClear.Add_Click({
    $lst.Items.Clear()
    $script:SearchResults.Clear()
    $lblStatus.Text = "已清空"
    Write-Log "结果已清空"
})

$lst.Add_DoubleClick({
    $it = $lst.SelectedItems[0]
    if ($it) { &$openUrl $it.Tag }
})

$bw = New-Object System.ComponentModel.BackgroundWorker
$bw.WorkerReportsProgress = $true

$bw.Add_DoWork({
    $script:SearchResults.Clear()
    $total = $Engines.Count
    $i = 0
    foreach ($e in $Engines) {
        $i++
        $pct = [math]::Round(($i - 1) / $total * 100)
        $bw.ReportProgress($pct, ("正在搜索 {0} ..." -f $e.Name))
        $links = Get-EngineLinks $e $script:ImagePath
        foreach ($l in $links) { $script:SearchResults.Add($l) }
    }
    $bw.ReportProgress(100, "搜索完成")
})

$bw.Add_ProgressChanged({
    param($sender,$e)
    $lblStatus.Text = [string]$e.UserState
    $progress.Value = $e.ProgressPercentage
    if (-not $progress.Visible) { $progress.Visible = $true }
    if ($e.UserState -is [string] -and $e.UserState -like "正在搜索*") {
        Write-Log $e.UserState
    }
})

$bw.Add_RunWorkerCompleted({
    param($sender,$e)
    $btnSearch.Enabled = $true
    $progress.Visible = $false
    $lblStatus.Text = ("共找到 {0} 条结果" -f $script:SearchResults.Count)
    $lst.Items.Clear()
    foreach ($r in $script:SearchResults) {
        $it = New-Object System.Windows.Forms.ListViewItem(@($r.Engine,$r.Title,$r.Url))
        $it.Tag = $r.Url
        $lst.Items.Add($it)
    }
    $lst.Columns[0].Width = 100
    $lst.Columns[1].Width = 340
    $lst.Columns[2].Width = 560
    Write-Log ("搜索结束，共 {0} 条结果" -f $script:SearchResults.Count)
})

[System.Windows.Forms.Application]::Run($form)
