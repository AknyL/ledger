# ============================================================
#  记账助手 · GitHub Pages 一键部署脚本
#  作用：把本文件夹推送到 GitHub，并自动开启 Pages 静态托管，
#        得到一个公网链接，iPhone 用 Safari 打开后「添加到主屏幕」即可。
#  用法：在 PowerShell 里执行
#        powershell -ExecutionPolicy Bypass -File deploy-github.ps1
# ============================================================
$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $here

Write-Host ""
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host "  记账助手 · GitHub Pages 一键部署" -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host ""

# ---------- 1. 检查 git ----------
if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
    Write-Host "[错误] 未检测到 git。" -ForegroundColor Red
    Write-Host "请先安装 Git：https://git-scm.com/download/win" -ForegroundColor Yellow
    Write-Host "安装时一路 Next 即可，装完重新打开本窗口再运行。" -ForegroundColor Yellow
    exit 1
}

# ---------- 2. 检查 gh (GitHub CLI，可选) ----------
$gh = Get-Command gh -ErrorAction SilentlyContinue
if ($gh) {
    Write-Host "[√] 已检测到 git 和 gh（GitHub CLI）" -ForegroundColor Green
} else {
    Write-Host "[√] 已检测到 git（未安装 gh，走手动创建仓库流程，脚本会打印步骤）" -ForegroundColor Green
}
Write-Host ""

# ---------- 3. 询问信息 ----------
$username = (Read-Host "GitHub 用户名").Trim()
if ([string]::IsNullOrWhiteSpace($username)) {
    Write-Host "[错误] 用户名不能为空。" -ForegroundColor Red
    exit 1
}
$reponame = (Read-Host "仓库名（直接回车用 ledger）").Trim()
if ([string]::IsNullOrWhiteSpace($reponame)) { $reponame = "ledger" }

# ---------- 4. 初始化 git 仓库（若尚未初始化） ----------
if (-not (Test-Path ".git")) {
    git init | Out-Null
    Write-Host "[√] 已初始化 git 仓库" -ForegroundColor Green
}
git add .
git commit -m "记账助手" 2>$null | Out-Null
if ($LASTEXITCODE -ne 0) {
    # 没有变化也会返回非零，忽略
}
Write-Host "[√] 已提交本地文件" -ForegroundColor Green
git branch -M main

# ---------- 5. 创建远程仓库并推送 ----------
if ($gh) {
    Write-Host ""
    Write-Host ">> 正在用 gh 登录并创建仓库…" -ForegroundColor Cyan

    $authOk = $false
    try {
        gh auth status | Out-Null
        if ($LASTEXITCODE -eq 0) { $authOk = $true }
    } catch { $authOk = $false }

    if (-not $authOk) {
        Write-Host "首次使用需要登录 GitHub（浏览器会弹出授权）：" -ForegroundColor Yellow
        gh auth login --hostname github.com --git-protocol https --web
    }

    # 创建仓库（已存在则跳过报错）
    gh repo create "$username/$reponame" --public --source . --push 2>$null
    if ($LASTEXITCODE -ne 0) {
        Write-Host "仓库可能已存在，尝试直接设置远程并推送…" -ForegroundColor Yellow
        git remote remove origin 2>$null
        git remote add origin "https://github.com/$username/$reponame.git"
        git push -u origin main
    } else {
        Write-Host "[√] 已推送代码到 GitHub" -ForegroundColor Green
    }

    # 开启 Pages
    Write-Host ">> 正在开启 GitHub Pages…" -ForegroundColor Cyan
    gh api --method POST "repos/$username/$reponame/pages" `
        -f "source[branch]=main" -f "source[path]=/" | Out-Null
    Write-Host "[√] Pages 已开启" -ForegroundColor Green
}
else {
    Write-Host ""
    Write-Host ">> 未安装 gh，请手动完成下面三步（约 1 分钟）：" -ForegroundColor Cyan
    Write-Host "  1) 打开 https://github.com/new 新建仓库，名字填：$reponame（Public 公开）"
    Write-Host "     创建时【不要】勾选 'Add a README'，然后点 Create repository。"
    Write-Host "  2) 回到本窗口，回车后会自动帮你推送代码："
    Read-Host "     按回车继续推送"
    git remote remove origin 2>$null
    git remote add origin "https://github.com/$username/$reponame.git"
    git push -u origin main
    Write-Host "[√] 已推送代码到 GitHub" -ForegroundColor Green
    Write-Host "  3) 打开仓库页面 → Settings → Pages → Source 选 main /(root) → Save"
    Write-Host "     （新版界面在 Settings → Pages → Deploy from a branch → 选 main 和 /(root)）"
}

# ---------- 6. 打印结果 ----------
$url = "https://$username.github.io/$reponame/"
Write-Host ""
Write-Host "==================================================" -ForegroundColor Green
Write-Host "  部署完成！你的公网地址：" -ForegroundColor Green
Write-Host "  $url" -ForegroundColor Yellow
Write-Host "==================================================" -ForegroundColor Green
Write-Host ""
Write-Host "  iPhone 使用：Safari 打开上面地址 → 分享按钮 → 添加到主屏幕。" -ForegroundColor Cyan
Write-Host "  首次部署后 Pages 需 1-2 分钟生效，若打不开稍等再刷新。" -ForegroundColor Cyan
Write-Host ""
