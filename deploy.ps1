# Arbitraj Botu — Otomatik Deploy Scripti
# Bu script yereldeki degisiklikleri GitHub'a push eder ve sunucuda botu gunceller.

$SERVER_IP = "64.227.124.74"
$SSH_KEY = "$env:USERPROFILE\.ssh\arbitraj_bot"
$REPO_PATH = "/root/bot"

Write-Host "--- [1/3] Yerel degisiklikler GitHub'a gonderiliyor... ---" -ForegroundColor Cyan
git add .
$status = git status --porcelain
if ($status) {
    git commit -m "Auto-deploy: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    git push origin main
} else {
    Write-Host "Degisiklik yok, push atlandi." -ForegroundColor Yellow
}

Write-Host "`n--- [2/3] Sunucuya baglaniliyor ve kod guncelleniyor... ---" -ForegroundColor Cyan
ssh -i $SSH_KEY -o StrictHostKeyChecking=no root@$SERVER_IP "cd $REPO_PATH && git pull origin main"

Write-Host "`n--- [3/3] Bot yeniden baslatiliyor... ---" -ForegroundColor Cyan
ssh -i $SSH_KEY -o StrictHostKeyChecking=no root@$SERVER_IP "screen -S arbitrajbot -X quit; screen -dmS arbitrajbot bash -c 'cd $REPO_PATH && python3 watchdog.py 2>&1 | tee $REPO_PATH/watchdog.log'"

Write-Host "`n--- [4/4] Sunucudaki Guncel Excel Yerel Klasore Senkronize Ediliyor... ---" -ForegroundColor Cyan
scp -i $SSH_KEY -o StrictHostKeyChecking=no "root@${SERVER_IP}:${REPO_PATH}/Finansal_Radar_*.xlsx" "$PSScriptRoot\"

Write-Host "`n✅ ISLEM TAMAMLANDI VE EXCEL GUNCEL!" -ForegroundColor Green
