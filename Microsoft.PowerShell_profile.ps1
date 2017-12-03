function Prompt
{
    $date = "`n" + $(Get-Date)
    Write-Host $date -ForegroundColor Blue
    $computer = "User::" + $env:ComputerName
    Write-Host $computer -ForegroundColor Magenta
    $ps = "PS"
    Write-Host $ps -NoNewline -ForegroundColor Cyan
    $location = " " + $(Get-Location) + ">"
    Write-Host $location -NoNewline -ForegroundColor Yellow
    return " "
}