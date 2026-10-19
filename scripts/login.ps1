<#
.SYNOPSIS
    Sign in to Microsoft Graph for one MCP server profile and verify access.

.DESCRIPTION
    Runs the device-code sign-in in this terminal (the code is copied to the
    clipboard and the sign-in page opens), caches the token encrypted for the
    current Windows user, then runs the health check. Identity settings come
    from .env / .env.<profile> next to package.json, or from the environment.

.EXAMPLE
    .\scripts\login.ps1                     # corporate profile
    .\scripts\login.ps1 -Profile personal   # second profile
    .\scripts\login.ps1 -SignOut            # forget the corporate sign-in
#>
param(
    [string]$Profile = 'corporate',
    [switch]$SignOut
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot

if (-not (Test-Path (Join-Path $root 'build\auth-cli.js'))) {
    Write-Host 'Building first...' -ForegroundColor Yellow
    Push-Location $root
    try { npm run build | Out-Host } finally { Pop-Location }
}

$env:MICROSOFT_PROFILE = $Profile
$cliArgs = @((Join-Path $root 'build\auth-cli.js'))
if ($SignOut) { $cliArgs += '--sign-out' }

Write-Host "`n=== Microsoft Graph MCP sign-in: $Profile ===" -ForegroundColor Cyan
& node @cliArgs
switch ($LASTEXITCODE) {
    0 { Write-Host "`n[OK] Profile '$Profile' is signed in and healthy." -ForegroundColor Green }
    2 { Write-Host "`n[WARN] Signed in, but the mailbox is not reachable (see Verdict above)." -ForegroundColor Yellow }
    default { Write-Host "`n[FAIL] Sign-in failed (exit $LASTEXITCODE)." -ForegroundColor Red }
}
exit $LASTEXITCODE
