[CmdletBinding()]
param(
    [string]$Repository = "lols8000/autoPsexec",
    [string]$RulesetPath = (
        Join-Path $PSScriptRoot "..\.github\rulesets\master-protection.json"
    ),
    [switch]$ValidateOnly
)

$ErrorActionPreference = "Stop"

function Assert-Condition {
    param(
        [bool]$Condition,
        [string]$Message
    )

    if (-not $Condition) {
        throw $Message
    }
}

$resolvedPath = (Resolve-Path $RulesetPath).Path
$raw = Get-Content -Raw -Encoding UTF8 $resolvedPath
$config = $raw | ConvertFrom-Json

Assert-Condition ($config.name -eq "Protect master - Central N2") "Nome inesperado no ruleset."
Assert-Condition ($config.target -eq "branch") "O ruleset precisa ter target=branch."
Assert-Condition ($config.enforcement -eq "active") "O ruleset precisa estar ativo."
Assert-Condition (@($config.conditions.ref_name.include) -contains "refs/heads/master") "O ruleset precisa atingir refs/heads/master."

$ruleTypes = @($config.rules | ForEach-Object { $_.type })
foreach ($requiredRule in @("deletion", "non_fast_forward", "pull_request", "required_status_checks")) {
    Assert-Condition ($ruleTypes -contains $requiredRule) "Regra obrigatória ausente: $requiredRule"
}

$statusRule = @($config.rules | Where-Object { $_.type -eq "required_status_checks" })[0]
Assert-Condition ($null -ne $statusRule) "Regra required_status_checks ausente."
Assert-Condition ($statusRule.parameters.strict_required_status_checks_policy -eq $true) "Os checks precisam exigir branch atualizada."

$contexts = @($statusRule.parameters.required_status_checks | ForEach-Object { $_.context })
foreach ($requiredContext in @("test (3.10)", "test (3.12)", "test (3.13)", "build")) {
    Assert-Condition ($contexts -contains $requiredContext) "Status check obrigatório ausente: $requiredContext"
}

$pullRequestRule = @($config.rules | Where-Object { $_.type -eq "pull_request" })[0]
Assert-Condition ($null -ne $pullRequestRule) "Regra pull_request ausente."
Assert-Condition ($pullRequestRule.parameters.required_review_thread_resolution -eq $true) "Conversas de review precisam ser resolvidas antes do merge."

if ($ValidateOnly) {
    Write-Host "Ruleset válido: $resolvedPath"
    exit 0
}

if (-not (Get-Command gh -ErrorAction SilentlyContinue)) {
    throw "GitHub CLI (gh) não encontrado. Instale-o e execute gh auth login antes de aplicar a proteção."
}

& gh auth status --hostname github.com *> $null
if ($LASTEXITCODE -ne 0) {
    throw "GitHub CLI não está autenticado em github.com."
}

$apiVersion = "2026-03-10"
$headers = @(
    "-H", "Accept: application/vnd.github+json",
    "-H", "X-GitHub-Api-Version: $apiVersion"
)

$existingJson = & gh api "repos/$Repository/rulesets" @headers
if ($LASTEXITCODE -ne 0) {
    throw "Não foi possível listar os rulesets de $Repository."
}

$existing = @($existingJson | ConvertFrom-Json)
$current = @($existing | Where-Object { $_.name -eq $config.name })

if ($current.Count -gt 1) {
    throw "Mais de um ruleset com o mesmo nome foi encontrado."
}

if ($current.Count -eq 1) {
    $endpoint = "repos/$Repository/rulesets/$($current[0].id)"
    $method = "PUT"
    Write-Host "Atualizando ruleset existente #$($current[0].id)..."
}
else {
    $endpoint = "repos/$Repository/rulesets"
    $method = "POST"
    Write-Host "Criando ruleset de proteção do master..."
}

$raw | & gh api $endpoint --method $method @headers --input -
if ($LASTEXITCODE -ne 0) {
    throw "Falha ao aplicar o ruleset. Confirme Administration: write no repositório."
}

Write-Host ""
Write-Host "Proteção do master aplicada com sucesso."
Write-Host "Validando estado remoto..."

$finalJson = & gh api "repos/$Repository/rulesets" @headers
if ($LASTEXITCODE -ne 0) {
    throw "Não foi possível validar o estado final dos rulesets."
}

$final = @($finalJson | ConvertFrom-Json)
$applied = @($final | Where-Object { $_.name -eq $config.name })[0]
Assert-Condition ($null -ne $applied) "Ruleset não apareceu após a aplicação."
Assert-Condition ($applied.enforcement -eq "active") "Ruleset existe, mas não está ativo."

Write-Host "OK: ruleset #$($applied.id) ativo para refs/heads/master."
