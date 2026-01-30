


# 1. Define o caminho da pasta
$path = "C:\inventario_otimizado_AtomoTek"

# 2. Cria a pasta se não existir
if (!(Test-Path $path)) {
    New-Item -Path $path -ItemType Directory
    Write-Host "Diretório criado: $path" -ForegroundColor Cyan
}

# 3. Obtém o SID do grupo 'Todos' / 'Everyone' (S-1-1-0)
# Isso evita o erro de tradução do nome do grupo
$sid = New-Object System.Security.Principal.SecurityIdentifier("S-1-1-0")
$everyone = $sid.Translate([System.Security.Principal.NTAccount])

# 4. Define a regra de 'Controle Total'
$acl = Get-Acl $path
$permission = $everyone.Value, "FullControl", "ContainerInherit, ObjectInherit", "None", "Allow"
$accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule($permission)

# 5. Aplica as permissões
$acl.SetAccessRule($accessRule)
Set-Acl $path $acl

Write-Host "Permissões de acesso total configuradas com sucesso para: $($everyone.Value)" -ForegroundColor Green



# ============================================================
# INVENTÁRIO OTIMIZADO - ATOMOTEK SOLUÇÕES TECNOLÓGICAS
# ============================================================

# URL de envio (formResponse) do seu formulário específico
$formUrl = "https://docs.google.com/forms/d/e/1FAIpQLSdicQ5C6uXWy8PgNEiN0GQ6J8KDtlAyQFyv2JxyoRB0s1Odng/formResponse"

# 1. Verificação de Privilégios de Administrador
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "ERRO: VOCÊ PRECISA EXECUTAR ESTE SCRIPT COMO ADMINISTRADOR." -ForegroundColor Red
    Write-Host "Por favor, feche e abra o PowerShell clicando com o botão direito em 'Executar como Administrador'."
    Pause
    exit
}

Clear-Host
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "   ATOMOTEK SOLUÇÕES TECNOLÓGICAS - SISTEMA DE INVENTÁRIO   " -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "`nIniciando Coleta de Dados..." -ForegroundColor Yellow

# --- DEFINIÇÃO DO DIRETÓRIO LOCAL ---
$diretorioDestino = "C:\inventario_otimizado_AtomoTek"
if (-not (Test-Path $diretorioDestino)) {
    try {
        New-Item -Path $diretorioDestino -ItemType Directory -ErrorAction Stop | Out-Null
    } catch {
        Write-Host "Erro ao criar diretório em C:\. Verifique permissões." -ForegroundColor Red
    }
}

# --- MENUS DE SELEÇÃO MANUAL ---

# Seleção de Cliente
$clienteMenu = @"
SELECIONE O CLIENTE:
1. AtomoTek
2. Sheila Contabilidade
3. MIC Contabilidade
4. Espaço Cont
5. Colégio Curumim
6. Avulso
"@
Write-Host $clienteMenu -ForegroundColor White
$opCliente = Read-Host "`nOpção"
$cliente = switch($opCliente){ "1"{"AtomoTek"}; "2"{"Sheila Contabilidade"}; "3"{"MIC Contabilidade"}; "4"{"Espaço Cont"}; "5"{"Colégio Curumim"}; default{"Avulso"} }

# Seleção de Localização
$localMenu = @"
`nPRINCIPAL LOCAL DE USO:
1. Matriz
2. Filial 01
3. Filial 02
4. Home Office
5. No Cliente
6. Híbrido
"@
Write-Host $localMenu -ForegroundColor White
$opLocal = Read-Host "`nOpção"
$localizacao = switch($opLocal){ "1"{"Matriz"}; "2"{"Filial 01"}; "3"{"Filial 02"}; "4"{"Home Office"}; "5"{"No Cliente"}; "6"{"Híbrido"}; default{"Externo"} }

# Seleção de Equipamento
$tipoMenu = @"
`nTIPO DE EQUIPAMENTO:
1. Desktop
2. Servidor / Desktop
3. Notebook
4. Servidor
5. All-in-One
6. Mini Desktop
7. Tablet
"@
Write-Host $tipoMenu -ForegroundColor White
$opTipo = Read-Host "`nOpção"
$tipoManual = switch($opTipo){ "1"{"Desktop"}; "2"{"Servidor / Desktop"}; "3"{"Notebook"}; "4"{"Servidor"}; "5"{"All-in-One"}; "6"{"Mini Desktop"}; "7"{"Tablet"}; default{"Estação de Trabalho"} }

$setor = Read-Host "`nDigite o SETOR (ex: RH, TI)"
$usuario = Read-Host "Digite o UTILIZADOR (Nome do Usuário)"

# --- COLETA AUTOMÁTICA DE SISTEMA ---
Write-Host "`nColetando informações de hardware..." -ForegroundColor Gray

$data = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
$hostname = $env:COMPUTERNAME
$bios = Get-CimInstance Win32_Bios
$comp = Get-CimInstance Win32_ComputerSystem
$os = Get-CimInstance Win32_OperatingSystem
$arquitetura = $os.OSArchitecture

# Processador e Geração
$cpuObj = Get-CimInstance Win32_Processor
$cpuNome = $cpuObj.Name
$geracao = if ($cpuNome -match "-(\d{1,2})") { $Matches[1] + "ª Geração" } else { "N/A" }

# Memória RAM Detalhada
$ramModulos = Get-CimInstance Win32_PhysicalMemory
$ramTotal = [math]::round(($ramModulos | Measure-Object -Property Capacity -Sum).Sum / 1GB, 0)
$ramVelocidade = ($ramModulos | Select-Object -First 1).ConfiguredClockSpeed
$ramTipoRaw = ($ramModulos | Select-Object -First 1).SmbiosMemoryType
$ramDDR = switch($ramTipoRaw){
    20 {"DDR"} 21 {"DDR2"} 24 {"DDR3"} 26 {"DDR4"} 34 {"DDR5"} default {"DDR/Outro"}
}

# Armazenamento e Saúde
$particaoC = Get-Partition -DriveLetter C
$disco = Get-Disk -Number $particaoC.DiskNumber
$volumeC = Get-Volume -DriveLetter C
$espacoTotal = [math]::Round($volumeC.Size / 1GB, 0)
$espacoLivre = [math]::Round($volumeC.SizeRemaining / 1GB, 0)
$espacoStr = "Total: $espacoTotal GB / Livre: $espacoLivre GB"

$saudeTraduzida = switch($disco.HealthStatus) {
    "Healthy"   { "Saudável" }
    "Warning"   { "Atenção (Alerta)" }
    "Unhealthy" { "Crítico (Falha)" }
    default     { "Desconhecido" }
}

# Bitlocker
$bitStatus = try {
    $bl = Get-BitLockerVolume -MountPoint "C:" -ErrorAction Stop
    if ($bl.ProtectionStatus -eq "On") { "Protegido" } else { "Desprotegido" }
} catch {
    "N/A ou Sem Permissão"
}

# --- MONTAGEM DO RELATÓRIO PARA EXIBIÇÃO ---
$relatorioTxt = @"
============================================================
  INVENTÁRIO OTIMIZADO - ATOMOTEK SOLUÇÕES TECNOLÓGICAS
============================================================
************************************************************

Data:                     $data
Cliente:                  $cliente
Local de Uso:             $localizacao
Setor:                    $setor
Utilizador:               $usuario
------------------------------------------------------------
Hostname:                 $hostname
Fabricante:               $($comp.Manufacturer)
Serial:                   $($bios.SerialNumber)
Tipo:                     $tipoManual
S.O:                      $($os.Caption) ($arquitetura)
------------------------------------------------------------
Processador (CPU):        $cpuNome
Geração (CPU):            $geracao
RAM Total:                $ramTotal GB
Tipo/Freq RAM:            $ramDDR @ $ramVelocidade MHz
------------------------------------------------------------
Armazenamento:            $($disco.FriendlyName)
Saúde do Armazenamento:   $saudeTraduzida
Espaço C:                 $espacoStr
Bitlock C:                $bitStatus
============================================================
"@

# --- EXIBIÇÃO NO CONSOLE ---
Clear-Host
Write-Host $relatorioTxt -ForegroundColor White

# --- ENVIO PARA O GOOGLE FORMS ---
$postParams = @{
    "entry.117297982"  = $data
    "entry.1417077877" = $cliente
    "entry.954384993"  = $localizacao
    "entry.774303326"  = $setor
    "entry.1198847817" = $usuario
    "entry.1602215666" = $hostname
    "entry.1818807027" = $comp.Manufacturer
    "entry.344416170"  = $bios.SerialNumber
    "entry.1328999387" = $tipoManual
    "entry.1124472686" = "$($os.Caption) ($arquitetura)"
    "entry.720697888"  = $cpuNome
    "entry.1517327484" = $geracao
    "entry.1082474464" = "$ramTotal GB"
    "entry.60910780"   = "$ramDDR @ $ramVelocidade MHz"
    "entry.1730575164" = $disco.FriendlyName
    "entry.1121982261" = $saudeTraduzida
    "entry.974233344"  = $espacoStr
    "entry.1744253886" = $bitStatus
}

Write-Host "`nSincronizando com Google Sheets (AtomoTek Cloud)..." -ForegroundColor Cyan
try {
    Invoke-WebRequest -Uri $formUrl -Method Post -Body $postParams -ErrorAction Stop | Out-Null
    Write-Host "SUCESSO: Dados enviados para a nuvem." -ForegroundColor Green
} catch {
    Write-Host "AVISO: Falha no envio para nuvem. Verifique a internet." -ForegroundColor Red
}

# --- SALVAMENTO DO ARQUIVO LOCAL ---
$nomeArquivo = "Inventario_$($cliente -replace '[^a-zA-Z0-9]', '_')_$($hostname).txt"
$caminhoRelatorio = Join-Path -Path $diretorioDestino -ChildPath $nomeArquivo
$relatorioTxt | Out-File -FilePath $caminhoRelatorio -Encoding UTF8

Write-Host "Arquivo salvo localmente em: $caminhoRelatorio" -ForegroundColor Cyan
Write-Host "`nScript finalizado. Pressione qualquer tecla para abrir o relatório..." -ForegroundColor Yellow
Pause
notepad.exe $caminhoRelatorio