$pythonInstallerUrl = "https://www.python.org/ftp/python/3.10.7/python-3.10.7-amd64.exe" 
$pythonScriptUrl = "https://raw.githubusercontent.com/svictorsena/scripts-sdk/refs/heads/master/planilha.py"

$pythonInstallerPath = "$env:TEMP\python-installer.exe"
$pythonScriptPath = "$env:TEMP\script.py"

function Check-PythonInstalled {
    try {
        python --version | Out-Null
        return $true
    } catch {
        return $false
    }
}

if (!(Check-PythonInstalled)) {
    Write-Output "Python não encontrado. Baixando e instalando..."
    Invoke-WebRequest -Uri $pythonInstallerUrl -OutFile $pythonInstallerPath
    Start-Process -FilePath $pythonInstallerPath -ArgumentList "/quiet InstallAllUsers=1 PrependPath=1" -Wait
    Remove-Item -Path $pythonInstallerPath
    Write-Output "Python instalado com sucesso."
} else {
    Write-Output "Python já está instalado."
}

Write-Output "Baixando o arquivo Python..."
Invoke-WebRequest -Uri $pythonScriptUrl -OutFile $pythonScriptPath

Write-Output "Executando o script Python..."
python $pythonScriptPath

Write-Output "Excluindo o arquivo Python..."
Remove-Item -Path $pythonScriptPath

Write-Output "Processo concluído!"

Write-Output "Desinstalando o Python..."

$pythonUninstallPath = "C:\Program Files\Python310\uninstall.exe"
$pythonUninstallPath2 = "C:\Program Files (x86)\Python310\uninstall.exe" 

if (Test-Path $pythonUninstallPath) {
    Start-Process -FilePath $pythonUninstallPath -ArgumentList "/quiet" -Wait
    Write-Output "Python foi desinstalado."
} else {
    Write-Output "Caminho de desinstalação do Python não encontrado."
    Write-Output "Tentando por outro caminho."
    if (Test-Path $pythonUninstallPath2) {
    	Start-Process -FilePath $pythonUninstallPath2 -ArgumentList "/quiet" -Wait
    	Write-Output "Python foi desinstalado."
    } else {
    	Write-Output "Caminho de desinstalação do Python não encontrado."
    }
}
