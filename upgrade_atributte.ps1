# Caminho para o arquivo Excel
$excelFilePath = "C:\Users\lgonzales\Desktop\shellscript\carlos\app\testead.xlsx"

# Nome da coluna que contém os nomes de usuários no Excel
$usuarioColumnName = "Usuários"

# Nome da coluna que contém as informações para atualizar no AD
$sobrenomeCargoColumnName = "Sobrenome do Cargo"

# Importar o módulo do Active Directory
Import-Module ActiveDirectory

# Ler o arquivo Excel e extrair os dados
$excelData = Import-Excel -Path $excelFilePath

# Iterar pelos dados do Excel
foreach ($row in $excelData) {
    $usuario = $row.$usuarioColumnName
    $sobrenomeCargo = $row.$sobrenomeCargoColumnName

    # Buscar o usuário no Active Directory
    $user = Get-ADUser -Filter {SamAccountName -eq $usuario}

    if ($user) {
        # Atualizar o atributo "comment" no Active Directory
        Set-ADUser -Identity $user -Replace @{comment = $sobrenomeCargo}
        Write-Host "Usuário $usuario encontrado no AD e atributo 'comment' atualizado."
    } else {
        Write-Host "Usuário $usuario não encontrado no AD."
    }
}
