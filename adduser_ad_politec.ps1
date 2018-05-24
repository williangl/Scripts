########################################################################
# ADAPTADO DE : Marius / Hican - http://www.hican.nl - @hicannl 
# FONTE       : https://www.class365.com.br/criar-usuarios-no-active-directory-a-partir-de-um-arquivo-csv/
# INSTRUÇÕES  : O arquivo CSV deve conter ao menos 4 colunas: RA, Aluno, Curso e Email. 
#               O CSV deve estar configurado com "," (virgula) como delimitador, senão não funcionará. 
# HINT        : Caso haja problemas com isso verificar nas configurações de idioma do Windows qual a pontuação de delimitação.
#               Por padrão o Windows instalado em PT-BR vem com o ";" (ponto e virgula) como delimitação, só trocar! 
########################################################################

Set-StrictMode -Version latest

#----------------------------------------------------------
# CARGA DOS MODULOS
#----------------------------------------------------------
Try
{
  Import-Module ActiveDirectory -ErrorAction Stop
}
Catch
{
  Write-Host "[ERRO]`t O modulo do Active Directory nao pode ser carregado. Nao ha como continuar!"
  Exit 1
}

#----------------------------------------------------------
# VARIAVEIS ESTATICAS
#----------------------------------------------------------
$path     = Split-Path -parent $MyInvocation.MyCommand.Definition
$newpath  = $path + "\2018-1teste.csv"
$log      = $path + "\UsuariosCriados.log"
$date     = Get-Date
$addn     = (Get-ADDomain).DistinguishedName
$dnsroot  = (Get-ADDomain).DNSRoot
$i        = 2

#----------------------------------------------------------
# FUNCOES
#----------------------------------------------------------
Function Start-Commands
{
  Create-Users
}

Function Create-Users
{
  "Processamento iniciado (em " + $date + "): " | Out-File $log -append
  "--------------------------------------------" | Out-File $log -append
  $users = Import-CSV $newpath
  ForEach ($user in $users) {
    If (($user.RA -eq "") -Or ($user.Aluno -eq ""))
    {
        Write-Host "[ERRO]`t Os campos 'RA' e/ou 'Aluno' sao nulos ou invalidos. Processamento nao executado para a linha $($i)`r`n"
        "[ERRO]`t Os campos 'RA' e/ou 'Aluno' sao nulos ou invalidos. Processamento nao executado para a linha $($i)`r`n" | Out-File $log -append
      }
      Else
      {
        # Definicao da OU destino
        $location = 'OU=Teste,OU=FAP' + ",$($addn)"

        # Criacao da propriedade 'sAMAccountName' seguindo o seguinte parent:
        # N° RA:
        # 2018123456
        $sam = $user.RA
        Try   { $exists = Get-ADUser -LDAPFilter "(sAMAccountName=$sam)" }
        Catch { }
        If(!$exists)
        {
          $setpass = $null

          Try
          {
            Write-Host "[INFO]`t Criando usuario: $($sam)"
            "[INFO]`t Criando usuario: $($sam)" | Out-File $log -append
            New-ADUser $user.Aluno -DisplayName:$user.Aluno -SamAccountName:$sam `
            -Description:$user.Curso -EmailAddress:$user.Email `
            -UserPrincipalName:($sam + "@" + $dnsroot) -Path:$location `
            -Company:'Faculdade de Santa Bárbara dOeste' -Department:'Aluno' -EmployeeID:$user.RA `
            -Title:'Aluno' -AccountPassword:$setpass `
            -Enabled:$True -ChangePasswordAtLogon:$True -CannotChangePassword:$True
            Write-Host "[INFO]`t Novo usuario criado: $($sam)"
            "[INFO]`t Novo usuario criado: $($sam)" | Out-File $log -append
          }
          Catch
          {
            Write-Host "[ERRO]`t Oops, algo deu errado: $($_.Exception.Message)`r`n"
          }
        }
        Else
        {
          Write-Host "[AVISO]`t Usuario $($sam) ($($user.Aluno)) ja existe ou retornou um erro!`r`n"
          "[AVISO]`t Usuario $($sam) ($($user.Aluno)) ja existe ou retornou um erro!`r`n" | Out-File $log -append
        }
      }

    $i++
  }
  "--------------------------------------------" + "`r`n" | Out-File $log -append
}

Write-Host "INICIADO`r`n"
Start-Commands
Write-Host "CONCLUIDO"