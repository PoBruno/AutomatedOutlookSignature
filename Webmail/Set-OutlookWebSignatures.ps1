#Requires -Module ExchangeOnlineManagement
#Requires -Module ActiveDirectory

<#
.SYNOPSIS
  Script de Assinatura do Outlook Web - https://github.com/PoBruno/AutomatedOutlookSignature
.DESCRIPTION
  Este script obtém os detalhes de cada usuário no AD, cria sua assinatura HTML e define como sua assinatura no Outlook web.
.OUTPUTS
  Todas as assinaturas são armazenadas na pasta $signatureFolder
.NOTES
  Versão:           1.0
  Autor:            Bruno Gomes 
  Data de Criação:  10/06/2024
#>

#-----[ Configuração ]-----#

# Define a pasta para armazenar todas as assinaturas web e, se a pasta não existir, ela será criada.
$signatureFolder = "$psscriptroot\Web-Signatures"
if (!(test-path $signatureFolder)) {
    New-Item -ItemType "directory" -Path $signatureFolder
}

#-----[ Funções ]-----#

# Esta função cria o arquivo de assinatura se ele não existir ou o atualiza se houver diferenças
function Create-WebSignatures {
    # Array armazenando uma lista de todos os usuários que precisam de atualização de assinatura
    $signaturesToUpdate = @()

    # Obtém todos os usuários no grupo de assinaturas do Outlook
    $allStaff = Get-ADGroupMember "Outlook Web Signature"

    # Para cada usuário no grupo Todos os Funcionários, o seguinte será executado
    $allStaff | ForEach-Object {

        # Armazena os detalhes do usuário em $user
        $user = Get-Aduser -Identity $_.distinguishedname -Properties Title, MobilePhone, EmailAddress, extensionattribute1, extensionattribute2, streetaddress, st, l, postalcode, telephonenumber
        
        # Se o usuário estiver desativado, ele será ignorado
        if (!$user.Enabled) {
            Write-Host "$($user.Name) está desativado e será ignorado"
            return
        }
        
        # Salvando as propriedades do usuário em variáveis com nomes mais amigáveis
        $username = ($user.userprincipalName).Substring(0, $user.userprincipalname.IndexOf('@'))
        $displayName = $user.Name
        $jobTitle = $user.Title
        $mobileNumber = $user.MobilePhone
        $email = $user.EmailAddress
        $namePrefix = $user.extensionattribute1 # Dr etc.
        $namePostfix = $user.extensionattribute2 # Bs(Hons) etc.

        # Estes são detalhes que você pode obter do Active Directory ou, se forem os mesmos para toda a empresa, podem ser definidos estaticamente aqui. Cada um tem um exemplo estático comentado, basta trocar as linhas comentadas e alterar o exemplo.
        $companyName = $user.company # Nome da empresa
        $street = $user.streetaddress # Endereço
        $city = $user.l # Cidade
        $state = $user.st # Estado
        $zipCode = $user.postalcode # CEP 
        $telephone = $user.telephonenumber # Número de telefone
        $website = "www.example.co.uk" # Website
        $logo = "https://www.google.com/images/branding/googlelogo/2x/googlelogo_color_92x30dp.png" # Logo

        # Reúne uma lista de todos os grupos dos quais o usuário é membro
        $groups = Get-ADPrincipalGroupMembership $_ | select name

        # Exemplo de Verificação de Grupo
        $Group = [ADSI]"LDAP://cn=IT Staff,OU=Groups,DC=Example,DC=co,DC=uk"
        $Group.Member | ForEach-Object {
            if ($user.distinguishedname -match $_) {
                $ItStaff = $true
            }
        }

        $address = $null

        # Construindo endereço        
        if ($street) { $address = "$($street), " } 
        if ($city) { $address = $address + "$($city), " }
        if ($state) { $address = $address + "$($state), " }
        if ($zipCode) { $address = $address + $zipCode }
    
    
        # Construindo Folha de Estilo
        $style = 
        @"
  <style>
  p, table, td, tr, a, span { 
      font-family: Arial, Helvetica, sans-serif;
      font-size:  12pt;
      color: #28b8ce;
  }

  span.blue
  {
      color: #28b8ce;
  }

  table {
      margin: 0;
      padding: 0;
  }

  a { 
  text-decoration: none;
  }

  hr {
  border: none;
  height: 1px;
  background-color: #28b8ce;
  color: #28b8ce;
  width: 700px;
  }

  table.main {
      border-top: 1px solid #28b8ce;
  }
  </style>
"@

        # Construindo HTML
        $signature = 
        @"
    $(if($displayName){"<span><b>"+$displayName+"</b></span><br />"})
    $(if($jobTitle){"<span>"+$jobTitle+"</span><br /><br />"})

  <p>
    <table class='main'>
        <tr>
            <td style='padding-right: 75px;'>$(if($logo){"<img src='$logo' />"})</td>
            <td>
                <table>
                    <tr><td colspan='2' style='padding-bottom: 10px;'>
                      $(if($companyName){ "<b>"+$companyName+"</b><br />" })
                      $(if($street){ $street+", " })
                      $(if($city){ $city+", " })
                      $(if($state){ $state+", " })
                      $(if($zipCode){ $zipCode })
                    </td></tr>
                    $(if($ITMember){"<tr><td colspan='2'>IT Helpdesk: 0188887 55555 6666</tr></td>"})
                    $(if($telephone){"<tr><td>T: </td><td><a href='tel:$telephone'>$($telephone)</a></td></tr>"})
                    $(if($mobileNumber){"<tr><td>M: </td><td><a href='tel:$mobileNumber'>$($mobileNumber)</a></td></tr>"})
                    $(if($email){"<tr><td>E: </td><td><a href='mailto:$email'>$($email)</a></td></tr>"})
                    $(if($website){"<tr><td>W: <a href='https://$website'>$($website)</a></td></tr>"})
                </table>
            </td>
        </tr>
    </table>
  </p>
  <br />
"@

        # Se o arquivo existir, ele comparará as assinaturas e, se houver alterações, marcará como necessitando de uma atualização
        if (test-path "$signatureFolder\$email.html") {          
            $currentSig = (Get-Content "$signatureFolder\$email.html" | out-string).TrimEnd()

            if ($currentSig -eq $signature) {
                write-host "Assinatura encontrada para $displayname - Nenhuma atualização necessária." -ForegroundColor green
            }
            else {
                write-host "Assinatura encontrada para $displayname - Atualização necessária." -ForegroundColor yellow
                Remove-Item -Path "$signatureFolder\$email.html" -Force
                $signature | Out-File "$signatureFolder\$email.html"
                $signaturesToUpdate += $email
            }
        }
        else {
            write-host "Nenhuma assinatura encontrada para $displayName. Criando assinatura" -ForegroundColor Red
            $signature | Out-File "$signatureFolder\$email.html"
            $signaturesToUpdate += $email
        }
    }
    return $signaturesToUpdate
}

# Esta função recebe todos os usuários que não têm assinatura ou precisam de atualização e usa o arquivo de assinatura para atualizar sua assinatura web.
function Update-WebSignatures {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]
        $userEmailAddress
    )

    write-host "Definindo assinatura na caixa de correio: $userEmailAddress"
    
    # Obtém o conteúdo da assinatura do arquivo.
    $signature = Get-Content "$signatureFolder\$userEmailAddress.html"

    # Define a assinatura do usuário com o conteúdo do arquivo.
    Get-Mailbox $userEmailAddress | Set-MailboxMessageConfiguration -SignatureHTML $signature -AutoAddSignature:$true
}

#-----[ Execução ]-----#

# Cria todas as assinaturas novas/modificadas e gera uma lista de quem precisa delas alteradas online
$usersToUpdate = Create-WebSignatures

# Se algum usuário precisar de uma nova assinatura/modificação, esta seção será executada
if ($usersToUpdate.Count -gt 0) {
    try {
        Write-Host "Conectando ao Exchange Online"
        Connect-ExchangeOnline

        foreach ($usertoUpdate in $usersToUpdate) {
            # Isso chama a função para cada um dos usuários que precisa de uma assinatura atualizada
            Update-WebSignatures $usertoUpdate
        }
    }
    catch {
        write-host "Oh dear something went wrong" -ForegroundColor Red
    }
    finally {
        Write-Host "Desconectando do Exchange Online"
        Disconnect-ExchangeOnline -Confirm:$false
    }
}
