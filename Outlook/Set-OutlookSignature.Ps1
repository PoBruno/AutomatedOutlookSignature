<#
.SYNOPSIS
  Criação Dinâmica de Assinatura do Outlook - https://github.com/PoBruno/AutomatedOutlookSignature
.DESCRIPTION
  Este script utiliza propriedades do ActiveDirectory para preencher os arquivos .htm e .txt que são então armazenados na $folderlocation com o nome $filename.htm e $filename.txt. 
  O script pode ser executado como uma tarefa agendada na inicialização ou, de preferência, como um script de Logon de Política de Grupo. Mais detalhes podem ser encontrados no GitHub ReadMe.
  O script usa instruções if para garantir que, se um usuário, por exemplo, não tiver um número de celular, essa seção inteira da assinatura não seja incluída. Isso torna a assinatura muito mais profissional e, 
  se o usuário receber um número de celular, na próxima inicialização, ele será adicionado à sua assinatura. 
  O script também pode adicionar partes adicionais com base na associação ao grupo. O exemplo dado neste script é ser membro da Equipe de TI adicionando informações de contato do Helpdesk. 
.INPUTS
  Várias propriedades do usuário são extraídas diretamente do Active Directory. Para mais detalhes, consulte o ReadMe no GitHub.
.OUTPUTS
  $folderlocation\$filename.htm - Assinatura HTML para e-mails em rich text
  $folderlocation\$filename.txt - Assinatura em texto para e-mails sem rich text
.NOTES
  Versão:        4.0
  Autor:         Bruno Gomes
  Modificado:    24/04/2023
#>

# Tenta obter o usuário do Active Directory, se nenhum usuário for encontrado ou houver um problema, o script será encerrado
try {
  $user = (([adsisearcher]"(&(objectCategory=User)(samaccountname=$env:username))").FindOne().Properties)
} catch {
  Write-Host "Erro: Não foi possível consultar o Active Directory para obter informações do usuário. Detalhes: $($_.Exception.Message)"
  exit
}

# Cria a pasta de assinaturas e define o nome do arquivo de assinatura
$folderLocation = Join-Path -Path $Env:appdata -ChildPath 'Microsoft\signatures'
$filename = 'Signature'
$file  = Join-Path -Path $folderLocation -ChildPath $filename

# Se a pasta não existir, cria-a
if (-not (Test-Path -Path $folderLocation)) {
  try {
      New-Item -ItemType directory -Path $folderLocation
  } catch {
      Write-Host "Erro: Não foi possível criar a pasta de assinaturas. Detalhes: $($_.Exception.Message)"
      exit
  }
}

# Logotipo a ser usado - se diferentes usuários precisarem de logotipos diferentes, você pode considerar usar suas extensionAttributes ou outra propriedade no Active Directory para armazenar a URL de seus logotipos.
$logo = 'https://www.google.com/images/branding/googlelogo/2x/googlelogo_color_92x30dp.png' # Observe que, se você incluir um logotipo, ele deve estar localizado em algum lugar na internet ao qual o público tenha acesso; muitos usuários o enviam para o site deles.

# Obtém as propriedades dos usuários (Estas devem estar sempre no Active Directory e ser Únicas)
if($user.name.count -gt 0){$displayName = $user.name[0]} # Nome de Exibição
if($user.title.count -gt 0){$jobTitle = $user.title[0]} # Título do Trabalho
if($user.mail.count -gt 0){$email = $user.mail[0]} # Endereço de E-mail

if($user.mobile.count -gt 0){$mobileNumber = $user.mobile[0]} # Número de Celular
if($user.homephone.count -gt 0){$directDial = $user.homephone[0]} # Número Residencial
if($user.telephonenumber.count -gt 0){$telephone = $user.telephonenumber[0]} # Número de Escritório

# Nome da empresa e site
if($user.company.count -gt 0){$companyName = $user.company[0]} # Nome da Empresa
if($user.wwwhomepage.count -gt 0){$website = $user.wwwhomepage[0]} # Endereço da Página Web

# Endereço
if($user.postofficebox.count -gt 0){$poBox = $user.postofficebox[0]} # Número da Caixa Postal
if($user.physicaldeliveryofficename.count -gt 0){$office = $user.physicaldeliveryofficename} # Nome do Escritório
if($user.streetaddress.count -gt 0){$street = $user.streetaddress[0]} # Endereço
if($user.l.count -gt 0){$city = $user.l[0]} # Cidade
if($user.st.count -gt 0){$state = $user.st[0]} # Estado
if($user.postalcode.count -gt 0){$zipCode = $user.postalcode[0]} # Código Postal / CEP

# Atributos Estendidos
if($user.extensionAttribute1.count -gt 0){$attribute1 = $user.extensionAttribute1[0]} # Atributo Personalizado 1
if($user.extensionAttribute2.count -gt 0){$attribute2 = $user.extensionAttribute2[0]} # Atributo Personalizado 2
if($user.extensionAttribute3.count -gt 0){$attribute3 = $user.extensionAttribute3[0]} # Atributo Personalizado 3
if($user.extensionAttribute4.count -gt 0){$attribute4 = $user.extensionAttribute4[0]} # Atributo Personalizado 4
if($user.extensionAttribute5.count -gt 0){$attribute5 = $user.extensionAttribute5[0]} # Atributo Personalizado 5

# Exemplo de Verificação de Grupo
$Group = [ADSI]"LDAP://cn=IT Staff,OU=Groups,DC=Example,DC=co,DC=uk"
$Group.Member | ForEach-Object {
if ($user.distinguishedname -match $_) {
      $ItStaff = $true
    }
}

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
                  $(if($ITMember){"<tr><td td colspan='2'>Helpdesk de TI: 0188887 55555 6666</tr></td>"})
                  $(if($telephone){"<tr><td>T: </td><td><a href='tel:$telephone'>$($telephone)</a></td></tr>"})
                  $(if($mobileNumber){"<tr><td>M: </td><td><a href='tel:$mobileNumber'>$($mobileNumber)</a></td></tr>"})
                  $(if($email){"<tr><td>E: </td><td><a href='mailto:$email'>$($email)</a></td></tr>"})
                  $(if($website){"<tr><td>W:</td><td><a href='https://$website'>$($website)</a></td></tr>"})
              </table>
          </td>
      </tr>
  </table>
</p>
<br />
"@

# Salva o HTML no arquivo de assinatura
try {
  $style + $signature | Out-File -FilePath "$file.htm" -Encoding ascii
} catch {
  Write-Host "Erro: Não foi possível salvar o arquivo de assinatura HTML. Detalhes: $($_.Exception.Message)"
  exit
}

# Cria a versão txt para e-mails sem rich text
$signature = 
@"
$(if($displayName){ $displayName })
$(if($jobTitle){ $jobTitle })
___________________________________________________________

$(if($companyName){ $companyName })
$(if($street){ $street+", " })
$(if($city){ $city+", " })
$(if($state){ $state+", " })
$(if($zipCode){ $zipCode })

$(if($ITMember){"Para o Helpdesk de TI ligue 0191231 212313"})
$(if($telephone){"T: "+$telephone})
$(if($mobileNumber){"M: "+$mobileNumber})
$(if($email){"E: "+$email})
$(if($website){"W: "+$website})
"@

# Salva o texto na pasta de assinaturas
try {
  $signature | out-file "$file.txt" -encoding ascii
} catch {
  Write-Host "Erro: Não foi possível salvar o arquivo de assinatura de texto. Detalhes: $($_.Exception.Message)"
  exit
}

# Definindo as chaves de registro para Outlook 2016
if (test-path "HKCU:\\Software\\Microsoft\\Office\\16.0\\Common\\General") 
{
    get-item -path HKCU:\\Software\\Microsoft\\Office\\16.0\\Common\\General | new-Itemproperty -name Signatures -value signatures -propertytype string -force
    get-item -path HKCU:\\Software\\Microsoft\\Office\\16.0\\Common\\MailSettings | new-Itemproperty -name NewSignature -value $filename -propertytype string -force
    get-item -path HKCU:\\Software\\Microsoft\\Office\\16.0\\Common\\MailSettings | new-Itemproperty -name ReplySignature -value $filename -propertytype string -force
    Remove-ItemProperty -Path HKCU:\\Software\\Microsoft\\Office\\16.0\\Outlook\\Setup -Name "First-Run" -ErrorAction silentlycontinue
}

# Definindo as chaves de registro para Outlook 2010 - Obrigado AJWhite1970 pelas chaves de registro de 2010
if (test-path "HKCU:\\Software\\Microsoft\\Office\\14.0\\Common\\General") 
{
    get-item -path HKCU:\\Software\\Microsoft\\Office\\14.0\\Common\\ General | new-Itemproperty -name Signatures -value signatures -propertytype string -force
    get-item -path HKCU:\\Software\\Microsoft\\Office\\14.0\\Common\\ MailSettings | new-Itemproperty -name NewSignature -value $filename -propertytype string -force
    get-item -path HKCU:\\Software\\Microsoft\\Office\\14.0\\Common\\ MailSettings | new-Itemproperty -name ReplySignature -value $filename -propertytype string -force
    Remove-ItemProperty -Path HKCU:\\Software\\Microsoft\\Office\\14.0\\Outlook\\Setup -Name "First-Run" -ErrorAction silentlycontinue
}
