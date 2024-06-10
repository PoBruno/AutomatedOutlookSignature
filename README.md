[![Contribuidores][contributors-shield]][contributors-url]
[![Forks][forks-shield]][forks-url]
[![Stargazers][stars-shield]][stars-url]
[![Issues][issues-shield]][issues-url]
[![Licença MIT][license-shield]][license-url]

# Scripts Automatizados de Assinatura do Outlook
Este projeto contém dois scripts:
* Set-OutlookSignature.ps1 - Usado para gerar e definir a assinatura de um usuário para o Outlook desktop.
* Set-OutlookWebSignatures.ps1 - Este script ainda está em desenvolvimento.

## Scripts
Recomendo usar o script como um script de logon em Política de Grupo. - [Configuring Logon PowerShell Scripts with Group Policy - 4Sysops](https://4sysops.com/archives/configuring-logon-powershell-scripts-with-group-policy/)

Durante o processo de logon do usuário, o script é executado em segundo plano, recupera os detalhes necessários do usuário, gera um novo arquivo de assinatura e substitui o existente. Além disso, o script define chaves de registro para configurar a nova assinatura criada como a assinatura padrão do Outlook do usuário. Isso garante que, se houver alterações nos detalhes, como o título do cargo, a assinatura será atualizada automaticamente no próximo logon.

[EduGeek Post](http://www.edugeek.net/forums/scripts/205976-outlook-email-signature-automation-ad-attributes.html#post1760284)

### Active Directory
Uma seleção de atributos do Active Directory já está configurada no script e listada abaixo, mas mais atributos podem ser facilmente adicionados.

As seguintes propriedades são usadas do Active Directory no script:

| Variável no Script | Campo no AD | Notas |
|-------------| ------------- | ------------- |
| $displayName | Nome de exibição | Nome de exibição dos usuários |
| $jobTitle | Cargo | Cargo dos usuários |
| $email | E-mail | Endereço de e-mail dos usuários |
| $telephone | Telefone | Número de telefone principal do site/filial |
| $directDial | Telefone residencial | Número de discagem direta dos usuários |
| $mobileNumber | Celular | Número de celular dos usuários |
| $street | Rua | Rua / Primeira linha do endereço |
| $poBox | Caixa Postal | Nome do site / filial que aparecerá em negrito acima do endereço, por exemplo, Sede |
| $city | Cidade | Cidade |
| $state | Estado/Província | Estado / Condado |
| $zipCode | CEP | Código postal |
| $office | physicaldeliveryofficename | Escritório |
| $website | Site | Endereço do site |
| $companyName | Empresa | Nome da empresa |

Variáveis adicionais que não dependem do Active Directory e estão definidas estaticamente:

| Variável no Script | Uso |
|-------------| ------------- |
| $logo | Variável contendo a URL de uma imagem para usar como logotipo na assinatura |



[contributors-shield]: https://img.shields.io/github/contributors/PoBruno/AutomatedOutlookSignature.svg?style=for-the-badge
[contributors-url]: https://github.com/PoBruno/AutomatedOutlookSignature/graphs/contributors
[forks-shield]: https://img.shields.io/github/forks/PoBruno/AutomatedOutlookSignature.svg?style=for-the-badge
[forks-url]: https://github.com/PoBruno/AutomatedOutlookSignature/network/members
[stars-shield]: https://img.shields.io/github/stars/PoBruno/AutomatedOutlookSignature.svg?style=for-the-badge
[stars-url]: https://github.com/PoBruno/AutomatedOutlookSignature/stargazers
[issues-shield]: https://img.shields.io/github/issues/PoBruno/AutomatedOutlookSignature.svg?style=for-the-badge
[issues-url]: https://github.com/PoBruno/AutomatedOutlookSignature/issues
[license-shield]: https://img.shields.io/github/license/PoBruno/AutomatedOutlookSignature.svg?style=for-the-badge
[license-url]: https://github.com/PoBruno/AutomatedOutlookSignature/blob/master/LICENSE
