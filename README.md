[![Contributors][contributors-shield]][contributors-url]
[![Forks][forks-shield]][forks-url]
[![Stargazers][stars-shield]][stars-url]
[![Issues][issues-shield]][issues-url]
[![MIT License][license-shield]][license-url]

# Automated Outlook Signature Scripts
This project contains two scripts: 
* Set-OutlookSignature.ps1 - Used to generate and set a user's signature for desktop Outlook.
* Set-OutlookWebSignatures.ps1 - This script is Currently a work in progress.

Outlook desktop signature script currently tested and working with Outlook 2010, 2016, 2019, 2021.

## Changes
* Breaking change of the scipt being renamed to Set-OutlookSignature.ps1
* Breaking change - Company now takes the value from Active Directory
* Added some better practice
* Added more error handling

## The Scripts
This is a very basic description on how to use the scripts and how they work. For more detail please see the YouTube videos linked earlier 

I recommend using the script in Group Policy as a logon script. If you're unfamiliar with this process, you can follow the detailed instructions provided in this article - [Configuring Logon PowerShell Scripts with Group Policy - 4Sysops](https://4sysops.com/archives/configuring-logon-powershell-scripts-with-group-policy/)

During the user's logon process, the script runs in the background, retrieves the necessary user details, generates a new signature file, and replaces the existing one. Additionally, the script sets registry keys to configure the newly created signature as the user's default Outlook signature. This ensures that if any details such as job title change, the signature will be automatically updated during the next logon.

[EduGeek Post](http://www.edugeek.net/forums/scripts/205976-outlook-email-signature-automation-ad-attributes.html#post1760284)

### Active Directory
A selection of Active Directory attribute are already configured in the script and listed below however more attributes can be easily added. 

The following properties are used from Active Directory within the script:

| Variable in Script | AD Field  | Notes |
|-------------| ------------- | ------------- |
| $displayName | Display name | Users display name |
| $jobTitle | Job title | Users job title |
| $email | Email | Users email address  |
| $telephone | Telephone  | The main site/branch telephone number |
| $directDial | Home | The users direct dial number |
| $mobileNumber | Mobile | The users mobile number |
| $street | Street | Street / First line of address |
| $poBox | P.O. Box | Site / Branch name which will appear in bold above the address e.g. Head Office |
| $city | City | City / Town |
| $state | State/Province | State / County |
| $zipCode | Zip/Postal Code | Post Code / Zip Code |
| $office | physicaldeliveryofficename | Office |
| $website | Website | Website address |
| $companyName | company | The name of the company |

Additional variables that do not rely on Active Directory and are currently set statically

| Variable in Script | Usage |
|-------------| ------------- |
| $logo | Variable containing the URL of a image to use as a logo in the signature |


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
