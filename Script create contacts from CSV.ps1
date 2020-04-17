#Programador: Uellington
#Data: 17/04/2020
#Versão: 1.0

clear-host
$mails = Import-Csv -encoding UTF8 "C:\Temp\Contatos.csv"
clear-host
foreach ($mail in $mails)
    {
        $name = $mail.DisplayName
        $SMTP = $mail.PrimarySmtpAddress
        $nameAD = Get-ADObject -Filter * -SearchBase "CN=$name,OU=CLIENTE,OU=CONTACTS,OU=OFFICE365,OU=UsersEmpresa,OU=empresa,DC=empresa,DC=net" -Properties *
        if ( $nameAD.DisplayName -eq $name)
            {
                $mail.DisplayName >> C:\temp\naocriados.txt
            }else{
                New-ADObject -type contact -path "OU=CLIENTE,OU=CONTACTS,OU=OFFICE365,OU=UsersEMPRESA,OU=EMPRESA,DC=empresa,DC=net"  -Name $name -otherAttributes @{'displayName'=$name;'mail'=$SMTP;'extensionAttribute10'='ATRIBUTO'}
                $mail.DisplayName >> C:\temp\criados.txt
            }
    }
