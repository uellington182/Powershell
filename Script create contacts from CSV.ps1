
clear-host
$mails = Import-Csv -encoding UTF8 "C:\Temp\Contatos-Cemig.csv"
clear-host
foreach ($mail in $mails)
    {
        $name = $mail.DisplayName
        $SMTP = $mail.PrimarySmtpAddress
        $nameAD = Get-ADObject -Filter * -SearchBase "CN=$name,OU=CEMIG,OU=CONTACTS,OU=OFFICE365,OU=UsuariosForluz,OU=FORLUZ,DC=forluznet,DC=net" -Properties *
        if ( $nameAD.DisplayName -eq $name)
            {
                $mail.DisplayName >> C:\temp\naocriados.txt
            }else{
                New-ADObject -type contact -path "OU=CEMIG,OU=CONTACTS,OU=OFFICE365,OU=UsuariosForluz,OU=FORLUZ,DC=forluznet,DC=net"  -Name $name -otherAttributes @{'displayName'=$name;'mail'=$SMTP;'extensionAttribute10'='CEMIG_SA'}
                $mail.DisplayName >> C:\temp\criados.txt
            }
    }
