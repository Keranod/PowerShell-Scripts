#encrypted and could be used only with user it was created
Read-Host -AsSecureString | ConvertFrom-SecureString | Out-File -FilePath "C:\temp\psswd.encrypted"

#Creating encryption key
$EncryptionKeyBytes = New-Object Byte[] 32
[Security.Cryptography.RNGCryptoServiceProvider]::Create().GetBytes($EncryptionKeyBytes)
$EncryptionKeyBytes | Out-File "C:\temp\encryption.key"

#Encrypting with Encryption key
$EncryptionKeyData = Get-Content "C:\temp\encryption.key"
Read-Host -AsSecureString | ConvertFrom-SecureString -Key $EncryptionKeyData | Out-File -FilePath "C:\temp\psswd.encrypted"

#Read using key
$EncryptionKeyData = Get-Content "C:\temp\encryption.key"
$PasswordSecureString = Get-Content "C:\temp\psswd.encrypted" | ConvertTo-SecureString -Key $EncryptionKeyData
$PlainTextPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($PasswordSecureString))