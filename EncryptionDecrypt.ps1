<#
AES encryption and decryption sample
#>
$userName = "Shubham"
#$password = "2ytcub5431783!@"
$password = "ExamplePassword"

Function GetAESObject()
{
    $aes = New-Object System.Security.Cryptography.AesManaged
    $aes.Mode = [System.Security.Cryptography.CipherMode]::CBC
    $aes.Padding = [System.Security.Cryptography.PaddingMode]::Zeros
    $aes.BlockSize = 128
    $aes.KeySize = 256
    return $aes
}

Function EncryptData($aes, $key, $plainText)
{
    $encryptor = $aes.CreateEncryptor()
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($plainText)
    $encryptedValue = $encryptor.TransformFinalBlock($bytes, 0, $bytes.length)
    [byte[]] $fullValue = $aes.IV + $encryptedValue
    $encryptedValue = [System.Convert]::ToBase64String($fullValue)
    return $encryptedValue
}

Function DecryptData($aes, $key, $encryptedValue)
{
    $bytes = [System.Convert]::FromBase64String($encryptedValue)
    $iv = $bytes[0..15]
    $aes.iv = $iv
    $aes.Key = [System.Convert]::FromBase64String($key)
    $decryptor = $aes.CreateDecryptor()
    $decryptedValue = $decryptor.TransformFinalBlock($bytes, 16, $bytes.length - 16)
    $decryptedValue = [System.Text.Encoding]::UTF8.GetString($decryptedValue).Trim([char]0)
    return $decryptedValue
}

#Get the AES object and generate the random key
$aes = GetAESObject
$aes.GenerateKey()
$key = [System.Convert]::ToBase64String($aes.Key)
#Encrypt the values
#$encryptedUserName = EncryptData $aes $key $userName
#$encryptedPassword = EncryptData $aes $key $password
$encryptedUserName = EncryptData $aes "9CohqCiVP8Nt6IAM223/39Wucp2du+sPGZDW90kDMQc=" $userName
$encryptedPassword = EncryptData $aes "9CohqCiVP8Nt6IAM223/39Wucp2du+sPGZDW90kDMQc=" $password
$aes.Dispose()
#Decrypt the values
$aes = GetAESObject
#$decryptedUserName = DecryptData $aes $key $encryptedUserName
#$decryptedPassword = DecryptData $aes $key $encryptedPassword
$decryptedUserName = DecryptData $aes "9CohqCiVP8Nt6IAM223/39Wucp2du+sPGZDW90kDMQc=" $encryptedUserName
$decryptedPassword = DecryptData $aes "9CohqCiVP8Nt6IAM223/39Wucp2du+sPGZDW90kDMQc=" $encryptedPassword
$aes.Dispose()
Write-Host "Plain text Username:" $userName
Write-Host "Plain text Password:" $password
Write-Host "Key:" $key
Write-Host "Encrypted Username:" $encryptedUserName
Write-Host "Encrypted Password:" $encryptedPassword
Write-Host "Decrypted Username:" $decryptedUserName
Write-Host "Decrypted Password:" $decryptedPassword
