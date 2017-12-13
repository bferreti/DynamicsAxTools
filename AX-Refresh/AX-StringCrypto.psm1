function Write-EncryptedString {
param ([String]$InputString, [String]$DTKey, [Switch]$Compress, [Switch]$GnuPG, [String]$Recipient)

if (($args -contains '-?') -or (-not $InputString) -or (-not $DTKey -and -not $GnuPG)) {
'NAME'
'    Write-EncryptedString'
''
'SYNOPSIS'
'    Encrypts a string using another string as a password.'
''
'SYNTAX'
'    Write-EncryptedString [-inputString] <string> [-password] <string>'
'                          [-compress]'
'    Write-EncryptedString [-inputString] <string> [[-password] <string>] -gnupg'
'    Write-EncryptedString [-inputString] <string> -gnupg -recipient <string>'
''
'PARAMETERS'
'    -inputString <string>'
'        The string to be encrypted.'
''
'    -password <string>'
'        The password to use for encryption.'
''
'    -compress'
'        Compress data before encryption.'
'        Is not used with GnuPG.'
''
'    -gnupg'
'        Perform encryption using GnuPG. Requires gpg.exe to alias as gpg.'
'        Can only use one single line passphrase.'
''
'    -recipient'
'        Available only with GnuPG. This is the UID search string to be used'
'        with asymmetric encryption.'
'RETURN TYPE'
'    string'
return
}

	if ($GnuPG) {
		if ($Recipient) {
			$InputString | gpg --encrypt --recipient $Recipient --armor --quiet --batch | Join-String -newline
		}
		elseif ($DTKey) {
			$DTKey, $InputString | gpg --symmetric --armor --quiet --batch --passphrase-fd 0 | Join-String -newline
		}
		else {
			$InputString | gpg --symmetric --armor | Join-String -newline
		}
	}
	else {
		$Rfc2898 = New-Object System.Security.Cryptography.Rfc2898DeriveBytes($DTKey,32)
		$Salt = $Rfc2898.Salt
		$AESKey = $Rfc2898.GetBytes(32)
		$AESIV = $Rfc2898.GetBytes(16)
		$Hmac = New-Object System.Security.Cryptography.HMACSHA1(,$Rfc2898.GetBytes(20))
		
		$AES = New-Object Security.Cryptography.RijndaelManaged
		$AESEncryptor = $AES.CreateEncryptor($AESKey, $AESIV)
		
		$InputDataStream = New-Object System.IO.MemoryStream
		if ($Compress) { $InputEncodingStream = (New-Object System.IO.Compression.GZipStream($InputDataStream, 'Compress', $True)) }
		else { $InputEncodingStream = $InputDataStream }
		$StreamWriter = New-Object System.IO.StreamWriter($InputEncodingStream, (New-Object System.Text.Utf8Encoding($true)))
		$StreamWriter.Write($InputString)
		$StreamWriter.Flush()
		if ($Compress) { $InputEncodingStream.Close() }
		$InputData = $InputDataStream.ToArray()
		
		$EncryptedEncodedInputString = $AESEncryptor.TransformFinalBlock($InputData, 0, $InputData.Length)
		
		$AuthCode = $Hmac.ComputeHash($EncryptedEncodedInputString)
		
		$OutputData = New-Object Byte[](52 + $EncryptedEncodedInputString.Length)
		[Array]::Copy($Salt, 0, $OutputData, 0, 32)
		[Array]::Copy($AuthCode, 0, $OutputData, 32, 20)
		[Array]::Copy($EncryptedEncodedInputString, 0, $OutputData, 52, $EncryptedEncodedInputString.Length)
		
		$OutputDataAsString = [Convert]::ToBase64String($OutputData)
		
		$OutputDataAsString
	}
}

function Read-EncryptedString {
param ([String]$InputString, [String]$DTKey)

if (($args -contains '-?') -or (-not $InputString) -or (-not $DTKey -and -not $InputString.StartsWith('-----BEGIN PGP MESSAGE-----'))) {
'NAME'
'    Read-EncryptedString'
''
'SYNOPSIS'
'    Decrypts a string using another string as a password.'
''
'SYNTAX'
'    Read-EncryptedString [-inputString] <string> [-password] <string>'
''
'PARAMETERS'
'    -inputString <string>'
'        The string to be decrypted.'
''
'    -password <string>'
'        The password to use for decryption.'
''
'RETURN TYPE'
'    string'
return
}

	if ($InputString.StartsWith('-----BEGIN PGP MESSAGE-----')) {
		# Decrypt with GnuPG
		if ($DTKey) {
			$DTKey, $InputString | gpg --decrypt --quiet --batch --passphrase-fd 0 | Join-String -newline
		}
		else {
			$InputString | gpg --decrypt | Join-String -newline
		}
	}
	else {
		# Decrypt with custom algo
		$InputData = [Convert]::FromBase64String($InputString)
		
		$Salt = New-Object Byte[](32)
		[Array]::Copy($InputData, 0, $Salt, 0, 32)
		$Rfc2898 = New-Object System.Security.Cryptography.Rfc2898DeriveBytes($DTKey,$Salt)
		$AESKey = $Rfc2898.GetBytes(32)
		$AESIV = $Rfc2898.GetBytes(16)
		$Hmac = New-Object System.Security.Cryptography.HMACSHA1(,$Rfc2898.GetBytes(20))
		
		$AuthCode = $Hmac.ComputeHash($InputData, 52, $InputData.Length - 52)
		
		if (Compare-Object $AuthCode ($InputData[32..51]) -SyncWindow 0) {
			throw 'Checksum failure.'
		}
		
		$AES = New-Object Security.Cryptography.RijndaelManaged
		$AESDecryptor = $AES.CreateDecryptor($AESKey, $AESIV)
		
		$DecryptedInputData = $AESDecryptor.TransformFinalBlock($InputData, 52, $InputData.Length - 52)
		
		$DataStream = New-Object System.IO.MemoryStream($DecryptedInputData, $false)
		if ($DecryptedInputData[0] -eq 0x1f) {
			$DataStream = New-Object System.IO.Compression.GZipStream($DataStream, 'Decompress')
		}
		
		$StreamReader = New-Object System.IO.StreamReader($DataStream, $true)
		$StreamReader.ReadToEnd()
	}
}
