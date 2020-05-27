# Function to update Outlook signature

function Update-Signature {

    # Get username
    $strName =  $env:username
	
    # Collect AD user attributes
    $strFilter = "(&(objectCategory=User)(samAccountName=$strName))"
    $objSearcher = New-Object System.DirectoryServices.DirectorySearcher
    $objSearcher.Filter = $strFilter
    $objPath = $objSearcher.FindOne()
    $objUser = $objPath.GetDirectoryEntry()

	# Assign variables to AD user attributes
	$strFirst = $objUser.FirstName.ToString()
	$strLast = $objUser.LastName.ToString()
    $strTitle = $objUser.Title.ToString()
    $strPhone = $objUser.telephoneNumber.ToString()
    $strEmail = $objUser.mail.ToString()

    # User signature creation & styling
    $signaturehtml = @"
    <head>
    <style>
    /* Style Definitions */
    p.MsoNormal, li.MsoNormal, div.MsoNormal
	{mso-style-unhide:no;
     mso-style-qformat:yes;
     mso-style-parent:"";
     margin:0cm;
     margin-bottom:.0001pt;
     mso-pagination:widow-orphan;
     font-size:11.0pt;
     font-family:"Calibri",sans-serif;
     mso-ascii-font-family:Calibri;
     mso-ascii-theme-font:minor-latin;
     mso-fareast-font-family:"Times New Roman";
     mso-fareast-theme-font:minor-fareast;
     mso-hansi-font-family:Calibri;
     mso-hansi-theme-font:minor-latin;
     mso-bidi-font-family:"Times New Roman";
     mso-bidi-theme-font:minor-bidi;}
     p {
       margin-top: 0px;
       line-height: 0px;
     }
     span {
       line-height: 0px;
     }
     </style>
     </head>
	 
	 <!-- User full name -->
     <body lang=EN-GB link=#023a98 vlink=#023a98 style='tab-interval:0.0pt'>
     <p class=MsoNormal><b><span
     style='font-size:11.0pt;font-family:"Calibri",sans-serif;
     color:#002d6a;LINE-HEIGHT:1pt'>$strFirst $strLast</span></b></p>
     
	 <!-- User title & company -->
	 <p class=MsoNormal><span
     style='font-size:10.0pt;font-family:"Calibri",sans-serif;color:#000000;LINE-HEIGHT:1pt'>
     $strtitle</span></p>
	 
	 <!-- User phone number -->
	 <p class=MsoNormal><span
     style='font-size:10.0pt;font-family:"Calibri",sans-serif;color:#000000;LINE-HEIGHT:1pt'>
	 $strPhone</span></p>
	 
	 <!-- User email address -->
	 <p class=MsoNormal><span
     style='font-size:10.0pt;font-family:"Calibri",sans-serif;color:#023a98;LINE-HEIGHT:1pt'><a href="mailto:$stremail">
	 $stremail</span></p>

     <font style="font-size: 5pt"><br></font>

	 <!-- Embed encoded base64 signature image -->
     <img border=0
     src="data:image/png;base64,##############PASTE BASE64 ENCODED IMAGE HERE. I use "https://www.base64-image.de/" ##############"</p>
	  </body>
      </html>
"@

    # Output the file to Outlook signature location.
    $signaturehtml | Out-File "$FolderLocation\$env:username.htm"
	
	#Set the signature in outlook as default.
    $MSWord = New-Object -com word.application
    $EmailOptions = $MSWord.EmailOptions
    $EmailSignature = $EmailOptions.EmailSignature
    $EmailSignatureEntries = $EmailSignature.EmailSignatureEntries
    $EmailSignature.NewMessageSignature = $env:username
    $MSWord.Quit()
	
}

#Set export path for signature, force dir creation
$UserDataPath = $Env:appdata
$FolderLocation = $UserDataPath + '\Microsoft\Signatures'
mkdir $FolderLocation -force

#Exit if username.htm already exists
if ((Test-Path -Path $FolderLocation\$env:username.htm)){exit}

#Invoke update-signature function
Update-Signature
