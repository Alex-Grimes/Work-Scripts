$removal = @("user emails")







foreach ($i in $removal){


	$userUPN= $i
	$userList = Get-AzureADUser -ObjectID $userUPN
	$Skus = $userList | Select -ExpandProperty AssignedLicenses | Select SkuID
		if($userList.Count -ne 0) {
    			if($Skus -is [array])
    		{
        		$licenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
        		for ($i=0; $i -lt $Skus.Count; $i++) {
            		$Licenses.RemoveLicenses +=  (Get-AzureADSubscribedSku | Where-Object -Property SkuID -Value $Skus[$i].SkuId -EQ).SkuID   
        	}
        	Set-AzureADUserLicense -ObjectId $userUPN -AssignedLicenses $licenses
    	} else {
        	$licenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
        	$Licenses.RemoveLicenses =  (Get-AzureADSubscribedSku | Where-Object -Property SkuID -Value $Skus.SkuId -EQ).SkuID
       		Set-AzureADUserLicense -ObjectId $userUPN -AssignedLicenses $licenses
    		}
	}
}