$removal = @()


foreach ($i in $removal){


    $userUPN= $i
    $planName="VISIOCLIENT"
    $license = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
    $License.RemoveLicenses = (Get-AzureADSubscribedSku | Where-Object -Property SkuPartNumber -Value $planName -EQ).SkuID
    Set-AzureADUserLicense -ObjectId $userUPN -AssignedLicenses $license
}