Param(
	[Parameter(Mandatory=$true)]
	[String] $Department,
	[Parameter(Mandatory=$true)]
	[String] $Description

)

if($DisplayName -eq ""){

    Exit
}

#region Connections to MS Grapn and Azure AD

$testConnection = try{Get-AzureADDomain}Catch{}

if($testconnection -eq $null){

	Connect-AzureAD 

}

try{

    $testConnection = Get-AADGroup -Top 1

}
Catch{

    $conError = $_

    if($conError -like '*Not authenticated.  Please use the "Connect-MSGraph" command to authenticate.*'){

        Connect-MSGraph
    }
}

#endregion Connections

#Set Default Display Name
$Displayname = "Intune_W10_UG_$($Department)"
#set Default groups Array
$basegroups = @("Intune_W10_All_Users","Intune_W10_App_AdobeAcrobatReader","Intune_W10_App_CiscoJabber","Intune_W10_App_DisplayLinkGraphicsDriver","Intune_W10_App_GoogleChrome","Intune_W10_App_O365x86","Intune_W10_App_VMWareHorizon","Intune_W10_Compliance_Base","Intune_W10_PG_BaseBuild","Intune_W10_SU_Standard30DayDefer")
#Create Blank Array
$addGroup = @()


$AADGroup = Get-AzureADGroup -Filter "Displayname eq '$($DisplayName)'" #Check to see if group already exists


if(!$AADGroup){ #create group if it does not exist

    $AADGroup = New-AzureADGroup -DisplayName $DisplayName -Description $Description -SecurityEnable $true -MailEnabled $false -MailNickname "NotSet"
    Write-Output "Created group"
}


$groupMembersOf = Get-AADGroupMemberOf -groupId $AADGroup.Objectid #Get members of the group



foreach($basegroup in $basegroups){#nested for loop to check 2 arrays and  add to new array
    $check = 0
    
    foreach($group in $groupMembersOf){
        
        if($group.displayname -like $basegroup){
              $check = 1      
        }    
    }
    
    if($check -eq 0){
        $addGroup += $basegroup
    }     
}

Write-Output "Adding groups:"


foreach($add in $addGroup){ #get AzureGroups of the missing groups and add to existing group object

    $addGroupID = Get-AzureADGroup -Filter "DisplayName eq '$($add)'"

    Add-AzureADGroupMember -ObjectId $addGroupID.ObjectId -RefObjectId $AADGroup.ObjectId

    Write-Output $add
}


#Update-AADGroup -groupId $AADGroup.ObjectId -memberOf (Get-AADGroup -groupId $addGroupID.objectID)
