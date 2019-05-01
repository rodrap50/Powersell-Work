Param(

	[Parameter(Mandatory=$true)]
	[string]$domain,
	[Parameter(Mandatory=$true)]
	[string]$secGroup

)
    [array]$final = $null
    
    $gpos = Get-GPO -All -Domain $domain
 
    foreach ($gpo in $gpos) 
    { 
        $secinfo = $gpo.GetSecurityInfo() 
        foreach ($sec in $secinfo) 
        { 
            if ($sec.Trustee.Name -eq $secGroup) 
            { 
                $final += $gpo 
            } 
        } 
    } 
 
