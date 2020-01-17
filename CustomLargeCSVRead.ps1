
# will need to create custom class to store Header's
class CSV {
    [String]$File_Server_Domain
    [String]$Operation_By
    [String]$Path
    [String]$Object
    [String]$Event_Description
    [String]$Device_IP_Address
    [String]$Device_Name
    [String]$Event_Count
    [String]$Event_Time
    [String]$Last_Occurrence
    [String]$File_Type
    [String]$Event_Operation
    [String]$Operation_Source
    [String]$Event_Type
    [String]$Mail_Source
    [String]$Mail_Recipients
    [String]$Mail_Date
    [String]$Attachment_Name
    [String]$Object_Type
    [String]$Event_Status
}

$reader = [System.IO.File]::OpenText('c:\temp\SMB.csv') #open IO to csv

$data = new-object CSV #create new CSV Object

$line = $reader.ReadLine() # read in first line ( headers)

$finalData = New-Object -TypeName "System.Collections.ArrayList"
for(;;){  #process

    $tempData = $null
    $data = new-object CSV
    $line = $reader.ReadLine()


    if ($null -eq $line) {
        break
    }

    [System.Collections.ArrayList]$tempData = $line.split(',')

    for($i= 0; $i -lt $tempData.count; $i++){

        if($tempData[$i] -like '"*'){

            $tempData[$i] = $tempData[$i] + $tempData[$i+1]
            $tempData.RemoveAt($i+1)

        }

    }

    for ($i= 0; $i -lt $tempData.count; $i++){
        switch ($i) {
            0 {$data.File_Server_Domain = $tempData[$i]; break}
            1 {$data.Operation_By = $tempData[$i]; break}
            2 {$data.Path = $tempData[$i]; break}
            3 {$data.Object = $tempData[$i]; break}
            4 {$data.Event_Description = $tempData[$i]; break}
            5 {$data.Device_IP_Address = $tempData[$i]; break}
            6 {$data.Device_Name = $tempData[$i]; break}
            7 {$data.Event_Count = $tempData[$i]; break}
            8 {$data.Event_Time = $tempData[$i]; break}
            9 {$data.Last_Occurrence = $tempData[$i]; break}
            10 {$data.File_Type = $tempData[$i]; break}
            11 {$data.Event_Operation = $tempData[$i]; break}
            12 {$data.Operation_Source = $tempData[$i]; break}
            13 {$data.Event_Type = $tempData[$i]; break}
            14 {$data.Mail_Source = $tempData[$i]; break}
            15 {$data.Mail_Recipients = $tempData[$i]; break}
            16 {$data.Mail_Date = $tempData[$i]; break}
            17 {$data.Attachment_Name = $tempData[$i]; break}
            18 {$data.Object_Type = $tempData[$i]; break}
            19 {$data.Event_Status = $tempData[$i]; break}

            Default {break}
        }
    }
    $finalData.add($data) 
}
$reader.Close() #close IO