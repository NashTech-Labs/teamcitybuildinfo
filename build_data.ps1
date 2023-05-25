param(
    [Parameter(Mandatory = $true)]
    [string] $TeamcityToken,
    [Parameter(Mandatory = $true)]
    [string] $Server_Url,
    [Parameter(Mandatory = $true)]
    [string] $startdate = "20230101" # equivalent to 1 january 2023
)

Get-Module ImportExcel -ListAvailable | Import-Module -Force -Verbose

function get_headers {
    $headers = @{
        Authorization="Bearer $TeamcityToken"
    }
    return $headers
}

function get_project_info {
    [CmdletBinding()]
    param (
        [string] $url = "https://teamcity.com",
        [string] $suburl = "/app/rest/projects/id:_Root"
    )
    $uri = $url + $suburl
    Write-Host $uri
    $headers = get_headers
    try {
        $data =Invoke-RestMethod -Uri $uri -Method Get -Headers $headers
    }
    catch {
        $data = $_.ErrorDetails
    }
    return $data    
}

$projects_xml_data = get_project_info -url $Server_Url -suburl "/app/rest/projects"
$projects = $projects_xml_data.projects.project

$Builds = @{
    Serial_No=""
    build_id=""
    build_number=""
    project_name=""
    project_id=""
    project_parentId=""
    project_webUrl=""
    build_url=""
    build_status=""
    build_finish_date=""
    build_finish_time=""
}

$count=0
foreach ($project_data in $projects) {
    Write-Host "--------------- Processing FOR $($project_data.name)-------"
    
    if($project_data.id -eq "_Root"){
        Write-Host "--------------SKIPPED Root--------------------"
        continue
    }
    
    $build_url = "/app/rest/builds?locator=project:$($project_data.id),sinceDate:$($startdate)T000000IST"
    $build_xml_data = get_project_info -url $Server_Url -suburl $build_url
    if(([string]$build_xml_data).Contains("404")  -or $build_xml_data.builds.count -eq "0"){
        continue
    }
    else{
        $build_data = $build_xml_data.builds.build
             
        foreach($build in $build_data){
            $count = $count + 1
            $Builds['Serial_No']=$count
            $Builds['build_id'] = $build.id
            $Builds['build_number'] = $build.number
            $Builds["project_name"] = $project_data.name
            $Builds["project_id"] = $project_data.id
            $Builds["project_parentId"] = $project_data.parentProjectId
            $Builds["project_webUrl"] = $project_data.webUrl
            $Builds["build_url"] = $build.webUrl
            $build_state = $build_data.state
            if($build_state -eq "running"){
                $Builds["build_status"] = "RUNNING"
                $Builds["build_finish_date"] = Get-Date
                $Builds["build_finish_date"] = $Builds["build_finish_date"].ToString("yyyy/MM/dd")
                $Builds["build_finish_time"] = Get-Date
                $BUilds["build_finish_time"] = $Builds["build_finish_time"].ToString("HH:mm:ss:ff:zzzz")
            }
            else{
                $Builds["build_status"] = $build_data.status
                $Builds["build_finish_date"] = [datetime]::ParseExact($build_data.finishOnAgentDate.Trim(),"yyyyMMddTHHmmsszzz",[Globalization.CultureInfo]::CurrentCulture).ToString("yyyy/MM/dd") 
                $Builds["build_finish_time"] = [datetime]::ParseExact($build_data.finishOnAgentDate.Trim(),"yyyyMMddTHHmmsszzz",[Globalization.CultureInfo]::CurrentCulture).ToString("HH:mm:ss:ff:zzzz")
            }
            $Builds
            # Sending data to json file
            $jsonfile = "$PSScriptRoot\build_data.json"
            $build_data_json = Get-Content $jsonfile | Out-String | ConvertFrom-Json
            $build_file_data = [System.Collections.ArrayList] $build_data_json
            $build_file_data.Add($Builds) | out-null
            $build_file_data | ConvertTo-Json | Set-Content $jsonfile 
        }
        
    }   
}


Get-Content "$PSScriptRoot\build_data.json" | ConvertFrom-Json | Export-Excel -Path "$PSScriptRoot\target\builds_data.xlsx"