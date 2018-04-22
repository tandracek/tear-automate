Param(
    [string]$disc,
    [string]$out,
    [string]$finalName,
    [string]$title,
    [string]$url,
    [string]$username,
    [string]$password
)

function runProgram {
    Param ([string] $command, [string] $program, [boolean] $suppress)

    Try {
        if ($suppress) {
            Invoke-Expression $command | Out-Null
        } else {
            Invoke-Expression $command
        }        
    } Catch [System.Management.Automation.CommandNotFoundException] {
        Write-Host "Could not find $program.  Make sure it is installed and set on the path."
        exit
    } Catch {
        $ErrorMessage = $_.Exception.Message
        Write-Host "Error occured running $program.`n$ErrorMessage"
        exit
    }
}

function checkRequiredArgument {
    Param ([string] $passed, [string] $argument)

    if (!$passed) {
        Write-Host "Did not find required argument: $argument"
        Write-Host "Usage: tear.ps1 [options]"
        Write-Host "Options:"
        Write-Host "--disc: required disc number"
        Write-Host "--out: required location to write mkv to"
        Write-Host "--title: optional title number, otherwise rips all"
        Write-Host "--url: optional remote location to copy to"
        Write-Host "--username: username for the remote location"
        Write-Host "--password: password for the remote location"
        exit
    } 
}

#Test that the required programs exist and correctly set on the path
runProgram -command "makemkvcon" -program "MakeMkv" -suppress true
runProgram -command "HandBrakeCLI" -program "HandBrake" -supress true
runProgram -command "pscp.exe" -program "Putty" -supress true

#Check if provided correct arguments for remote copy
if ($url) {
    if (!$username -or !$password) {
        Write-Host "Must provide username and password if copying to remote location"
        exit
    }
}

#Check that all the required arguments were passed
checkRequiredArgument -passed $disc -argument "disc"
checkRequiredArgument -passed $out -argument "out"

$directory = Split-Path $out
$filename = Split-Path $out -Leaf
$extension = $filename.Split(".")[1]
if ($extension -ne "mkv") {
    Write-Host "Must provide full path and name to a mkv: $out"
    exit
}

if (!$title) {
    $title="all"
}
Write-Host "Ripping title $title on disc $dist to location $directory."
runProgram -command "makemkvcon -r mkv disc:$disc $title $directory" -program "MakeMKV"

$mkvFilename = Get-ChildItem $directory\*.mkv | Sort-Object CreationTime -Descending | Select-Object -First 1
Write-Host "Running handbrake for file $mkvFilename and saving it to $filename"
runProgram -command "HandBrakeCLI -i $mkvFilename -o $out" -program "HandBrake"

Write-Host "Copying file $out to remote location $url"
runProgram -command "pscp.exe -pw $password $out $username@$url"

Write-Host "All done!"