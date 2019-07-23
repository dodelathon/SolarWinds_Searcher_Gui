#Passable parameters to the script.
Param(
[switch]$FileName,
[switch]$SheetName,
[switch]$Name,
[String]$FilePath
)

#Results array
$Global:Results = @()

#Function to stop a process that is no longer needed
Function Stop-SelectedProcess($Process) 
{
    if($Process -ne $null)
    {
        Get-Process $Process -ErrorAction SilentlyContinue | Stop-Process -ErrorAction SilentlyContinue
    }
}

#Returns a list of the files in the search_repo folder to the the program that calls it.
Function Get-Files()
{
    $user = $env:UserName
    $items = Get-ChildItem "C:\Users\$User\Desktop\Search_Repo"

    #Finds all the items in the same directory as the script, and then selects the Excel files from there
    $ExcelBooks = @()
    foreach( $f in $items)
    {
        if($f -ne $null -and $f.'name' -like "*.xls*")
        {
             $Global:Results += $f
        }
    }
    Stop-SelectedProcess "Excel"
    if($Global:Results[0] -eq $null -or $Global:Results[0] -eq "")
    {
        Write-Output -1
    }
    else
    {
        Write-Output $Global:Results
    }
}

Function Get-Sheets()
{     
    $excel = New-Object -comobject Excel.Application
    $wb = $excel.Workbooks.Open($FilePath)
    
    $items = $wb.worksheets
        
    foreach( $f in $items)
    { 
        $Global:Results += $f
    }

    $excel.Close()
    Write-Output $Global:Results
}

Function Get-SheetNames()
{     
    $excel = New-Object -comobject Excel.Application
    $wb = $excel.Workbooks.Open($FilePath)
    
    $items = $wb.worksheets
        
    foreach( $f in $items)
    { 
        $Global:Results += $f.Name
    }

    Stop-SelectedProcess "Excel"
    Write-Output $Global:Results
}



if($FileName)
{
    Get-Files
}
elseif($SheetName -and $Name)
{
    Get-SheetNames
}
elseif($SheetName)
{
    Get-Sheets
}
else
{
    Write-Output "Script called at least"
}