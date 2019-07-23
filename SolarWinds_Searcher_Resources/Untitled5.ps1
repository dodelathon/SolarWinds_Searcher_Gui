Function Stop-SelectedProcess($Process) 
{
    $processess = Get-Process
    foreach($p in $processess)
    {
        if($p -ne $null -and $p -eq $Process )
        {
            Get-Process $Process -ErrorAction SilentlyContinue | Stop-Process -ErrorAction SilentlyContinue
        }
    }
}