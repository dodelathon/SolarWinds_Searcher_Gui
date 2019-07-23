Function Stop-SelectedProcess($Process) 
{
    
    $processess = Get-Process
    foreach($p in $processess)
    {
        
       # Write-Host $p.ProcessName
        if($p -ne $null -and $p.ProcessName -eq $Process )
        {
            Get-Process $Process -ErrorAction SilentlyContinue | Stop-Process -ErrorAction SilentlyContinue
        }
    }
}

Stop-SelectedProcess "Excel"