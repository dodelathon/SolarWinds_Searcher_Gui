Param(
[Parameter(Mandatory=$true)][Int]$Col,
[Parameter(Mandatory=$true)][string]$FileName,
[Parameter(Mandatory=$true)][string]$SheetName,
[Int]$ColsAcross,
[Int]$RowsAcross,
[string]$Attribute
)

$Global:Attributes = ("Node Name", "IP Address", "IP Version", "DNS", "Machine Type", "Vendor", "Desciption", "Location", "Contact", "Status", "Software Image",
               "Software Version", "Asset_Environment", "Asset_Location", "Asset_Model", "Asset_State", "Cyber_Security_Classification", "Cybersecurity_Function",
               "Decomm_Year", "eFIApplication", "EOL_Date_HW", "EOL_Date_SW", "Hardware_Owner", "Holiday_Readiness", "Impact", "Imported_From_NCM", "InServiceDate",
               "Internet_Facing", "Legacy_Environment", "Local_Contact", "Management_Server", "Model_Category", "Network_Diagram", "Owner", "Physical_Host", "PONumber",
               "PurchaseDate", "PurchasePrice", "PurchasePrice_Maintenance", "QueueEmail", "Rack", "Rack_DataCenter", "Region", "Replacement_Cost", "Serial_Number", 
               "SNOW_Assignment_Group", "SNOW_Configuration_Item", "SNOW_Product_Name", "Splunk_Index", "Splunk_Sourcetypes", "Term_End_Date", "Term_Start_Date", "Vendor")

#Function to stop a process
Function Stop-SelectedProcess($Process) 
{
    if($Process -ne $null)
    {
        Get-Process $Process -ErrorAction SilentlyContinue | Stop-Process -ErrorAction SilentlyContinue
    }
}

Function Get-Sheet($Val, $Sheets, $SheetNames)
{        
    #Takes the sheet name in as input from the user, Loops until its valid of 
    for($i = 0; $i -lt $SheetNames.Count; $i++)
    {
        if($SheetNames[$i] -eq $Val)
        {
            $sh = $Sheets[$i]
            return $sh
        }        
    }
 
    return $sh                  
}


Function Get-Path($Val, $books)
{
    if($Val -like "*\*")
    {
        if(Test-Path($Val))
        {
            return $Val
        } 
    }
    else
    {
       foreach($b in $books)
       {
        `	if($b.name -eq $Val)
            {
                Write-Host $b 555
                $book = $b.name
            }
       }

       $Val = $global:File_Repo + '\' + $book;
       return $Val
    }
}

Function Ready-Components()
{
    #Closes Chrome and the webdriver to stop an error cause by chrome driver trying to open up the same session again   
    if(Get-Process | Where-Object{$_.Name -eq "ChromeDriver"})
    {
        Stop-SelectedProcess("ChromeDriver")
    }
    
    if(Get-Process | Where-Object{$_.Name -eq "GoogleChrome"})
    {
        Stop-SelectedProcess("GoogleChrome")
    }
 
    #Tells the webdriver to open chrome as a background process 
    $ChromeOptions = New-Object OpenQA.Selenium.Chrome.ChromeOptions
    $ChromeOptions.addArguments('headless')
    $Global:ChromeDriver = New-Object OpenQA.Selenium.Chrome.ChromeDriver($ChromeOptions)

    #Opens Excel as a background process
    $Global:excel = New-Object -comobject Excel.Application
    
    $tempArr = @()
    $tempArr += $Global:ChromeDriver
    $tempArr += $Global:excel
    return $tempArr
}

#Detects if an element is loaded on the page yet, REF: https://powershellone.wordpress.com/2015/02/12/waiting-for-elements-presence-with-selenium-web-driver/
Function isElementPresent($locator,[switch]$byClass,[switch]$byName)
{
    try
    {
        if($byClass)
        {
            $null=$ChromeDriver.FindElementByClassName($locator)
        }
        elseif($byName)
        {
            $null=$ChromeDriver.FindElementByName($locator)
        }
        else
        {
            $null=$ChromeDriver.FindElementById($locator)
        }
        return $true
    }
    catch
    {
        return $false
    }
}

#Auto search functionality to find the fist column containing the search word. The user can choose between either it auto searching for the serial numbers,
#or enter their own title along with a range in both the columns and rows to search.
Function AutoSearch-Column($Sheet, $title, $ColRange, $RowRange, [switch]$UserParams)
{
    $tempArr = @()
    $found = $false
    if($UserParams)
    {
        for($RowIter = 1; (($RowIter -lt $RowRange) -and ($found -eq $false)); $RowIter++)
        {
            for($ColIter = 1; (($ColIter -lt $ColRange) -and ($found -eq $false)); $ColIter++)
            {
                if($Sheet.Cells.Item($RowIter, $ColIter).Value2 -like $title)
                {
                    Write-Host ""
                    $found = $true
                    $tempArr += $ColIter
                    $tempArr += ($RowIter + 1)
                    Write-Host "Found at Row: $RowIter | Column: $ColIter "
                    Write-Host ""
                    return $tempArr
                }
            }
        }
        return $null
    }
    else
    {
        for($RowIter = 1; (($RowIter -lt $RowRange) -and ($found -eq $false)); $RowIter++)
        {
            for($ColIter = 1; (($ColIter -lt $ColRange) -and ($found -eq $false)); $ColIter++)
            {
                if($Sheet.Cells.Item($RowIter, $ColIter).Value2 -like $title)
                {
                    Write-Host ""
                    $found = $true
                    $tempArr += $ColIter
                    $tempArr += ($RowIter + 1)
                    Write-Host "Found at Row: $RowIter | Column: $ColIter "
                    Write-Host ""
                    return $tempArr
                }
            }
        }
        return $null
    }
}

Function Handle-AutoSearch($Col, $sh, $ChromeDriver)
{
    if($Col -eq 11)
    {
        Write-Host "Beginning Search for Column containing 'Serial'..."
        $tempArr = @()
        $tempArr = AutoSearch-Column $sh "*Serial*" 100 100
        if($tempArr -eq $null)
        {
            Stop-SelectedProcess "Excel"
            $ChromeDriver.Close()
            Stop-SelectedProcess "Chrome" 
            Stop-SelectedProcess"ChromeDriver" 
        }
        return $tempArr
    }
    elseif($Col -eq 12)
    {
        $valid = $false
        while($valid -eq $false)
        {
            $tempTitle = Read-Host "Enter the Title of the colum you want to search for"
            $TempColRange = Read-Host "Enter the amount of columns across you want to search (Higher the number, the longer it will take)"
            $TempRowRange = Read-Host "Enter the amount of Rows Deep you want to search (Higher the number, the longer it will take)"

            if(($tempTitle -ne $null -and $tempTitle -ne "") -and ($TempColRange -ne $null -and $TempColRange -ne "" ) -and ($TempRowRange -ne $null -and $TempRowRange -ne ""))
            {
                Write-Host "Beginning Search for Column containing '" $tempTitle "'..."
                $tempTitle = "*" + $tempTitle + "*"
                $tempArr = @()
                $tempArr = AutoSearch-Column $sh $tempTitle $TempColRange $TempRowRange -UserParams
                if($tempArr -eq $null)
                {
                    Write-Host "Column not found, Exiting..." -ForegroundColor Red
                    Stop-SelectedProcess "Excel"
                    $ChromeDriver.Close() # Close selenium browser session method
                    Stop-SelectedProcess("ChromeDriver") 
                    Stop-SelectedProcess("Chrome")
                    return -1 
                }
                else
                {
                    $valid = $true
                }
            }
            else
            {
                Write-Host "One of the above parameters is Empty, Please input all values!"
                Write-Host ""
            }
        }
        return $tempArr
    }
}

#Loads the Solarwinds page, Searches for the input boxes and enters the paramenters determind by the Passed into it, then determines whether or not it exists on solarwinds, Then writes the results to an Excel sheet
function SearchSolarWinds($Snum, $Attribute)
{
    #Navigates to the solarwinds page
    $ChromeDriver.Navigate().GoToURL("https://solarwindscs.dell.com/Orion/SummaryView.aspx?ViewID=1")

    $check = $false
    while($check -eq $false)
    {
        $check = isElementPresent 'ctl00$ctl00$ctl00$BodyContent$ContentPlaceHolder1$MainContentPlaceHolder$ResourceHostControl1$resContainer$rptContainers$ctl00$rptColumn1$ctl00$ctl01$Wrapper$txtSearchString' -byName
        sleep 1
    }
    
    #Selects the elements needed to interact with the page
    $TextBox = $ChromeDriver.FindElementByName('ctl00$ctl00$ctl00$BodyContent$ContentPlaceHolder1$MainContentPlaceHolder$ResourceHostControl1$resContainer$rptContainers$ctl00$rptColumn1$ctl00$ctl01$Wrapper$txtSearchString')
    $DropBox = $ChromeDriver.FindElementByName('ctl00$ctl00$ctl00$BodyContent$ContentPlaceHolder1$MainContentPlaceHolder$ResourceHostControl1$resContainer$rptContainers$ctl00$rptColumn1$ctl00$ctl01$Wrapper$lbxNodeProperty')
    $Btn = $ChromeDriver.FindElementsById('ctl00_ctl00_ctl00_BodyContent_ContentPlaceHolder1_MainContentPlaceHolder_ResourceHostControl1_resContainer_rptContainers_ctl00_rptColumn1_ctl00_ctl01_Wrapper_btnSearch')
    $SelectElement = [OpenQA.Selenium.Support.UI.SelectElement]::new($DropBox)

    #Inputs the data in the fields and selects the correct value from the dropdown
    $TextBox.SendKeys($Snum);
    $SelectElement.selectByValue($Global:Attributes[$Attribute])

    #Clicks search and loads the results page
    $Btn.Click()
    $check = $false

    #Calls the above method and loops until the page elements have loaded
    while($check -eq $false)
    {
        $check = isElementPresent 'StatusMessage' -byClass
        sleep 1
    }

    #Gets the result message from the page and then determines whether it exists, and whether or not there are duplicates if it does, then enters the data into the Excel file
    $Result = $ChromeDriver.FindElementByClassName('StatusMessage').Text
    if($Result -like "Nodes with* *similar to*")
    {
        
        $Amount = @() 
        $ResultSheet.Cells.Item($RowLabel, 2).Value2 = "Y"
        $Amount = $ChromeDriver.FindElementsByClassName('StatusIcon')
        $ResultSheet.Cells.Item($RowLabel, 3).Value2 = ($Amount.count - 1).ToString()
       
    }
    else
    {
       
        $ResultSheet.Cells.Item($RowLabel, 2).Value2 = "N"
        $ResultSheet.Cells.Item($RowLabel, 3).Value2 = "0"
    }
}

Function Ready-ResultSheet($wb)
{
    if(("Results" -notin $($wb.worksheets).Name))
    {
        $ResultSheet = $wb.worksheets.Add()
        $ResultSheet.Name = "Results"
    }
    else
    {
        $ResultSheet = $wb.worksheets | Where-Object {$_.Name -eq "Results"}
    }

    #Adds headers to the Result sheet
    $headerLine = 1
    $ResultSheet.Cells.Item($headerLine, $SerialCol).Value2 = "Serial_Number"
    $ResultSheet.Cells.Item($headerLine, $ExistsCol).Value2 = "On_Solarwinds?"
    $ResultSheet.Cells.Item($headerLine, $DupCol).Value2 = "Duplicates?"

    $ResultSheet.Cells.Item($headerLine, $SerialCol).Interior.ColorIndex = 6
    $ResultSheet.Cells.Item($headerLine, $ExistsCol).Interior.ColorIndex = 6
    $ResultSheet.Cells.Item($headerLine, $DupCol).Interior.ColorIndex = 6
    $ResultSheet.Cells.ColumnWidth = "On_Solarwinds?".Length
    return $ResultSheet
}

######################################  PROGRAM BEGINS HERE  ########################################################
Function Main($Col, $FileName, $SheetName, $ColsAcross, $RowsAcross, $Attribute)
{
    $user = $env:UserName #Getting The user
    $env:Path += ";C:\Users\$user\Desktop\SolarWinds_Searcher_Resources\Selenium" #Set up involved for use of the webdriver for chrome
    Add-Type -Path "C:\Users\$user\Desktop\SolarWinds_Searcher_Resources\Selenium\WebDriver.dll"
    Add-Type -Path "C:\Users\$User\Desktop\SolarWinds_Searcher_Resources\Selenium\WebDriver.Support.dll"
    $global:location = Get-Location #Gets the Path of this script
    $global:File_Repo = "C:\Users\$User\Desktop\Search_Repo"
    $items = Get-ChildItem $location #Gets All files and shortcuts surrounding the script

    #Initializaton of some variables
    $ExcelBooks = @()
    $global:Worksheets = @()
    $global:Sheets = @()
    $labeller = 0
    $startRow = 2
    $SerialCol = 1
    $ExistsCol = 2
    $DupCol = 3
    $RowLabel = 2


    $Col = [Int]$Col
    
    #Calls the Fetch-Data script with -FileName. This returns all the Excel files in the Search_Repo folder.
    $ExcelBooks = invoke-expression -Command "C:\Users\$User\Desktop\SolarWinds_Searcher_Resources\Fetch-Data.ps1 -FileName" 
    
    #Calls get path whick concatinates the path to the file name to pass as a complete path to the Excel construction.
    $FilePath = Get-Path $FileName $ExcelBooks
    if($FilePath -ne -1)
    {
        try
        {    
            $tempArr = @()
            $global:Sheets = invoke-expression -Command "C:\Users\$User\Desktop\SolarWinds_Searcher_Resources\Fetch-Data.ps1 -SheetName $FilePath"
            $global:SheetNames = invoke-expression -Command "C:\Users\$User\Desktop\SolarWinds_Searcher_Resources\Fetch-Data.ps1 -Name -SheetName $FilePath"
            $tempArr = Ready-Components
            $ChromeDriver = $tempArr[0]
            $excel = $tempArr[1]
            $wb = $excel.Workbooks.Open($FilePath)
            #$excel.Visible = $true

            #$sh = Get-Sheet $SheetName $global:Sheets $global:SheetName
            $sh = $wb.worksheets | Where-Object {$_.Name -eq $SheetName}

            $Value = $sh.Cells.Item($startRow, $Col).Value2
            
            $ResultSheet = Ready-ResultSheet $wb $global:SheetNames

           <# if(("Results" -notin $($wb.worksheets).Name))
            {
                Write-Output here
                $ResultSheet = $wb.worksheets.Add()
                $ResultSheet.Name = "Results"
                Write-Output here2
            }
            else
            {
                Write-Output "Found it"
                $ResultSheet = $wb.worksheets | Where-Object {$_.Name -eq "Results"}
            }

            #Adds headers to the Result sheet
            $headerLine = 1
            $ResultSheet.Cells.Item($headerLine, $SerialCol).Value2 = "Serial_Number"
            $ResultSheet.Cells.Item($headerLine, $ExistsCol).Value2 = "On_Solarwinds?"
            $ResultSheet.Cells.Item($headerLine, $DupCol).Value2 = "Duplicates?"

            $ResultSheet.Cells.Item($headerLine, $SerialCol).Interior.ColorIndex = 6
            $ResultSheet.Cells.Item($headerLine, $ExistsCol).Interior.ColorIndex = 6
            $ResultSheet.Cells.Item($headerLine, $DupCol).Interior.ColorIndex = 6
            $ResultSheet.Cells.ColumnWidth = "On_Solarwinds?".Length
            Write-Output "RSheet ready"
            Write-Output "$ResultsSheet"#>
             
            if($Col -eq 11 -or $Col -eq 12)
            {
               $tempArr = @()
               $tempArr = Handle-AutoSearch $Col $sh $ChromeDriver 
               $Col = $tempArr[0]
               $startRow = $tempArr[1]
            }
                   
            if($Col -ne 0)
            {
                #Reads the excel column containing the serial numbers, and searches Solarwinds for them
                Do
                {
                    $Value = $sh.Cells.Item($startRow, $Col).Value2
                    if($Value -notlike $null -or $Value -notlike "")
                    {
                        Write-Host "Searching for $Value"
                        $ResultSheet.Cells.Item($RowLabel, $SerialCol).Value2 = $Value
                        SearchSolarWinds $Value $Attribute
                    }
                    $startRow++
                    $RowLabel++
                } 
                while($Value -notlike $null -or $Value -notlike "")
                                           
                #Saves the Excel and then shows it to the user, Then proceeds to close the driver and Chrome
                $wb.save()
                $excel.Visible = $true
                $Global:ChromeDriver.Close() # Close selenium browser session method
                Stop-SelectedProcess "Chrome"
                Stop-SelectedProcess("ChromeDriver") 
            }                
        }
        catch
        {

            Write-Output $_.Exception.Message
            Write-output "Something has failed, Closing..."
            Stop-SelectedProcess("Excel")
        }
        finally
        {
            #$Global:ChromeDriver.Close() # Close selenium browser session method
            Stop-SelectedProcess( "Chrome")
            Stop-SelectedProcess("ChromeDriver")
        }
    }
    else
    {
        Write-output "Exiting"

    }
}
  

Function Process-Args($Col, $FileName, $SheetName, $ColsAcross, $RowsAccros, $Attribute)
{
    if(($Col -eq "" -or $Col -eq $null) -or ($FileName -eq "" -or $FileName -eq $null) -or ($SheetName -eq "" -or $SheetName -eq $null))
    {
        Write-output "Usage: SolarWindsSearcher -Col [Num > 0 and Num < 10 Unless 11 or 12(Activates autosearch)] -FileName [Name] -SheetName [Name] (Following are optional if Col = 12) -ColsAcross [Num] -RowsAcross [Num] -Attribute [Name]"
    }
    
    if($Attribute -eq $null -or $Attribute -eq "" )
    {
        $Attribute = 44
    }

    if($ColsAcross -eq $null -or $ColsAcross -eq "" )
    {
        $ColsAcross = 10
    }

    if($RowsAccros -eq $null -or $RowsAccros -eq "" )
    {
        $RowsAccros = 10
    }

    Main $Col $FileName $SheetName $ColsAcross $RowsAcross $Attribute
}

    
Process-Args $Col $FileName $SheetName $ColsAcross $RowsAcross $Attribute       


