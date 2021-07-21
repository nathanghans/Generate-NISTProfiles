#css
$css = @"
<style>
h1 {

        font-family: Arial, Helvetica, sans-serif;
        color: #e68a00;
        font-size: 28px;
    }

    
    h2 {

        font-family: Arial, Helvetica, sans-serif;
        color: #000099;
        font-size: 20px;

    }
table {
		font-size: 14px;
		border: 2px solid black; 
        border-collapse: collapse; 
		font-family: Arial, Helvetica, sans-serif;
	} 
	
    td {
		padding-top: 15px;
        padding-bottom: 15px;
        padding-left: 5px;
		margin: 0px;
		border: 2px solid black; 
	}
	
    th {
        background: #002061;
        color: #FFF;
        font-size: 16px;
        text-transform: uppercase;
        padding: 10px 15px;
        vertical-align: middle;
	}

    tbody tr:nth-child(even) {
        background: #f0f0f2;
    }
    .Identify {
        background: #006fbf;
    }
    .Protect {
        background: #6f30a0;
    }
    .Detect {
        background: #ffff00;
    }
    .Respond {
        background: #ff0000;
    }
    .Recover {
        background: #00b04f;
    }
</style>
"@
#path to files
$basepath = 'C:\Users\nate\Documents\GitHub\Powershell\'
#output path
$outputpath = "C:\Users\nate\Documents\GitHub\Powershell\NIST\"
#get CSF to 800 mappings
$csfto800csv = 'csf-pf-to-sp800-53r5-mappings.csv'
#get Control Catalog
$80053csv ='sp800-53r5-control-catalog.csv'
#get 800 53 b
$80053bcsv ='sp800-53b.csv'

#import csvs
$csfto800 = import-csv $($basepath + $csfto800csv)
$80053b = import-csv $($basepath + $80053bcsv)
$80053 = Import-Csv $($basepath + $80053csv)

$counter=0
#Cycle through CSF Categories
foreach($subcat in $csfto800)
{
    #new ps object to hold category information
    $subcategoryobject = New-Object -TypeName PSObject    
    #if there is a function field, set it.
    if($subcat.Function)
    {
        $subcatfunction = $subcat.Function
        $subcatclass = $subcatfunction.split(" ")[0]
    }
    #if there is a category field, set it. 
    if($subcat.Category)
    {
        $subcatcat = $subcat.Category
    }
    #gather subcategory
    $start = 0 #$subcat.Subcategory.IndexOf(".")
    $end = $subcat.Subcategory.IndexOf(":")
    $end = $end - $start
    #add the subcategory to an array
    #Not Used $Subcats += $subcat.Subcategory.Substring($start+1,$end-1)
    #get the sub category name for the HTML header
    $subcatname = $subcat.Subcategory.Substring($start,$end) 
    
    #$HashTable.$subcatname = @()
    #get the subcategory name
    $subcategoryname = $subcat.Subcategory.Split(":")[0]
    #get subcategory description
    $subcategorydescription = $subcat.Subcategory.Split(":")[1]
    #get controls for current subcategory
    $currentcontrols = $subcat.'NIST SP 800-53, Revision 5 Control'.Split(",").trim()   
    #Add members to Subcategory information PS Object
    $subcategoryobject | Add-Member -MemberType NoteProperty -Name Function -Value $subcatfunction
    $subcategoryobject | Add-Member -MemberType NoteProperty -Name Category -Value $subcatcat
    $subcategoryobject | Add-Member -MemberType NoteProperty -Name Subcategory -Value $subcategoryname
    $subcategoryobject | Add-Member -MemberType NoteProperty -Name Description -Value "$subcategorydescription"
    #convert subcategory ps object into HTML for report
    $subcatreport =  $subcategoryobject | ConvertTo-Html -Fragment -PreContent "<h2>$subcatname</h2>"
    #Add an HTML class attribute on the table for css format matching(Makes the cell color match the NIST colors)
    $subcatreport = $subcatreport -replace "<td>$subcatclass","<td class=""$subcatclass"">$subcatclass"

    #set arrays.
    $controltablearray = @()
    $currentmaturityarray = @()
    $futurematurityarray = @()
    #loop through controls for current subcategory
    foreach($currentcontrol in $currentcontrols)
    {
        #set HTML output name
        $cathtml = $outputpath + $subcatname + "-Low.html"
        #create hashtable with current control.
        $HashTable = @{}
        $hashtable.$currentcontrol #+= $currentcontrol
        #get controls that are low
        $control = $80053b | where {$_.'Control Identifier' -like "$currentcontrol*" -and ($_.'Security Control Baseline Low' -like 'N/A - Deployed organiation-wide' -or $_.'Security Control Baseline Low' -like '*x*')} | select 'Control Identifier'
        #$control = $80053b | where {$_.'Control Identifier' -like "$currentcontrol*" -and $_.'Security Control Baseline Low' -like '*x*'} | select 'Control Identifier'
        #loop through each control
        foreach($controlitems in $control)
        {
            #get control ID, name, text, and discussion
            $controloutput = $80053 | where {$_.'Control Identifier' -eq $controlitems.'Control Identifier'} | select 'Control Identifier', 'Control (or Control Enhancement) Name', 'Control Text', 'Discussion'
            #create new object for control information
            $controltableobject= New-Object -TypeName PSObject    
            $controltableobject| Add-Member -MemberType NoteProperty -Name "Control Identifier" -Value $controloutput.'Control Identifier'
            $controltableobject| Add-Member -MemberType NoteProperty -Name "Control Name" -Value $controloutput.'Control (or Control Enhancement) Name'
            $controltableobject| Add-Member -MemberType NoteProperty -Name "Control Text" -Value $controloutput.'Control Text'
            $controltableobject| Add-Member -MemberType NoteProperty -Name Discussion -Value "See NIST Doc"  -force
            $controltablearray += $controltableobject

            #create object for current maturity
            $currentmaturityobject = New-Object -TypeName PSObject    
            $controlheading =  $controloutput.'Control Identifier' + "::" + $controloutput.'Control (or Control Enhancement) Name'
            $currentmaturityobject | Add-Member -MemberType NoteProperty -Name "Control" -value $controlheading
            $currentmaturityobject | Add-Member -MemberType NoteProperty -Name "Maturity" -value ""
            $currentmaturityobject | Add-Member -MemberType NoteProperty -Name "Notes" -value ""
            $currentmaturityarray += $currentmaturityobject

            #create object for future maturity
            $futurematurityobject = New-Object -TypeName PSObject    
            $controlheading =  $controloutput.'Control Identifier' + "::" + $controloutput.'Control (or Control Enhancement) Name'
            $futurematurityobject | Add-Member -MemberType NoteProperty -Name "Control" -value $controlheading
            $futurematurityobject | Add-Member -MemberType NoteProperty -Name "Current Maturity" -value ""
            $futurematurityobject | Add-Member -MemberType NoteProperty -Name "Target Maturity" -value ""
            $futurematurityarray += $futurematurityobject
        }
    }
    #if there are no controls, output an HTML file that says so(So we don't think we are missing any due to a script error)
    if(!$controltablearray)
    {
        $cathtml = $cathtml -replace '.html', '-NoControls.html'
        New-Item $cathtml -ErrorAction SilentlyContinue | out-null
    }
    else
    {
        #Convert reports to htmls.
        $controlreport = $controltablearray | ConvertTo-Html -Fragment -PreContent "<h2>Controls</h2>"
        $currentmaturityreport = $currentmaturityarray | ConvertTo-Html -Fragment -PreContent "<h2>Current Maturity</h2>"
        $currentmaturityreport = $currentmaturityreport -replace "::","<br />"
        $futurematurityreport = $futurematurityarray | ConvertTo-Html -Fragment -PreContent "<h2>Target Maturity</h2>"
        $futurematurityreport = $futurematurityreport -replace "::","<br />"
        #combine reports into 1 html report
        $allreports = ConvertTo-Html -Body "$subcatreport $controlreport $currentmaturityreport $futurematurityreport" -head $css -Title $subcatname
    }
    $counter++
    #Write-Progress -Activity "Controls Processed" -CurrentOperation "$subcatname" -PercentComplete (($counter / $csfto800.count) * 100)
    $subcatname
    #pause
    #output to html and docx if there are controls
    if($cathtml -notlike '*-NoControls*')
    {
        
        $allreports | Out-File $cathtml 
        #convert to word
        [ref]$SaveFormat = "microsoft.office.interop.word.WdSaveFormat" -as [type] 
        $word = New-Object -ComObject word.application 
        $word.visible = $false 
        
        $docx = $cathtml -replace ".html", ".docx"
        #"Converting $html to $docx..." 
        
        $doc = $word.documents.open($cathtml) 
        
        $doc.saveas([ref] $docx, [ref]$SaveFormat::wdFormatDocumentDefault) 
        $doc.close() 
        $word.Quit() 
        $word = $null 
        [gc]::collect() 
        [gc]::WaitForPendingFinalizers()
        #convert to word end
    }
}
