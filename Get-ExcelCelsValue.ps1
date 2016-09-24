# Usage
$Usages = @(
	"To get the contents of a specific cell number from the Excel file in the specified path.";
	" Usage : $MyInvocation.MyCommand.Name  -i <InputPath> -cell <cell range no>";
	"  OPTIONS";
	"  -i    : InputFiles Directory. (FileExtension is *.xls,*.xlsx)";
	"  -Cell : Cell number to get the value.";
)

function Output-Usage {
	$Usages
}

#Argument Get & Check
function Get-Argument($ArgList) {
	if($ArgList.Count -lt $ArgMap.Count*2) {
		return $false
	}


	$ArgKey=""
	$ArgList | ForEach-Object {
		if($ArgKey -eq "") {
			if($ArgMap.ContainsKey($_)) {
				$ArgKey=$_
			} else {
				return $false
			}
		} else {
			$ArgMap[$ArgKey]=$_
			$ArgKey=""
		}
	}
	return $true
}

$ArgMap = @{ "-i" = ""; "-cell" = ""}

if(-not(Get-Argument($Args))){
	[Console]::WriteLine("ERROR: Invalid Arguments.")
	Output-Usage
	exit
}


if(-not(Test-Path -LiteralPath $ArgMap["-i"] -PathType Container)) {
	[Console]::WriteLine("ERROR: Not Exist InputDirectory.")
	Output-Usage
	exit
}
if($ArgMap["-i"].LastIndexOf("\")+1 -eq $ArgMap["-i"].Length) {
	$ArgMap["-i"].Remove($ArgMap["-i"].LastIndexOf("\"), 1)
}

$path = $ArgMap["-i"]
$cell = $ArgMap["-cell"]

try{
	# Excelオブジェクト作成
	$excel = New-Object -ComObject Excel.Application
	$excel.Visible = $false

	Get-ChildItem $path | Where-Object { 
		$_.Name -match '((.+)\.(xlsx|xls)$)'} | ForEach-Object {

		$file_name = $Matches[0]; 
		$name = $Matches[2]; 
		$ext = $Matches[3]; 

		$book = $excel.Workbooks.Open("${path}\${file_name}")
		$book.Sheets | ForEach-Object{
			if($_.visible) {
				$value = ""
				$_.RANGE($cell) | ForEach-Object{
				 	$value = "${value}	" + $_.text
				}
				if( $value.trim().length -gt 0 ){
					Write-Output "${file_name}${value}"
				}
			}
		}
		$book.close($false)
		$book = $null
	}

} finally {
	if($book -ne $null) {
		$book.close($false)
	}

	if($excel -ne $null) {
		$excel.Quit()
	}

	# null破棄
	$excel,$book | foreach{$_ = $null}

	$dummy = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
	Remove-Variable excel

	[System.GC]::Collect()
}
