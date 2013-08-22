<#
.Description
	Converts a Yaml document to PowerShell custom objects.
	
.Parameter Yaml
	The Yaml document to convert.
	
.Parameter InitialProperties
	A hash table of default values to set for the top-level object. The values in the Yaml document will override the default values provided by this parameter. For this to work the top-level needs to be a map.

.Example
$Yaml = @"
---
Prop1: "value1"
Prop2: "value2"
...
"@
	
ConvertFrom-Yaml -Yaml $Yaml

Prop2    Prop1
-----    -----
value2   value1
#>
function ConvertFrom-Yaml {
	param(
		[Parameter(Mandatory=$true,ValueFromPipeline=$true)]
		[String]$Yaml,
		
		[Parameter(Mandatory=$false)]
		[System.Collections.HashTable]$InitialProperties
	)
	process {
		$sr = New-object IO.StringReader($Yaml)
		$YamlStream = new-object YamlDotNet.RepresentationModel.YamlStream
		$YamlStream.Load($sr)
		if($InitialProperties) {
			$YamlStream.Documents | %{ ParseNode -YamlNode $_ -InitialProperties $InitialProperties }
		} else {
			$YamlStream.Documents | %{ ParseNode -YamlNode $_ }
		}
		$sr.Close()
	}
}
Export-ModuleMember -Function ConvertFrom-Yaml

function ParseNode {
	param(
		$YamlNode,
		[System.Collections.HashTable]$InitialProperties = @{}
	)
	switch($YamlNode.GetType().FullName) {
		"YamlDotNet.RepresentationModel.YamlMappingNode" {
			Write-Debug "Creating new object"
			$Props = $InitialProperties
			$YamlNode | Foreach-Object {
				Write-Debug "Recursing for property $($_.Key.Value)"
				$Props[$_.Key.Value] = (ParseNode -YamlNode $_.Value)
			}
			[PsCustomObject]$Props
		}
		"YamlDotNet.RepresentationModel.YamlScalarNode" {
			Write-Debug "Processing Scalar Node"
			switch($YamlNode.Style) {
				([YamlDotNet.Core.ScalarStyle]::Plain) {
					$Value = $null
					if([int]::TryParse($YamlNode.Value, [ref]$Value)) {
						$Value
					} elseif([int64]::tryparse($YamlNode.Value, [ref]$Value)) {
						$Value
					} elseif([single]::tryparse($YamlNode.Value, [ref]$Value)) {
						$Value
					} elseif([double]::tryparse($YamlNode.Value, [ref]$Value)) {
						$Value
					} elseif([bool]::tryparse($YamlNode.Value, [ref]$Value)) {
						$Value
					} else {
						#If all else fails, just return the string
						$YamlNode.Value
					}
				}
				default {
					$YamlNode.Value
				}
			}
		}
		"YamlDotNet.RepresentationModel.YamlSequenceNode" {
			Write-Debug "Processing Sequence Node"
			$YamlNode | %{ ParseNode -YamlNode $_}
		}
		"YamlDotNet.RepresentationModel.YamlDocument" {
			Write-Debug "Processing Document Node"
			if($InitialProperties) {
				ParseNode -YamlNode $YamlNode.RootNode -InitialProperties $InitialProperties
			} else {
				ParseNode -YamlNode $YamlNode.RootNode
			}
		}
		default {
			Write-Warning "Unknown node type $($YamlNode.GetType().Fullname)"
		}
	}
}

<#
.Description
	Convert an object to a Yaml document.
	
.Parameter InputObject
	The object to convert.

.Parameter Depth
	Specifies how many levels of contained objects are included in the Yaml representation. The default value is 2.
#>
function ConvertTo-Yaml {
	param(
		[Parameter(Mandatory=$true,ValueFromPipeline=$true)]
		[ValidateNotNullOrEmpty()]
		[Object[]]$InputObject,
		
		[Int32]$Depth
	)
	begin {
		$PsBoundParameters.Remove("InputObject") > $null
		$JsonPipeline = { ConvertTo-Json @PSBoundParameters | FixJSONDate }.GetSteppablePipeline()
		$JsonPipeline.Begin($True)
		$YamlSb = New-Object Text.StringBuilder
	}
	process {
		$InputObject | ForEach-Object {
			$JsonPipeline.Process($_) | ForEach-Object {
				$YamlSb.Append($_) > $null
			}
		}
	}
	end {
		$JsonPipeline.End() | ForEach-Object {
			$YamlSb.Append($_) > $null
		}
		"---`r`n{0}`r`n..." -f $YamlSb.ToString().Trim("{}`r`n")
	}
}
Export-ModuleMember -Function ConvertTo-Yaml
		

function FixJSONDate {
	param(
		[Parameter(ValueFromPipeline=$true)]
		$Json
	)
	process {
		[Regex]::Matches($json, "\\/Date\((?<Ticks>\d+)\)\\/") | ForEach-Object {
			$ms = [double]$_.groups[1].value
			$Iso8601 = ([DateTime]"1/1/1970").AddMilliseconds($ms).ToString("o")
			$json = $json.replace($_.Groups[0].value, $Iso8601)
		}
		$json
	}
}