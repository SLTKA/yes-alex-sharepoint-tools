<# 
.SYNOPSIS 
 Finds and tries to fix brocken Terms cache in TaxonomyHiddenList.
 Can fix problems which appear in result of incorrect taxonomy field provisioning
 or due to usage of some incorrect characters in term labels
 Please fix your terms in Term store first
 
.DESCRIPTION 
 Please NOTE: this script can break you site completely, use it as a reference only!
 Test it carefully by running without `-fix` flag.
 The author in NOT responsible for any damage this script may do!
 
.NOTES 
License:

MIT License

Copyright (c) 2020 Alexandr Yeskov

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
 
.PARAMETER siteUrl
Site URL where brocken TaxonomyHiddenList is located
 
.PARAMETER fix 
Use this flag to apply all updates, without it only report will be generated
 
.EXAMPLE 

#To run in REPORT ONLY mode
.\TaxonomyHiddenListAnalyzer.ps1 -siteUrl <YOUR_SITE_COLLECTION_URL>

#To run in FIX mode and apply all fixes
.\TaxonomyHiddenListAnalyzer.ps1 -siteUrl <YOUR_SITE_COLLECTION_URL> -fix
 
#>

Param(
    [string]
    $siteUrl,
    [switch]
    $fix
)

$keywordTermSetId = "3ea0ac55-abd1-41a6-890a-6df10ee10901"

Connect-PnPOnline -Url $siteUrl

Function Get-InvalidTaxonomyItems {
    $taxList = Get-PnPList -Identity "Lists/TaxonomyHiddenList"

    $position = $null

    Do {
        $q = new-object Microsoft.SharePoint.Client.CamlQuery
        $q.ListItemCollectionPosition = $position
        $q.ViewXml = "<View Scope='RecursiveAll'><Query><OrderBy><FieldRef Name='ID' Ascending='True' /></OrderBy></Query><RowLimit Paged='True'>200</RowLimit></View>"
        $items = $taxList.GetItems($q)
        $taxList.Context.Load($items)
        $taxList.Context.ExecuteQuery()
        $position = $items.ListItemCollectionPosition

        foreach ($item in $items)
        {
            if ($item["CatchAllData"] -eq "UNVALIDATED") {
                @{
                    Item = $item
                    Title = $item["Title"];
                    IdForTermSet = $item["IdForTermSet"];
                    IdForTerm = $item["IdForTerm"];
                    Term = $item["Term"];
                    Path = $item["Path"];
                    CatchAllData = $item["CatchAllData"];
                    CatchAllDataLabel = $item["CatchAllDataLabel"];
                    Term1033 = $item["Term"];
                    Path1033 = $item["Path"];
                }
            }
        }
        Write-Host "Loop"
    } While ($position -ne $null)
}

Function Get-CompressedGuid {
    Param(
        [Parameter(Mandatory,
            Position = 0,
            ValueFromPipeline)]
        [ValidateNotNullOrEmpty()]
        [Guid]
        $id
    )
    return [Convert]::ToBase64String($id.ToByteArray(), [Base64FormattingOptions]::None)
}

Function Get-TermFromItem {
    Param(
        [Parameter(Mandatory,
            Position = 0,
            ValueFromPipeline)]
        [ValidateNotNullOrEmpty()]
        $item
    )
    try {
        $term = Get-PnPTerm -Identity $item["IdForTerm"]
        $term.Context.Load($term.TermSet)
        $term.Context.Load($term.TermSet.Group)
        $term.Context.Load($term.TermSet.TermStore)
        $term.Context.ExecuteQuery()
        $groupId = $term.TermSet.Group.Id
        $term = Get-PnPTerm -Identity $item["IdForTerm"] -TermSet $item["IdForTermSet"] -TermGroup $groupId -TermStore $term.TermSet.TermStore.Id -ErrorAction Ignore
        if ($term -eq $null) { 
            Write-Warning -Message "Term with Identity $($item["IdForTerm"]) in TermSet $($item["IdForTermSet"]) TermGroup $groupId TermStore $($item["IdForTermStore"]) was not found for item with ID $($item.Id). You may want to remove it from TaxonomyHiddenList manually!"
            return $null
        }
        $term.Context.Load($term.Labels)
        $term.Context.Load($term.TermStore)
        $term.Context.Load($term.TermSet)
        $term.Context.Load($term.TermSet.Group)
        $term.Context.Load($term.TermSet.TermStore)
        $term.Context.ExecuteQuery()
    } catch {
        Write-Warning -Message "Term with Identity $($item["IdForTerm"]) in TermSet $($item["IdForTermSet"]) TermGroup $groupId TermStore $($item["IdForTermStore"]) was not found for item with ID $($item.Id). You may want to remove it from TaxonomyHiddenList manually!"
    }
    return $term
}

Function Get-CatchAllData {
    Param(
        [Parameter(Mandatory,
            Position = 0,
            ValueFromPipeline)]
        [ValidateNotNullOrEmpty()]
        $term
    )
    $isKeyword = $item["IdForTermSet"] -like $keywordTermSetId

    $result = ""
    $result += Get-CompressedGuid -id $term.TermSet.TermStore.Id
    $result += "|"
    $result += Get-CompressedGuid -id $term.TermSet.Id
    if (-not $isKeyword)
    {
        $parent = $term;
        do
        {
            $result += "|"
            $result += Get-CompressedGuid -id $parent.Id
            $term.Context.Load($parent.Parent)
            $term.Context.ExecuteQuery()
            $parent = $parent.Parent
        }
        while ($parent -ne $null -and $parent.Id -ne $null -and $result.Length -lt 230)
    }
    return $result
}

Function Get-CatchAllLabel {
    Param(
        [Parameter(Mandatory,
            Position = 0,
            ValueFromPipeline)]
        [ValidateNotNullOrEmpty()]
        $term
    )
    $term.Context.ExecuteQuery()
    
    $result = ""
    if ($term.Name.Length > 50)
    {
        $result += $term.Name.Substring(0, 50)
    }
    else
    {
        $result += $term.Name
    }
    $result += "#"
    $result += [char]$term.TermStore.DefaultLanguage
    $result += "|"
    foreach ($label in $term.Labels)
    {
        if (-not $label.IsDefaultForLanguage)
        {
            continue;
        }
        if ($label.Value -ne $term.Name)
        {
            if ($label.Value.Length -gt 50)
            {
                $result += $label.Value.Substring(0, 50)
            }
            else
            {
                $result += $label.Value
            }
            $result += "#"
            $result += [char]$label.Language
            $result += "|"
        }
    }
    if ($result.Length > 255)
    {
        $result = $result.Substring(0, 254)
    }
    return $result
}

Function Get-ConcatenatedPath {
    Param(
        [Parameter(Mandatory,
            Position = 0,
            ValueFromPipeline)]
        [ValidateNotNullOrEmpty()]
        [string]
        $path
    )

    if ([String]::IsNullOrEmpty($path))
    {
        return $path
    }

    $newPath = $path.Replace(';', ':')
    if ($newPath.Length -lt 255)
    {
        return $newPath
    }
    
    $suffix = [string]::Empty
    $prefix = [string]::Empty
    $strArray = $newPath.Split(':')
    $fromEnd = $true
    $index = 0
    $isFull = $false
    $newPath = [string]::Empty
    while (-not $isFull) {
        if ($fromEnd) {
            $pathPart = $strArray[$strArray.Length - 1 - $index]
        } else {
            $pathPart = $strArray[$index]
        }

        if ([string]::IsNullOrEmpty($suffix))
        {
            $suffix = $strArray[$strArray.Length - 1];
        } elseif (($pathPart.Length + $prefix.Length + $suffix.Length + 6) -lt 255) {
            if ($fromEnd)
            {
                $suffix = $pathPart + ':' + $suffix;
            }
            else
            {
                if ([string]::IsNullOrEmpty($prefix)) {
                    $prefix = $pathPart
                } else {
                    $prefix += ':' + $pathPart
                }
            }
        } else {
            if ($prefix.Length -eq 0)
            {
                if ($suffix.Length -ge 251) {
                    $newPath =  $suffix
                } else {
                    $newPath = "...:" + $suffix
                }
            } else {
                $newPath = [string]::Concat(@( $prefix, ':...:', $suffix ))
            }
            $isFull = $true;
        }
        if (-not $fromEnd)
        {
            $index++;
        }
        $fromEnd = -not $fromEnd;
    }
    return $newPath
}

Function Get-TermEffectiveLabel {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [array]
        $labels,
        [Parameter(Mandatory)]
        [int]
        $language,
        [Parameter(Mandatory)]
        [int]
        $defaultLanguage
    )

    $label = $labels | ? {$_.Language -eq $language}
    if ($label -eq $null) {
        $label = $term.Labels | ? {$_.Language -eq $defaultLanguage}
    }
    return $label
}

Function Fix-InvalidTaxonomyItems {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory,
            Position = 0,
            ValueFromPipeline)]
        [ValidateNotNullOrEmpty()]
        [array]
        $items
    )
    PROCESS {
        foreach ($item in $items) {
            Write-Host Processing $item.Id

            $updates = @{}

            $term = Get-TermFromItem -item $item

            if ($term -eq $null) { continue }
            $catchAllData = Get-CatchAllData -term $term
            $catchAllLabel = Get-CatchAllLabel -term $term
            $isKeyword = $term.TermSet.Id.ToString() -like $keywordTermSetId

            #Fixing stor if it is incorrect. Sometimes happence when a Managed Metadata field deployed incorrectly with a wrong Term Store ID
            if ($item["IdForTermStore"] -ne $term.TermSet.TermStore.Id.ToString()) {
                Write-Host "IdForTermStore is different, value '$($item["IdForTermStore"])' will be replaced with '$($term.TermSet.TermStore.Id.ToString())'"
                $updates.Add("IdForTermStore", $term.TermSet.TermStore.Id.ToString())
            }

            if ($item["CatchAllData"] -ne $catchAllData) {
                Write-Host "CatchAllData is different, value '$($item["CatchAllData"])' will be replaced with '$catchAllData'"
                $updates.Add("CatchAllData", $catchAllData)
            }
            if ($item["CatchAllDataLabel"] -ne $catchAllLabel) {
                Write-Host "CatchAllDataLabel is different, value '$($item["CatchAllDataLabel"])' will be replaced with '$catchAllLabel'"
                $updates.Add("CatchAllDataLabel", $catchAllLabel)
            }

            $defaultLabel = $term.GetDefaultLabel($term.TermStore.DefaultLanguage)
            $defaultPath = $term.GetPath($term.TermStore.DefaultLanguage)
            $term.Context.ExecuteQuery()

            if ($item["Title"] -ne $defaultLabel.Value) {
                Write-Host "Title is different, value '$($item["Title"])' will be replaced with '$($defaultLabel.Value)'"
                $updates.Add("Title", $defaultLabel.Value)
            }
            if ($item["Term"] -ne $defaultLabel.Value) {
                Write-Host "Term is different, value '$($item["Term"])' will be replaced with '$($defaultLabel.Value)'"
                $updates.Add("Term", $defaultLabel.Value)
            }


            if ($isKeyword) {
                $path = $defaultLabel.Value
            } else {
                $path = Get-ConcatenatedPath -path $defaultPath.Value
            }

            if ($item["Path"] -ne $path) {
                Write-Host "Path is different, value '$($item["Path"])' will be replaced with '$path'"
                $updates.Add("Path", $path)
            }
            
            $keys = $item.FieldValues.Keys | % {$_}
            foreach ($fieldName in $keys) {
                if ($fieldName.StartsWith("Term", [StringComparison]::Ordinal))
                {
                    $result = 0;
                    if (-not [int]::TryParse($fieldName.Substring("Term".Length), [ref] $result))
                    {
                        continue;
                    }
                    
                    $label = Get-TermEffectiveLabel -labels $term.Labels -language $result -defaultLanguage $term.TermStore.DefaultLanguage

                    if ($item[$fieldName] -ne $label.Value) {
                        Write-Host "$fieldName is different, value '$($item[$fieldName])' will be replaced with '$($label.Value)'"
                        $updates.Add($fieldName, $label.Value)
                    }
                    continue;
                }
                if ($fieldName.StartsWith("Path", [StringComparison]::Ordinal)) {
                    $result = 0;
                    
                    if (-not [int]::TryParse($fieldName.Substring("Path".Length), [ref] $result)) {
                        if ($isKeyword) {
                            $path = $defaultLabel.Value
                        } else {
                            $path = Get-ConcatenatedPath -path $defaultPath.Value
                        }
                    } else {
                        if ($isKeyword) {
                            $path = Get-TermEffectiveLabel -labels $term.Labels -language $result -defaultLanguage $term.TermStore.DefaultLanguage
                        } else {
                            $langPath = $term.GetPath($result)
                            $term.Context.ExecuteQuery()
                            $path = Get-ConcatenatedPath -path $langPath.Value
                        }
                    }
                    if ($item[$fieldName] -ne $path) {
                        Write-Host "$fieldName is different, value '$($item[$fieldName])' will be replaced with '$path'"
                        if (-not $updates.ContainsKey($fieldName)) {
                            $updates.Add($fieldName, $path)
                        }
                    }
                }
            }
            if ($fix) {
                Write-Host "Saving changes..."
                foreach ($key in $updates.Keys) {
                    $item[$key] = $updates[$key]
                }
                
                $item.UpdateOverwriteVersion()
                $item.Context.ExecuteQuery()
                Write-Host "Done"
            } else {
                Write-Host "The changes were not applied, use '-fix' flag to fix all terms"
            }
            Write-Host "-----------"
        }
    }
}

Get-InvalidTaxonomyItems | % {$_.Item} | Fix-InvalidTaxonomyItems
