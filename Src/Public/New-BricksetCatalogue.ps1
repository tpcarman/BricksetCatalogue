function New-BricksetCatalogue {
    <#
    .SYNOPSIS
        Creates an inventory catalogue of a Brickset collection in HTML & Word formats using the Brickset API.
    .DESCRIPTION
        Creates an inventory catalogue of a Brickset collection in HTML & Word formats using the Brickset API - https://brickset.com/article/52664/api-version-3-documentation.
    .PARAMETER Format
        Specifies the output format of the catalogue.
        The supported output formats are HTML & WORD.
        Multiple output formats may be specified, separated by a comma.
    .PARAMETER Credential
        Specifies the stored credential for the Brickset API.
    .PARAMETER Username
        Specifies the username for the Brickset API.
    .PARAMETER Password
        Specifies the password for the Brickset API.
    .PARAMETER ApiKey
        Specifies an API key to authenticate to the Brickset API.
    .PARAMETER Timestamp
        Specifies whether to append a timestamp string to the catalogue filename.
        By default, the timestamp string is not added to the catalogue filename.
    .PARAMETER OutputFolder
        Specifies the folder path to save the catalogue file.
    .PARAMETER Filename
        Specifies a filename for the catalogue.
    .PARAMETER OrderBy
        Specifies the sort order for the sets.
        The supported sort orders are 'Name', 'Theme', 'Number', 'Pieces', 'QtyOwned', 'Rating'
    .PARAMETER ExcludeWantedSets
        Excludes wanted sets from the catalogue.
    .PARAMETER ExcludeOwnedSets
        Excludes owned sets from the catalogue.
    .PARAMETER ExcludeWantedMinifigs
        Excludes wanted minifigs from the catalogue.
    .PARAMETER ExcludeOwnedMinifigs
        Excludes owned minifigs from the catalogue.
    .PARAMETER ExcludeToC
        Excludes the Table of Contents from the catalogue.
    .EXAMPLE
        PS C:\>New-BricksetCatalogue -Username 'tim@lego.com' -Password 'LEGO!' -ApiKey 'cgY-67-tYUip' -OutputFolder 'C:\MyDocs'
        Creates a Brickset catalogue in HTML format using the specified username, password and API key.
    .EXAMPLE
        PS C:\>New-BricksetCatalogue -Format Word -Credential (Get-Credential) -ApiKey 'cgY-67-tYUip' -OutputFolder 'C:\MyDocs'
        Creates a Brickset catalogue in Word format using a PSCredential and API key.
    .EXAMPLE
        PS C:\>New-BricksetCatalogue -Format HTML,Word -Username 'tim@lego.com' -Password 'LEGO!' -ApiKey 'cgY-67-tYUip' -OutputFolder 'C:\MyDocs'
        Creates a Brickset catalogue in HTML and Word formats using the specified username, password and API key.
    .NOTES
        Version:        0.1.2
        Author:         Tim Carman
        Twitter:        @tpcarman
        Github:         tpcarman
        Credits:        Iain Brighton (@iainbrighton) - PScribo module
                        Jonathan Medd (@jonathanmedd) - Brickset module
    .LINK
        https://github.com/tpcarman/BricksetCatalogue
    #>

    [CmdletBinding()]
    param (
        [Parameter(
            Position = 0,
            Mandatory = $false,
            HelpMessage = 'Please provide the Brickset Catalogue output format'
        )]
        [ValidateNotNullOrEmpty()]
        [ValidateSet('Word', 'HTML')]
        [Array] $Format = 'HTML',

        [Parameter(
            Position = 1,
            Mandatory = $true,
            HelpMessage = 'Please provide credentials to connect to the Brickset API',
            ParameterSetName = 'Credential'
        )]
        [ValidateNotNullOrEmpty()]
        [PSCredential] $Credential,

        [Parameter(
            Position = 1,
            Mandatory = $true,
            HelpMessage = 'Please provide the username to connect to the Brickset API',
            ParameterSetName = 'UsernameAndPassword'
        )]
        [ValidateNotNullOrEmpty()]
        [String] $Username,

        [Parameter(
            Position = 2,
            Mandatory = $true,
            HelpMessage = 'Please provide the password to connect to the Brickset API',
            ParameterSetName = 'UsernameAndPassword'
        )]
        [ValidateNotNullOrEmpty()]
        [String] $Password,

        [Parameter(
            Position = 3,
            Mandatory = $true,
            HelpMessage = 'Please provide Brickset API key'
        )]
        [ValidateNotNullOrEmpty()]
        [String] $ApiKey,

        [Parameter(
            Position = 4,
            Mandatory = $false,
            HelpMessage = 'Specify the Brickset Catalogue filename'
        )]
        [ValidateNotNullOrEmpty()]
        [String] $Filename = 'BricksetCatalogue',

        [Parameter(
            Position = 5,
            Mandatory = $true,
            HelpMessage = 'Please provide the folder path to save the Brickset Catalogue file'
        )]
        [ValidateNotNullOrEmpty()]
        [String] $OutputFolder,

        [Parameter(
            Mandatory = $false,
            HelpMessage = 'Specify whether to append a timestamp to the document filename'
        )]
        [Switch] $Timestamp = $false,

        [Parameter(
            Mandatory = $false,
            HelpMessage = 'Specify whether to append a timestamp to the document filename'
        )]
        [ValidateSet('Name', 'Theme', 'Number', 'Pieces', 'QtyOwned', 'Rating')]
        [String] $OrderBy = 'Theme',

        [Parameter(
            Mandatory = $false,
            HelpMessage = 'Specify whether to exclude wanted sets from the catalogue'
        )]
        [Switch] $ExcludeWantedSets = $false,

        [Parameter(
            Mandatory = $false,
            HelpMessage = 'Specify whether to exclude wanted minifigs from the catalogue'
        )]
        [Switch] $ExcludeWantedMinifigs = $false,

        [Parameter(
            Mandatory = $false,
            HelpMessage = 'Specify whether to exclude owned sets from the catalogue'
        )]
        [Switch] $ExcludeOwnedSets = $false,

        [Parameter(
            Mandatory = $false,
            HelpMessage = 'Specify whether to exclude owned minifigs from the catalogue'
        )]
        [Switch] $ExcludeOwnedMinifigs = $false,

        [Parameter(
            Mandatory = $false,
            HelpMessage = 'Specify whether to exclude the Table of Contents from the catalogue'
        )]
        [Switch] $ExcludeToC = $false
    )

    # If Username & Password are used, convert to PSCredential
    if (($Username -and $Password)) {
        $SecurePassword = ConvertTo-SecureString $Password -AsPlainText -Force
        $Credential = New-Object System.Management.Automation.PSCredential ($Username, $SecurePassword)
    }

    # If Timestamp parameter is specified, add the timestamp to the catalogue filename
    if ($Timestamp) {
        $FileName = $Filename + " - " + (Get-Date -Format 'yyyy-MM-dd_HH.mm.ss')
    }

    # Check for output folder
    if (!(Test-Path -Path $OutputFolder)) {
        Write-Error -Message "Output folder not found [$($OutputFolder)]" -ErrorAction Stop
    }

    # Connect to Brickset API
    Try {
        $BrickSet = Connect-Brickset -apiKey $ApiKey -credential $Credential -ErrorAction Stop
    } Catch {
        Write-Error $_
    }

    if ($BrickSet) {
        # Brickset Collection Sets
        if (!($ExcludeOwnedSets)) {
            $BrickSetSetOwned = Get-BricksetSetOwned -orderBy $OrderBy
            $TotalSetsOwned = ($BrickSetSetOwned.collection.qtyowned | Measure-Object -Sum).Sum
            $UniqueSetsOwned = ($BrickSetSetOwned.collection.qtyowned | Measure-Object -Sum).Count
            $TotalSetPieceCount = ($BrickSetSetOwned.pieces | Measure-Object -Sum).Sum
            $BrickSetSetOwnedThemes = $BrickSetSetOwned.theme | Select-Object -Unique | Sort-Object
        }

        if (!($ExcludeWantedSets)) {
            $BrickSetSetWanted = Get-BricksetSetWanted -orderBy $OrderBy
            $BrickSetSetWantedThemes = $BrickSetSetWanted.theme | Select-Object -Unique | Sort-Object
        }

        # Brickset Collection Minifigs
        if (!($ExcludeOwnedMinifigs)) {
            $BricksetMinifigOwned = Get-BricksetMinifigCollectionOwned | Sort-Object name
        }
        if (!($ExcludeWantedMinifigs)) {
            $BricksetMinifigWanted = Get-BricksetMinifigCollectionWanted | Sort-Object name
        }

        # Create Brickset Catalogue
        $BrickSetCatalogue = Document -Name $FileName {
            # Set Document Style
            Set-DocStyle

            # Set Header & Footer
            Header -Default {
                Paragraph -Style Header "Brickset Collection - $($Filename)"
            }

            Footer -Default {
                Paragraph -Style Footer 'Page <!# PageNumber #!> of <!# TotalPages #!>'
            }

            # Cover Page
            Set-CoverPage

            # Table of Contents
            if (!($ExcludeToC)) {
                TOC
                PageBreak
            }

            # Summary Sections
            if ($BrickSetSetOwned) {
                Section -Style Heading1 "Summary" {
                    $BrickSetCollection = [PSCustomObject]@{
                        'Total Sets Owned' = $TotalSetsOwned
                        'Unique Sets Owned' = $UniqueSetsOwned
                        'Total Set Piece Count' = $TotalSetPieceCount
                    }
                    $TableParams = @{
                        Name = 'Brickset Collection Summary'
                        List = $false
                        ColumnWidths = 33, 34, 33
                        Caption = '- Brickset Collection Summary'
                    }
                    $BrickSetCollection | Table @TableParams
                }
            }

            # Sets Sections
            if (($BrickSetSetOwned) -or ($BrickSetSetWanted)) {
                Section -Style Heading1 "Sets" {
                    if (($BrickSetSetOwned) -and (!($ExcludeOwnedSets))) {
                        Section -Style Heading2 "Owned" {
                            foreach ($SetTheme in $BrickSetSetOwnedThemes) {
                                Section -Style Heading3 $($SetTheme) {
                                    $BrickSetSetOwnedByTheme = $BrickSetSetOwned | Where-Object {$_.theme -eq $SetTheme}
                                    foreach ($SetOwned in $BrickSetSetOwnedByTheme) {
                                        Section -Style Heading4 -ExcludeFromTOC "$($SetOwned.Number): $($SetOwned.Name)" {
                                            $Instructions = Get-BricksetSetInstructions -setId $SetOwned.setId
                                            Image -Uri $SetOwned.image.thumbnailURL -Align Center
                                            Blankline
                                            $SetOwnedInfo = [PSCustomObject] @{
                                                'Set Number' = $SetOwned.Number
                                                'Name' = $SetOwned.Name
                                                'Set Type' = $SetOwned.category
                                                'Theme Group' = $SetOwned.themeGroup
                                                'Theme' = $SetOwned.Theme
                                                'Year Released' = $SetOwned.Year
                                                'Pieces' = $SetOwned.Pieces
                                                'Minifigs' = Switch ($SetOwned.Minifigs) {
                                                    $null { '0' }
                                                    default { $SetOwned.Minifigs }
                                                }
                                                'Age Range' = Switch ($SetOwned.ageRange.max) {
                                                    $null { "$($SetOwned.ageRange.min)+" }
                                                    default { "$($SetOwned.ageRange.min) - $($SetOwned.ageRange.max)" }
                                                }
                                                'Packaging' = $SetOwned.PackagingType
                                                'Availability' = $SetOwned.Availability
                                                'Qty Owned' = $SetOwned.collection.qtyOwned
                                                'Instructions' = $Instructions.URL -join [Environment]::NewLine
                                                'Rating' = $SetOwned.Rating
                                                'Notes' = Switch ($SetOwned.collection.notes) {
                                                    $null { '' }
                                                    default { $SetOwned.collection.notes }
                                                }
                                                'BrickSet URL' = $SetOwned.bricksetURL
                                            }
                                            $TableParams = @{
                                                Name = "$($SetOwned.Number): $($SetOwned.Name)"
                                                List = $true
                                                ColumnWidths = 50, 50
                                                Caption = "- $($SetOwned.Number): $($SetOwned.Name)"
                                            }
                                            $SetOwnedInfo | Table @TableParams
                                        }
                                    }
                                }
                            }
                        }
                    }
                    if (($BrickSetSetWanted) -and (!($ExcludeWantedSets))) {
                        Section -Style Heading2 "Wanted" {
                            foreach ($SetTheme in $BrickSetSetWantedThemes) {
                                Section -Style Heading3 $($SetTheme) {
                                    $BrickSetSetWantedByTheme = $BrickSetSetWanted | Where-Object {$_.theme -eq $SetTheme}
                                    foreach ($SetWanted in $BrickSetSetWantedByTheme) {
                                        Section -Style Heading4 -ExcludeFromTOC "$($SetWanted.Number): $($SetWanted.Name)" {
                                            $Instructions = Get-BricksetSetInstructions -setId $SetWanted.setId
                                            Image -Uri $SetWanted.image.thumbnailURL -Align Center
                                            Blankline
                                            $SetWantedInfo = [PSCustomObject] @{
                                                'Set Number' = $SetWanted.Number
                                                'Name' = $SetWanted.Name
                                                'Set Type' = $SetWanted.category
                                                'Theme Group' = $SetWanted.themeGroup
                                                'Theme' = $SetWanted.Theme
                                                'Year Released' = $SetWanted.Year
                                                'Pieces' = $SetWanted.Pieces
                                                'Minifigs' = Switch ($SetWanted.Minifigs) {
                                                    $null { '0' }
                                                    default { $SetWanted.Minifigs }
                                                }
                                                'Age Range' = Switch ($SetWanted.ageRange.max) {
                                                    $null { "$($SetWanted.ageRange.min)+" }
                                                    default { "$($SetWanted.ageRange.min) - $($SetWanted.ageRange.max)" }
                                                }
                                                'Packaging' = $SetWanted.PackagingType
                                                'Availability' = $SetWanted.Availability
                                                'Qty Owned' = $SetWanted.collection.qtyOwned
                                                'Instructions' = $Instructions.URL -join [Environment]::NewLine
                                                'Rating' = $SetWanted.Rating
                                                'Notes' = Switch ($SetWanted.collection.notes) {
                                                    $null { '' }
                                                    default { $SetWanted.collection.notes }
                                                }
                                                'BrickSet URL' = $SetWanted.bricksetURL
                                            }
                                            $TableParams = @{
                                                Name = "$($SetWanted.Number): $($SetWanted.Name)"
                                                List = $true
                                                ColumnWidths = 50, 50
                                                Caption = "- $($SetWanted.Number): $($SetWanted.Name)"
                                            }
                                            $SetWantedInfo | Table @TableParams
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            # Minifigs Sections
            if (($BricksetMinifigOwned) -or ($BricksetMinifigWanted)) {
                Section -Style Heading1 "Minfigs" {
                    if (($BricksetMinifigOwned) -and (!($ExcludeOwnedMinifigs))) {
                        Section -Style Heading2 "Owned" {
                            foreach ($MinfigOwned in $BricksetMinifigOwned) {
                                Section -Style Heading3 -ExcludeFromTOC "$($MinfigOwned.minifigNumber): $($MinfigOwned.Name)" {
                                    Image -Uri "https://img.bricklink.com/ItemImage/MN/0/$($MinfigOwned.minifigNumber).png" -Align Center -Percent 50 -Text $($MinfigOwned.minifigNumber)
                                    Blankline
                                    $MinifigOwnedInfo = [PSCustomObject]@{
                                        'Name' = $MinfigOwned.Name
                                        'Number' = $MinfigOwned.minifigNumber
                                        'Category' = $MinfigOwned.Category
                                        'Owned in Sets' = $MinfigOwned.OwnedInSets
                                        'Owned Loose' = $MinfigOwned.OwnedLoose
                                    }
                                    $TableParams = @{
                                        Name = "$($MinfigOwned.minifigNumber): $($MinfigOwned.Name)"
                                        List = $true
                                        ColumnWidths = 50, 50
                                        Caption = "- $($MinfigOwned.minifigNumber): $($MinfigOwned.Name)"
                                    }
                                    $MinifigOwnedInfo | Table @TableParams
                                }
                            }
                        }
                    }
                    if (($BricksetMinifigWanted) -and (!($ExcludeWantedMinifigs))) {
                        Section -Style Heading2 "Wanted" {
                            foreach ($MinifigWanted in $BricksetMinifigWanted) {
                                Section -Style Heading3 -ExcludeFromTOC "$($MinifigWanted.minifigNumber): $($MinifigWanted.Name)" {
                                    Image -Uri "https://img.bricklink.com/ItemImage/MN/0/$($MinifigWanted.minifigNumber).png" -Align Center -Percent 50 -Text $($MinifigWanted.minifigNumber)
                                    Blankline
                                    $MinifigOwnedInfo = [PSCustomObject]@{
                                        'Name' = $MinifigWanted.Name
                                        'Number' = $MinifigWanted.minifigNumber
                                        'Category' = $MinifigWanted.Category
                                        'Owned in Sets' = $MinifigWanted.OwnedInSets
                                        'Owned Loose' = $MinifigWanted.OwnedLoose
                                    }
                                    $TableParams = @{
                                        Name = "$($MinifigWanted.minifigNumber): $($MinifigWanted.Name)"
                                        List = $true
                                        ColumnWidths = 50, 50
                                        Caption = "- $($MinifigWanted.minifigNumber): $($MinifigWanted.Name)"
                                    }
                                    $MinifigOwnedInfo | Table @TableParams
                                }
                            }
                        }
                    }
                }
            }
        }
        Export-Document -Document $BrickSetCatalogue -Format $Format -Path $OutputFolder
        Disconnect-Brickset -Confirm:$false
    }
}