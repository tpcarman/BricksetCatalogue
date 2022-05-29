<p align="center">
    <a href="https://www.powershellgallery.com/packages/BricksetCatalogue/" alt="PowerShell Gallery Version">
        <img src="https://img.shields.io/powershellgallery/v/BricksetCatalogue.svg" /></a>
    <a href="https://www.powershellgallery.com/packages/BricksetCatalogue/" alt="PS Gallery Downloads">
        <img src="https://img.shields.io/powershellgallery/dt/BricksetCatalogue.svg" /></a>
    <a href="https://www.powershellgallery.com/packages/BricksetCatalogue/" alt="PS Platform">
        <img src="https://img.shields.io/powershellgallery/p/BricksetCatalogue.svg" /></a>
</p>
<p align="center">
    <a href="https://github.com/AsBuiltReport/BricksetCatalogue/graphs/commit-activity" alt="GitHub Last Commit">
        <img src="https://img.shields.io/github/last-commit/AsBuiltReport/BricksetCatalogue/master.svg" /></a>
    <a href="https://raw.githubusercontent.com/AsBuiltReport/BricksetCatalogue/master/LICENSE" alt="GitHub License">
        <img src="https://img.shields.io/github/license/AsBuiltReport/BricksetCatalogue.svg" /></a>
    <a href="https://github.com/AsBuiltReport/BricksetCatalogue/graphs/contributors" alt="GitHub Contributors">
        <img src="https://img.shields.io/github/contributors/AsBuiltReport/BricksetCatalogue.svg"/></a>
</p>

# BricksetCatalogue
BricksetCatalogue creates a catalogue of LEGO sets and minifigs from a [Brickset](https://brickset.com/) inventory using the [Brickset API](https://brickset.com/article/52664/api-version-3-documentation).

Brickset catalogues can be created in HTML, Word and/or Text formats.

# :beginner: Getting Started

This simple list of instructions will get you started with the Brickset Catalogue module.
## :one: Pre-requisites
A [Brickset](https://brickset.com/) account and API key is required to use this module. If you do not have an existing Brickset account, you may sign up [here](https://brickset.com/signup). If you do not have an existing API key, you may request one [here](https://brickset.com/tools/webservices/requestkey).

Once you have registered an account with Brickset, add the LEGO sets and minifigs which you own and want to your online Brickset inventory.

## :floppy_disk: Supported Versions
### **PowerShell**
The Brickset Catalogue module is compatible with the following PowerShell versions;

| Windows PowerShell 5.1 | PowerShell 7 |
|:----------------------:|:------------------:|
|   :x:   |  > 7.1 :white_check_mark:|

## :wrench: System Requirements

The following required modules will be installed automatically. These modules may also be manually installed.

| Module Name        | Minimum Required Version |                          PS Gallery                           |                                   GitHub | Author                                    |
|--------------------| :-----: | :------------------------:|:---------------------------------------------------------------------:|:---------------------------------------------------------------------------:|
| PScribo            |          0.10.0           |      [Link](https://www.powershellgallery.com/packages/PScribo)       |         [Link](https://github.com/iainbrighton/PScribo/) | [@iainbrighton](https://twitter.com/iainbrighton)
| BricksetModule            |         2.1.0           |      [Link](https://www.powershellgallery.com/packages/Brickset)       |         [Link](https://github.com/jonathanmedd/BricksetModule/) | [@jonathanmedd](https://twitter.com/jonathanmedd)

## :package: Module Installation

### **PowerShell**
Open a PowerShell terminal window and install the required module as follows;
```powershell
install-module BricksetCatalogue
```

## :pencil2: Commands

### **New-BricksetCatalogue**
The `New-BricksetCatalogue` cmdlet is used to generate a [Brickset](https://brickset.com/) inventory catalogue using the [Brickset API](https://brickset.com/article/52664/api-version-3-documentation). User credentials for the Brickset API are specifed using the `Credential`, or the `Username` and `Password` parameters. One or more document formats, such as `HTML`, `Word` or `Text` can be specified using the `Format` parameter. Additional parameters are outlined below.

```powershell
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
```

## :computer: Examples
Here are some examples to get you going.

```powershell
Get-Help New-BricksetCatalogue -Examples

# Creates a Brickset catalogue in HTML format using the specified username, password and API key.
PS C:\>New-BricksetCatalogue -Username 'tim@lego.com' -Password 'LEGO!' -ApiKey 'cgY-67-tYUip' -OutputFolder 'C:\MyDocs'

# Creates a Brickset catalogue in Word format using a PSCredential and API key.
PS C:\>New-BricksetCatalogue -Format Word -Credential (Get-Credential) -ApiKey 'cgY-67-tYUip' -OutputFolder 'C:\MyDocs'

# Creates a Brickset catalogue in HTML & Word formats using the specified username, password and API key.
PS C:\>New-BricksetCatalogue -Format HTML,Word -Username 'tim@lego.com' -Password 'LEGO!' -ApiKey 'cgY-67-tYUip' -OutputFolder 'C:\MyDocs'
```

## :pencil: Notes
- Table Of Contents (TOC) may be missing in Word formatted catalogue

    When opening a DOCX catalogue, MS Word prompts the following

    `"This document contains fields that may refer to other files. Do you want to update the fields in this document?"`

    `Yes / No`

    Clicking `No` will prevent the TOC fields being updated and leaving the TOC empty.

    Always reply `Yes` to this message when prompted by MS Word.
