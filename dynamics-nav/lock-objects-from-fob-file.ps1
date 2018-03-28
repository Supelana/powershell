$defaultDataSource = "UDV-NAVSQL01"
$defaultLockedBy = $env:USERDOMAIN + '\' + $env:USERNAME

function PromptDataSource () {
    $dataSource = Read-Host -Prompt "Database Server (Default: '$defaultDataSource')"
    if ($dataSource.Length -eq 0) {
        $dataSource = $defaultDataSource
    }
    return $dataSource
}

function PromptDatabase () {
    $database = ""
    while ($database.Length -eq 0) {
        $database = Read-Host -Prompt "Database Name"
        if ($database.Length -eq 0) {
            Write-Host 'You have to specify a "Database Name"' 
        }
    }
    return $database
}

function PromptLockedBy () {
    $lockedBy = Read-Host -Prompt "Locked By (Default: '$defaultLockedBy')"
    if ($lockedBy.Length -eq 0) {
        $lockedBy = $defaultLockedBy
    }
    return $lockedBy  
  }

function PromptFobFilePath($initialDirectory) {
    Write-Host 'Choose a .fob file...'
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $fileDialog.InitialDirectory = $initialDirectory
    $fileDialog.Filter = "*.fob|*.fob"
    $fileDialog.ShowHelp = $true
    $fileDialog.ShowDialog() | Out-Null

    return $fileDialog.FileName
}

function GetFobFile () {
    $fobFilePath = PromptFobFilePath
    if ($fobFilePath.Length -ne 0) {
        return Get-Content $fobFilePath
    }
}

$dataSource = PromptDataSource
$database = PromptDatabase
$lockedBy = PromptLockedBy
$fobFile = GetFobFile

$connectionString = "Server=$dataSource;Database=$database;Integrated Security=True;"
$connection = New-Object System.Data.SqlClient.SqlConnection
$connection.ConnectionString = $connectionString
$NAVObjects = New-Object System.Collections.ArrayList

Add-Type -TypeDefinition @"
public enum ObjectType
{
   Table = 1,
   Page = 8,
   Report = 3,
   Codeunit = 5,
   Query = 9,
   XMLport = 6,
   MenuSuite = 7
}
"@

function GetObjectID ($text) {
    $id = ""
    for ($i = 19; $i -ne 0; $i--) {
        if ($text[$i] -eq " ") {
            break
        }
        $id = $text[$i] + $id
    }
    return $id
}

function HandleObject ($text, $objectType) {
    $object = New-Object psobject -Property @{
        TypeID = $objectType.value__
        Type   = $objectType
        ID     = GetObjectID($text)
    }

    $NAVObjects.Add($object) > $null
}

function GetObjectsFromFobFile ($fobFile) {
    if ($fobFile.Length -eq 0) {
        return
    }

    $numberOfLinesSinceLastObject = 0
    foreach ($line in $fobFile) {
        if ($numberOfLinesSinceLastObject -eq 2) {
            return
        }

        $objectType = $line.Split()[0]

        switch ($objectType) {
            $([ObjectType]::Table) { 
                HandleObject $line ([ObjectType]::Table)
                $numberOfLinesSinceLastObject = 0
            }
            $([ObjectType]::Page) { 
                HandleObject $line ([ObjectType]::Page)
                $numberOfLinesSinceLastObject = 0
            }
            $([ObjectType]::Report) { 
                HandleObject $line ([ObjectType]::Report)
                $numberOfLinesSinceLastObject = 0
            }
            $([ObjectType]::Codeunit) { 
                HandleObject $line ([ObjectType]::Codeunit)
                $numberOfLinesSinceLastObject = 0
            }
            $([ObjectType]::Query) { 
                HandleObject $line ([ObjectType]::Query)
                $numberOfLinesSinceLastObject = 0
            }
            $([ObjectType]::XMLport) { 
                HandleObject $line ([ObjectType]::XMLport)
                $numberOfLinesSinceLastObject = 0
            }
            $([ObjectType]::MenuSuite) { 
                HandleObject $line ([ObjectType]::MenuSuite)
                $numberOfLinesSinceLastObject = 0
            }
            Default {
                $numberOfLinesSinceLastObject++
            }
        }
    }
}

function LockObjectsInDatabase ($filter) {
    try {
        if ($filter.Length -eq 0) {
            return
        }

        $connection.Open()
        $query = "UPDATE [dbo].[Object]
                  SET [Locked] = 0,
                  [Locked By] = ''
                  WHERE UPPER([Locked By]) = UPPER('$lockedBy')
                   
                  UPDATE [dbo].[Object]
                  SET [Locked] = 1,
                  [Locked By] = '$lockedBy'
                  WHERE $filter"

        $command = $connection.CreateCommand()
        $command.CommandText = $query
        $command.ExecuteReader()
    }
    finally {
        $connection.Close()
    }
}

function CreateObjectFilter ($objects) {
    $filter = ""
    foreach ($object in $objects) {
        if ($filter.Length -eq 0) {
            $filter = $filter + "([Type] = $($object.TypeID) AND ID = $($object.ID))"
        }
        else {
            $filter = $filter + " OR ([Type] = $($object.TypeID) AND ID = $($object.ID))"
        }
    }
    return $filter
}

function GetObjectFilter ($objects) {
    $filter = ''

    if ($objects.Count -eq 0) {
        return
    }

    $filter = CreateObjectFilter $objects

    return $filter
}

function LockObjects ($fobFile) {
    GetObjectsFromFobFile $fobFile
    $filter = GetObjectFilter $NAVObjects
    LockObjectsInDatabase($filter)
}

LockObjects $fobFile

Write-Host 'Script completed!'
cmd /c pause
