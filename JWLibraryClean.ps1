$excludedPubs= "'sfl','rr','bhs','nwtsty'"


Clear-Host

###############################################
#Check if PSSQLite installed and import module
if (Get-Module -ListAvailable -Name PSSQLite) {
    Write-Host "Importing SQLite module..."
    Import-Module PSSQLite
    Write-Host "SQLite module imported."
} 
else {
    Write-Host "Installing SQLite module..."
    Install-Module -Name PSSQLite -Force
    Write-Host "Importing SQLite module..."
    Import-Module PSSQLite
    Write-Host "SQLite module imported."
}


Write-Host ""

###############################################
#Search JwLibrary app folder
try
{
    Write-Host "A pesquisar pasta..." -ForegroundColor Yellow

    $jwPath= $env:LOCALAPPDATA + "\Packages\"
    $jwFolder = Get-ChildItem $jwPath | Where-Object {$_.PSIsContainer -eq $true -and $_.Name -match "Watchtower"}
    Write-Host "Folder found!" -ForegroundColor Green
}
catch
{
    Write-Host "Folder not found" -ForegroundColor Red
    Break
}

###############################################
#Search app database
try
{
    Write-Host "A pesquisar base de dados..." -ForegroundColor Yellow

    $jwDB = Get-ChildItem -Path $jwFolder.FullName -Filter userData.db -Recurse -ErrorAction SilentlyContinue -Force
    Write-Host "Database found!" -ForegroundColor Green
}
catch
{
    Write-Host "Database not found" -ForegroundColor Red
    Break
}

###############################################
#Check if JWLibrary is open and close
Get-Process -Name JWLibrary -ErrorAction Ignore | Stop-Process -Force

###############################################
#Create connection to sqlite database
$sqlCon = New-SQLiteConnection -DataSource $jwDB.FullName


Write-Host "Searching for notes (pubs older than 1 month)" -ForegroundColor Yellow
Write-Host "Exclude publications: " $excludedPubs -ForegroundColor Yellow

###############################################
#Query: publications older than one month
$QueryList = "SELECT DISTINCT l.KeySymbol as Publication, l.IssueTagNumber as Date
FROM Note AS n
LEFT JOIN Location AS l ON l.LocationId = n.LocationId
WHERE l.KeySymbol NOT IN ($excludedPubs)                                   --exclude some publications
and n.LastModified < date('now', '-1 month')                               --does not include recent publications
and n.LocationId <> 'Null'                                                 --Does not inclued notes without publications associated
ORDER BY l.KeySymbol, l.IssueTagNumber ASC"

Invoke-SqliteQuery -SQLiteConnection $sqlCon -Query $QueryList | Format-Table

$pub=Read-Host "Which publication do you want to delete?"
if ($pub -eq "mwb" -or $pub -eq "w"){
    $pubDate=Read-Host "Which publication date do you want to delete?"
}
else{
    $pubDate='0'
}

###############################################
#Query: notes in publications older than one month (FULL_LIST)

#criar query baseada nos filtros acima.
$Query = "SELECT n.NoteId, n.LocationId, l.KeySymbol as Publication, l.IssueTagNumber as Date, n.LastModified AS LastModified, l.title as PublicationTitle,  n.Title as NoteTitle, n.Content
FROM Note AS n
LEFT JOIN Location AS l ON l.LocationId = n.LocationId
WHERE l.KeySymbol = '$pub' and l.IssueTagNumber = '$pubDate'
ORDER BY l.KeySymbol, l.IssueTagNumber ASC"

$queryResult = New-Object System.Data.DataTable
$queryResult = Invoke-SqliteQuery -SQLiteConnection $sqlCon -Query $Query -As DataTable

foreach ($row in $queryResult){

    #elimina notes
    
    if( ($row.Publication.ToString() -eq $pub.ToString()) -and ($row.Date.ToString() -eq $pubDate.ToString()) ){
        Write-Host "Deleting" $row.Publication "|" $row.Date " - NoteID: " $row.NoteID -ForegroundColor Red
        
        $QueryDelete = "DELETE FROM Note WHERE NoteId='$row.NoteID' "
        Invoke-SqliteQuery -SQLiteConnection $sqlCon -Query $QueryDelete

        Write-Host "Deleting" $pub -ForegroundColor Red
        $QueryDeleteLoc = "DELETE FROM Location WHERE KeySymbol = '$pub' and IssueTagNumber = '$pubDate' "
        Invoke-SqliteQuery -SQLiteConnection $sqlCon -Query $QueryDeleteLoc
    }
}








$sqlCon.Close()
$sqlCon.Dispose()

#########################################