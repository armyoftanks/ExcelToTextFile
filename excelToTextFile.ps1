######
#
#  Create bulk templated files with unique values from an excel worksheet. 
#  Example: create signature files for each contact in the Excel sheet
#  Another Example: create html pages based on excel data. such fancy. 
#
######

#community made library to handle logic between excel and powershell. Documentation here --> https://github.com/dfinke/ImportExcel
Import-Module ImportExcel 

#template you want to use for batching 
$template = "local\file\path\file.name"

#open spreadsheet containing data you want to save into individual files and specify which sheet in the file to use (default is Sheet1)
$collection = Import-Excel -Path "local\file\path\file.xlsx" -WorksheetName Sheet1

# uncomment next line to test if $collection is reading the excel file column names and cells by writing to the console
# $collection | Select 'NAME', 'EMAIL'

#1. loops through each row in the $collection Excel file
#2. for each column name found copy its cell value to a variable ($local_variable = $current_rows_cell.by_columnname)
foreach ($item in $collection) {
    $name = $item.NAME
    $title = $item.TITLE
    $work = $item.WORK 
    $mobile = $item.MOBILE
    $email = $item.EMAIL
    $filename = $item.FILENAME

    #3. create a new file for storing the current Excel items values
    $destination_file = "\path\to\your\"+$fname+".extension"

    #4  read the $template file and search for specified keywords
    #5. replace keywords in template with current iterations variable values
    #6. save modified template file as a new file with a specific $filename and location ($destination_file)
    (Get-Content $template) | ForEach-Object {
        $_  -replace 'NAME', $name `
            -replace 'NAME', $title `
            -replace 'NAME', $work `
            -replace 'NAME', $mobile `
            -replace 'NAME', $email `
            -replace 'NAME', $filename 
    } | Set-Content $destination_file
}   #7. move on to the next row to create next file



