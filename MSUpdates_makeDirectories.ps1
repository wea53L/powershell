# this script will build a series of directories relating to the current date
# formatted in the style of MS updates, "MSYY-###"
# built this for work, and it's the first time i made something significang in PS
# Copyright 2015 - Warren Kopp

# get current year from system and format for directories
$year = get-date -format yyyy
$year_short = get-date -f yy
$month = get-date -format MMMM
$month_short = get-date -format MM


# constants; current month name and prefix
$parent_dir = "$month_short - $month"
$prefix = "\MS$year_short-"

#options to verify directory for output
$checkifdir = "\\testhost.domain\software\patches\$year\$parent_dir"


# check if current month directory exists, if not, create (w/loop)
# this happens without user input

if ($checkifdir -eq $false){ 
   md $checkifdir
   }
   
# ask user for range of this month's bulletins(user input)
# create the names for directories needed. 
# create directories

$st = Read-Host "Please enter the number of the first bulletin: "
$end = Read-Host "Please enter the number of the last bulletin: "

$bulletin_range = $st..$end

for ($i = 0; $i -lt $bulletin_range.length; $i++){
    $string_name = $prefix + $bulletin_range[$i].ToString("000")
    $folder = $checkifdir + $string_name

    Write-Host $folder

    if ($folder -ne $false){
        md $folder
        }
    }
	
