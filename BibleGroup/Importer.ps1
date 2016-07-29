$rootDirector = "c:/BibleGroup/"

$homeworkDirectory = $rootDirector + "作业"
$excelDirectory = $rootDirector + "Excel"

$arrFileContent = @()
$arrUnmatched = @()

$fileExcel = Get-ChildItem $excelDirectory -Filter *.*

$xl = New-Object -COM "Excel.Application"
$xl.Visible = $true
$wb = $xl.Workbooks.Open($fileExcel.fullname)
$ws1 = $wb.Sheets.Item(1)


function GetNumber($l)
{
    $index = $l.IndexOf("桌")
    if(-1 -eq $index)
    {
        return "";
    }
    $number = $l.SubString(0, $index)
    
    return $number
}

function GetName($l)
{
    $index = $l.IndexOf("-")
    $name = $l
    if(-1 -ne $index)
    {
        $name = $l.SubString($index + 1, $l.length - ($index + 1))
    }
    
    $index = $name.IndexOf("F")
    if(-1 -eq $index)
    {
        $index = $name.IndexOf("M")
        if(-1 -eq $index)
        {
            $index = $name.IndexOf(" ")
            if(-1 -eq $index)
            {
                $index = $name.IndexOf("弟兄")
                if(-1 -eq $index)
                {
                    $index = $name.IndexOf("姊妹")
                }                 
            }    
                      
        }        
    }
    
    if($index -ne -1)
    {
        $name = $name.SubString(0, $index);
    }
    
    return $name;
}

function GetSex($l)
{
    $index = $l.IndexOf("F")
    if(-1 -ne $index)
    {
        return "F"
    }
    
    $index = $l.IndexOf("M")
    if(-1 -ne $index)
    {
        return "M"
    } 
    
    $index = $l.IndexOf("姊妹")
    if(-1 -ne $index)
    {
        return "F"
    }
    
    $index = $l.IndexOf("弟兄")
    if(-1 -ne $index)
    {
        return "M"
    }     
    
    return "";   
}

function GetYear($filename)
{
    $year = $filename.SubString(0, 4);
    return $year;
}

function GetMonth($filename)
{
    $month = $filename.SubString(4, 2);
    return $month;
}

function GetDay($filename)
{
    $day = $filename.SubString(6, 2);
    return $day;
}

function FindColInSheet1($year, $month, $day, $ws1 )
{
    $col = 1
    while($true)
    {
        $text = $ws1.Cells.Item(1, $col).text;
        if(0 -eq $text.length)
        {
            break;
        }
        
        $col++
    }
    
    return $col;
}


function InsertToSheet1($ws1, $number, $name, $sex, $col)
{
    $row = FindRowInSheet1 $ws1 $number $name $sex
    if(-1 -eq $row)
    {
        return 0
    }
    
    $ws1.Cells.Item($row, $col) = "√"
    
    return 1

}


function FindRowInSheet1($ws1, $number, $name, $sex)
{
    $nameShouldInSheet = $name
    if($sex -eq "F")
    {
        $nameShouldInSheet = $nameShouldInSheet + " " + "姊妹"
    }
    
    if($sex -eq "M")
    {
        $nameShouldInSheet = $nameShouldInSheet + " " + "弟兄"
    }

    $rows = $ws1.UsedRange.Rows.Count
    $row = -1;
    for($i=2; $i -le $rows; $i++)
    {
        $curName = $ws1.Cells.Item($i, 2).text;
        if($nameShouldInSheet -eq $curName)
        {
            $row = $i;
            $curNumber = $ws1.Cells.Item($i, 1).text;
            if($curNumber -eq $number)
            {
                break;
            }
        }
    }
    
    return $row

}


$file = Get-ChildItem $homeworkDirectory -Filter *.*
$arrFileContent = Get-Content $file.fullname

$year = GetYear $file.fullname
$month = GetMonth $file.fullname
$day = GetDay $file.fullname

$col = FindColInSheet1 $year $month $day $ws1

$NumberCount = 0;
$totalCount = 0;
$sucCount = 0;
for($i=0; $i -lt $arrFileContent.length; $i++)
{
    $line = $arrFileContent[$i]
    if(0 -eq $line.length)
    {
        continue
    }
    
    $line = $line.tostring();
     
    $number = GetNumber $line
    if($number.length -gt 0)
    {
        $NumberCount++
    }
    
    $name = GetName $line
    $sex = GetSex $line
    
    $totalCount++;
    
    $result = InsertToSheet1 $ws1 $number $name $sex $col 
     
    if( 0 -eq $result)
    {
        $arrUnmatched += $line
    
    }
    else
    {
        $sucCount++;
        "成功导入" + $name + “的作业， 导入数量：” + $sucCount
    }
}


$ws1.Cells.Item(1, $col) = $year + "." + $month + "." + $day + "\r\n（" + $totalCount + "）"


$wb.Save()
$wb.Close()

$xl.Quit()

if($arrUnmatched.length -gt 0)
{
    "成功导入" + $sucCount +　"人，" + "未匹配成功的弟兄姐妹有："
    $arrUnmatched

}
else
{
    "所有弟兄姐妹（" + $sucCount + "人）的作业已经成功导入EXCEL表。"
}

