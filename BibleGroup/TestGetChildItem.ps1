$homeworkDirectory = "C:\BibleGroup\��ҵ"
$file = Get-ChildItem $homeworkDirectory -Filter *.*


$arrFileContent = Get-Content $file.fullname

$arrFileContent.length

$arrFileContent[0].length