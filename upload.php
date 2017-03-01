<!DOCTYPE html>
<html>
<head>
    <title>测试excel表导入</title>
    <meta charset="utf-8">
</head>
<body>
	<form method="post" action="upload.php" enctype="multipart/form-data">
		<input type="file" name="file" value="文件选择"/>
		<input type="submit" value="导入"/>
	</form>
</body>
</html>
<?php
	if(!empty($_FILES['file']['tmp_name'])){
		include './Excel.php';
		$aa = $_FILES['file']['tmp_name'];
		$str = file_get_contents($aa);
		
		$name = $_FILES['file']['name'];
		$extend = strrchr ($name, '.');
		if($_FILES['file']['type'] == "application/vnd.ms-excel" || $_FILES['file']['type'] == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"){
			$excel = new Excel();
			if($extend == ".xls"){
				$bb = $excel->Import($aa, 1);
			}else if($extend == ".xlsx"){
				$bb = $excel->Import($aa, 2);
			}else{
				$bb = "必须是excel文件111111";
			}
		}else{
			$bb = "必须是excel文件";
		}
		
		print_r("<pre>");
		print_r($bb);
	}
	
	/* if($_FILES[''][]){
		
	} */
?>