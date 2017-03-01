<?php 
header("Content:Content-type:text/html;charset=utf-8");
include_once './GetEquipment.php';
//测试
$info = new get_equipment_info();

echo json_decode($info -> GetLang());
echo "</br>";
echo json_decode($info -> GetOs());
echo "</br>";
echo json_decode($info -> GetBrowser());
echo "</br>";
echo $_SERVER['HTTP_USER_AGENT'];
echo "</br>";
$ipa = $info -> Getip();
print_r($ipa);
echo "</br>";
$ip = "180.97.33.108";
print_r($info -> Getaddress($ip));
echo "</br>";
//echo $_SERVER['REMOTE_HOST'];
die;  
?>