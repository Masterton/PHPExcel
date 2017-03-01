<?php
header("Content-type: text/html; charset=utf-8");
include './Excel.php';
include './DB.php';

$download = new Excel();

$db = new DB();
$data = $db->connect();

$header = array('id', '姓名', '年龄', 'QQ', '电话');
/* $data = array(
		array('1','小王','男','20','100'),
		array('2','小李','男','20','101'),
		array('3','小张','女','20','102'),
		array('4','小赵','女','20','103')
); */
$title = "呆呆.xlsx";
$download->Export($data, $header, $title);