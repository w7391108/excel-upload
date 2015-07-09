<?php
	include_once("PHPExcel.php");//引入PHP EXCEL类
	include_once("medoo.php");//引入数据库类
	include_once("UploadFile.php");//引入上传类
	define ('UPLOAD_PATH','./Uploads/');
	$fieldArr = array('shenfenzhenghao', 'zhigongbianhao', 'gongjijinzhanghao', 'danwei', 'banzu', 'xingming');

	if (isset($_FILES['excel']['size']) && $_FILES['excel']['size'] != null) {
		$upload = new UploadFile();
		$upload->maxSize = 10240000;
		$upload->allowExts  = array('xls');
		$dirname = UPLOAD_PATH . date('Ym', time()).'/'.date('d', time()).'/';
		if (!is_dir($dirname) && !mkdir($dirname, 0777, true)) {
			echo '<script type="text/javascript">alert("目录没有写入权限!!");</script>';
		}
		$upload->savePath = $dirname;
		$message = $upload->getErrorMsg();
		if(!$upload->upload()) {
			echo '<script type="text/javascript">alert("{$message}");</script>';
		}else{
			$info =  $upload->getUploadFileInfo();
		}

		if(is_array($info[0]) && !empty($info[0])){
			$savePath = $dirname . $info[0]['savename'];
		}else{
			echo '<script type="text/javascript">alert("上传失败");</script>';
		};

		if(empty($savePath) or !file_exists($savePath)){die('file not exists');}
		$PHPReader = new PHPExcel_Reader_Excel2007();        //建立reader对象
		if(!$PHPReader->canRead($savePath)){
				$PHPReader = new PHPExcel_Reader_Excel5();
				if(!$PHPReader->canRead($savePath)){
						echo 'no Excel';
						return ;
				}
		}
		$PHPExcel = $PHPReader->load($savePath);        //建立excel对象
		$currentSheet = $PHPExcel->getSheet(0);        //**读取excel文件中的指定工作表*/
		$allColumn = $currentSheet->getHighestColumn();        //**取得最大的列号*/
		$allRow = $currentSheet->getHighestRow();        //**取得一共有多少行*/
// var_dump($allColumn, $allRow);die();
		$data = array();
		$row = 1;
		$rowOne = $rowArr = $main = $time = array();
		$i = 0;
		// 取出excel第一行全部字段
		while(stringFromColumnIndex($i) != $allColumn) {
			$addr = stringFromColumnIndex($i) . $row;
			$cell = (String)$currentSheet->getCell($addr)->getValue();
			if($cell instanceof PHPExcel_RichText){ //富文本转换字符串
				$cell = $cell->__toString();
			}
			$rowOne[$row][stringFromColumnIndex($i)] = $cell;
			$i++;
		}
		$cell = (String)$currentSheet->getCell($allColumn . $row)->getValue();
		$rowOne[$row][$allColumn] = $cell;
		
		$newArr = array();
		foreach($rowOne[1] as $key => $value) {
			$tmp = Pinyin($value,'utf-8');
			if(!in_array($tmp, $fieldArr)) {
				$newArr[$key] = $tmp; 
			}
		}
		$db = new medoo(array(
			'database_type' => 'mysql',
			'database_name' => 'gzoa',
			'server' => '127.0.0.1',
			'username' => 'root',
			'password' => '',
			'port' => 3306,
			'charset' => 'utf8',
			'option' => array(PDO::ATTR_CASE => PDO::CASE_NATURAL)
		));
		
		$time = date("Ym", time());
		$result = $db->select("fields", ["field_id","field","name"], ["time[=]" => $time]);
		if(!empty($result)) {
			$db->query("delete from fields where time = {$time}");
		}
		foreach($newArr as $key => $value) {
			$insertData = array(
				'is_main'	=>	0,
				'field'		=>	$value,
				'name'		=>	$rowOne[1][$key],
				'form_type'	=>	'number',
				'time'		=>	$time
			);
			$db->insert("fields", $insertData);
		}
		
		
		$infoArr = array();
		foreach($newArr as $key => $value) {
			foreach($rowOne[1] as $list => $content) {
				if($key == $list) {
					$infoArr[$value] = $content;
				}
			}
		}
// var_dump($rowArr, $infoArr);die();
		$infoSql = '';
		foreach($infoArr as $key => $value) {
			if(!empty($value)) {
				$infoSql .= "`{$key}` float(25,2) NOT NULL COMMENT '{$value}',";
			}
		}
		$infoSql = rtrim($infoSql, ',');

		$db->query("DROP TABLE `info_{$time}`");

		$db->query("CREATE TABLE IF NOT EXISTS `info_{$time}` (
  `userid` int(10) unsigned NOT NULL COMMENT '用户id',
  `groupid` int(10) unsigned NOT NULL COMMENT '用户分组id',  {$infoSql}
) ENGINE=InnoDB  DEFAULT CHARSET=utf8 AUTO_INCREMENT=1;");

		$field_list = $db->select("fields", ["field_id","field","name"], ["OR" => ["is_main[=]" => 1,"time[=]" => $time]]);
		foreach($field_list as $key => $value) {
			foreach($rowOne[1] as $list => $content) {
				if($content == $value['name']) {
					var_dump($value['name']);
					$rowArr[$list] = $value['field'];
				}
			}
		}
var_dump($field_list, $rowOne, $rowArr);die();
		$db->query("delete from info where time = {$time}");
		for($rowIndex=2;$rowIndex<=$allRow;$rowIndex++){        //循环读取每个单元格的内容。注意行从1开始，列从A开始
					
				$i = 0;
				// 取出excel第一行全部字段
				while(stringFromColumnIndex($i) != $allColumn) {
					$colnum = stringFromColumnIndex($i);
					$addr = stringFromColumnIndex($i) . $rowIndex;
					$cell = (String)$currentSheet->getCell($addr)->getValue();
					if($cell instanceof PHPExcel_RichText){ //富文本转换字符串
						$cell = $cell->__toString();
					}
					if(!empty($cell)) {
						if(in_array($rowArr[$colnum], $fieldArr)) {
							$data1[$rowArr[$colnum]] = $cell;
						} else {
							$data2[$rowArr[$colnum]] = $cell;
						}
					}
					$i++;
				}
				$cell = (String)$currentSheet->getCell($allColumn . $allRow)->getValue();
				if(!empty($cell)) {
					if(in_array($rowArr[$allColumn], $fieldArr)) {
						$data1[$rowArr[$allColumn]] = $cell;
					} else {
						$data2[$rowArr[$allColumn]] = $cell;
					}
				}
		
				$data1['time'] = $time;
				$data1['groupid'] = $data2['groupid'] = 0;//设置信息分组id
				$name = isset($data1['xingming']) ? $data1['xingming'] : '';//判断如果帐号不存在，则创建帐号，默认密码123456
				$result = $db->select("admin", ["id","uid","username"], ["username[=]" => $name]);
				if(empty($result)) {
					$adminData = array(
						'uid'		=>	3,
						'username'	=>	$name,
						'password'	=>	md5('123456')
					);
					$db->insert("admin", $adminData);
				}
				$userid = $db->insert("info", $data1);
				if($userid) {
					$data2['userid'] = $userid;
					$last_user_id = $db->insert("info_{$time}", $data2);
				}
	// var_dump($data2);die();
		}
echo "<script language=javascript>" .
 				"alert('上传成功！'),parent.location.href='../main.php' " .
 				"</script>";
	}

 function stringFromColumnIndex($pColumnIndex = 0)  
{  
        static $_indexCache = array();  
   
        if (!isset($_indexCache[$pColumnIndex])) {  
            if ($pColumnIndex < 26) {  
                $_indexCache[$pColumnIndex] = chr(65 + $pColumnIndex);  
            } elseif ($pColumnIndex < 702) {  
                $_indexCache[$pColumnIndex] = chr(64 + ($pColumnIndex / 26)) . chr(65 + $pColumnIndex % 26);  
            } else { //开源软件:phpfensi.com 
                $_indexCache[$pColumnIndex] = chr(64 + (($pColumnIndex - 26) / 676)) . chr(65 + ((($pColumnIndex - 26) % 676) / 26)) . chr(65 + $pColumnIndex % 26);  
            }  
        }  
        return $_indexCache[$pColumnIndex];  
} 


?>
