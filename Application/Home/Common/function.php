<?php

header("Content-type:text/html;charset=utf-8");

/**
 * 数组转xls格式的excel文件
 * @param  string $expTitle  生成的excel文件名
 * @param  array  $expCellName  生成的excel文件的表头
 * @param  array  $data      需要生成excel文件的数组
 */
function exportExcel($expTitle,$expCellName,$expTableData){
    $xlsTitle = iconv('utf-8', 'gb2312', $expTitle);//文件名称
    $fileName = $_SESSION['account'].date('_YmdHis');//or $xlsTitle 文件名称可根据自己情况设定
    $cellNum = count($expCellName);
    $dataNum = count($expTableData);

    vendor("PHPExcel.PHPExcel");
    $objPHPExcel = new PHPExcel();


    $cellName = array('A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE','AF','AG','AH','AI','AJ','AK','AL','AM','AN','AO','AP','AQ','AR','AS','AT','AU','AV','AW','AX','AY','AZ');
    $objPHPExcel->getActiveSheet(0)->mergeCells('A1:'.$cellName[$cellNum-1].'1');//合并单元格
    // $objPHPExcel->setActiveSheetIndex(0)->setCellValue('A1', $expTitle.'  Export time:'.date('Y-m-d H:i:s'));
    for($i=0;$i<$cellNum;$i++){
        $objPHPExcel->setActiveSheetIndex(0)->setCellValue($cellName[$i].'2', $expCellName[$i][1]);
    }
    // Miscellaneous glyphs, UTF-8
    for($i=0;$i<$dataNum;$i++){
        for($j=0;$j<$cellNum;$j++){
            $objPHPExcel->getActiveSheet(0)->setCellValue($cellName[$j].($i+3), $expTableData[$i][$expCellName[$j][0]]);
        }
    }

    header('pragma:public');
    header('Content-type:application/vnd.ms-excel;charset=utf-8;name="'.$xlsTitle.'.xls"');
    header("Content-Disposition:attachment;filename=$fileName.xls");//attachment新窗口打印inline本窗口打印

    $objWriter = \PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
    $objWriter->save('php://output');
    exit;
}

//Excel导入
function import_excel(){
    $config=array(
        'allowExts'=>array('xlsx','xls'),
        'savePath'=>'./Public/upload/',
        'saveRule'=>'time',
    );
    $upload = new Org\Util\UploadFile($config);
    if (!$upload->upload()) {
        $this->error($upload->getErrorMsg());
    } else {
        $info = $upload->getUploadFileInfo();
    }

    vendor("PHPExcel.PHPExcel");
    vendor("PHPExcel.PHPExcel.Writer.Excel5");
    vendor("PHPExcel.PHPExcel.IOFactory");
    $file_name=$info[0]['savepath'].$info[0]['savename'];
    $ext = strtolower(pathinfo($file_name,PATHINFO_EXTENSION));//获取扩展名

    //同时兼容两种后缀
    if($ext == 'xlsx'){
        $objReader = \PHPExcel_IOFactory::createReader('Excel2007');
        $objPHPExcel = $objReader->load($file_name,'utf-8');
    }elseif($ext == 'xls'){
        $objReader = \PHPExcel_IOFactory::createReader('Excel5');
        $objPHPExcel = $objReader->load($file_name,$encode='utf-8');
    }

    $sheet = $objPHPExcel->getSheet(0);
    $highestRow = $sheet->getHighestRow(); // 取得总行数
    $highestColumn = $sheet->getHighestColumn(); // 取得总列数
    for($i=2;$i<=$highestRow;$i++)
    {
        $data['name'] = $objPHPExcel->getActiveSheet()->getCell("A".$i)->getValue();
        $data['sex'] = $objPHPExcel->getActiveSheet()->getCell("B".$i)->getValue();
        $data['old'] = $objPHPExcel->getActiveSheet()->getCell("C".$i)->getValue();
        $data['hobby'] = $objPHPExcel->getActiveSheet()->getCell("D".$i)->getValue();
        $data['add_time']= time();

        //TODO:请自行设计数据库文件

        var_dump($data);die;

        M('chanpin')->add($data);
    }
}

//本地上传文件的IO操作
function file_upload($src_file,$dest_file){
    $pdir=dirname($dest_file);
    if(!is_dir($pdir)) @mkdir($pdir,0777);
    return copy($src_file,$dest_file);
}


/**
 * 数据转csv格式的excle
 * @param  array $data      需要转的数组
 * @param  string $header   要生成的excel表头
 * @param  string $filename 生成的excel文件名
 *      示例数组：
$data = array(
'1,2,3,4,5',
'6,7,8,9,0',
'1,3,5,6,7'
);
 */
function create_csv($data,$header=null,$filename='simple.csv'){
    // 如果手动设置表头；则放在第一行,不要表头可以注释掉
    $header='用户名,密码,头像,性别,手机号';
    if (!is_null($header)) {
        array_unshift($data, $header);
    }
    // 防止没有添加文件后缀
    $filename=str_replace('.csv', '', $filename).'.csv';
    ob_clean();
    Header( "Content-type:  application/octet-stream ");
    Header( "Accept-Ranges:  bytes ");
    Header( "Content-Disposition:  attachment;  filename=".$filename);
    foreach( $data as $k => $v){
        // 如果是二维数组；转成一维
        if (is_array($v)) {
            $v=implode(',', $v);
        }
        // 替换掉换行
        $v=preg_replace('/\s*/', '', $v);
        // 解决导出的数字会显示成科学计数法的问题
        $v=str_replace(',', "\t,", $v);
        // 转成gbk以兼容office乱码的问题
        echo iconv('UTF-8','GBK',$v)."\t\r\n";
    }
}
