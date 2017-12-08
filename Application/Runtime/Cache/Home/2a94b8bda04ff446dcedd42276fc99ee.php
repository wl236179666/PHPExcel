<?php if (!defined('THINK_PATH')) exit();?><!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport"
          content="width=device-width, user-scalable=no, initial-scale=1.0, maximum-scale=1.0, minimum-scale=1.0">
    <meta http-equiv="X-UA-Compatible" content="ie=edge">
    <title>Document</title>
</head>
<body>
<a href="<?php echo U('export');?>">点击导出Excel文件</a>

<form action="<?php echo U('import_excel');?>" method="post" enctype="multipart/form-data">
    <input type="file" name="import_file"/>
    <input type="hidden" name="names" value="import_file"/>
    <input type="submit" value="导入"/>
</form>

<hr>
<a href="<?php echo U('csv');?>">生成cvs表格</a>

<form action="<?php echo U('import_csv');?>" method="post" enctype="multipart/form-data">
    <input type="file" name="import_csv"/>
    <input type="hidden" name="namess" value="import_csv"/>
    <input type="submit" value="导入"/>
</form>
</body>
</html>