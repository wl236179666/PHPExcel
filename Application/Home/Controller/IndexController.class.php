<?php
namespace Home\Controller;
use Think\Controller;
class IndexController extends Controller {

    public function index(){
        $this -> display();
    }

    /*
     * 导出Excel
     */
    public function export(){

        $xlsCell  = array(
            array('id','序列号'),
            array('order_num','订单号'),
            array('uid','收件人'),
            array('phone','联系方式'),
            array('address','收货地址'),
            array('addtime','下单时间'),
            array('pay_status','订单状态'),
            array('pay_time','支付时间'),
            array('total','总金额'),
        );

        $xlsData[0]['id']="1";
        $xlsData[0]['order_num']="TXT2236655214523565";
        $xlsData[0]['uid']="李玲";
        $xlsData[0]['phone']="13613999999";
        $xlsData[0]['address']="河南省郑州市";
        $xlsData[0]['addtime']=date('Y-m-d H:i',time());
        $xlsData[0]['pay_status']="已支付";
        $xlsData[0]['pay_time']=date('Y-m-d H:i',time() + 100);
        $xlsData[0]['total']=500.00;

        $xlsData[1]['id']="2";
        $xlsData[1]['order_num']="TXT2236655564256485";
        $xlsData[1]['uid']="王强";
        $xlsData[1]['phone']="13613888888";
        $xlsData[1]['address']="河南省郑州市";
        $xlsData[1]['addtime']=date('Y-m-d H:i',time());
        $xlsData[1]['pay_status']="已支付";
        $xlsData[1]['pay_time']=date('Y-m-d H:i',time() + 100);
        $xlsData[1]['total']=128.50;

        exportExcel("订单表",$xlsCell,$xlsData);
    }

    /*
     * 导入Excel文件
     */
    public function import_excel(){
        if(!empty($_FILES[I('post.names')]['name'])){
            import_excel();
        }else{
            $this->error("请选择上传的文件");
        }
    }

    //生成CSV表格
    public function csv()
    {
        $data=array(
            '1,2,3,4,5',
            '6,7,8,9,0',
            '1,3,5,6,7'
        );
        create_csv($data);
    }

    /**
     * 导入csv格式的数据
     */
    public function import_csv(){
        if(!empty($_FILES[I('post.namess')]['name'])){
            $text=file_get_contents($_FILES[I('post.namess')]['tmp_name']);

            //---------------------------这些是为读取中文乱码准备的 start ----------------------------
            define('UTF32_BIG_ENDIAN_BOM', chr(0x00) . chr(0x00) . chr(0xFE) . chr(0xFF));
            define('UTF32_LITTLE_ENDIAN_BOM', chr(0xFF) . chr(0xFE) . chr(0x00) . chr(0x00));
            define('UTF16_BIG_ENDIAN_BOM', chr(0xFE) . chr(0xFF));
            define('UTF16_LITTLE_ENDIAN_BOM', chr(0xFF) . chr(0xFE));
            define('UTF8_BOM', chr(0xEF) . chr(0xBB) . chr(0xBF));
            $first2 = substr($text, 0, 2);
            $first3 = substr($text, 0, 3);
            $first4 = substr($text, 0, 3);
            $encodType = "";
            if ($first3 == UTF8_BOM)
                $encodType = 'UTF-8 BOM';
            else if ($first4 == UTF32_BIG_ENDIAN_BOM)
                $encodType = 'UTF-32BE';
            else if ($first4 == UTF32_LITTLE_ENDIAN_BOM)
                $encodType = 'UTF-32LE';
            else if ($first2 == UTF16_BIG_ENDIAN_BOM)
                $encodType = 'UTF-16BE';
            else if ($first2 == UTF16_LITTLE_ENDIAN_BOM)
                $encodType = 'UTF-16LE';
            //下面的判断主要还是判断ANSI编码的·
            if ($encodType == '') {//即默认创建的txt文本-ANSI编码的
                $content = iconv("GBK", "UTF-8", $text);
            } else if ($encodType == 'UTF-8 BOM') {//本来就是UTF-8不用转换
                $content = $text;
            } else {//其他的格式都转化为UTF-8就可以了
                $content = iconv($encodType, "UTF-8", $text);
            }
            //---------------------------这些是为读取中文乱码准备的 end ----------------------------

            $data=explode("\r\n", $content);
            var_dump($data);die;
        }else{
            $this->error("请选择上传的文件");
        }
    }
}