<?php
header('Content-Type:text/json;charset=UTF-8');
$id = $_GET['id']??'cctv1';
$n = [
    'cctv1' => 'cctv1hd8m/8000000',//CCTV-1HD
    'cctv2' => 'cctv2hd8m/8000000',//CCTV-2
    'cctv3' => 'cctv38m/8000000',//CCTV-3HD
    'cctv4' => 'cctv4hd8m/8000000',//CCTV-4
    'cctv5' => 'cctv58m/8000000',//CCTV-5HD
    'cctv5p' => 'cctv5hd8m/8000000',//CCTV-5+HD
    'cctv5p2' => 'cctv5phd8m/8000000',//CCTV-5+HD2
    'cctv6' => 'cctv6hd8m/8000000',//CCTV-6HD
    'cctv7' => 'cctv7hd8m/8000000',//CCTV-7
    'cctv8' => 'cctv8hd8m/8000000',//CCTV-8HD
    'cctv9' => 'cctv9hd8m/8000000',//CCTV-9
    'cctv10' => 'cctv10hd8m/8000000',//CCTV-10HD
    'cctv11' => 'cctv11hd8m/8000000',//CCTV-11
    'cctv12' => 'cctv12hd8m/8000000',//CCTV-12
    'cctv13' => 'cctv13xwhd8m/8000000',//CCTV-13
    'cctv14' => 'cctvsehd8m/8000000',//CCTV-14HD
    'cctv15' => 'cctv15hd8m/8000000',//CCTV-15
    'cctv16' => 'cctv16hd8m/8000000',//CCTV-16HD
    'cctv164k' => 'cctv16hd4k/15000000',//CCTV-16HD4k
    'cctv17' => 'cctv17hd8m/8000000',//CCTV-17HD
    'cctv4k' => 'cctv4k/15000000',//CCTV4K
    'cgtn' => 'ottcctvnews/1300000',//CGTN

    'jswshd' => 'jswshd8m/8000000',//��������
    'gxwshd' => 'gxwshd8m/8000000',//��������
    'scwshd' => 'scwshd8m/8000000',//�Ĵ�����
    'hunanwshd' => 'hunanwshd8m/8000000',//��������
    'zjwshd' => 'zjwshd8m/8000000',//�㽭����
    'dfwshd' => 'dfwshd8m/8000000',//��������
    'bjwshd' => 'bjwshd8m/8000000',//��������
    'tjwshd' => 'tjwshd8m/8000000',//�������
    'lnwshd' => 'lnwshd8m/8000000',//��������
    'ahwshd' => 'ahwshd8m/8000000',//��������
    'hljwshd' => 'hljwshd8m/8000000',//����������
    'gdwshd' => 'gdwshd8m/8000000',//�㶫����
    'szwshd' => 'szwshd8m/8000000',//��������
    'hubeiwshd' => 'hubeiwshd8m/8000000',//��������
    'jlwshd' => 'jlwshd8m/8000000',//��������
    'sdwshd' => 'sdws8m/8000000',//ɽ������
    'jxwshd' => 'jxwshd8m/8000000',//��������
    'hnwshd' => 'hnwshd8m/8000000',//��������
    'hbwshd' => 'hbwshd8m/8000000',//�ӱ�����
    'gswshd' => 'gswshd8m/8000000',//��������
    'cqwshd' => 'cqwshd8m/8000000',//��������
    'dnwshd' => 'dnwshd8m/8000000',//��������
    'ynwshd' => 'ynwshd8m/8000000',//��������
    'gzwshd' => 'gzwshd8m/8000000',//��������
    'hainanwshd' => 'hainanwshd8m/8000000',//��������

    'cetv1' => 'zgjy1t8m/8000000',//CETV1HD
    'cetv2' => 'cetv2/2500000',//CETV2
    'cetv4' => 'zgjy4hd8m/8000000',//CETV4HD

    'zgtq' => 'zgqx/1300000',//�й�����
    'xqjx' => 'xqjx8m/8000000',//Ϸ����ѡ
    'bjjskj' => 'dajs8m/8000000',//��ʵ�ƽ�
    'bjkk' => 'kkse8m/8000000',//�����ٶ�
    'qjshd' => 'qjshd8m/8000000',//ȫ��ʵ
    'rmzy' => 'rmzy8m/8000000',//��������
    'qcsj' => 'qcsj8m/8000000',//��������
    'jkys' => 'jkys8m/8000000',//��������
    'ktxjx' => 'ktxjx8m/8000000',//�����¾�ѡ
    'bbkt' => 'bbkt8m/8000000',//�ٱ����
    'qqdp' => 'qqdp8m/8000000',//ȫ���Ƭ
    'hyyy' => 'hyyy8m/8000000',//����ӰԺ
    'djtt' => 'djtt8m/8000000',//�羺����
    'qcdm' => 'qcdm8m/8000000',//�ഺ����
    'bbdh' => 'bbdh8m/8000000',//��������
    'xgyy' => 'xgyy8m/8000000',//�ǹ�ӰԺ
    'dzjc' => 'dzjc8m/8000000',//��ս�糡
    'rmjc' => 'rmjc8m/8000000',//���ž糡
    'dfcj' => 'dfcjhd8m/8000000',//�����ƾ�
    'dfys' => 'dfyshd8m/8000000',//����Ӱ��
    'dfgw' => 'dfgwsxhd8m/8000000',//��������
    'dycj' => 'dycjhd8m/8000000',//��һ�ƾ�
    'shxwzh' => 'xwzhhd8m/8000000',//�Ϻ�����
    'shds' => 'dshd8m/8000000',//�Ϻ�����
    'shjsrw' => 'jspdhd/4000000',//�Ϻ���ʵ
    'shics' => 'icshd8m/8000000',//��������
    'hhxd' => 'hhxd8m/8000000',//�Ŷ���ͨ
    'wxty' => 'wxtyhd8m/8000000',//�������� 
    'fztd' => 'fztd8m/8000000',//�������
    'hxjc' => 'hxjc8m/8000000',//��Ц�糡
    'hxjc4k' => 'hxjc4k/1500000',//��Ц�糡4K
    'dsjc' => 'dsjc8m/8000000',//���о糡
    'qcxj' => 'qcxjhd8m/8000000',//�߲�Ϸ��
    'dmxc' => 'dmxc8m/8000000',//�����㳡
    'jbty' => 'jbtyhd8m/8000000',//��������
    'xsj' => 'xsjhd8m/8000000',//���Ӿ���
    'yxfy' => 'yxfy8m/8000000',//��Ϸ����
    'shss' => 'shss8m/8000000',//����ʱ��
    'jspd' => 'jingsepd8m/8000000',//��ɫѧ��
    'qjs' => 'qjshd8m/8000000',//����HD
    'mlzq' => 'mlyyhd8m/8000000',//��������
    'shjy' => 'setvhd/8000000',//�Ϻ�����
    'pdtv' => 'hhse/2500000',//�ֶ�����̨

    'zqpd' => 'zqpd8m/8000000',//����Ƶ��

    'cpd' => 'cpdhdavs8m/8000000',//��Ƶ��
    'klcd' => 'klcd8m/8000000',//���ִ���
    'jyjs' => 'jyjs8m/8000000',//��ӥ��ʵ
    'jykt' => 'jykt/1300000',//��ӥ��ͨ
    'tcpd' => 'taocihd/8000000',//�մ�Ƶ��
    'jjkt' => 'jjkt/1300000',//�μѿ�ͨ

    'cftx' => 'cftx/2500000',//�Ƹ�����
     ];
date_default_timezone_set("Asia/Shanghai");
$date = date('YmdH');
$timestamp = intval((time()-50)/10);
$stream = 'http://121.12.115.154/liveplay-kk.rtxapp.com/live/program/live/'.$n[$id].'/'.$date.'/';
//$stream = 'http://dongfang5g-mobile-v5-live.bestvcdn.com.cn/live/program/live/'.$n[$id].'/'.$date.'/';
$current = "#EXTM3U"."\r\n";
$current.= "#EXT-X-VERSION:3"."\r\n";
$current.= "#EXT-X-TARGETDURATION:3"."\r\n";
$current.= "#EXT-X-MEDIA-SEQUENCE:{$timestamp}"."\r\n";
for ($i=0; $i<3; $i++) {
    $current.= "#EXTINF:3.000,"."\r\n";
    $current.= $stream.$timestamp.".ts"."\r\n";
    $timestamp = $timestamp + 1;
    }
print_r($current);
?>
