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

    'jswshd' => 'jswshd8m/8000000',//江苏卫视
    'gxwshd' => 'gxwshd8m/8000000',//广西卫视
    'scwshd' => 'scwshd8m/8000000',//四川卫视
    'hunanwshd' => 'hunanwshd8m/8000000',//湖南卫视
    'zjwshd' => 'zjwshd8m/8000000',//浙江卫视
    'dfwshd' => 'dfwshd8m/8000000',//东方卫视
    'bjwshd' => 'bjwshd8m/8000000',//北京卫视
    'tjwshd' => 'tjwshd8m/8000000',//天津卫视
    'lnwshd' => 'lnwshd8m/8000000',//辽宁卫视
    'ahwshd' => 'ahwshd8m/8000000',//安徽卫视
    'hljwshd' => 'hljwshd8m/8000000',//黑龙江卫视
    'gdwshd' => 'gdwshd8m/8000000',//广东卫视
    'szwshd' => 'szwshd8m/8000000',//深圳卫视
    'hubeiwshd' => 'hubeiwshd8m/8000000',//湖北卫视
    'jlwshd' => 'jlwshd8m/8000000',//吉林卫视
    'sdwshd' => 'sdws8m/8000000',//山东卫视
    'jxwshd' => 'jxwshd8m/8000000',//江西卫视
    'hnwshd' => 'hnwshd8m/8000000',//河南卫视
    'hbwshd' => 'hbwshd8m/8000000',//河北卫视
    'gswshd' => 'gswshd8m/8000000',//甘肃卫视
    'cqwshd' => 'cqwshd8m/8000000',//重庆卫视
    'dnwshd' => 'dnwshd8m/8000000',//东南卫视
    'ynwshd' => 'ynwshd8m/8000000',//云南卫视
    'gzwshd' => 'gzwshd8m/8000000',//贵州卫视
    'hainanwshd' => 'hainanwshd8m/8000000',//海南卫视

    'cetv1' => 'zgjy1t8m/8000000',//CETV1HD
    'cetv2' => 'cetv2/2500000',//CETV2
    'cetv4' => 'zgjy4hd8m/8000000',//CETV4HD

    'zgtq' => 'zgqx/1300000',//中国天气
    'xqjx' => 'xqjx8m/8000000',//戏曲精选
    'bjjskj' => 'dajs8m/8000000',//纪实科教
    'bjkk' => 'kkse8m/8000000',//卡酷少儿
    'qjshd' => 'qjshd8m/8000000',//全纪实
    'rmzy' => 'rmzy8m/8000000',//热门综艺
    'qcsj' => 'qcsj8m/8000000',//超级体育
    'jkys' => 'jkys8m/8000000',//健康养生
    'ktxjx' => 'ktxjx8m/8000000',//看天下精选
    'bbkt' => 'bbkt8m/8000000',//百变课堂
    'qqdp' => 'qqdp8m/8000000',//全球大片
    'hyyy' => 'hyyy8m/8000000',//华语影院
    'djtt' => 'djtt8m/8000000',//电竞天堂
    'qcdm' => 'qcdm8m/8000000',//青春动漫
    'bbdh' => 'bbdh8m/8000000',//宝宝动画
    'xgyy' => 'xgyy8m/8000000',//星光影院
    'dzjc' => 'dzjc8m/8000000',//谍战剧场
    'rmjc' => 'rmjc8m/8000000',//热门剧场
    'dfcj' => 'dfcjhd8m/8000000',//东方财经
    'dfys' => 'dfyshd8m/8000000',//东方影视
    'dfgw' => 'dfgwsxhd8m/8000000',//东方购物
    'dycj' => 'dycjhd8m/8000000',//第一财经
    'shxwzh' => 'xwzhhd8m/8000000',//上海新闻
    'shds' => 'dshd8m/8000000',//上海都市
    'shjsrw' => 'jspdhd/4000000',//上海纪实
    'shics' => 'icshd8m/8000000',//上视外语
    'hhxd' => 'hhxd8m/8000000',//炫动卡通
    'wxty' => 'wxtyhd8m/8000000',//五星体育 
    'fztd' => 'fztd8m/8000000',//法治天地
    'hxjc' => 'hxjc8m/8000000',//欢笑剧场
    'hxjc4k' => 'hxjc4k/1500000',//欢笑剧场4K
    'dsjc' => 'dsjc8m/8000000',//都市剧场
    'qcxj' => 'qcxjhd8m/8000000',//七彩戏剧
    'dmxc' => 'dmxc8m/8000000',//动漫秀场
    'jbty' => 'jbtyhd8m/8000000',//劲爆体育
    'xsj' => 'xsjhd8m/8000000',//新视觉、
    'yxfy' => 'yxfy8m/8000000',//游戏风云
    'shss' => 'shss8m/8000000',//生活时尚
    'jspd' => 'jingsepd8m/8000000',//金色学堂
    'qjs' => 'qjshd8m/8000000',//乐游HD
    'mlzq' => 'mlyyhd8m/8000000',//魅力足球
    'shjy' => 'setvhd/8000000',//上海教育
    'pdtv' => 'hhse/2500000',//浦东电视台

    'zqpd' => 'zqpd8m/8000000',//足球频道

    'cpd' => 'cpdhdavs8m/8000000',//茶频道
    'klcd' => 'klcd8m/8000000',//快乐垂钓
    'jyjs' => 'jyjs8m/8000000',//金鹰纪实
    'jykt' => 'jykt/1300000',//金鹰卡通
    'tcpd' => 'taocihd/8000000',//陶瓷频道
    'jjkt' => 'jjkt/1300000',//嘉佳卡通

    'cftx' => 'cftx/2500000',//财富天下
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
