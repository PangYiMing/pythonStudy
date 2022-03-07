--同步数据
. /etc/profile
. setEnvHome.sh
set -e
dd=`date -d yesterday +%Y%m%d`
dt=`date -d yesterday +%Y-%m-%d`
dm=`date +%Y%m%d`
sign=`date +%s`
resfile1=/home/fx/fx_tmp_data/xy/arealength_${dm}.txt
resfile2=/home/fx/fx_tmp_data/xy/arealength_log_${dm}.txt
echo "小蜜蜂来报——">$resfile1
hive -e "select concat('${dt}','门店面积/陈列长度数据情况如下：')
" >> $resfile1

hive -e "select concat('① ',' ',nullarea_tips,'，',nulllength_tips)
from fxetl.fxetl_shop_arealength_tips1 where dt='2022-01-04'
" >> $resfile1

hive -e "select concat('② ',changearea_tips,'\n',changearea_total)
from fxetl.fxetl_shop_arealength_tips2 where dt='2021-12-02'
" >> $resfile1

hive -e "select concat('③ ',changedislength_tips,'\n',changedislength_total)
from fxetl.fxetl_shop_arealength_tips3 where dt='2022-01-04'
" >> $resfile1

hive -e "select concat('④ ',addarea_tips,'\n',changeaddarea_total)
from fxetl.fxetl_shop_arealength_tips4 where dt='2021-12-01'
" >> $resfile1

hive -e "select concat('⑤ ',adddislength_tips,'\n',adddislength_total)
from fxetl.fxetl_shop_arealength_tips5 where dt='2022-01-04'
" >> $resfile1
echo "报告完毕！">>$resfile1

time2=$(date "+%Y/%m/%d %H:%M:%S")
echo $time2 > resfile2
# message=`cat $resfile1`
# echo "$message"
# echo "{\"msgtype\":\"text\",\"text\":{\"mentioned_mobile_list\":[\"@all\"],\"content\":\"$message\"}}"
# curl "https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=b70126af-b24a-4a3e-957c-e3c03363a963"  -H 'Content-Type: application/json' -d "{\"msgtype\":\"text\",\"text\":{\"mentioned_mobile_list\":[\"@all\"],\"content\":\"$message\"}}"