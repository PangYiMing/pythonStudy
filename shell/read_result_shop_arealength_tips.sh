###
 # @Author: xy
 # @Date: 2022-03-07 21:14:35
 # @Description: TODO xxx
 # @LastEditTime: 2022-03-07 21:17:15
### 
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
message=`cat $resfile2`


echo "$message"
echo "{\"msgtype\":\"text\",\"text\":{\"mentioned_mobile_list\":[\"@all\"],\"content\":\"$message\"}}"
curl "https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=b70126af-b24a-4a3e-957c-e3c03363a963"  -H 'Content-Type: application/json' -d "{\"msgtype\":\"text\",\"text\":{\"mentioned_mobile_list\":[\"@all\"],\"content\":\"$message\"}}"