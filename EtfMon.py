import datetime
import math

import numpy as np

from DataUtil import DataUtil
from pydantic import BaseModel
import pandas as pd

from Util import Util


class Dict(BaseModel):
    code: str
    name:str
    type: int;

class EtfMon:
    def getDownOrUp(self,codes,end_day,days):
        result = []
        for code in codes:
            code_rt={}
            df_etf=DataUtil().getEtfHis(str(code),end_day,days)
            df_etf.set_index('日期', inplace=True)
            df_etf['ma5'] = df_etf['收盘'].rolling(5).mean()
            df_etf_desc = df_etf.sort_values(by=['日期'], ascending=False)
            #print(df_etf_desc)
            min = 0
            minDay = ''
            firstDay = ''
            stop_flag = True
            flag=''
            for row in df_etf_desc.itertuples(index=True, name='Pandas'):
                Ma5=getattr(row, "ma5")
                if(min!=0 and Ma5 >= min and flag==''):
                    stop_flag = False
                    flag='Down'
                elif(min!=0 and Ma5 < min and flag==''):
                    stop_flag=False
                    flag='Up'
                elif(min != 0 and Ma5 >= min and flag == 'Down'):
                    stop_flag = False
                elif(min != 0 and Ma5 < min and flag == 'Up'):
                    stop_flag = False
                else:
                    stop_flag=True
                if(stop_flag and min!=0):break
                if(stop_flag and min==0):firstDay=getattr(row, "Index")

                min = Ma5
                minDay = getattr(row, "Index")

            #print(firstDay,minDay)
            #print(df_etf_desc.loc[firstDay]['收盘'])
            code_rt['code']=code
            code_rt['start_day'] = minDay
            code_rt['end_day'] = firstDay
            code_rt['start_close'] = '%.3f' % float(df_etf_desc.loc[minDay]['收盘'])
            code_rt['end_close'] = '%.3f' % float(df_etf_desc.loc[firstDay]['收盘'])
            code_rt['up_or_down'] = flag
            code_rt['ma5'] = '%.3f' % float(df_etf_desc.loc[firstDay]['ma5'])
            code_rt['close_rate'] = '%.3f' % float(df_etf_desc.loc[firstDay]['涨跌幅'])
            result.append(code_rt)
            #print(result)
        return  result


    def getEtfCodes(self):
        codes = ['516950','510500']
        file_path = r'D:/pythonProject/etfdict.xlsx'
        df = pd.read_excel(file_path, sheet_name="ETF")

        return df

    def getEtfList(self):
        etf_dict=EtfMon().getEtfCodes()
        codes=np.array(etf_dict['code']).tolist()
        etfDt=DataUtil().getEtfNow(codes)
        df_a = pd.DataFrame(etfDt)
        df_b = df_a.stack()
        df = df_b.unstack(0)
        #print(df)
        result = []
        for row in df.itertuples(index=True, name='Pandas'):
            etf={}
            etf['code']=getattr(row, "Index")
            etf['name'] = getattr(row, "name")
            etf['now'] = getattr(row, "now")
            etf['code'] = getattr(row, "Index")
            etf['rate'] = '%.4f' % float((getattr(row, "now") - getattr(row, "close")) / getattr(row, "close")*100) if float(getattr(row, "close")) != 0 else 0
            result.append(etf)
            #print(row)
        # avg = df['volume'][0] / df['turnover'][0] if int(df['turnover'][0]) != 0 else 0
        # now = df['now'][0]
        # rate = '%.4f' % float((now - avg) / now) if float(now) != 0 else 0
        # zdf = '%.4f' % float((now - float(df['close'][0])) / float(df['close'][0])) if float(df['close'][0]) != 0 else 0
        # avg = '%.3f' % avg
        # return {'code': id, 'now': now, 'close': df['close'][0], 'avg': avg, 'rate': rate, 'zdf': zdf,
        #         'name': str(df['name'][0])}

        return result


    def getBgColor(self,flag):
        if(flag=='Down'):
            return 'green'
        if(flag=='Up'):
            return 'red'
        if(flag>=0):
            return 'red'
        if(flag<0):
            return 'green'
    def insUpOrDown(self,now,days):
        if(Util().isHoliday(now)):return
        etf_dict = EtfMon().getEtfCodes()
        codes = np.array(etf_dict['code']).tolist()
        etf_dict.set_index('code', inplace=True)
        #now=datetime.datetime.now().strftime("%Y%m%d")
        rs=EtfMon().getDownOrUp(codes,now,days)
        print(rs)
        email_comment=[]
        for item in rs:
            code=item['code']
            event_day=item['end_day']
            start_day=item['start_day']
            end_day=item['end_day']
            end_close=float(item['end_close'])
            start_close=float(item['start_close'])
            ma5=float(item['ma5'])
            close_rate = float(item['close_rate'])
            uddays=Util().workdays(item['start_day'],item['end_day'])-1
            angle_norm = round(math.atan((end_close / start_close - 1)*100)*180 / 3.1415926,2)
            #ud_rate='%.2f' %float((end_close-start_close)/start_close*100)
            ud_rate = round((end_close - start_close) / start_close * 100,2)
            ma5_rate=round((end_close-ma5)/ma5*100,2)
            up_or_down=item['up_or_down']
            name=etf_dict.loc[code]['name']

            email_comment.append('<tr>')
            email_comment.append('<td align="center">' + str(code) + '</td>')
            email_comment.append('<td align="center">' + str(name) + '</td>')
            email_comment.append('<td align="center">' + str(event_day) + '</td>')
            email_comment.append('<td align="center" style="color:' + EtfMon().getBgColor(up_or_down) + '">' + str(up_or_down) + '</td>')
            email_comment.append('<td align="center">'  + str(start_day) + '</td>')
            email_comment.append('<td align="center">' + str(end_day) + '</td>')
            email_comment.append('<td align="center">' + str(end_close) + '元</td>')
            email_comment.append('<td align="center">' + str(start_close) + '元</td>')
            email_comment.append('<td align="center">' + str(ma5) + '元</td>')
            email_comment.append('<td align="center" style="color:' + EtfMon().getBgColor(close_rate) + '">' + str(close_rate) + '%</td>')
            email_comment.append('<td align="center" style="color:' + EtfMon().getBgColor(up_or_down) + '">' + str(uddays) + '天</td>')
            email_comment.append('<td align="center">' + str(angle_norm) + '%</td>')
            email_comment.append('<td align="center" style="color:' + EtfMon().getBgColor(ud_rate) + '">' + str(ud_rate) + '%</td>')
            email_comment.append('<td align="center" style="color:' + EtfMon().getBgColor(ma5_rate) + '">' + str(ma5_rate) + '%</td>')
            email_comment.append('</tr>')
        send_msg = '\n'.join(email_comment)
        EtfMon().email_msg(send_msg)


    def email_msg(self,msg):
        email_comment = []
        email_comment.append('<html>')
        email_comment.append('<b><p><h3><font size="2" color="black">您好：</font></h4></p></b>')
        email_comment.append('<p><font size="2" color="#000000">根据设置参数，现将监控到证券价格异动消息汇报如下：</font></p>')
        email_comment.append(
            '<table border="1px" cellspacing="0px"   width="600" bgcolor="#FFFFFF" style="border-collapse:collapse">')

        email_comment.append('<tr>')
        email_comment.append('<td align="center"><b>代码</b></td>')
        email_comment.append('<td align="center"><b>名称</b></td>')
        email_comment.append('<td align="center"><b>统计日期</b></td>')
        email_comment.append('<td align="center"><b>连续涨跌</b></td>')
        email_comment.append('<td align="center"><b>开始日期</b></td>')
        email_comment.append('<td align="center"><b>结束日期</b></td>')
        email_comment.append('<td align="center"><b>收盘价</b></td>')
        email_comment.append('<td align="center"><b>开始日期收盘价</b></td>')
        email_comment.append('<td align="center"><b>5日均价</b></td>')
        email_comment.append('<td align="center"><b>涨跌幅</b></td>')
        email_comment.append('<td align="center"><b>连续涨跌天数</b></td>')
        email_comment.append('<td align="center"><b>斜率</b></td>')
        email_comment.append('<td align="center"><b>连续涨跌幅</b></td>')
        email_comment.append('<td align="center"><b>Ma5幅率</b></td>')
        email_comment.append('</tr>')
        email_comment.append(msg)
        email_comment.append('</table>')
        email_comment.append('</html>')
        send_msg = '\n'.join(email_comment)
        Util().send_Email('541085504@qq.com', send_msg)


if __name__ == "__main__":
    codes=['516950','510500']
    print(type(codes))
    now = datetime.datetime.now().strftime("%Y%m%d")
    EtfMon().insUpOrDown(now, 40)
    rt=EtfMon().getBgColor(1)
   #  dd=20230331
   # # d = datetime.strptime(d, '%Y-%m-%d').date()
   #  i=0
   #  while i<20:
   #      EtfMon().insUpOrDown(str(d), 40)
   #      d=d-1
   #      i=i+1

    #print(rt)
