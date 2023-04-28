import akshare as ak
import datetime
import easyquotation as eq


class DataUtil:
    def getEtfHis(self,code,end_day,day):
        #now = datetime.datetime.now().strftime("%Y%m%d")
        now=end_day
        start_date=(datetime.datetime.strptime(now, "%Y%m%d")-datetime.timedelta(days=day)).strftime("%Y%m%d")
        df=ak.fund_etf_hist_em(symbol=code,period='daily',start_date=start_date,end_date=now,adjust='qfq')
        return df

    def getEtfNow(self,codes):
        qs = eq.use('sina')  # 新浪 ['sina'] 腾讯 ['tencent', 'qq']
        etf = qs.stocks(codes)
        return etf


if __name__ == "__main__":
    codes=['516950','510500']
   # df =DataUtil().getEtfNow(codes)
    code='510500'
    end_day='20230424'
    day=10
    df = DataUtil().getEtfHis(code,end_day,day)
    print(df)