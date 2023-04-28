import datetime as dt
from datetime import datetime, timedelta
from chinese_calendar import is_holiday
import smtplib
from email.mime.text import MIMEText


class Util:
    def workdays(self,start, end):
        '''
        计算两个日期间的工作日
        start:开始时间
        end:结束时间
        '''
        # 字符串格式日期的处理
        if type(start) == str:
            start = datetime.strptime(start, '%Y-%m-%d').date()
        if type(end) == str:
            end = datetime.strptime(end, '%Y-%m-%d').date()
            # 开始日期大，颠倒开始日期和结束日期
        if start > end:
            start, end = end, start
        counts = 0
        while True:
            if start > end:
                break
            if is_holiday(start) or start.weekday() == 5 or start.weekday() == 6:
                start += timedelta(days=1)
                continue
            counts += 1
            start += timedelta(days=1)
        return counts

    def isHoliday(self,day):
        if type(day) == str:
            day = datetime.strptime(day, '%Y%m%d').date()
        if is_holiday(day) or day.weekday() == 5 or day.weekday() == 6:
            return True
        else:
            return False

    def send_Email(self,Email_address, email_text):
        from_addr = '240713753@qq.com'  # 发出电子邮件的地址
        password = 'asrtsrdxvodrbiib'  # 发出电子邮件的密码
        title = '价格异动监控消息-' + dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')  # 电子邮件的标题
        msg = MIMEText(email_text, 'html', 'utf-8')  # 电子邮件的格式是HTML
        msg['From'] = from_addr
        msg['To'] = Email_address
        msg['Subject'] = title

        try:
            server = smtplib.SMTP_SSL('smtp.qq.com', 465)
            server.login(from_addr, password)  # 发送邮件
            server.send_message(msg)
            server.quit()

            # print(Email_address+'  send success!')
            # send_info.append(Email_address + '  send success!\n')
        except Exception as e:
             print(e)
            # send_info.append(e + '\n')
            # send_info.append(Email_address + ' send failed!\n')
            # print(Email_address+' send failed!')



if __name__ == "__main__":
    #rt=Util().workdays('2023-03-15','2023-03-31')
    rt=Util().isHoliday('20230401')
    print(rt)