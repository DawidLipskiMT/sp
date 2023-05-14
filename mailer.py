import smtplib
from collections import namedtuple
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.header import Header
from email.utils import formataddr
import pandas as pd

fromaddr_hard = 'natalia.rzezniczak@eu.sirowa.com'
#password_hard = 'lclfhxopqzbwmtwj'
password_hard = 'Vuv82503'
server = 'smtp-mail.outlook.com'
port = '587'


def send_mail_only_body(fromaddr, toaddr, subject, body, password, server, port):
    msg = MIMEMultipart()
    msg['From'] = formataddr(('Natalia Lipska (SIR PL)', 'natalia.lipska@sirowa.com'))
    msg['To'] = toaddr
    msg['Subject'] = subject

    body = body
    msg.attach(MIMEText(body, 'html'))

    server = smtplib.SMTP(server, port)
    server.starttls()
    server.login(fromaddr, password)
    text = msg.as_string()
    server.sendmail(fromaddr, toaddr, text)
    server.quit()


TestInfo = namedtuple('TestInfo', ['mail', 'text'])


class AfterTestMailer:
    def __init__(self, gifts, test_df, title):
        self.gifts = self.prepare_gifts(gifts)
        self.test_df = test_df
        self.title = title
        self.info_list = self.create_info_for_molls()
        self.send_mail()

    def create_info_for_molls(self):
        return list(self.test_df.apply(lambda x: self.create_test_info(x), axis=1))

    def send_mail(self):
        if self.title is None:
            raise AttributeError('Nie podałeś tytułu Dejwie')
        for info in self.info_list:
            if info is None:
                continue
            send_mail_only_body(fromaddr_hard, info.mail, self.title,
                                info.text, password_hard, server, port)

    def create_test_info(self, x):
        m = f'mag{self.create_mail_number(x.iloc[0])}@sephora.pl'
        if len(x.iloc[1:].dropna()) == 0:
            return None
        print(f'Sending mail to {m}')
        return TestInfo(m, self.create_text_for_employees(x.iloc[1:].dropna()))

    @staticmethod
    def create_mail_number(param):
        if len(str(param)) == 4:
            return param
        elif len(str(param)) == 3:
            return f'0{param}'
        else:
            return f'00{param}'

    def create_text_for_employees(self, param):
        param[0] = f'&nbsp; - {param[0]}'
        names = '<br>    &nbsp; - '.join(param)

        return f"""
            <html>
              <head></head>
              <body>
                <p>
                    Dzień dobry, <br>
                    <br>
                    W związku z uczestnictwem personelu Waszej perfumerii w lutowym szkoleniu z nowości zapachowych WIOSNA’22, przesłaliśmy poprzez naszą agencję INCENTIVA do perfumerii giftpack dla osób: <br>
                   {names} <br>
                    <br>
                    W paczce znajduje się: <br>
                   {self.gifts} <br>
                    <br>
                    Dziękuję za udział personelu perfumerii w szkoleniu i za ich uważność - gratuluję 
                    zaangażowanego personelu 😊. <br>
                    <br>
                    Pozdrawiam <br>
                    Natalia Lipska
                </p>
              </body>
            </html>
            """

    @staticmethod
    def prepare_gifts(gifts):
        gifts[0] = f'&nbsp; - {gifts[0]}'
        gifts = f'<br>    &nbsp; - '.join(gifts)
        return gifts


if __name__ == '__main__':
    df = pd.read_excel('D://Nauka//Projects//SephoraParser//web_2724_sir_wiosna_gifts.xls', header=0)
    gifty = ['BS THE SCENT SG 50ML PWP',
             'BS T/SCEN F/HER BL 50 PWP',
             'GU GUILT PH PARF EP5 MINI',
             'GU GUIL PF INTEN EP5 MINI',
             'MJ PERFECT INT BC 30 GWP']

    mailer = AfterTestMailer(gifty, df, title='Wysyłka giftów za szkolenie lutowe Sirowa')
