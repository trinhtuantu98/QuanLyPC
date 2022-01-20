from email import encoders
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from pandas.core.frame import DataFrame
import smtplib
from datetime import datetime, timedelta
import os
###--------- Send Email ----------#########
ngay_bc = datetime.today() - timedelta(days=1)
ngay_bc_file_str = ngay_bc.strftime('%Y.%m.%d')
ngay_bc_str = ngay_bc.strftime('%d/%m/%Y')


user = 'nltt@nldc.evn.vn'         # Email userID
password = 'Abc@1234567'      # Email password
from_addr = 'nltt@nldc.evn.vn' ### = username

# recipients_addr = ['hangnm.nldc@gmail.com']
# cc_addr = ['hangnm.nldc@gmail.com']
# recipients_addr = ['ninhnd@nldc.evn.vn', 'khuvx@nldc.evn.vn', 'trungnq@nldc.evn.vn', 'dungpt@nldc.evn.vn', 'tuanna@nldc.evn.vn', 'truongdx@nldc.evn.vn', 'cuongnt@nldc.evn.vn', 'anhkt@nldc.evn.vn', 'ducdx@nldc.evn.vn', 'longvm@nldc.evn.vn', 'quynhp@nldc.evn.vn', 'quangnm@nldc.evn.vn', 'tuanha@nldc.evn.vn', 'dieudo@nldc.evn.vn']
# cc_addr = ['tungnm@nldc.evn.vn','hoainb@nldc.evn.vn','linhbd@nldc.evn.vn','chiennd@nldc.evn.vn','binhntt@nldc.evn.vn', 'hungtv@nldc.evn.vn','thonglv@nldc.evn.vn','Tuna01@nldc.evn.vn','tungns@nldc.evn.vn', 'tutt@nldc.evn.vn','hangnm@nldc.evn.vn','phongnk@nldc.evn.vn', 'tuonghm@nldc.evn.vn']
# recipients_addr = ['hungtv@nldc.evn.vn','tutt@nldc.evn.vn']
dict_PC = {
    'NPC' : ['bankythuatnpc@gmail.com','quangnm@evn.com.vn'],
    'SPC' : ['thanhvt.a2@nldc.evn.vn','quangnm@evn.com.vn'],
    'CPC' : ["danhnh.a3@nldc.evn.vn" ,"thotq1@cpc.vn",'quangnm@evn.com.vn']
}


# cc_addr = ['hangnt@nldc.evn.vn']
# receivers = recipients_addr + cc_addr
# receivers = recipients_addr 




def add_file(path_file,msg): 

    with open(path_file, 'rb') as f:  ## path của ảnh
        # set attachment mime and file name, the image type is png
        mime0 = MIMEBase('application','vnd.ms-excel', filename=os.path.basename(path_file))
        # add required header data:
        mime0.add_header('Content-Disposition', 'attachment', filename=os.path.basename(path_file))
        mime0.add_header('X-Attachment-Id', '0')
        mime0.add_header('Content-ID', '<BCSXNgay_excel>')  ## Đặt tên cho ảnh để sau đó gắn ảnh mang tên đó vào thân html 
        # read attachment file content into the MIMEBase object
        mime0.set_payload((f).read())
        # encode with base64
        encoders.encode_base64(mime0)
        # add MIMEBase object to MIMEMultipart object
        msg.attach(mime0)

def Send_email(path_file,pc):
    ngaylc = datetime.today()
    recipients_addr = dict_PC[pc]
    cc_addr = ['hungtv@nldc.evn.vn','tutt@nldc.evn.vn']
    receivers = recipients_addr + cc_addr

    msg = MIMEMultipart('mixed')   ################# Dạng mixed là để bao gồm nhiều dạng văn bản khác nhau (text, html) ##################
    msg['From'] = from_addr                      
    msg['To'] = ','.join(recipients_addr)
    server ='smtp.office365.com'         # server name
    subject = f"Dữ liệu BCSX smov.vn của các PC ngày {(ngaylc - timedelta(days = 1)).strftime('%d/%m/%Y')}"
    msg['Subject'] = subject
    msg_html = MIMEMultipart('alternative') ###### Dạng alternative để định dạng html #####
    add_file(path_file,msg)
    
    with open('P:/3. Bao cao va Hau kiem/Báo cáo/Bao cao NLTT ngay/EVNNLDC.jpeg', 'rb') as f:  ## path của ảnh
        # set attachment mime and file name, the image type is png
        mime1 = MIMEBase('image', 'png', filename='EVNNLDC.jpeg')
        # add required header data:
        mime1.add_header('Content-Disposition', 'attachment', filename='EVNNLDC.jpeg')
        mime1.add_header('X-Attachment-Id', '0')
        mime1.add_header('Content-ID', '<EVNNLDC>')  ## Đặt tên cho ảnh để sau đó gắn ảnh mang tên đó vào thân html 
        # read attachment file content into the MIMEBase object
        mime1.set_payload((f).read())
        # encode with base64
        encoders.encode_base64(mime1)
        # add MIMEBase object to MIMEMultipart object
        msg.attach(mime1)
    body_html = f"""
    <html>
        <body>            
            <p style="color:black">Kính gửi anh,</p>
            <p style="color:black">Phòng Năng lượng tái tạo - A0 gửi anh số liệu BCSX trên smov.vn ngày {(ngaylc - timedelta(days = 1)).strftime('%d/%m/%Y')} của các PC như file đính kèm.</i></p>

            <p style="color:black">Trân trọng.<br>

            <p style="color:#FF8B33;">Renewable Energy Management Department</p>

            <img src="cid:EVNNLDC" alt="pic0" width="130" height="40"></img>

            <p style="color:#16365c;"><u><b>National Load Dispatch Centre</b></u><br>
            <i><small>Add:     Room 9.03-9.05 – Floor 9, No 11 Cua Bac Street, Ba Dinh Dist., Ha Noi, Vietnam<br>
            Phone: +8424 37173271<br>
            E-Mail: nltt@nldc.evn.vn<br>
            Website:    http://www.nldc.evn.vn</small></i></p>
        </body>
    </html>"""

    msg_html.attach(MIMEText(body_html, 'html', 'utf-8'))  ####### attach thân mail dạng html
    
    msg.attach(msg_html)    #################### attach thân mail vào mail

    if server == 'localhost':   # send mail from local server
        # Start local SMTP server
        server = smtplib.SMTP(server)
        text = msg.as_string()
        server.send_message(msg)
    else:
        # Start SMTP server at port 587
        server = smtplib.SMTP(server, 587)
        server.starttls()
        # Enter login credentials for the email you want to sent mail from
        server.login(user, password)
        text = msg.as_string()
        # Send mail
        server.sendmail(from_addr, receivers, text)

    server.quit()

