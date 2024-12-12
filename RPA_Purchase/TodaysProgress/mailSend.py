import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.utils import formataddr
import os
import util as ut

config = ut.load_config()

def SMTP_send(output_file_name):
    # 이메일 정보
    sender_name = config["EMAIL"]["Sender_NAME"]
    sender_email = config["EMAIL"]["Sender_address"]  # 발신자 이메일 주소
    password = config["EMAIL"]["PW"]  # 발신자 이메일 비밀번호
    # 받는 사람 목록
    recipients = [
        {"name":config["EMAIL_RECEIVER"]["NAME1"], "email":config["EMAIL_RECEIVER"]["address1"]},
        {"name":config["EMAIL_RECEIVER"]["NAME2"], "email":config["EMAIL_RECEIVER"]["address2"]},
        {"name":config["EMAIL_RECEIVER"]["NAME3"], "email":config["EMAIL_RECEIVER"]["address3"]}
    ]

    # 엑셀 파일 경로
    excel_file = os.path.join(os.getcwd(), 'result', output_file_name)

    # 이메일 메시지 기본 구성
    base_msg = MIMEMultipart()
    base_msg['From'] = formataddr((sender_name, sender_email))
    base_msg['Subject'] = '일일 구매진행현황 보고드립니다.'

    # 본문 내용
    body = """
    \n 안녕하십니까 선생님,
    \n 오늘의 구매진행현황 알림드립니다.
    \n 즐거운 하루 되세요^^
    """
    base_msg.attach(MIMEText(body, 'plain'))

    # 엑셀 파일 첨부
    with open(excel_file, 'rb') as attachment:
        part = MIMEApplication(attachment.read(), Name="구매진행현황_")
        part['Content-Disposition'] = f'attachment; filename="{os.path.basename(excel_file)}"'
        base_msg.attach(part)

    # SMTP 서버에 연결하고 이메일 전송
    try:
        with smtplib.SMTP('smtp.sejong.ac.kr', 25) as server:  # SMTP 서버와 포트 번호 입력
            server.starttls()  # TLS 보안 연결 시작
            server.login(sender_email, password)  # 로그인
            for recipient in recipients:
                # 각 수신자에 대해 개별 이메일 메시지 생성
                individual_msg = MIMEMultipart()
                individual_msg['From'] = formataddr((sender_name, sender_email))
                individual_msg['To'] = formataddr((recipient["name"], recipient["email"]))
                individual_msg['Subject'] = base_msg['Subject']
                individual_msg.attach(MIMEText(body, 'plain'))  # 본문 내용
                individual_msg.attach(part)  # 첨부파일 추가

                # 이메일 전송
                server.send_message(individual_msg)
                print(f"이메일이 {recipient['name']}에게 성공적으로 전송되었습니다.")
    except Exception as e:
        print(f"이메일 전송 중 오류가 발생했습니다: {e}")