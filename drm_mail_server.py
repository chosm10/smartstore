import os
from flask import Flask, request  # 서버 구현을 위한 Flask 객체 import
from flask_restx import Api, Resource  # Api 구현을 위한 Api 객체 import
from flask_cors import CORS

app = Flask(__name__)  # Flask 객체 선언, 파라미터로 어플리케이션 패키지의 이름을 넣어줌.
CORS(app)
api = Api(app)  # Flask 객체에 Api 객체 등록
# SFTP 서버에서 파일을 받기로 설정된 폴더 경로를 적어줘야함
path = r'C:\Users\Administrator\Desktop\drm_need\rpa'
@api.route('/sendmail')
class SendMail(Resource):
    def post(self):
        to = ["chosm10@kakao.com", "chosm10@hyundai-ite.com", "cindy@hyundaihmall.com", "move@hyundai-ite.com"]
        to.append(request.json.get('to'))
        subject = request.json.get('subject')
        msg = request.json.get('msg')
        files = request.json.get('files')

        attach = []
        for file in files:
            filepath = r'{}\{}'.format(path, file)
            attach.append(filepath)
            if 'csv' in file:
                continue

            try:
                naver.setDRM(filepath)
            except Exception as e:
                print('{} drm 설정 실패: {}'.format(filepath, e))
        
        try:
            # subject = '테스트메일 입니다.'
            mail.sendmail(to, subject, msg, attach)
            print('메일 발송 성공')
        except Exception as e:
            print('메일 발송 실패: {}'.format(e))
if __name__ == "__main__":

    from comm import naver
    from comm import mail
    app.run(debug=True, host='0.0.0.0', port=9000)