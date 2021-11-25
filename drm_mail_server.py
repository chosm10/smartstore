import os
from flask import Flask  # 서버 구현을 위한 Flask 객체 import
from flask_restx import Api, Resource  # Api 구현을 위한 Api 객체 import
from flask_cors import CORS
from comm import naver
from comm import mail

app = Flask(__name__)  # Flask 객체 선언, 파라미터로 어플리케이션 패키지의 이름을 넣어줌.
CORS(app)
api = Api(app)  # Flask 객체에 Api 객체 등록

path = 'C:\Program Files (x86)\freeFTPd\sftproot\rpa\'
@api.route('/sendmail/<string:params>')
class SendMail(Resource):
    def get(self, params):
        to = ["chosm10@kakao.com", "chosm10@hyundai-ite.com", "cindy@hyundaihmall.com", "move@hyundai-ite.com"]
        #서버에서 받을 때 첫번째는 받는 사람, 두번째는 메일 제목, 세번째는 메일 내용, 네번째부터 파일 이름
        params = params.split(' ')
        to.append(params[0])
        subject = params[1]
        msg = params[2]
        files = params[3:]

        attach = []
        for file in files:
            filepath = r'{}{}'.format(path, file)
            naver.setDRM(filepath)
            attach.append(filepath)

        mail.sendmail(to, subject, msg, attach)
        

if __name__ == "__main__":
    app.run(debug=True, host='0.0.0.0', port=9000)