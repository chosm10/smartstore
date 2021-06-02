import os
from flask import Flask  # 서버 구현을 위한 Flask 객체 import
from flask_restx import Api, Resource  # Api 구현을 위한 Api 객체 import
from flask_cors import CORS

app = Flask(__name__)  # Flask 객체 선언, 파라미터로 어플리케이션 패키지의 이름을 넣어줌.
CORS(app)
api = Api(app)  # Flask 객체에 Api 객체 등록
task = {
    "일매출정리":"naver_daily_selling",
    "구매확정":"naver_purchase_selling",
    "정산":"naver_settle_selling",
    "외상매출금":"naver_carryover",
    "고객부담배송비":"naver_client_bill"
}
shops = {
        "김포(아)":"gimpo",
        "동대문(아)":"ddp",
        "송도(아)":"songdo",
        "가든(아)":"garden",
        "대구(아)":"daeguOut",
        "가산(아)":"gasan",
        "디큐브(백)":"decube",
        "미아(백)":"mia",
        "본점(백)":"head",
        "무역(백)":"muyeog",
        "천호(백)":"chunho",
        "신촌(백)":"sinchon",
        "목동(백)":"mdong",
        "중동(백)":"jdong",
        "킨텍스(백)":"kintex",
        "부산(백)":"busan",
        "울산(백)":"ulsan",
        "동구(백)":"donggu",
        "대구(백)":"daegu",
        "충청(백)":"cchung",
        "판교(백)":"pangyo"
}

@api.route('/runtask/<string:task_name>')
class RunTask(Resource):
    def get(self, task_name):
        # 네이버_일매출정리_동대문(아) -> ['네이버', '일매출정리', '동대문(아)']
        task_arr = task_name.split('_')
        os.system("C:\\Users\\Administrator\\rpa_naver_brandmall\\batch\\{}\\{}_{}.bat".format(task[task_arr[1]], task[task_arr[1]], shops[task_arr[2]]))

if __name__ == "__main__":
    app.run(debug=True, host='0.0.0.0', port=9000)