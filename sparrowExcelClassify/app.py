import os
import sys
import time
from flask import Flask
from flask import request
from flask import jsonify
from flask import url_for
from flask import redirect
from flask_cors import CORS
from sparrowSql import SqLite
from flask import render_template
from flask import send_from_directory
from sparrowExcelClassify.tools import create
from sparrowExcelClassify.tools import open_url
from sparrowExcelClassify.tools import is_login
from sparrowExcelClassify.tools import DealExcel
from sparrowExcelClassify.tools import open_xlsx
from sparrowExcelClassify.tools import allowed_file
from sparrowExcelClassify.tools import compress_folder
from sparrowExcelClassify.tools import return_sheet_name
from sparrowExcelClassify.tools import classify_worksheet
from sparrowExcelClassify.tools import get_split_information
from sparrowEncryptionDecryption import SparrowEncryptionDecryption
from sparrowExcelClassify.tools import get_split_choose_information

app = Flask(__name__)
CORS(app)
UPLOAD_FOLDER = 'static/uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER


class SparrowExcelClassify(object):
    def __init__(self, flask_app):
        self.sed = SparrowEncryptionDecryption()
        create()
        self.app = flask_app
        self.user = {
            'username': None,
            'password': None
        }
        self.sheet_name = None
        self.choose_data = None
        self.sheet_data = None
        self.time = 0
        self.run()
        open_url()

    def run(self):
        @self.app.get('/')
        def index():
            """
            主页面
            :return: 返回主页面模板
            """
            if is_login(self.user):
                return render_template("index.html")
            else:
                return render_template("login.html")

        @self.app.post("/")
        def login():
            data = request.form
            username = data['username']
            password = data['password']
            e_username = self.sed.easy_encryption(username, "SparrowExcelClassify", 0)
            e_password = self.sed.easy_encryption(password, "SparrowExcelClassify", 0)
            sqlite = SqLite("SparrowExcelClassify.db")
            db_password = sqlite.select("user", {"username": e_username}, ['password'])
            sqlite.close()
            if len(db_password) == 0:
                return render_template("login.html", status="用户名错误")
            if db_password[0][0] == e_password:
                self.user['username'] = e_username
                self.user['password'] = e_password
                return redirect("http://127.0.0.1:5000")
            else:
                return render_template("login.html", status="密码错误")

        @self.app.post('/upload')
        def upload_file():
            """
            文件上传页面
            :return: 返回上传状态
            """
            if is_login(self.user):
                try:
                    if 'excelFile' not in request.files:
                        return jsonify({'message': '没有文件部分'}), 499
                    file = request.files['excelFile']
                    if file.filename == '':
                        return jsonify({'message': '没有选择文件'}), 498
                    if file and allowed_file(file.filename):
                        filename = os.path.join(self.app.config['UPLOAD_FOLDER'], f"upload.xlsx")
                        file.save(filename)
                        self.sheet_data = open_xlsx()
                        return redirect(url_for('get_deal'))
                    else:
                        return jsonify({'message': '错误：只允许上传Excel文件。'}), 497
                except Exception as e:
                    return render_template("index.html")
            else:
                return render_template("login.html")

        @self.app.get('/deal')
        def get_deal():
            """
            获取表名
            :return: 表名和表现在页面
            """
            if is_login(self.user):
                try:
                    self.sheet_name = return_sheet_name(self.sheet_data)
                    return render_template("deal.html", sheet_name=self.sheet_name)
                except Exception as e:
                    return render_template("index.html")
            else:
                return render_template("login.html")

        @self.app.post("/sheet_deal")
        def sheet_deal():
            """
            表处理页面
            :return: 返回表具体如何处理的页面和需要被处理的表信息
            """
            if is_login(self.user):
                try:
                    self.choose_data = get_split_information(request.form)[0]
                    return render_template("sheet_deal.html", choose_data=self.choose_data)
                except Exception as e:
                    return render_template("index.html")
            else:
                return render_template("login.html")

        @self.app.post('/classify')
        def classify_worksheet_deal():
            """
            分拣表
            :return: 分拣状态
            """
            if is_login(self.user):
                try:
                    choose_information = get_split_choose_information(request.form)
                    data = classify_worksheet(self.choose_data, self.sheet_data)
                    de_run = DealExcel(data, self.choose_data, choose_information)
                    de_run.to_excel()
                    compress_folder()
                    return render_template("download.html")
                except Exception as e:
                    return render_template("index.html")
            else:
                return render_template("login.html")


if __name__ == '__main__':
    run = SparrowExcelClassify(app)
    app.run(debug=True, host='0.0.0.0')
