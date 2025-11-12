import os
from flask import Flask, render_template, request, redirect, url_for, session, flash
import pandas as pd

# 1. 初始化 Flask 应用
app = Flask(__name__)

# --- 核心配置区 ---

# 2. 设置 SECRET_KEY (修复 session 错误的关键！)
#    请务-必把它换成一个没人知道的、复杂的字符串！
app.secret_key = 'qazplm714119'

# 3. 设置你的管理员密码
ADMIN_PASSWORD = '18580508131' # <-- 在这里修改成你的真实密码！

# 4. 允许上传的文件类型
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

def allowed_file(filename):
    """检查文件名后缀是否在允许的范围内"""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# --- 路由和视图函数 ---

# 主页路由 (现在受登录保护)
@app.route('/', methods=['GET', 'POST'])
def index():
    # 检查 session 中是否已有 'logged_in' 标记
    if not session.get('logged_in'):
        # 如果没有，重定向到登录页面
        return redirect(url_for('login'))

    # 如果是 POST 请求，意味着登录用户提交了文件
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('请求中没有文件部分', 'warning')
            return redirect(request.url)
        
        file = request.files['file']

        if file.filename == '':
            flash('没有选择文件', 'warning')
            return redirect(request.url)

        if file and allowed_file(file.filename):
            try:
                df = pd.read_excel(file)

                if df.shape[1] < 3:
                    error_message = "上传的 Excel 文件必须至少包含 A, B, C 三列。"
                    return render_template('error.html', error=error_message)

                sum_a = pd.to_numeric(df.iloc[:, 0], errors='coerce').sum()
                sum_b = pd.to_numeric(df.iloc[:, 1], errors='coerce').sum()
                sum_c = pd.to_numeric(df.iloc[:, 2], errors='coerce').sum()

                return render_template('result.html', 
                                       sum_a=f"{sum_a:,.2f}", 
                                       sum_b=f"{sum_b:,.2f}", 
                                       sum_c=f"{sum_c:,.2f}")
            except Exception as e:
                error_message = f"处理文件时出错: {e}。请确保上传的是有效的 Excel 文件。"
                return render_template('error.html', error=error_message)
        else:
            flash('不允许的文件类型，请上传 .xlsx 或 .xls 文件', 'danger')
            return redirect(request.url)

    # 如果是 GET 请求，显示上传表单
    return render_template('upload.html')

# 登录路由
@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        password_attempt = request.form['password']
        if password_attempt == ADMIN_PASSWORD:
            session['logged_in'] = True
            return redirect(url_for('index'))
        else:
            flash('密码错误，请重试！', 'danger')
            
    return render_template('login.html')

# 登出路由
@app.route('/logout')
def logout():
    session.pop('logged_in', None)
    flash('您已成功登出。', 'info')
    return redirect(url_for('login'))

# --- 程序入口 ---
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)