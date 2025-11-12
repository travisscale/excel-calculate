import os
from flask import Flask, render_template, request, redirect, url_for, session, flash
import pandas as pd
from datetime import datetime, timedelta

# 1. 初始化 Flask 应用
app = Flask(__name__)

# --- 核心配置区 ---

# 2. 设置 SECRET_KEY
app.secret_key = 'qazplm714119' # 为这个项目设置一个独立的密钥

# 3. 设置管理员密码
ADMIN_PASSWORD = '333222111' # <-- 在这里为Excel计算工具设置密码

# 4. 登录锁定配置
LOGIN_ATTEMPTS = {}
MAX_ATTEMPTS = 10
LOCKOUT_MINUTES = 30

# 5. 允许上传的文件类型
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# --- 路由和视图函数 ---

# 主页路由 (包含上传逻辑，受登录保护)
@app.route('/', methods=['GET', 'POST'])
def index():
    if not session.get('logged_in'):
        return redirect(url_for('login'))

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
                    return render_template('error.html', error="上传的 Excel 文件必须至少包含 A, B, C 三列。")

                sum_a = pd.to_numeric(df.iloc[:, 0], errors='coerce').sum()
                sum_b = pd.to_numeric(df.iloc[:, 1], errors='coerce').sum()
                sum_c = pd.to_numeric(df.iloc[:, 2], errors='coerce').sum()

                return render_template('result.html', 
                                       sum_a=f"{sum_a:,.2f}", 
                                       sum_b=f"{sum_b:,.2f}", 
                                       sum_c=f"{sum_c:,.2f}")
            except Exception as e:
                return render_template('error.html', error=f"处理文件时出错: {e}。")
        else:
            flash('不允许的文件类型，请上传 .xlsx 或 .xls 文件', 'danger')
            return redirect(request.url)

    return render_template('upload.html')

# 登录路由 (带锁定功能)
@app.route('/login', methods=['GET', 'POST'])
def login():
    ip_address = request.headers.get('X-Forwarded-For', request.remote_addr)

    # 检查IP是否被锁定
    if ip_address in LOGIN_ATTEMPTS and LOGIN_ATTEMPTS[ip_address]['failures'] >= MAX_ATTEMPTS:
        last_attempt_time = LOGIN_ATTEMPTS[ip_address]['last_attempt_time']
        lockout_duration = timedelta(minutes=LOCKOUT_MINUTES)
        
        if datetime.now() < last_attempt_time + lockout_duration:
            time_remaining = (last_attempt_time + lockout_duration) - datetime.now()
            minutes_remaining = (time_remaining.seconds // 60) + 1
            flash(f'登录已被锁定，请在 {minutes_remaining} 分钟后重试。', 'danger')
            return render_template('login.html')
        else:
            del LOGIN_ATTEMPTS[ip_address]

    if request.method == 'POST':
        password_attempt = request.form['password']
        if password_attempt == ADMIN_PASSWORD:
            if ip_address in LOGIN_ATTEMPTS:
                del LOGIN_ATTEMPTS[ip_address]
            session['logged_in'] = True
            return redirect(url_for('index'))
        else:
            attempt_info = LOGIN_ATTEMPTS.get(ip_address, {'failures': 0})
            attempt_info['failures'] += 1
            attempt_info['last_attempt_time'] = datetime.now()
            LOGIN_ATTEMPTS[ip_address] = attempt_info

            if attempt_info['failures'] >= MAX_ATTEMPTS:
                flash(f'尝试次数过多。您的IP已被锁定{LOCKOUT_MINUTES}分钟。', 'danger')
            else:
                remaining = MAX_ATTEMPTS - attempt_info['failures']
                flash(f'密码错误！您还有 {remaining} 次尝试机会。', 'danger')
            return render_template('login.html')
            
    return render_template('login.html')

# 登出路由
@app.route('/logout')
def logout():
    session.pop('logged_in', None)
    flash('您已成功登出。', 'info')
    return redirect(url_for('login'))

# --- 程序入口 ---
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True) # 使用默认端口 5000```
