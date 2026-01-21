from flask import Flask, render_template, request, redirect, url_for, flash, send_file, jsonify
from flask_wtf import FlaskForm
from wtforms import StringField, DateField, TextAreaField, SubmitField, PasswordField
from wtforms.validators import DataRequired, Length, Regexp
import pandas as pd
import json
import os
import uuid
from datetime import datetime, timedelta
import re
from flask_login import LoginManager, UserMixin, login_user, logout_user, login_required, current_user

app = Flask(__name__)
app.config['SECRET_KEY'] = 'your-secret-key-here'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# 初始化Flask-Login
login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login'
login_manager.login_message = '请先登录以访问该页面'

# 用户模型
class User(UserMixin):
    def __init__(self, id, phone, name, password='123456', role='user'):
        self.id = id
        self.phone = phone
        self.name = name
        self.password = password
        self.role = role  # 'user' 或 'admin'

# 用户数据文件
USERS_FILE = 'data/users.json'

# 确保用户数据文件存在
if not os.path.exists(USERS_FILE):
    # 创建默认管理员用户
    default_users = {
        '1': {
            'id': '1',
            'phone': '13800138000',
            'name': '管理员',
            'password': '123456',
            'role': 'admin'
        }
    }
    with open(USERS_FILE, 'w') as f:
        json.dump(default_users, f, ensure_ascii=False, indent=2)

# 加载用户数据
def load_users():
    with open(USERS_FILE, 'r') as f:
        return json.load(f)

# 保存用户数据
def save_users(users):
    with open(USERS_FILE, 'w') as f:
        json.dump(users, f, ensure_ascii=False, indent=2)

# 根据用户ID加载用户
@login_manager.user_loader
def load_user(user_id):
    users = load_users()
    if user_id in users:
        user_data = users[user_id]
        return User(user_data['id'], user_data['phone'], user_data['name'], user_data.get('password', '123456'), user_data['role'])
    return None

# 登录表单
class LoginForm(FlaskForm):
    phone = StringField('手机号', validators=[
        DataRequired('手机号不能为空'),
        Regexp(r'^1[3-9]\d{9}$', message='请输入有效的手机号')
    ])
    password = PasswordField('密码', validators=[
        DataRequired('密码不能为空'),
        Length(min=6, message='密码长度至少为6位')
    ])
    submit = SubmitField('登录')

# 用户信息修改表单
class UserInfoForm(FlaskForm):
    name = StringField('姓名', validators=[DataRequired('姓名不能为空')])
    password = PasswordField('密码', validators=[
        DataRequired('密码不能为空'),
        Length(min=6, message='密码长度至少为6位')
    ])
    submit = SubmitField('保存修改')

# 表单模型
class WorkSummaryForm(FlaskForm):
    name = StringField('姓名', validators=[DataRequired()])
    department = StringField('部门', validators=[DataRequired()])
    start_date = DateField('本周开始日期', format='%Y-%m-%d', validators=[DataRequired()])
    end_date = DateField('本周结束日期', format='%Y-%m-%d', validators=[DataRequired()])
    core_work = TextAreaField('本周核心工作内容', validators=[DataRequired()])
    completion = TextAreaField('完成情况（可量化优先）', validators=[DataRequired()])
    problems = TextAreaField('遇到的问题')
    next_week_plan = TextAreaField('下周工作计划', validators=[DataRequired()])
    submit = SubmitField('提交')

# 数据存储路径
DATA_FILE = 'data/summaries.json'

# 确保数据文件存在
if not os.path.exists(DATA_FILE):
    with open(DATA_FILE, 'w') as f:
        json.dump({}, f)

# 加载数据
def load_data():
    with open(DATA_FILE, 'r') as f:
        return json.load(f)

# 保存数据
def save_data(data):
    with open(DATA_FILE, 'w') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

# 规范化完成情况描述
def normalize_completion(text):
    if not text:
        return text
    
    # 规范化模糊表述
    text = text.replace('完成一部分', '推进中，完成度50%')
    text = text.replace('差不多完成了', '接近完成，完成度90%')
    text = text.replace('刚起步', '启动阶段，完成度10%')
    text = text.replace('还没开始', '未开始，完成度0%')
    text = text.replace('完成了', '已完成，完成度100%')
    
    # 如果没有明确的完成度，尝试提取或添加默认值
    if '完成度' not in text and '完成' in text:
        if '已完成' in text:
            text += '，完成度100%'
        else:
            text += '，完成度XX%'
    
    return text

# 登录页面
@app.route('/login', methods=['GET', 'POST'])
def login():
    if current_user.is_authenticated:
        return redirect(url_for('dashboard'))
    
    form = LoginForm()
    if form.validate_on_submit():
        phone = form.phone.data
        password = form.password.data
        users = load_users()
        
        # 查找用户
        for user_id, user_data in users.items():
            if user_data['phone'] == phone:
                # 验证密码
                if user_data.get('password', '123456') == password:
                    user = User(user_data['id'], user_data['phone'], user_data['name'], user_data.get('password', '123456'), user_data['role'])
                    login_user(user)
                    flash('登录成功！', 'success')
                    return redirect(url_for('dashboard'))
                else:
                    flash('密码错误！', 'error')
                    return redirect(url_for('login'))
        
        # 如果用户不存在，自动创建新用户
        new_user_id = str(len(users) + 1)
        new_user = {
            'id': new_user_id,
            'phone': phone,
            'name': f'用户{new_user_id}',
            'password': password,
            'role': 'user'
        }
        users[new_user_id] = new_user
        save_users(users)
        
        login_user(User(new_user_id, phone, new_user['name'], new_user['password'], 'user'))
        return redirect(url_for('dashboard'))
    
    return render_template('login.html', form=form)

# 登出
@app.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('login'))

# 修改用户信息
@app.route('/user_info', methods=['GET', 'POST'])
@login_required
def user_info():
    form = UserInfoForm()
    users = load_users()
    user_data = users[current_user.id]
    
    if form.validate_on_submit():
        # 验证密码
        if form.password.data != user_data.get('password', '123456'):
            flash('密码错误！', 'error')
            return redirect(url_for('user_info'))
        
        # 更新用户信息
        user_data['name'] = form.name.data
        users[current_user.id] = user_data
        save_users(users)
        
        # 更新当前用户对象
        current_user.name = form.name.data
        
        flash('用户信息修改成功！', 'success')
        return redirect(url_for('dashboard'))
    
    # 预填表单
    form.name.data = current_user.name
    
    return render_template('user_info.html', form=form)

# 仪表盘（根据用户角色显示不同内容）
@app.route('/')
@login_required
def dashboard():
    if current_user.role == 'admin':
        # 管理员仪表盘
        return redirect(url_for('admin_dashboard'))
    else:
        # 普通用户仪表盘
        return redirect(url_for('user_dashboard'))

# 普通用户仪表盘
@app.route('/user_dashboard')
@login_required
def user_dashboard():
    # 获取用户历史提交记录
    data = load_data()
    user_records = [v for v in data.values() if v.get('name') == current_user.name]
    
    # 判断是否为本周
    def is_current_week(start_date_str, end_date_str):
        today = datetime.now()
        current_monday = today - timedelta(days=today.weekday())
        current_friday = current_monday + timedelta(days=4)
        
        # 格式化日期字符串为日期对象
        try:
            start_date = datetime.strptime(start_date_str, '%Y-%m-%d').date()
            end_date = datetime.strptime(end_date_str, '%Y-%m-%d').date()
            return start_date == current_monday.date() and end_date == current_friday.date()
        except:
            return False
    
    # 为每条记录添加是否为本周的标记
    for record in user_records:
        record['is_current_week'] = is_current_week(record.get('start_date'), record.get('end_date'))
    
    # 按提交时间排序
    user_records.sort(key=lambda x: x.get('submission_time', ''), reverse=True)
    
    return render_template('user_dashboard.html', user=current_user, records=user_records)

# 管理员仪表盘
@app.route('/admin_dashboard')
@login_required
def admin_dashboard():
    if current_user.role != 'admin':
        flash('您没有权限访问此页面')
        return redirect(url_for('user_dashboard'))
    
    # 获取所有提交记录
    data = load_data()
    all_records = list(data.values())
    # 按提交时间排序
    all_records.sort(key=lambda x: x.get('submission_time', ''), reverse=True)
    
    return render_template('admin_dashboard.html', user=current_user, records=all_records)

# 生成汇报
@app.route('/generate_report', methods=['GET', 'POST'])
@login_required
def generate_report():
    if current_user.role != 'admin':
        flash('您没有权限使用此功能')
        return redirect(url_for('user_dashboard'))
    
    if request.method == 'POST':
        data_source = request.form.get('data_source')
        
        if data_source == 'file':
            # 从文件上传获取数据
            if 'file' not in request.files:
                flash('请选择一个文件上传')
                return redirect(request.url)
            
            file = request.files['file']
            if file.filename == '':
                flash('请选择一个文件上传')
                return redirect(request.url)
            
            # 保存上传的文件
            filename = f"{uuid.uuid4()}_{file.filename}"
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)
            
            try:
                # 读取文件
                if file.filename.endswith(('.xlsx', '.xls')):
                    df = pd.read_excel(filepath)
                elif file.filename.endswith('.csv'):
                    df = pd.read_csv(filepath)
                else:
                    flash('不支持的文件格式，请上传Excel或CSV文件')
                    return redirect(url_for('admin_dashboard'))
            except Exception as e:
                flash(f'文件读取失败：{str(e)}')
                return redirect(url_for('admin_dashboard'))
            finally:
                # 清理文件
                if os.path.exists(filepath):
                    os.remove(filepath)
        elif data_source == 'database':
            # 从数据库获取数据
            data = load_data()
            if not data:
                flash('没有找到提交数据')
                return redirect(url_for('admin_dashboard'))
            
            # 转换数据，将英文字段名映射到中文
            mapped_data = []
            for record in data.values():
                mapped_record = {
                    '姓名': record.get('name', ''),
                    '本周工作周期': f"{record.get('start_date', '')} - {record.get('end_date', '')}",
                    '本周核心工作内容': record.get('core_work', ''),
                    '完成情况': record.get('completion', ''),
                    '遇到的问题': record.get('problems', ''),
                    '下周工作计划': record.get('next_week_plan', ''),
                    '提交时间': record.get('submission_time', '')
                }
                mapped_data.append(mapped_record)
            
            df = pd.DataFrame(mapped_data)
        else:
            flash('请选择数据来源')
            return redirect(request.url)
        
        try:
            # 处理数据
            df = df.dropna(how='all')  # 忽略空白行
            
            # 标准化表头
            df.columns = [col.strip() for col in df.columns]
            
            # 检查必需字段
            required_columns = ['姓名', '本周工作周期', '本周核心工作内容', '完成情况', '遇到的问题', '下周工作计划']
            for col in required_columns:
                if col not in df.columns:
                    flash(f'数据缺少必需字段：{col}')
                    return redirect(url_for('admin_dashboard'))
            
            # 解析工作周期
            period = df['本周工作周期'].iloc[0] if len(df) > 0 else ''
            if '-' in period:
                start_str, end_str = period.split('-')[:2]
                try:
                    # 尝试不同的日期格式
                    for fmt in ['%Y-%m-%d', '%Y/%m/%d', '%Y%m%d']:
                        try:
                            start_date = datetime.strptime(start_str.strip(), fmt)
                            end_date = datetime.strptime(end_str.strip(), fmt)
                            break
                        except ValueError:
                            continue
                    else:
                        raise ValueError("无法解析日期格式")
                    
                    # 格式化为YYYYMMDD-YYYYMMDD
                    period_formatted = f"{start_date.strftime('%Y%m%d')}-{end_date.strftime('%Y%m%d')}"
                    
                    # 计算第X周（简单计算：以当前日期为准，每周一为一周开始）
                    today = datetime.now()
                    weeks_diff = (today - start_date).days // 7 + 1
                    week_number = weeks_diff
                    
                except ValueError:
                    period_formatted = period
                    week_number = 'X'
            else:
                period_formatted = period
                week_number = 'X'
            
            # 处理重复提交（以最后一次为准）
            df = df.sort_values(by='提交时间') if '提交时间' in df.columns else df
            df = df.drop_duplicates(subset=['姓名'], keep='last')
            
            # 按姓名首字母排序
            df = df.sort_values(by='姓名')
            
            # 生成汇报内容
            report_content = f"# 开源鸿蒙系统研发能力提升第{week_number}周工作总结（{period_formatted}）\n\n"
            
            # 上周工作总结
            report_content += "## 上周工作总结：\n\n"
            for idx, row in df.iterrows():
                name = row['姓名']
                work_content = row['本周核心工作内容']
                completion = normalize_completion(row['完成情况'])
                problems = row['遇到的问题'] if pd.notna(row['遇到的问题']) else ''
                
                summary = f"{work_content}，{completion}"
                if problems:
                    summary += f"。遇到的问题：{problems}"
                
                report_content += f"（{idx+1}）{name}：{summary}。\n"
            
            # 本周工作计划
            report_content += "\n## 本周工作计划：\n\n"
            for idx, row in df.iterrows():
                name = row['姓名']
                plan = row['下周工作计划']
                report_content += f"（{idx+1}）{name}：{plan}。\n"
            
            # 保存汇报内容到文件
            report_filename = f"report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.md"
            report_path = os.path.join('uploads', report_filename)
            with open(report_path, 'w', encoding='utf-8') as f:
                f.write(report_content)
            
            return render_template('result.html', report_content=report_content, report_filename=report_filename)
            
        except Exception as e:
            flash(f'数据处理失败：{str(e)}')
            return redirect(url_for('admin_dashboard'))
    
    # GET请求，显示数据来源选择页面
    return render_template('report_source.html')

# 下载汇报文件
@app.route('/download_report/<filename>')
@login_required
def download_report(filename):
    if current_user.role != 'admin':
        flash('您没有权限下载此文件')
        return redirect(url_for('user_dashboard'))
    
    return send_file(os.path.join('uploads', filename), as_attachment=True)

# 提交表单
@app.route('/submit_form', methods=['POST'])
@login_required
def submit_form():
    form_data = request.form.to_dict()
    
    # 保存到数据文件
    data = load_data()
    
    # 获取编辑ID
    edit_id = form_data.pop('edit_id', None)
    
    if edit_id:
        # 修改现有记录
        if edit_id in data:
            # 检查是否是当前用户的记录
            if data[edit_id].get('name') == current_user.name:
                # 检查是否为本周
                today = datetime.now()
                monday = today - timedelta(days=today.weekday())
                friday = monday + timedelta(days=4)
                
                try:
                    start_date = datetime.strptime(data[edit_id].get('start_date'), '%Y-%m-%d').date()
                    end_date = datetime.strptime(data[edit_id].get('end_date'), '%Y-%m-%d').date()
                    if start_date == monday.date() and end_date == friday.date():
                        # 更新记录
                        form_data['id'] = edit_id
                        form_data['submission_time'] = datetime.now().isoformat()
                        form_data['user_id'] = current_user.id
                        data[edit_id] = form_data
                        save_data(data)
                        return jsonify({'status': 'success', 'message': '修改成功！'})
                except:
                    pass
    
    # 生成唯一ID
    entry_id = str(uuid.uuid4())
    form_data['id'] = entry_id
    form_data['submission_time'] = datetime.now().isoformat()
    form_data['user_id'] = current_user.id
    
    # 更新或添加记录（按用户和周期）
    period_key = f"{form_data.get('start_date')}_{form_data.get('end_date')}"
    existing_keys = [k for k, v in data.items() 
                    if v.get('name') == current_user.name 
                    and f"{v.get('start_date')}_{v.get('end_date')}" == period_key]
    
    for key in existing_keys:
        del data[key]
    
    data[entry_id] = form_data
    save_data(data)
    
    return jsonify({'status': 'success', 'message': '提交成功！'})

# 生成表单页面
@app.route('/form')
@login_required
def form_page():
    form = WorkSummaryForm()
    # 获取编辑ID
    edit_id = request.args.get('edit_id')
    
    # 设置默认姓名为当前用户名
    form.name.data = current_user.name
    # 设置默认日期（本周一到本周五）
    today = datetime.now()
    monday = today - timedelta(days=today.weekday())
    friday = monday + timedelta(days=4)
    form.start_date.data = monday
    form.end_date.data = friday
    
    # 如果是编辑模式，加载现有数据
    if edit_id:
        data = load_data()
        if edit_id in data:
            record = data[edit_id]
            # 检查是否是当前用户的记录
            if record.get('name') == current_user.name:
                # 检查是否为本周
                try:
                    start_date = datetime.strptime(record.get('start_date'), '%Y-%m-%d').date()
                    end_date = datetime.strptime(record.get('end_date'), '%Y-%m-%d').date()
                    if start_date == monday.date() and end_date == friday.date():
                        # 加载现有数据
                        form.department.data = record.get('department', '')
                        form.core_work.data = record.get('core_work', '')
                        form.completion.data = record.get('completion', '')
                        form.problems.data = record.get('problems', '')
                        form.next_week_plan.data = record.get('next_week_plan', '')
                        return render_template('form.html', form=form, edit_id=edit_id)
                except:
                    pass
    
    return render_template('form.html', form=form)

# 导出数据为Excel
@app.route('/export_excel')
@login_required
def export_excel():
    if current_user.role != 'admin':
        flash('您没有权限导出数据')
        return redirect(url_for('user_dashboard'))
    
    data = load_data()
    if not data:
        flash('没有数据可以导出')
        return redirect(url_for('admin_dashboard'))
    
    # 转换为DataFrame
    df = pd.DataFrame(list(data.values()))
    
    # 保存为Excel文件
    export_filename = f"work_summaries_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    export_path = os.path.join('uploads', export_filename)
    df.to_excel(export_path, index=False)
    
    return send_file(export_path, as_attachment=True)

# 导出数据为CSV
@app.route('/export_csv')
@login_required
def export_csv():
    if current_user.role != 'admin':
        flash('您没有权限导出数据')
        return redirect(url_for('user_dashboard'))
    
    data = load_data()
    if not data:
        flash('没有数据可以导出')
        return redirect(url_for('admin_dashboard'))
    
    # 转换为DataFrame
    df = pd.DataFrame(list(data.values()))
    
    # 保存为CSV文件
    export_filename = f"work_summaries_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    export_path = os.path.join('uploads', export_filename)
    df.to_csv(export_path, index=False, encoding='utf-8-sig')
    
    return send_file(export_path, as_attachment=True)

# 生成新表单（管理员功能）
@app.route('/create_form')
@login_required
def create_form():
    if current_user.role != 'admin':
        flash('您没有权限使用此功能')
        return redirect(url_for('user_dashboard'))
    
    return render_template('create_form.html')

# 查看提交统计（管理员功能）
@app.route('/submission_stats')
@login_required
def submission_stats():
    if current_user.role != 'admin':
        flash('您没有权限使用此功能')
        return redirect(url_for('user_dashboard'))
    
    # 获取所有用户和提交记录
    users = load_users()
    data = load_data()
    
    # 获取当前周期（本周）
    today = datetime.now()
    monday = today - timedelta(days=today.weekday())
    friday = monday + timedelta(days=4)
    current_period_start = monday.strftime('%Y-%m-%d')
    current_period_end = friday.strftime('%Y-%m-%d')
    current_period_key = f"{current_period_start}_{current_period_end}"
    
    # 统计已提交和未提交的用户
    submitted_users = set()
    submission_times = {}
    
    for record in data.values():
        record_period_key = f"{record.get('start_date', '')}_{record.get('end_date', '')}"
        if record_period_key == current_period_key:
            submitted_users.add(record.get('name', ''))
            submission_times[record.get('name', '')] = record.get('submission_time', '')
    
    all_users = {user_data['name'] for user_data in users.values() if user_data['role'] == 'user'}
    not_submitted_users = all_users - submitted_users
    
    return render_template('submission_stats.html', 
                          current_period_start=current_period_start,
                          current_period_end=current_period_end,
                          submitted_users=submitted_users,
                          not_submitted_users=not_submitted_users,
                          submission_times=submission_times)

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5001)
