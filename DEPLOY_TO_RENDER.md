# 问卷转换器 - Render 部署指南

## 🚀 快速部署到 Render

### 前提条件
1. GitHub 账户
2. Render 账户 (免费注册: https://render.com)

### 📋 部署步骤

#### 1. 准备代码仓库
```bash
# 1. 初始化 Git 仓库
git init

# 2. 添加所有文件
git add .

# 3. 提交代码
git commit -m "Initial commit: Survey Converter Web App"

# 4. 创建 GitHub 仓库并推送
# 在 GitHub 创建新仓库，然后：
git remote add origin https://github.com/你的用户名/survey-converter.git
git branch -M main
git push -u origin main
```

#### 2. 在 Render 部署

1. **登录 Render**
   - 访问 https://render.com
   - 使用 GitHub 账户登录

2. **创建新的 Web Service**
   - 点击 "New +" → "Web Service"
   - 连接你的 GitHub 仓库
   - 选择 `survey-converter` 仓库

3. **配置部署设置**
   - **Name**: `survey-converter` (或你喜欢的名称)
   - **Environment**: `Python 3`
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn --bind 0.0.0.0:$PORT web_app:app`
   - **Plan**: `Free` (免费计划)

4. **设置环境变量**
   在 Environment Variables 部分添加：
   ```
   FLASK_ENV=production
   SECRET_KEY=your-super-secret-key-here-change-this
   ```

5. **部署**
   - 点击 "Create Web Service"
   - 等待部署完成（通常需要 5-10 分钟）

### 🔧 配置说明

#### 文件结构
```
survey-converter/
├── web_app.py              # 主应用文件
├── requirements.txt        # Python 依赖
├── render.yaml            # Render 配置文件
├── Procfile               # 备选启动配置
├── .env.example           # 环境变量模板
├── .gitignore             # Git 忽略文件
├── survey_converter.py    # 转换器核心
├── survey_parser.py       # 解析器
├── word_to_json.py        # Word 处理
├── xml_generator.py       # XML 生成器
├── static/                # 静态文件
├── templates/             # HTML 模板
├── uploads/               # 上传目录 (自动创建)
└── outputs/               # 输出目录 (自动创建)
```

#### 环境变量说明
- `FLASK_ENV`: 设置为 `production` 启用生产模式
- `SECRET_KEY`: Flask 应用密钥，请使用强密码
- `PORT`: 端口号 (Render 自动设置)

### 📱 使用说明

部署成功后，你将获得一个类似这样的 URL：
```
https://survey-converter-xxxx.onrender.com
```

#### 功能特性
- ✅ 在线上传 Word 文档 (.doc, .docx)
- ✅ 自动转换为 JSON 和 XML 格式
- ✅ 实时转换进度显示
- ✅ 结果文件下载
- ✅ 响应式 Web 界面

### 🔍 故障排除

#### 常见问题

1. **部署失败**
   - 检查 `requirements.txt` 是否包含所有依赖
   - 确保代码没有语法错误
   - 查看 Render 部署日志

2. **应用无法启动**
   - 检查环境变量设置
   - 确认 `gunicorn` 命令正确
   - 查看应用日志

3. **文件上传失败**
   - 检查文件大小限制 (50MB)
   - 确认文件格式为 .doc 或 .docx
   - 查看浏览器控制台错误

#### 查看日志
在 Render 控制台中：
1. 进入你的 Web Service
2. 点击 "Logs" 标签
3. 查看实时日志输出

### 🔄 更新部署

当你修改代码后：
```bash
# 1. 提交更改
git add .
git commit -m "Update: 描述你的更改"

# 2. 推送到 GitHub
git push origin main

# 3. Render 会自动重新部署
```

### 💰 费用说明

**免费计划限制**:
- ✅ 750 小时/月 (足够个人使用)
- ✅ 自动 HTTPS
- ✅ 自定义域名支持
- ⚠️ 应用会在 15 分钟无活动后休眠
- ⚠️ 冷启动可能需要 30 秒

### 🎯 生产环境优化建议

1. **安全性**
   - 使用强密码作为 `SECRET_KEY`
   - 定期更新依赖包
   - 启用 HTTPS (Render 自动提供)

2. **性能**
   - 考虑升级到付费计划避免冷启动
   - 优化文件处理逻辑
   - 添加文件缓存机制

3. **监控**
   - 设置 Render 通知
   - 监控应用性能
   - 定期检查日志

### 📞 支持

如果遇到问题：
1. 查看 Render 官方文档: https://render.com/docs
2. 检查项目 GitHub Issues
3. 联系开发者

---

🎉 **恭喜！你的问卷转换器现在已经在云端运行了！**