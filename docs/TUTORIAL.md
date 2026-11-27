# 1) 先发 TestPyPI 演练
python -m pip install --upgrade build twine
python -m twine upload -r testpypi dist/*
python -m pip install -i https://test.pypi.org/simple/ pdf2word-pro

# 2) 正式发布（token 只在你自己的终端里设置，不要贴到任何对话/命令历史）
set TWINE_USERNAME=**your_pypi_name>
set TWINE_PASSWORD=*** PyPI API Token>
python -m twine upload dist/*

# 3) 验证
pip install pdf2word-pro
pdf2word-pro --help


==========

## 获取步骤

1. 打开 https://test.pypi.org/account/register/ ，用邮箱注册（用户名、密码随便设，和正式站无关）。
2. 去邮箱点验证链接，**未验证邮箱无法创建 token**。
3. 登录后打开 https://test.pypi.org/manage/account/token/ ，点 **Add API token**。
4. 填写：
   - **Token name**：随便写，例如 `local-upload`；
   - **Scope**：首次上传必须选 **Entire account (all projects)** —— 因为 `pdf2word-pro` 这个项目还不存在，没有项目级 token 可选。等项目建好后，建议改成 Project: `pdf2word-pro` 的最小权限 token。
5. 点 **Add token**，页面会显示一次完整 token（`pypi-` 开头）。**立刻复制保存**，关掉页面就再也看不到，只能重新生成。

==========

## test pypi
pip install --index-url https://test.pypi.org/simple/ --extra-index-url https://pypi.org/simple/ pdf2word-pro

## pypi
pip install pdf2word-pro==0.1.0
