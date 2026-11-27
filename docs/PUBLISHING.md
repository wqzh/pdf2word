# 发布到 PyPI / TestPyPI 完整教程

本文以本仓库的 `pdf2word-pro` 为例，从零开始走完「构建 → TestPyPI 演练 → 正式 PyPI 发布 → 安装验证」的全流程。命令以 Windows PowerShell 为主，其他平台差异会单独标注。

> 约定：
> - 发行名（PyPI 上的名字）：`pdf2word-pro`
> - 导入名（Python 里的名字）：`pdf2word_pro`
> - 命令：`pdf2word-pro`
> - 构建产物目录：`dist/`

---

## 目录

1. [前置准备](#1-前置准备)
2. [理解三套凭据](#2-理解三套凭据)
3. [构建发布产物](#3-构建发布产物)
4. [TestPyPI 全流程（先演练）](#4-testpypi-全流程先演练)
5. [正式 PyPI 发布](#5-正式-pypi-发布)
6. [Token 的三种配置方式](#6-token-的三种配置方式)
7. [安装与验证](#7-安装与验证)
8. [版本号与重复上传](#8-版本号与重复上传)
9. [用 GitHub Actions 免密发布（Trusted Publishing）](#9-用-github-actions-免密发布trusted-publishing)
10. [常见报错对照表](#10-常见报错对照表)
11. [发布前检查清单](#11-发布前检查清单)

---

## 1. 前置准备

### 1.1 环境

| 项目 | 要求 | 检查命令 |
| --- | --- | --- |
| Python | ≥ 3.9 | `python -V` |
| pip | 最新 | `python -m pip -V` |
| 构建工具 | `build` | `python -m build --version` |
| 上传工具 | `twine` | `python -m twine --version` |

安装工具（只在本机装一次）：

```powershell
python -m pip install --upgrade build twine
```

> 本机的 Python 运行时**没有** `venv` 模块，所以本文不使用虚拟环境；如果你自己的机器正常，建议先 `python -m venv .venv` 再激活。

### 1.2 账号

需要**两个互相独立的账号**：

| 站点 | 地址 | 用途 |
| --- | --- | --- |
| TestPyPI | https://test.pypi.org/account/register/ | 演练，随便发 |
| PyPI | https://pypi.org/account/register/ | 正式发布 |

**它们不互通**：TestPyPI 的账号、token、项目在正式站都看不到，反之亦然。两个站点都要：

1. 注册；
2. 点邮件里的验证链接（**未验证邮箱无法创建 token**）；
3. 建议开启双因素认证（2FA）：账号设置 → Two-factor authentication。

### 1.3 确认名字可用

上传前先确认名字没被占用（404 = 空闲）：

```powershell
try { Invoke-WebRequest "https://pypi.org/pypi/pdf2word-pro/json" -UseBasicParsing | Out-Null; "TAKEN" }
catch { "FREE" }
```

输出 `FREE` 表示名字空闲；`TAKEN` 表示已被占用（此时请求会返回 404，被 `catch` 捕获）。

> 名字一旦首次上传成功就**永久归属**，不能删除、不能改名。测试站不影响正式站。

---

## 2. 理解三套凭据

发布过程中会出现三种「密码」，别混：

| 凭据 | 用在哪 | 长什么样 | 说明 |
| --- | --- | --- | --- |
| 账号密码 | 登录网页 | 自定义 | 只用于登录网站，**不能**用于 `twine` |
| API token | `twine upload` | `pypi-` 开头的一长串 | 上传凭据，只显示一次 |
| Trusted Publishing | GitHub Actions 里的 OIDC | 无 | CI 里免密，本地用不了 |

**API token 的两个作用域：**

- `Entire account (all projects)`：账号级，能发布到该账号下**任何**项目。首次上传新项目时**必须**用它（项目还不存在，无法选项目级）。
- `Project: <name>`：项目级，只能发布指定项目。项目建好后建议改用这个，最小权限。

**token 只显示一次**，关掉页面就看不到，丢了只能 Revoke 旧的再建新的。

---

## 3. 构建发布产物

### 3.1 清理旧产物

```powershell
Remove-Item -Recurse -Force dist, build -ErrorAction SilentlyContinue
```

### 3.2 构建

```powershell
python -m build
```

成功后会得到两个文件：

```text
dist/
  pdf2word_pro-0.1.0-py3-none-any.whl    # 二进制分发包，pip 优先用它
  pdf2word_pro-0.1.0.tar.gz              # 源码分发包
```

### 3.3 校验元数据与内容

```powershell
python -m twine check dist/*
```

应输出 `Checking ... PASSED`。

再手动确认 sdist 里没有多余内容（不应出现 `data-local/`、`dist/`、`_*.py` 等）：

```powershell
python -c "import tarfile;t=tarfile.open('dist/pdf2word_pro-0.1.0.tar.gz');print('\n'.join(t.getnames()))"
```

应有：`LICENSE`、`README.md`、`README_EN.md`、`pyproject.toml`、`src/pdf2word_pro/*.py`、`PKG-INFO` 等，仅此而已。

### 3.4 检查 pyproject 关键字段

```powershell
python -c "import zipfile;z=zipfile.ZipFile('dist/pdf2word_pro-0.1.0-py3-none-any.whl');print(z.read('pdf2word_pro-0.1.0.dist-info/METADATA').decode()[:900])"
```

重点确认：`Name`、`Version`、`Requires-Python`、`Requires-Dist`（4 个依赖）、`License-Expression: MIT`、`Description-Content-Type: text/markdown`。

---

## 4. TestPyPI 全流程（先演练）

### 4.1 创建 TestPyPI token

1. 登录 https://test.pypi.org/ ；
2. 打开 https://test.pypi.org/manage/account/token/ ；
3. 点 **Add API token**；
4. **Token name**：`local-upload`（随便写，便于日后撤销）；
5. **Scope**：选 **Entire account (all projects)**（首次发布必须）；
6. 点 **Add token**，复制 `pypi-...` 整串 —— 只显示这一次。

### 4.2 上传

```powershell
python -m twine upload --repository testpypi dist/*
```

提示 `Enter your API token:` 时：**直接粘贴 token 回车**。粘贴时屏幕上不会显示任何字符（连 `*` 都没有），这是正常的；用户名不用输，`twine` 自动用 `__token__`。

看到 `View at: https://test.pypi.org/project/pdf2word-pro/0.1.0/` 即成功。

> **那句 WARNING 可以忽略**：`This environment is not supported for trusted publishing` 只是说当前不在 CI 里，无法走 OIDC 免密；本地用 token 上传完全正常。

### 4.3 从 TestPyPI 安装验证

TestPyPI 上没有第三方依赖（pdfplumber 等），必须补正式源：

```powershell
python -m pip install --index-url https://test.pypi.org/simple/ `
  --extra-index-url https://pypi.org/simple/ pdf2word-pro

pdf2word-pro --help
python -c "import pdf2word_pro; print(pdf2word_pro.__version__)"
```

验收通过后再执行第 5 步；验收完可以卸载：

```powershell
python -m pip uninstall -y pdf2word-pro
```

### 4.4 在网页上确认

打开 https://test.pypi.org/project/pdf2word-pro/ ，检查：

- 项目描述是否正常渲染（Markdown 表格、代码块、链接）；
- 版本、依赖、许可证、Python 版本要求是否正确；
- 下载两个文件能正常下载。

---

## 5. 正式 PyPI 发布

流程与 TestPyPI 完全一致，只是站点和 token 换掉：

### 5.1 创建正式站 token

1. 登录 https://pypi.org/ （**不是** test.pypi.org）；
2. 打开 https://pypi.org/manage/account/token/ ；
3. **Add API token** → Scope 选 **Entire account**（首次）；
4. 复制 `pypi-...`。

### 5.2 上传

```powershell
python -m twine upload dist/*
```

不带 `--repository` 时默认发到正式 PyPI。粘贴 token 回车，出现 `View at: https://pypi.org/project/pdf2word-pro/0.1.0/` 即成功。

### 5.3 发布后建议做的事

1. **换成项目级 token**：在 https://pypi.org/manage/project/pdf2word-pro/settings/ 创建 scope 为该项目的 token，撤销账号级 token，降低泄露风险；
2. 在项目页 **Manage → Publishing** 里配置 Trusted Publisher（见第 9 节），之后 CI 发布就不再需要 token；
3. 在项目页 **Manage → Collaborators** 里确认维护者列表。

---

## 6. Token 的三种配置方式

三种任选其一，**按安全程度从高到低**排列。

### 方式 A：交互式粘贴（推荐，最安全）

```powershell
python -m twine upload dist/*
# Enter your API token: <粘贴>
```

token 不落盘、不进历史、不进代码。缺点是每次都要去复制。

### 方式 B：环境变量（仅当前会话）

```powershell
$env:TWINE_USERNAME = "__token__"
$env:TWINE_PASSWORD = "pypi-你的token"
python -m twine upload dist/*
Remove-Item Env:TWINE_PASSWORD   # 用完清掉
```

⚠️ 不要写成 `twine upload -p pypi-xxx`：命令行参数会进入 PowerShell 历史（`Get-History`）和进程列表。

若确实要用命令行参数，至少关掉历史记录：

```powershell
Set-PSReadLineOption -HistorySaveStyle SaveNothing   # 临时
```

### 方式 C：`~/.pypirc`（长期使用）

文件路径：`C:\Users\<用户名>\.pypirc`（**不要**放进仓库！）

```ini
[distutils]
index-servers =
    pypi
    testpypi

[pypi]
username = __token__
password = pypi-你的正式站token

[testpypi]
repository = https://test.pypi.org/legacy/
username = __token__
password = pypi-你的测试站token
```

配置好后直接：

```powershell
python -m twine upload dist/*             # 正式站
python -m twine upload -r testpypi dist/* # 测试站
```

建议限制文件权限（仅当前用户可读）：

```powershell
icacls "$env:USERPROFILE\.pypirc" /inheritance:r /grant:r "$env:USERNAME:(R,W)"
```

### 安全红线

- 不要把 token 贴到聊天、issue、commit、CI 日志里；
- 不要提交 `.pypirc`、`.env`；
- token 泄露：立刻到对应站点 **Revoke** 并重建；
- 换机器/换项目优先用项目级 token 或 Trusted Publishing。

---

## 7. 安装与验证

### 7.1 从正式 PyPI 安装

```powershell
python -m pip install --upgrade pdf2word-pro
pdf2word-pro --help
```

### 7.2 冒烟测试

```powershell
pdf2word-pro input.pdf -o output.docx
python -c "import pdf2word_pro; from pdf2word_pro import Config, pdf_to_docx; print(Config(image_dpi=150))"
```

### 7.3 确认安装来源

```powershell
python -m pip show -f pdf2word-pro
```

应显示 `Location: ...\site-packages` 与 `pdf2word_pro-0.1.0.dist-info`，且 `pdf2word-pro.exe` 在 `Scripts` 目录。

---

## 8. 版本号与重复上传

### 8.1 规则

- 同一个版本号**只能上传一次**，覆盖不了；
- 已发布版本**不能删除**（只能 yank，即隐藏但不删除）；
- 版本号格式遵循 PEP 440：`0.1.0`、`0.1.1`、`1.0.0`、`1.0.0rc1`、`1.0.0.post1`。

### 8.2 改版本号

版本号只有一个来源：`src/pdf2word_pro/_version.py`。

```python
__version__ = "0.1.1"
```

`pyproject.toml` 里是 `dynamic = ["version"]` + `{ attr = "pdf2word_pro._version.__version__" }`，不需要改第二处。

### 8.3 重新构建发布

```powershell
Remove-Item -Recurse -Force dist, build -ErrorAction SilentlyContinue
python -m build
python -m twine check dist/*
python -m twine upload dist/*
```

### 8.4 上传失败后想重传同一版本

- TestPyPI：同样不能覆盖，必须升版本号（`0.1.0.post1` 也算新版本）；
- 如果只是本地构建错了、还没上传成功，删掉 `dist/` 重建即可。

---

## 9. 用 GitHub Actions 免密发布（Trusted Publishing）

仓库已提供 `.github/workflows/publish.yml`，采用 PyPI 官方推荐的 **Trusted Publishing（OIDC）**：不需要保存任何 token，发布时由 GitHub 临时签发凭据。

### 9.1 在 PyPI 上登记发布者

1. 登录 PyPI → **Your projects** → 项目 → **Manage → Publishing**；
2. **Add a new publisher** → 选择 **GitHub**；
3. 填写：
   - Owner：`wqzh`
   - Repository：`pdf2word`
   - Workflow name：`publish.yml`
   - Environment name：`pypi`
4. 保存。

> 首次发布前项目还不存在时，用 **Pending publisher**（Your account → Publishing → Add pending publisher）先登记，发布后自动转为正式项目。

### 9.2 触发发布

```powershell
git tag v0.1.0
git push origin v0.1.0
```

然后在 GitHub 上 **Releases → Draft a new release** 选择该标签并发布，工作流会自动：

1. `python -m build` 构建 sdist + wheel；
2. 上传构建产物为 artifact；
3. 通过 OIDC 发布到 PyPI。

### 9.3 私有仓库 / 本地环境

Trusted Publishing 只支持 GitHub Actions、GitLab CI 等受支持的 CI 平台。本地命令行发布仍然用 token。

---

## 10. 常见报错对照表

| 报错 | 原因 | 处理 |
| --- | --- | --- |
| `Enter your API token:` 后回车报 `403` | token 错误、已撤销，或用了另一站点的 token | 重新复制；确认站点（test.pypi.org vs pypi.org） |
| `403 Forbidden ... username/password not valid` | 把账号密码当成了 token | 用户名固定 `__token__`，密码填 `pypi-...` |
| `400 File already exists` | 同版本重复上传 | 升版本号（见 8.2） |
| `400 Invalid value for requires-python` | 版本约束写法非法 | 用 `>=3.9` 这种 PEP 440 写法 |
| `400 ... description content type` | README 文件缺失或类型不对 | 检查 `readme = { file = "README_EN.md", content-type = "text/markdown" }` |
| `error: Invalid distribution ... invalid metadata` | `pyproject.toml` 字段非法（如 license 写法与 setuptools 版本不匹配） | 用 `python -m twine check dist/*` 定位具体字段 |
| `ModuleNotFoundError: No module named build` | 没装构建工具 | `python -m pip install build` |
| 上传成功但 `pip install` 找不到 | 刚发布，CDN/索引有延迟 | 等 1–2 分钟，或加 `--no-cache-dir` |
| `pip install` 从 TestPyPI 装不上依赖 | TestPyPI 没有第三方包 | 加 `--extra-index-url https://pypi.org/simple/` |
| Git Bash 里 `dist/*` 被展开成多个文件 | shell 通配符行为 | PowerShell 下直接写 `dist/*` 即可；Git Bash 用 `dist/*.whl dist/*.tar.gz` |

---

## 11. 发布前检查清单

- [ ] `src/pdf2word_pro/_version.py` 里的版本号是新的，且没有重复发布过
- [ ] `pyproject.toml` 的 name / description / requires-python / dependencies 正确
- [ ] `LICENSE` 存在，`license = "MIT"` 与文件一致
- [ ] `README_EN.md` 作为长描述，链接用绝对地址（PyPI 不解析相对路径）
- [ ] `dist/` 是干净重建的，`twine check` 全 PASSED
- [ ] sdist 内容不含 `data-local/`、`dist/`、`_*.py` 等中间文件
- [ ] TestPyPI 上传并安装验证通过
- [ ] 正式 token 使用项目级（或改用 Trusted Publishing）
- [ ] 发布后在 https://pypi.org/project/pdf2word-pro/ 目检描述渲染
- [ ] 打 tag 归档本次发布
