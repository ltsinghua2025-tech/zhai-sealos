# ZHAI Sealos Git 导入启动

这三个文件放在 Git 仓库根目录，旁边应有完整的 `zhai/` 运行目录，包括 Web 前端、服务端、文件引擎、`manifest.json`、`verify-package.py` 和两架构 Python wheels。服务器使用 Linux Python 3.12，选择 x86_64 或 ARM64，无需 GPU。

## 从 Git 导入

1. 在 Sealos Devbox 选择从 Git 仓库导入，填入本仓库地址。
2. 选择 Python 3.12，容器端口填 `8080`。
3. 在环境变量中设置自己的 `ZHAI_ACCESS_PASSWORD`。用户名默认为 `zhai`，可用 `ZHAI_ACCESS_USER` 自定。
4. 启动命令填 `bash start.sh`。首次从本地 wheels 安装依赖，后续启动检查版本并复用虚拟环境。服务以前台方式运行，由平台管理进程。
5. 打开 8080 对应的应用地址，输入自己设置的访问账号。进入模型服务页配置自己的 API 地址、模型名称和 Key。

若导入页面只提供启动命令字段，可填下面这一行，并把占位内容换成自己的访问密码。

```bash
ZHAI_ACCESS_PASSWORD='请替换成你自己设置的访问密码' bash start.sh
```

模型也可由部署者通过环境变量设置 `ZHAI_LLM_BASE_URL`、`ZHAI_LLM_MODEL`、`ZHAI_LLM_NAME` 和 `ZHAI_LLM_API_KEY`。公开仓库不保存 API Key 或公共访问密码。

## 正式部署与数据

完成安装并发布版本后，正式应用的启动命令用 `bash entrypoint.sh`，并在正式应用中设置访问密码环境变量。`entrypoint.sh` 只检查并启动，不安装依赖。

数据默认保存在仓库根目录下的 `zhai/data`，也可通过 `ZHAI_DATA_DIR` 指定持久卷路径。只挂载数据目录，首次挂载空卷前先备份和迁移已有会话。已有 `ZHAI_CONFIG_FILE` 与 `ZHAI_DATA_DIR` 配置会继续使用。`ZHAI_WEB_PORT` 优先，其次使用平台的 `PORT`，两者未设置时监听 `0.0.0.0:8080`。

服务健康检查路径为 `/api/health`。普通任务的 API 请求需要服务器能访问模型接口。当前 Web 版本使用已有文件处理能力，桌面版的 Codex CLI、临时 Python 隔离执行和扫描识字属于各自的桌面发布范围。
