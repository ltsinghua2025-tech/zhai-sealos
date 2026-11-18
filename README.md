# ZHAI Sealos Git 导入启动

本仓库包含构建好的 ZHAI 网页、Python 服务及依赖，供 Sealos 从 Git 导入。运行环境为 Linux Python 3.12，支持 x86_64 和 ARM64。

## 从 Git 导入

1. 在 Sealos DevBox 选择从 Git 仓库导入，地址填 `https://github.com/ltsinghua2025-tech/zhai-sealos.git`，分支选 `main`。
2. 如果表单提供运行环境和容器端口，选 Python 3.12、填 `8080`。如果提供启动命令，填 `bash entrypoint.sh`。仓库根目录已带 Sealos 发布所用的 `entrypoint.sh`。
3. 打开 8080 对应的访问地址即可进入网页。默认免登录，无须配置环境变量。首次启动从随包 wheels 安装依赖，后续启动复用虚拟环境。

Git 导入和启动是两个步骤。如果导入后服务还未启动，使用平台提供的启动或发布入口。Sealos 正式部署会运行 `entrypoint.sh`。旧版本启动命令 `bash start.sh` 也适用。

## 正式部署与数据

免登录实例的会话和成果由访问者共享，适合用合成材料演示。需要联网模型时，在网页「模型服务」填写自己的接口地址、模型 ID 和 API Key，配置仅在当前网页有效。普通文件处理能力可以直接试用。

数据默认保存在仓库根目录下的 `zhai/data`，也可通过 `ZHAI_DATA_DIR` 指定持久卷路径。只挂载数据目录，首次挂载空卷前先备份和迁移已有会话。已有 `ZHAI_CONFIG_FILE` 与 `ZHAI_DATA_DIR` 配置会继续使用。`ZHAI_WEB_PORT` 优先，其次使用平台的 `PORT`，两者未设置时监听 `0.0.0.0:8080`。

自行管理的实例若需密码保护，可选配 `ZHAI_ACCESS_PASSWORD`，用户名默认 `zhai`，可通过 `ZHAI_ACCESS_USER` 自定。配置网页访问密码后可以使用服务端的模型环境变量 `ZHAI_LLM_BASE_URL`、`ZHAI_LLM_MODEL`、`ZHAI_LLM_NAME`、`ZHAI_LLM_API_KEY`。

服务健康检查路径为 `/api/health`。普通任务的 API 请求需要服务器能访问模型接口。当前 Web 版本使用已有文件处理能力，桌面版的 Codex CLI、临时 Python 隔离执行和扫描识字属于各自的桌面发布范围。
