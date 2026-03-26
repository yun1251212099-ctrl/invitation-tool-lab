# 长期外网访问（本机 Streamlit + Cloudflare Tunnel）

## 你需要两条长期运行的进程

1. **Streamlit**（本机 `8599`）
2. **cloudflared tunnel**（把公网域名指到 `http://127.0.0.1:8599`）

## 0. Python 依赖（首次）

```bash
cd /tmp/invitation-tool-lab
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## 1. 启动 Streamlit（推荐）

```bash
cd /tmp/invitation-tool-lab
chmod +x scripts/start_streamlit_8599.sh
./scripts/start_streamlit_8599.sh
```

说明：脚本使用项目内 `.streamlit_home/` 存放 credentials，避免 `~/.streamlit` 权限或首次邮箱交互导致进程退出。

## 2. 一次性：创建持久隧道（Cloudflare）

在已安装 `cloudflared` 的 Mac 上执行：

```bash
cloudflared tunnel login
cloudflared tunnel create invitation-streamlit
cloudflared tunnel route dns invitation-streamlit invite.你的域名.com
```

记下输出的 **Tunnel UUID**，并确认 `~/.cloudflared/<UUID>.json` 已生成。

将 [`cloudflared/config.example.yml`](../cloudflared/config.example.yml) 复制为 `~/.cloudflared/config.yml`，把 `tunnel`、`credentials-file`、`hostname` 改成你的值。

## 3. 启动公网隧道

```bash
cd /tmp/invitation-tool-lab
chmod +x scripts/start_public_tunnel.sh
./scripts/start_public_tunnel.sh
```

或指定配置路径：

```bash
export CLOUDFLARED_CONFIG="$HOME/.cloudflared/config.yml"
./scripts/start_public_tunnel.sh
```

## 4. 分享前自检（必须先通过再给同事链接）

```bash
export PUBLIC_URL="https://invite.你的域名.com"   # 可选：测公网
./scripts/healthcheck_public_url.sh
```

仅测本机：

```bash
./scripts/healthcheck_public_url.sh
```

## 5. 备选：Streamlit Cloud（免隧道）

若你已有 `*.streamlit.app` 部署，见仓库 [`DEPLOY_FLOW.md`](../DEPLOY_FLOW.md)。该方式由平台托管，通常比本机隧道更省事。

## 6. macOS 长期驻留（可选）

可用 `launchd` 把上述两个脚本注册为开机启动；需要时可在「启动台 - 登录项」或自建 plist，此处不强制。
