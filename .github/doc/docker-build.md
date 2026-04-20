# 项目构建并打包到 Docker

## 结论

- 根目录 `docker-compose.yaml` 当前是直接运行官方镜像 `ghcr.io/requarks/wiki:2`，不是把当前源码打包成镜像。
- 如果要从当前仓库源码构建正式镜像，应使用 `dev/build/Dockerfile`。
- `dev/containers/Dockerfile` 仅用于开发容器，不适合生产打包。

## 直接运行官方镜像

如果目标只是把服务快速跑起来，直接在仓库根目录执行：

```powershell
docker compose up -d
```

这会使用根目录 `docker-compose.yaml` 中的官方镜像配置。

查看状态和日志：

```powershell
docker compose ps
docker compose logs -f wiki
```

## 从当前源码构建镜像

正式镜像入口是 `dev/build/Dockerfile`。

这个 Dockerfile 会在构建阶段执行：

```powershell
yarn build
```

对应的脚本定义在 `package.json`：

```json
"build": "cross-env NODE_OPTIONS=--openssl-legacy-provider webpack --profile --config dev/webpack/webpack.prod.js"
```

### 本地构建镜像

在仓库根目录执行：

```powershell
docker build -f dev/build/Dockerfile -t wiki-local:2.0.0 .
```

### 推送到镜像仓库

```powershell
docker tag wiki-local:2.0.0 your-registry/wiki:2.0.0
docker push your-registry/wiki:2.0.0
```

### 多架构构建

如果要一次构建 `amd64` 和 `arm64`：

```powershell
docker buildx build --platform linux/amd64,linux/arm64 -f dev/build/Dockerfile -t your-registry/wiki:2.0.0 --push .
```

## 用 Compose 运行你自己构建的镜像

如果想继续使用根目录 `docker-compose.yaml`，但改成运行你自己构建的镜像，可以把 `wiki` 服务改成这样：

```yaml
wiki:
  image: wiki-local:2.0.0
  build:
    context: .
    dockerfile: dev/build/Dockerfile
  depends_on:
    - db
  environment:
    DB_TYPE: postgres
    DB_HOST: db
    DB_PORT: 5432
    DB_USER: wikijs
    DB_PASS: wikijsrocks
    DB_NAME: wiki
  restart: unless-stopped
  ports:
    - "80:3000"
  volumes:
    - ./wiki-storage:/wiki/data/content
```

然后执行：

```powershell
docker compose up -d --build
```

## 持久化目录差异

这里有一个容易踩坑的点：

- 根目录 `docker-compose.yaml` 当前把数据挂载到 `/wiki/storage`。
- 但 `dev/build/Dockerfile` 中声明的持久化目录是 `/wiki/data/content`。

因此：

- 如果继续使用官方镜像，沿用当前 `docker-compose.yaml` 即可。
- 如果改为使用当前仓库源码构建的镜像，建议把卷挂载改成 `./wiki-storage:/wiki/data/content`。

## 构建前注意事项

当前仓库中存在 `db-data/` 和 `wiki-storage/` 目录，同时仓库里没有 `.dockerignore`。

这意味着执行：

```powershell
docker build -f dev/build/Dockerfile -t wiki-local:2.0.0 .
```

时，这些运行数据目录也会被发送到 Docker 构建上下文，可能导致构建明显变慢。

建议后续补一个 `.dockerignore`，至少排除：

```text
db-data
wiki-storage
.git
node_modules
```

## 推荐做法

- 只想部署并运行：直接使用根目录 `docker compose up -d`。
- 想把当前源码版本打成自己的镜像：使用 `dev/build/Dockerfile` 执行 `docker build`。
- 想用 Compose 统一管理自建镜像：在 `docker-compose.yaml` 中添加 `build`，并把卷改到 `/wiki/data/content`。
