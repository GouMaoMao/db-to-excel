#!/usr/bin/env bash
# 把 GitHub 的 main 和三段版本标签快进镜像到 Gitee。
# 安装包由本机发布脚本上传，这里不下载、也不传文件。
# 令牌只写在 url.insteadOf 中，推送地址本身不含令牌。Gitee 历史分叉时拒绝覆盖。

set -euo pipefail

if [[ -z "${GITEE_TOKEN:-}" || -z "${GITEE_REPO:-}" ]]; then
  echo "缺少 GitHub Secret GITEE_TOKEN 或 Variable GITEE_REPO。"
  echo "GITEE_REPO 要建在 Settings、Secrets and variables、Actions 的 Variables 页，不是 Secrets。"
  exit 1
fi
if [[ ! "$GITEE_REPO" =~ ^[^/[:space:]]+/[^/[:space:]]+$ ]]; then
  echo "GITEE_REPO 必须是 owner/repo，当前是：${GITEE_REPO}"
  exit 1
fi

git config --global --add url."https://oauth2:${GITEE_TOKEN}@gitee.com/".insteadOf "https://gitee.com/"
if git remote get-url gitee >/dev/null 2>&1; then
  git remote set-url gitee "https://gitee.com/${GITEE_REPO}.git"
else
  git remote add gitee "https://gitee.com/${GITEE_REPO}.git"
fi

git fetch origin --tags --prune

push_ref() {
  local src="$1"
  local dst="$2"
  echo "推送 ${dst}。"
  git push gitee "${src}:${dst}"
}

ensure_tag() {
  local tag="$1"
  if ! git rev-parse --verify --quiet "refs/tags/${tag}" >/dev/null; then
    git fetch origin "refs/tags/${tag}:refs/tags/${tag}"
  fi
}

push_tag() {
  local tag="$1"
  ensure_tag "$tag"
  push_ref HEAD refs/heads/main
  push_ref "refs/tags/${tag}" "refs/tags/${tag}"
}

on_push() {
  if [[ "$REF_TYPE" == "branch" && "$REF_NAME" == "main" ]]; then
    push_ref HEAD refs/heads/main
  elif [[ "$REF_TYPE" == "tag" ]]; then
    push_tag "$REF_NAME"
  else
    echo "不镜像 ${REF_TYPE} ${REF_NAME}。"
    exit 1
  fi
}

on_release() {
  if [[ -z "${RELEASE_TAG:-}" ]]; then
    echo "Release 事件没有标签名。"
    exit 1
  fi
  push_tag "$RELEASE_TAG"
}

on_backfill() {
  local tag
  if [[ -n "${MIRROR_TAG:-}" ]]; then
    if [[ ! "$MIRROR_TAG" =~ ^v[0-9]+\.[0-9]+\.[0-9]+$ ]]; then
      echo "MIRROR_TAG 必须是 v主.次.修订，当前是：${MIRROR_TAG}"
      exit 1
    fi
    ensure_tag "$MIRROR_TAG"
    push_ref "refs/tags/${MIRROR_TAG}" "refs/tags/${MIRROR_TAG}"
    return
  fi
  push_ref HEAD refs/heads/main
  while IFS= read -r tag; do
    [[ "$tag" =~ ^v[0-9]+\.[0-9]+\.[0-9]+$ ]] || continue
    push_ref "refs/tags/${tag}" "refs/tags/${tag}"
  done < <(git tag -l 'v*')
}

case "${EVENT_NAME:-}" in
  push)
    on_push
    ;;
  release)
    on_release
    ;;
  workflow_dispatch)
    on_backfill
    ;;
  *)
    echo "未知事件：${EVENT_NAME:-}"
    exit 1
    ;;
esac
