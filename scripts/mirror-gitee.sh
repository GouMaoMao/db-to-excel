#!/usr/bin/env bash
# 把 GitHub 的 main、三段版本标签和 db-to-excel.exe 快进镜像到 Gitee。
# 令牌只写在 url.insteadOf 中，推送地址本身不含令牌。Gitee 历史分叉时拒绝覆盖。

set -euo pipefail

ASSET_NAME="db-to-excel.exe"

if [[ -z "${GITEE_TOKEN:-}" || -z "${GITEE_REPO:-}" ]]; then
  echo "缺少 GitHub Secret GITEE_TOKEN 或 Variable GITEE_REPO。"
  echo "GITEE_REPO 要建在 Settings、Secrets and variables、Actions 的 Variables 页，不是 Secrets。"
  exit 1
fi
if [[ ! "$GITEE_REPO" =~ ^[^/[:space:]]+/[^/[:space:]]+$ ]]; then
  echo "GITEE_REPO 必须是 owner/repo，当前是：${GITEE_REPO}"
  exit 1
fi

OWNER="${GITEE_REPO%%/*}"
REPO="${GITEE_REPO#*/}"
API="https://gitee.com/api/v5/repos/${OWNER}/${REPO}"

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

# 单次请求无响应时失败后再试。查询最多三次，每次一分钟。
# 上传安装包最多两次，每次最多三十分钟，并记下已发送字节，便于判断是卡住还是链路太慢。
# 100 是 Expect 的中间状态，不是上传结果。诊断写到标准错误，标准输出只留 HTTP 状态码。
gitee_request() {
  local method="$1"
  local url="$2"
  local out="$3"
  shift 3
  local query=()
  local extra=()
  local speed=()
  local timeout=60
  local tries=3
  local attempt code arg result sent
  if [[ "$method" == "GET" || "$method" == "DELETE" ]]; then
    query=(-G)
  fi
  for arg in "$@"; do
    if [[ "$arg" == file=@* ]]; then
      timeout=1800
      tries=2
      extra=(--http1.1 -H "Expect:")
      speed=(--speed-limit 1024 --speed-time 45)
      break
    fi
  done
  for ((attempt = 1; attempt <= tries; attempt++)); do
    : > "$out"
    result="$(curl -sS --connect-timeout 20 --max-time "$timeout" "${speed[@]}" \
      -o "$out" -w "%{http_code} %{size_upload}" \
      -X "$method" "${extra[@]}" "${query[@]}" "$url" "$@" || true)"
    code="${result%% *}"
    sent="${result##* }"
    if [[ ! "$code" =~ ^[0-9]{3}$ ]]; then
      code="000"
    fi
    case "$code" in
      100|000|408|429|500|502|503|504)
        if [[ "$attempt" -lt "$tries" ]]; then
          echo "Gitee 接口暂时没有响应（HTTP ${code}，已发送 ${sent} 字节），5 秒后重试。" >&2
          sleep 5
          continue
        fi
        echo "Gitee 接口没有完成（HTTP ${code}，已发送 ${sent} 字节）。" >&2
        ;;
      *)
        printf '%s' "$code"
        return 0
        ;;
    esac
  done
  printf '%s' "$code"
}

show_api_error() {
  local tmp="$1"
  jq -r '.message // .' "$tmp" 2>/dev/null || true
}

read_release_id() {
  jq -r 'if type == "object" and (.id | type == "number") then .id else "" end' "$1"
}

require_release_id() {
  local release_id="$1"
  local tmp="$2"
  if [[ ! "$release_id" =~ ^[0-9]+$ ]]; then
    echo "Gitee 没有返回 Release 编号。"
    show_api_error "$tmp"
    exit 1
  fi
}

delete_same_asset() {
  local release_id="$1"
  local json="$2"
  local file_id tmp code
  while IFS= read -r file_id; do
    [[ -z "$file_id" ]] && continue
    tmp="$(mktemp)"
    echo "删除 Gitee 上的旧 ${ASSET_NAME}（${file_id}）。"
    code="$(gitee_request DELETE "${API}/releases/${release_id}/attach_files/${file_id}" "$tmp" \
      --data-urlencode "access_token=${GITEE_TOKEN}")"
    if [[ "$code" != "204" && "$code" != "200" ]]; then
      echo "删除 Gitee 附件 ${file_id} 失败，HTTP ${code}。"
      show_api_error "$tmp"
      exit 1
    fi
  done < <(jq -r --arg name "$ASSET_NAME" '
    (.assets // .attach_files // [])[]?
    | select((.name // .file_name // "") == $name)
    | .id
  ' "$json")
}

sync_release() {
  local tag="$1"
  local title="$2"
  local notes="$3"
  local file="$4"
  local tmp code release_id attempt
  tmp="$(mktemp)"

  if [[ -z "$title" || "$title" == "null" ]]; then
    title="$tag"
  fi
  if [[ -z "${notes//[[:space:]]/}" || "$notes" == "null" ]]; then
    notes="$title"
  fi

  code="$(gitee_request GET "${API}/releases/tags/${tag}" "$tmp" \
    --data-urlencode "access_token=${GITEE_TOKEN}")"
  release_id="$(read_release_id "$tmp")"
  # 没有发行版时 Gitee 返回 HTTP 200 和 JSON null，不返回 404。
  if [[ "$release_id" =~ ^[0-9]+$ ]]; then
    delete_same_asset "$release_id" "$tmp"
    code="$(gitee_request PATCH "${API}/releases/${release_id}" "$tmp" \
      --data-urlencode "access_token=${GITEE_TOKEN}" \
      --data-urlencode "tag_name=${tag}" \
      --data-urlencode "name=${title}" \
      --data-urlencode "body=${notes}" \
      --data-urlencode "target_commitish=main" \
      --data-urlencode "prerelease=false")"
    if [[ "$code" != "200" ]]; then
      echo "更新 Gitee Release ${tag} 失败，HTTP ${code}。"
      show_api_error "$tmp"
      exit 1
    fi
  elif [[ "$code" == "200" || "$code" == "404" ]]; then
    release_id=""
    for attempt in 1 2 3; do
      code="$(gitee_request POST "${API}/releases" "$tmp" \
        --data-urlencode "access_token=${GITEE_TOKEN}" \
        --data-urlencode "tag_name=${tag}" \
        --data-urlencode "name=${title}" \
        --data-urlencode "body=${notes}" \
        --data-urlencode "target_commitish=main" \
        --data-urlencode "prerelease=false")"
      release_id="$(read_release_id "$tmp")"
      if [[ "$release_id" =~ ^[0-9]+$ ]]; then
        break
      fi
      if [[ "$attempt" -lt 3 ]]; then
        echo "创建 Gitee Release ${tag} 未成功（HTTP ${code}），5 秒后重试。"
        show_api_error "$tmp"
        sleep 5
        continue
      fi
      echo "创建 Gitee Release ${tag} 失败，HTTP ${code}。"
      show_api_error "$tmp"
      exit 1
    done
    require_release_id "$release_id" "$tmp"
  else
    echo "查询 Gitee Release ${tag} 失败，HTTP ${code}。"
    show_api_error "$tmp"
    exit 1
  fi

  code="$(gitee_request POST "${API}/releases/${release_id}/attach_files" "$tmp" \
    -F "access_token=${GITEE_TOKEN}" \
    -F "file=@${file};filename=${ASSET_NAME}")"
  if [[ "$code" != "201" && "$code" != "200" ]]; then
    echo "上传 ${ASSET_NAME} 到 Gitee Release ${tag} 失败，HTTP ${code}。"
    show_api_error "$tmp"
    exit 1
  fi
  echo "已同步 Gitee Release ${tag}。"
}

download_asset() {
  local tag="$1"
  local dest_dir="$2"
  local attempt
  mkdir -p "$dest_dir"
  rm -f "${dest_dir}/${ASSET_NAME}"
  for attempt in 1 2 3 4 5 6; do
    if gh release download "$tag" --dir "$dest_dir" --pattern "$ASSET_NAME" --clobber \
      && [[ -s "${dest_dir}/${ASSET_NAME}" ]]; then
      return 0
    fi
    echo "第 ${attempt} 次还没有 ${tag} 的 ${ASSET_NAME}，10 秒后重试。"
    sleep 10
  done
  return 1
}

publish_tag_and_asset() {
  local tag="$1"
  local title="$2"
  local notes="$3"
  local file="$4"
  ensure_tag "$tag"
  push_ref HEAD refs/heads/main
  push_ref "refs/tags/${tag}" "refs/tags/${tag}"
  sync_release "$tag" "$title" "$notes" "$file"
}

on_push() {
  if [[ "$REF_TYPE" == "branch" && "$REF_NAME" == "main" ]]; then
    push_ref HEAD refs/heads/main
  elif [[ "$REF_TYPE" == "tag" ]]; then
    ensure_tag "$REF_NAME"
    push_ref HEAD refs/heads/main
    push_ref "refs/tags/${REF_NAME}" "refs/tags/${REF_NAME}"
  else
    echo "不镜像 ${REF_TYPE} ${REF_NAME}。"
    exit 1
  fi
}

on_release() {
  local dir
  if [[ -z "${RELEASE_TAG:-}" ]]; then
    echo "Release 事件没有标签名。"
    exit 1
  fi
  dir="$(mktemp -d)"
  if ! download_asset "$RELEASE_TAG" "$dir"; then
    echo "GitHub Release ${RELEASE_TAG} 上没有 ${ASSET_NAME}。"
    exit 1
  fi
  publish_tag_and_asset "$RELEASE_TAG" "${RELEASE_NAME:-}" "${RELEASE_BODY:-}" "${dir}/${ASSET_NAME}"
}

on_backfill() {
  local tag meta title notes has dir
  if [[ -n "${MIRROR_TAG:-}" ]]; then
    if [[ ! "$MIRROR_TAG" =~ ^v[0-9]+\.[0-9]+\.[0-9]+$ ]]; then
      echo "MIRROR_TAG 必须是 v主.次.修订，当前是：${MIRROR_TAG}"
      exit 1
    fi
    meta="$(gh release view "$MIRROR_TAG" --json name,body)"
    RELEASE_TAG="$MIRROR_TAG"
    RELEASE_NAME="$(jq -r '.name // ""' <<<"$meta")"
    RELEASE_BODY="$(jq -r '.body // ""' <<<"$meta")"
    on_release
    return
  fi
  push_ref HEAD refs/heads/main
  while IFS= read -r tag; do
    [[ "$tag" =~ ^v[0-9]+\.[0-9]+\.[0-9]+$ ]] || continue
    push_ref "refs/tags/${tag}" "refs/tags/${tag}"
  done < <(git tag -l 'v*')

  while IFS= read -r tag; do
    [[ -z "$tag" ]] && continue
    meta="$(gh release view "$tag" --json name,body,assets)"
    has="$(jq -r --arg name "$ASSET_NAME" '[(.assets // [])[].name] | index($name) != null' <<<"$meta")"
    if [[ "$has" != "true" ]]; then
      echo "跳过 ${tag}：GitHub Release 上没有 ${ASSET_NAME}。"
      continue
    fi
    title="$(jq -r '.name // ""' <<<"$meta")"
    notes="$(jq -r '.body // ""' <<<"$meta")"
    dir="$(mktemp -d)"
    if ! download_asset "$tag" "$dir"; then
      echo "下载 ${tag} 的 ${ASSET_NAME} 失败。"
      exit 1
    fi
    publish_tag_and_asset "$tag" "$title" "$notes" "${dir}/${ASSET_NAME}"
  done < <(gh release list --limit 500 --json tagName,isDraft \
    | jq -r '.[] | select(.isDraft == false) | .tagName')
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
