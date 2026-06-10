# ====================== 依赖导入 ======================
import os
import re
import json
import subprocess
import requests
import pandas as pd
from datetime import datetime

# ====================== 【配置项】 ======================
REPO_ROOT = os.path.expanduser("~/jenkins_agent/workspace/git_repos_stat/20260526")
REPOS = [
    "Public_ar9481_product",
    "External_Linux_AR9481_System",
    "Public_Glpp2.0",
    "External_Linux_ARS31_SDK",
    "Public_ars31_product",
    "Public_glpp",
]

SINCE = "2026-05-20"
UNTIL = "2026-06-09"

ZENTAO_URL = "http://192.168.100.52"
ZENTAO_USER = "lil"
ZENTAO_PWD  = "123456"

STATUS_MAP = {
    "active": "激活",
    "resolved": "已解决",
    "closed": "已关闭",
    "confirmed": "已确认",
}

TYPE_MAP = {
    "codeerror": "代码错误",
    "config": "配置相关",
    "install": "安装部署",
    "security": "安全相关",
    "performance": "性能问题",
    "standard": "标准规范",
    "automation": "自动化测试",
    "designdefect": "设计缺陷",
    "others": "其他",
}

today = datetime.now().strftime("%Y-%m-%d")
OUTPUT_EXCEL = f"BUG修复追踪报表-{today}.xlsx"
# =====================================================================

# ====================== Git ======================
def run_git(cmd, cwd):
    try:
        return subprocess.check_output(cmd, shell=True, text=True, encoding="utf-8", errors="ignore", cwd=cwd).strip()
    except:
        return ""

def fetch_all_branches(cwd):
    print("  → 拉取最新代码...")
    result = run_git("git fetch --all --prune 2>&1", cwd)
    if result:
        print(f"  → {result}")
    else:
        print("  → 拉取完成")

# ====================== 禅道 ======================
zentao_session = None

def zentao_login():
    global zentao_session
    try:
        zentao_session = requests.Session()
        url = f"{ZENTAO_URL}/zentao/user-login.json"
        data = {"account": ZENTAO_USER, "password": ZENTAO_PWD}
        resp = zentao_session.post(url, data=data, timeout=10)
        result = resp.json()
        if result.get("status") == "success":
            print("  → 禅道登录成功")
        else:
            print(f"  → 禅道登录失败: {result.get('errmsg', '')}")
            zentao_session = None
    except Exception as e:
        print(f"  → 禅道登录异常: {e}")
        zentao_session = None

def get_bug_detail(bug_id):
    if not zentao_session or not bug_id:
        return {"title": "", "product": "", "status": "", "handler": ""}
    try:
        resp = zentao_session.get(f"{ZENTAO_URL}/zentao/bug-view-{bug_id}.json", timeout=10)
        resp.raise_for_status()
        result = resp.json()
        if result.get("status") != "success":
            return {"title": "不存在", "product": "", "status": "不存在", "handler": ""}

        data = json.loads(result["data"])
        bug = data.get("bug", {})
        users = data.get("users", {})
        products = data.get("products", {})

        assigned_account = bug.get("assignedTo", "")
        handler_name = users.get(assigned_account, assigned_account)
        if handler_name:
            handler_name = f"{handler_name} ({assigned_account})" if assigned_account else "无"

        product_id = str(bug.get("product", ""))
        product_name = products.get(product_id, product_id)

        status_raw = bug.get("status", "")
        status_cn = STATUS_MAP.get(status_raw, status_raw)

        return {
            "title": bug.get("title", ""),
            "product": product_name,
            "status": f"{status_cn} ({status_raw})" if status_raw else "",
            "handler": handler_name,
        }
    except Exception as e:
        return {"title": f"查询失败: {e}", "product": "", "status": "", "handler": ""}

# ====================== 【核心修复：标题+正文一起解析】 ======================
def parse_commit(repo, commit):
    cwd = f"{REPO_ROOT}/{repo}"
    msg = run_git(f"git log -1 {commit} --pretty=%s", cwd)
    body = run_git(f"git log -1 {commit} --pretty=%b", cwd)
    author = run_git(f"git log -1 {commit} --pretty=%an", cwd)
    time_str = run_git(f"git log -1 {commit} --pretty=%ai", cwd)

    full_text = (msg + "\n" + body).lower()

    # 1. 必须是 fix 开头
    if not msg.lower().startswith("fix"):
        return []

    # 2. 从 标题+正文 里找所有 bug-xxx（支持 bug-102474 112540 这种空格分隔的多个ID）
    bug_matches = re.findall(r"bug-([\d\s]+)", full_text)
    bug_ids = []
    for m in bug_matches:
        bug_ids += m.split()
    if not bug_ids:
        return []

    # 模块
    module = "通用模块"
    sm = re.match(r"fix\((\w+)\)", msg, re.I)
    if sm:
        module = sm.group(1)

    rows = []
    for bid in bug_ids:
        bug = get_bug_detail(bid)
        rows.append({
            "仓库": repo,
            "提交人": author,
            "提交时间": time_str[:19] if len(time_str) >= 19 else "",
            "提交信息": msg,
            "BUG_ID": bid,
            "BUG标题": bug["title"],
            "产品": bug["product"],
            "BUG状态": bug["status"],
            "当前处理人": bug["handler"],
        })
    return rows

# ====================== 扫描 ======================
def scan_repo(repo):
    print(f"\n===== 扫描：{repo} =====")
    cwd = f"{REPO_ROOT}/{repo}"
    if not os.path.exists(cwd):
        print(f"  → 跳过")
        return []
    fetch_all_branches(cwd)
    cmd = f"git log --all --remotes=origin --since='{SINCE}' --until='{UNTIL}' --pretty=%H --no-merges"
    commits = [c for c in run_git(cmd, cwd).splitlines() if c]
    print(f"  → 总提交：{len(commits)}")
    rows = []
    for cm in commits:
        r = parse_commit(repo, cm)
        if r:
            for item in r:
                print(f"  ✅ 找到BUG：{cm[:8]} | bug-{item['BUG_ID']}")
            rows += r
    return rows

# ====================== 主程序 ======================
if __name__ == "__main__":
    print("=============================================")
    print("    BUG 修复提交统计（fix + bugID）")
    print("=============================================")
    zentao_login()
    all_data = []
    for repo in REPOS:
        all_data += scan_repo(repo)

    df = pd.DataFrame(all_data)
    print("\n=============================================")
    print(f"📊 最终找到 BUG 修复提交：{len(df)} 条")

    with pd.ExcelWriter(OUTPUT_EXCEL, engine="openpyxl") as w:
        df.to_excel(w, sheet_name="BUG修复明细", index=False)

    print(f"✅ 报表已生成：{OUTPUT_EXCEL}")