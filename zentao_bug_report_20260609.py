#!/usr/bin/env python3
import json
import requests
import pandas as pd
from datetime import datetime, timedelta

ZENTAO_URL = "http://192.168.100.52"
USERNAME = "lil"
PASSWORD = "123456"

# 团队成员列表
TEAM_MEMBERS = [
    "陈国栋", "何江涛", "柳城", "吕自成", "马力（智）",
    "王聪", "许鹏（研发）", "杨吉生", "杨义"
]

today = datetime.now().strftime("%Y-%m-%d")
OUTPUT_EXCEL = f"BUG数量周统计报表-{today}.xlsx"

PRIORITY_MAP = {
    "1": "紧急",
    "2": "高",
    "3": "中",
    "4": "低",
}


def login(session):
    url = f"{ZENTAO_URL}/zentao/user-login.json"
    resp = session.post(url, data={"account": USERNAME, "password": PASSWORD})
    resp.raise_for_status()
    result = resp.json()
    if result.get("status") != "success":
        raise RuntimeError(f"Login failed: {result}")
    user = result.get("user", {})
    print(f"登录成功: {user.get('realname', '')} ({user.get('account', '')})")
    return result


def get_products_and_users(session):
    url = f"{ZENTAO_URL}/zentao/bug-browse-528.json"
    resp = session.get(url)
    resp.raise_for_status()
    text = resp.text
    first_obj = text.split('{"status"')[1]
    result = json.loads('{"status"' + first_obj)
    data = json.loads(result["data"])
    return data.get("products", {}), data.get("users", {})


def parse_first_json(text):
    """从可能包含多个JSON对象的文本中解析第一个成功的JSON"""
    decoder = json.JSONDecoder()
    obj, _ = decoder.raw_decode(text)
    return obj


def get_bugs_for_product(session, product_id):
    bugs = []
    page = 1
    rec_total = None
    rec_per_page = 20

    while True:
        if rec_total is None:
            url = f"{ZENTAO_URL}/zentao/bug-browse-{product_id}.json"
        else:
            url = f"{ZENTAO_URL}/zentao/bug-browse-{product_id}---0--{rec_total}-{rec_per_page}-{page}.json"

        try:
            resp = session.get(url, timeout=30)
            resp.raise_for_status()
            result = parse_first_json(resp.text)
            if result.get("status") != "success":
                break

            data = json.loads(result["data"])
            if "message" in data or "locate" in data:
                break

            pager = data.get("pager", {})
            page_bugs = data.get("bugs", [])

            if rec_total is None:
                rec_total = pager.get("recTotal", 0)
                if rec_total == 0:
                    break

            if not page_bugs:
                break

            bugs.extend(page_bugs)

            if page * rec_per_page >= rec_total:
                break
            page += 1
        except Exception:
            break

    return bugs


def main():
    session = requests.Session()
    login(session)

    products, users = get_products_and_users(session)
    account_to_realname = {}
    for account, realname in users.items():
        account_to_realname[account] = realname

    print(f"共 {len(products)} 个产品，开始遍历...\n")

    all_bugs = []
    for pid, pname in products.items():
        bugs = get_bugs_for_product(session, pid)
        if bugs:
            print(f"  产品 {pname}: {len(bugs)} 条BUG")
        for bug in bugs:
            assigned_account = bug.get("assignedTo", "")
            realname = account_to_realname.get(assigned_account, assigned_account)
            deadline = bug.get("deadline", "")
            severity = bug.get("severity", "")

            overdue_days = 0
            if deadline and deadline != "0000-00-00":
                try:
                    deadline_date = datetime.strptime(deadline, "%Y-%m-%d")
                    overdue_days = (datetime.now() - deadline_date).days
                except ValueError:
                    overdue_days = 0

            opened_date = bug.get("openedDate", "")
            pri = bug.get("pri", "")
            status = bug.get("status", "")
            is_active = "是" if status == "active" else "否"

            all_bugs.append({
                "BUG_ID": bug.get("id", ""),
                "产品": pname,
                "BUG标题": bug.get("title", ""),
                "严重程度": severity,
                "状态": status,
                # "指派人账号": assigned_account,
                "指派人": realname,
                "截止日期": deadline,
                "超期天数": overdue_days,
                "优先级": pri,
                "是否激活": is_active,
                "创建日期": opened_date,
            })

    print(f"\n共收集 {len(all_bugs)} 条BUG")

    df = pd.DataFrame(all_bugs)

    # 按团队成员统计
    stats = []
    for member in TEAM_MEMBERS:
        member_bugs = df[(df["指派人"] == member) & (df["状态"] == "active")]
        total = len(member_bugs)
        sev1 = len(member_bugs[member_bugs["严重程度"] == "1"])
        sev2 = len(member_bugs[member_bugs["严重程度"] == "2"])
        sev34 = len(member_bugs[member_bugs["严重程度"].isin(["3", "4"])])
        overdue30 = len(member_bugs[member_bugs["超期天数"] > 30])
        overdue60 = len(member_bugs[member_bugs["超期天数"] > 60])
        stats.append({
            "团队成员": member,
            "bug总数": total,
            "一级bug": sev1,
            "二级bug": sev2,
            "三、四级bug": sev34,
            "超期30天": overdue30,
            "超期60天": overdue60,
        })

    stats_df = pd.DataFrame(stats)

    # 筛选团队成员的问题单记录
    # record_df = df[df["指派人"].isin(TEAM_MEMBERS)].copy()
    # record_df = record_df.rename(columns={"指派人": "指派给"})
    # record_df = record_df[["BUG_ID", "BUG标题", "产品", "严重程度", "优先级", "是否激活", "指派给", "创建日期", "截止日期", "超期天数"]].copy()
    # record_df["优先级"] = record_df["优先级"].map(PRIORITY_MAP).fillna(record_df["优先级"])
    # record_df = record_df.sort_values(by=["指派给", "创建日期"], ascending=[True, False])

    detail_df = df[df["指派人"].isin(TEAM_MEMBERS)].copy()

    with pd.ExcelWriter(OUTPUT_EXCEL, engine="openpyxl") as w:
        stats_df.to_excel(w, sheet_name="团队BUG统计", index=False)
        detail_df.to_excel(w, sheet_name="BUG明细", index=False)
        # record_df.to_excel(w, sheet_name="问题单记录表", index=False)

    print(f"\n{'='*70}")
    print(f"{'团队成员':<12} {'bug总数':<8} {'一级bug':<8} {'二级bug':<8} {'三、四级bug':<12} {'超期30天':<8} {'超期60天':<8}")
    print(f"{'='*70}")
    for _, row in stats_df.iterrows():
        print(f"{row['团队成员']:<12} {row['bug总数']:<8} {row['一级bug']:<8} {row['二级bug']:<8} {row['三、四级bug']:<12} {row['超期30天']:<8} {row['超期60天']:<8}")
    print(f"{'='*70}")
    print(f"\n报表已生成: {OUTPUT_EXCEL}")


if __name__ == "__main__":
    main()
