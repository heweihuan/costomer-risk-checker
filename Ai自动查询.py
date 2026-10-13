#第一次使用请先前往cmd输入该代码
# pip install pandas requests openpyxl
import pandas as pd
import requests
import os
import time
import sys


# 环境变量配置，不要写死密钥
QCC_KEY = os.environ.get("QCC_KEY")
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY")


def query_company(name):
    """
    查询单家公司信息。
    如果有企业数据接口密钥，就调用接口；
    如果没有密钥，就返回虚拟演示数据，方便课程项目直接运行。
    """
    result = {
        "公司名": name,
        "登记状态": "",
        "注册资本": "",
        "参保人数": "",
        "法定代表人": "",
        "成立时间": "",
        "经营异常": "",
    }

    if QCC_KEY:
        try:
            url = "https://api.example.com/company/detail"
            resp = requests.get(
                url,
                params={"keyword": name, "key": QCC_KEY},
                timeout=10,
            )
            data = resp.json().get("data", {})
            result["登记状态"] = data.get("regStatus", "")
            result["注册资本"] = data.get("regCapital", "")
            result["参保人数"] = data.get("socialNum", "")
            result["法定代表人"] = data.get("legalPerson", "")
            result["成立时间"] = data.get("estiblishTime", "")
            result["经营异常"] = data.get("abnormalInfo", "")
        except Exception as e:
            result["登记状态"] = f"查询失败：{e}"
    else:
        demo = {
            "张三科技有限公司": {
                "登记状态": "存续",
                "注册资本": "100万元",
                "参保人数": "3人",
                "法定代表人": "张三",
                "成立时间": "2021-05-10",
                "经营异常": "无",
            },
            "李四商贸有限公司": {
                "登记状态": "注销",
                "注册资本": "50万元",
                "参保人数": "0人",
                "法定代表人": "李四",
                "成立时间": "2018-03-12",
                "经营异常": "未按时公示年报",
            },
            "王五文化传媒有限公司": {
                "登记状态": "存续",
                "注册资本": "200万元",
                "参保人数": "12人",
                "法定代表人": "王五",
                "成立时间": "2019-07-22",
                "经营异常": "无",
            },
            "赵六电子商务有限公司": {
                "登记状态": "存续",
                "注册资本": "10万元",
                "参保人数": "0人",
                "法定代表人": "赵六",
                "成立时间": "2024-11-01",
                "经营异常": "地址异常",
            },
            "孙七网络科技有限公司": {
                "登记状态": "吊销",
                "注册资本": "500万元",
                "参保人数": "0人",
                "法定代表人": "孙七",
                "成立时间": "2015-09-18",
                "经营异常": "严重违法失信",
            },
        }

        info = demo.get(name, {
            "登记状态": "存续",
            "注册资本": "100万元",
            "参保人数": "5人",
            "法定代表人": "未知",
            "成立时间": "2020-01-01",
            "经营异常": "无",
        })
        result.update(info)

    return result


def ai_risk_summary(row):
    """
    对单家公司生成风险总结。
    如果配置了 OPENAI_API_KEY，就调用大模型接口；
    否则使用内置规则判断。
    """
    status = str(row.get("登记状态", ""))
    abnormal = str(row.get("经营异常", ""))
    insured = str(row.get("参保人数", ""))
    name = str(row.get("公司名", ""))

    if OPENAI_API_KEY:
        prompt = (
            f"你是客户背景核查助手。根据以下公司信息，输出风险等级、主要风险点和跟进建议。\n"
            f"公司名：{name}\n"
            f"登记状态：{status}\n"
            f"参保人数：{insured}\n"
            f"经营异常：{abnormal}\n"
        )
        try:
            resp = requests.post(
                "https://api.openai.com/v1/chat/completions",
                headers={
                    "Authorization": f"Bearer {OPENAI_API_KEY}",
                    "Content-Type": "application/json",
                },
                json={
                    "model": "gpt-3.5-turbo",
                    "messages": [{"role": "user", "content": prompt}],
                    "temperature": 0.3,
                },
                timeout=15,
            )
            text = resp.json()["choices"][0]["message"]["content"]
            return text.strip()
        except Exception:
            pass

    risk_level = "低风险"
    risk_note = "信息正常，可正常跟进。"

    if "注销" in status or "吊销" in status:
        risk_level = "高风险"
        risk_note = "公司已注销或吊销，不建议跟进。"
    elif "异常" in abnormal or "失信" in abnormal:
        risk_level = "中风险"
        risk_note = "存在经营异常或失信记录，需谨慎核实。"
    elif "0人" in insured or "0" == insured.strip():
        risk_level = "中风险"
        risk_note = "参保人数为0，可能为空壳或经营异常，建议进一步核实。"

    return f"风险等级：{risk_level}。{risk_note}"


def create_demo_excel(path="虚拟客户名单.xlsx"):
    """
    如果没有输入文件，就自动生成一份虚拟客户名单。
    """
    data = [
        {"公司名称": "张三科技有限公司", "联系人": "张三", "电话": "13800000001", "备注": "模拟客户A"},
        {"公司名称": "李四商贸有限公司", "联系人": "李四", "电话": "13800000002", "备注": "模拟客户B"},
        {"公司名称": "王五文化传媒有限公司", "联系人": "王五", "电话": "13800000003", "备注": "模拟客户C"},
        {"公司名称": "赵六电子商务有限公司", "联系人": "赵六", "电话": "13800000004", "备注": "模拟客户D"},
        {"公司名称": "孙七网络科技有限公司", "联系人": "孙七", "电话": "13800000005", "备注": "模拟客户E"},
    ]
    df = pd.DataFrame(data)
    df.to_excel(path, index=False)
    print(f"已生成虚拟客户名单：{path}")
    return path


def main():
    input_file = sys.argv[1] if len(sys.argv) > 1 else "虚拟客户名单.xlsx"

    if not os.path.exists(input_file):
        print(f"未找到文件：{input_file}")
        print("正在自动生成虚拟客户名单……")
        input_file = create_demo_excel()

    df = pd.read_excel(input_file)

    if "公司名称" not in df.columns:
        print("Excel 文件中缺少【公司名称】列，请检查表头。")
        return

    results = []

    for name in df["公司名称"].dropna().unique():
        name = str(name).strip()
        print(f"正在查询：{name}")
        info = query_company(name)
        summary = ai_risk_summary(info)
        info["风险总结"] = summary
        results.append(info)
        time.sleep(1)

    output_file = "客户背景核查结果.xlsx"
    pd.DataFrame(results).to_excel(output_file, index=False)
    print(f"核查完成，结果已保存到：{output_file}")


if __name__ == "__main__":
    main()