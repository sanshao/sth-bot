import pandas as pd
import os
from datetime import datetime


def process_taobao_file(file_path, target_directory):
    print("当前文件：", file_path)

    # 读取Excel文件
    xls = pd.ExcelFile(file_path)

    # 读取第一个工作表
    df = pd.read_excel(xls, sheet_name=0, header=4, skipfooter=4)

    # 打印列名以供检查
    print("读取的列名：", df.columns.tolist())

    # 清理列名，去除前后空格
    df.columns = df.columns.str.strip()

    # 计算净值并添加到新列
    df['净值'] = df['收入金额（+元）'] + df['支出金额（-元）']

    # 计算净值并添加到新列
    df['净值'] = df['收入金额（+元）'] + df['支出金额（-元）']

    # 定义分类字典
    category_mapping = {
        "结息":["基金代发任务"],
        "淘宝-百亿服务费":["百亿补贴软件服务费"],
        "淘宝-保险费":["消费者体验提升计划服务费", "保险承保"],
        # "淘宝-保证金充值":["淘宝消费者保证金-充值（代扣）", "淘宝消费者保证金-充值（代扣）"],
        "淘宝-保证金解冻":["淘宝消费者保证金-解冻"],
        "淘宝-大促软件服务费":["淘宝大促软件技术服务费"],
        "淘宝-官方竞价软件费":["官方竞价软件服务费"],
        "淘宝-基础软件服务费":["基础软件服务费", "淘宝天猫跨境服务基础费"],
        "淘宝-公益":["公益宝贝捐赠"],
        # "淘宝-交易收款":[""],
        "淘宝-交易退款":["淘宝消费者保证金-交易售后", "售后退款-"],
        "淘宝-保证金退款":["保证金退款"],
        "淘宝-保证金退邮费":["淘宝消费者保证金-充值（代扣）-退货邮费"],
        "淘宝-跨境服务费":["淘宝天猫跨境服务增值费"],
        "淘宝-赔付":[
            "支付宝转账小额打款-关联订单号",
            "淘宝消费者保证金-充值（代扣）-缺货", 
            "淘宝消费者保证金-充值（代扣）-红包冻结", 
            "淘宝消费者保证金-充值（代扣）-延迟发货", "商家权益红包-预算追加-卖家延迟发货赔付红包-赔付红包",
            "商家权益红包-预算追加-淘宝缺货赔付红包-赔付红包"
        ],
        "淘宝-虚假发货赔付":["预算追加-淘宝虚假发货赔付"],
        "淘宝-品牌护肤赤兔":["代扣-赤兔名"],
        "淘宝-首单拉新":["淘宝新客礼金技术服务费", "品牌新享淘宝老客礼金软件服务费", "品牌新享淘宝限时礼金软件服务费"],
        "淘宝-淘宝客佣金":["淘宝客佣金代扣款", "淘宝联盟推广佣金返还", "淘宝联盟佣金代扣"],
        "淘宝-万相台充值":["万相台无界版自动充值"],
        "淘宝-转运物流费":["商家集运"],
        "淘宝-每日必买服务费":["每日必买"],
        "淘宝-物流轨迹异常":["淘宝物流轨迹异常"],
        "淘工厂-托管充值":["账户充值-工作台充值","账户充值-手动充值","账户充值-自动充值"],
        "淘工厂-促销费":["直营&联营&营促销"],
        # "淘工厂-好评返现":["评价有礼"],
        "淘工厂-交易收款":["C2M订单交易货款分账", "C2M-退款赔付-申诉单号", "订单交易货款分账", "C2M-合作费用-订单号"],
        "淘工厂-交易退款":["分账退回", "淘特直营保证金履约险_追偿款"], 
        "淘工厂-赔付":["天猫售后赔付", "因物流轨迹异常-物流停滞原因 扣罚", "因延迟发货原因 扣罚", "扣款用途：C2M-处罚赔付"],
        "淘工厂-商家出资补贴":["正向扣款-商家出资补贴", "逆向退款-商家出资补贴"],
        "淘工厂-赠品营销费": ["正向扣款-买赠赠品营销费用", "逆向退款-买赠赠品营销费用"],
        "淘工厂-淘宝佣金":["正向扣佣-精选淘客-", "红包签到供应商cps佣金"],
        "淘工厂-托管费":["商品运营托管推广服务费"],
        "淘工厂-托管账户充值":["扣款用途：账户充值-工作台充值", "扣款用途：账户充值-自动充值"],
        "淘工厂-先用后付":["C2M-先用后付技术服务费"],
        "淘宝-先用后付服务费":["淘宝-先用后付服务费", "先用后付技术服务费"],
        "淘工厂-运费险":["退货包运费代扣"],
        # "提现":[""],
        "网商贷-放款":["网商贷-放款"],
        "网商贷-还款":["网商贷-还款", "网商银行扣款"],
        "支付宝-花呗还款": ["花呗|信用购"],
        "淘工厂-技术服务费": ["扣款用途：C2M-技术服务费"],
        "淘工厂-转运物流费": ["新疆物流集运"],
        # "淘工厂-违约赔付": ["淘宝物流轨迹异常"],
    }

    # 函数来根据关键字确定分类

    # 函数来根据关键字确定分类
    def assign_category(row):
        # 从各个相关列获取字符串
        item_name = str(row['商品名称']).strip()  # 请确认“商品名称”列名
        business_type = str(row['业务类型']).strip()  # 请确认“业务类型”列名
        remark = str(row['备注']).strip()  # 请确认“备注”列名
        counterparty = str(row['对方账号']).strip()  # 请确认“对方账号”列名        
        pay_amount = row['支出金额（-元）']; # 支出金额
        income_amount = row['收入金额（+元）']; # 收入金额
        
        # "淘宝-保证金充值":["淘宝消费者保证金-充值（代扣）", "淘宝消费者保证金-充值（代扣）"],
        if remark == "淘宝消费者保证金-充值（代扣）" or remark == "淘宝消费者保证金-充值":
            return "淘宝-保证金充值";
        
        if "万相台无界版扫码充值" in item_name:
            return "淘工厂-直通车充值";
        elif "门道商家助手-基础版-订单付款" in item_name:
            return "门道商家助手-基础版-订单付款";
        elif "赤兔名品客服绩效" in item_name:
            return "千橙食品赤兔";
        
        # 根据业务类型判断
        if business_type == "交易付款" or "基金代发任务" in remark:
            return "淘宝-交易收款";
        elif business_type == "提现":
            return "提现";
        elif business_type == "结息":
            return "结息";
        elif business_type == "在线支付":
            return "在线支付";
        
        if counterparty == "*骁(dux***@gmail.com)" and remark == "转账":
            return "海那边-转账";
        elif counterparty == "**飞(165***@qq.com)" and remark == "转账":
            return "报销-丁庆飞";
        elif counterparty == "*璁(cao***@aliyun.com)" and remark == "转账":
            return "小木登子-转入";
        elif counterparty == "**振(156******90)" and remark == "转账":
            return "赵振";
        elif counterparty == "杭州昌诚电子商务有限公司(ydbbzj@service.aliyun.com)" and remark == "转账":
            return "一大包零食交保证金";
        elif counterparty == "杭州淘宝直播严选电子商务有限公司(qdzfb@service.aliyun.com)" and remark == "转账":
            return "88VIP货款";
        elif counterparty == "**婧(150******97)" and remark == "转账" and pay_amount < 0:
            return "大C店-转出";
        elif counterparty == "**婧(150******97)" and remark == "转账" and income_amount > 0:
            return "大C店-转入";
        
        if "花呗" in remark and "还款" in remark:
            return "支付宝-花呗还款";
        
        if "淘特直营商家管理保证金" in remark and "违约金扣罚" in remark:
            return "淘工厂-违约赔付";
        
        # 检查备注中的关键字
        for category, keywords in category_mapping.items():
            if any(keyword in remark for keyword in keywords):
                return category

        return '其它-未分类'  # 默认分类

    # 添加分类列
    df['分类'] = df.apply(assign_category, axis=1)

    # 调整列顺序，将“净值”列放在“支出金额（-元）”后、账户余额之前
    columns_order = list(df.columns)
    # 找到“支出金额（-元）”和“账户余额（元）”的位置
    outflow_index = columns_order.index('支出金额（-元）')
    balance_index = columns_order.index('账户余额（元）')

    # 重新排列列顺序
    columns_order.insert(balance_index, columns_order.pop(outflow_index + 1))  # 移动“净值”到“支出金额（-元）”和“账户余额（元）”之间
    df = df[columns_order]

    # # 创建新的工作表“整理”
    # with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
    #     df.to_excel(writer, sheet_name='整理', index=False)

    # 创建透视表来计算分类下的净值和总和
    pivot_table = df.pivot_table(values='净值', index='分类', aggfunc='sum').reset_index()

    # 添加总和行
    total_row = pd.DataFrame({'分类': ['总和'], '净值': [pivot_table['净值'].sum()]})
    pivot_table = pd.concat([pivot_table, total_row], ignore_index=True)
    
    base_name = os.path.basename(file_path)
    new_file_name = f"{os.path.splitext(base_name)[0]}_整理.xlsx"
    
    new_file_path = os.path.join(target_directory, new_file_name)
    
    # print("新文件名：", new_file_path, base_name)

    # 创建新的工作表 “整理” “透视”
    with pd.ExcelWriter(new_file_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='整理', index=False)
        pivot_table.to_excel(writer, sheet_name='透视', index=False)

    print(f"处理完成！新文件生成: {new_file_path}")

