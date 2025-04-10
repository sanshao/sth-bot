import pandas as pd
import os
from datetime import datetime


def process_tmall_file(file_path, target_directory):
    print("当前文件：", file_path)

    # 读取Excel文件
    xls = pd.ExcelFile(file_path)

    # 读取第一个工作表
    df = pd.read_excel(xls, sheet_name=0, header=0, skipfooter=0)

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
        # "88VIP结算款":["88VIP结算款"],
        # "结息":["基金代发任务"],
        "天猫-百亿服务费":["百亿补贴软件服务费"],
        "天猫-保险费":["消费者体验提升计划服务费", "保险承保", "保险理赔"],
        "天猫-保证金解冻":["天猫保证金-解冻"],
        "天猫-保证金赔付": ["天猫保证金-充值（代扣）-缺货", "天猫保证金-充值（代扣）-红包冻结", "天猫保证金-充值（代扣）-未按时开具发票"],
        "天猫-延迟发货赔付": ["天猫保证金-充值（代扣）-延迟发货", "天猫保证金-充值（代扣）-物流轨迹异常", "天猫保证金履约险追偿_"],
        "天猫-退货包运费":["天猫保证金-充值（代扣）-退货邮费"],
        "天猫-虚假发货赔付":["虚假发货赔付"],
        "天猫-返点积分": ["代扣返点积分", "代扣交易退回积分"],
        "天猫-公益":["公益宝贝捐赠"],
        "天猫-官方竞价软件费":["官方竞价软件服务费"],
        "天猫-好评返现": ["支付宝转账小额打款-关联订单号", "支付宝转账小额打款-未关联"],
        "天猫-基础软件服务费":["基础软件服务费", "淘宝天猫跨境服务基础费", "C2M-技术服务费"],
        "天猫-跨境服务费":["淘宝天猫跨境服务增值费"],
        "天猫-交易退款":["消费者保证金-交易售后", "售后退款-", "售后支付"],
        "天猫-保证金退款":["保证金退款"],
        "天猫-首单拉新":["天猫新客礼金技术服务费", "品牌新享天猫老客礼金软件服务费", 
                   "品牌新享天猫限时礼金软件服务费", "品牌新享-首单拉新计划", "品牌新享新品孵化软件服务费", "品牌新享会员礼金", "首单技术礼金",
                   "品牌新享天猫超级老客加速软件服务费"
                   ],
        "天猫-淘宝客佣金":["淘宝客佣金代扣款", "淘宝联盟推广佣金返还", "淘宝联盟佣金代扣"],
        "天猫-天猫佣金": ["扣款用途：天猫佣金"],
        "天猫-万相台充值":["万相台无界版自动充值"],
        "天猫-转运物流费":["商家集运中转操作费", "商家集运物流服务费"],
        "天猫超市-扣款": ["扣款用途：DDD商家扣款", "扣款用途：DDD商家结算款 "],
        "天猫超市-交易收款": ["DDD商家结算款"], # 
        "天猫超市-淘客佣金": ["扣款用途：DDD淘客佣金"],
        # "天猫超市-延迟发货": ["扣款用途：DDD商家扣款"],
        # "淘工厂-充值":["账户充值-工作台充值","账户充值-手动充值","账户充值-自动充值"],
        "淘工厂-淘宝佣金":["正向扣佣-精选淘客-", "红包签到供应商cps佣金"],
        "淘工厂-促销费":["直营&联营&营促销"],
        "淘工厂-好评返现":["评价有礼"],
        "淘工厂-交易收款":["C2M订单交易货款分账", "C2M-退款赔付-申诉单号", "订单交易货款分账", "C2M-合作费用-订单号", "扣款用途：C2M-处罚赔付"],
        "淘工厂-交易退款":["分账退回"], 
        "淘工厂-赔付":["天猫售后赔付", "因物流轨迹异常-物流停滞原因 扣罚", "因延迟发货原因 扣罚", "淘特直营保证金履约险_追偿款"],
        "淘工厂-保证金履约险":["淘特直营保证金履约险_保险费", "淘特直营保证金履约险_退保保险费"],
        "淘工厂-商家出资补贴":["正向扣款-商家出资补贴", "正向扣款-买赠赠品营销费用", "逆向退款-买赠赠品营销费用", "逆向退款-商家出资补贴"],
        "淘工厂-托管费":["商品运营托管推广服务费"],
        "淘工厂-托管账户充值":["扣款用途：账户充值-工作台充值", "扣款用途：账户充值-自动充值"],
        "淘工厂-先用后付":["C2M-先用后付技术服务费"],
        "天猫-先用后付服务费":["天猫-先用后付服务费", "先用后付技术服务费"],
        "淘工厂-运费险":["退货包运费代扣"],
        "网商贷-放款":["网商贷-放款"],
        "网商贷-还款":["网商贷-还款", "网商银行扣款"],
        "天猫-一大包零食推广费":["天猫一大包零食广告套餐费用"],
        "天猫-千橙食品店ERP接口费": ["代扣款（扣款用途：按订单收费服务费"],
        "淘宝买菜-保证金履约险":["淘宝买菜保证金履约险_保费"],
        "淘宝买菜-新品套餐货款抵扣":["扣款用途：新品套餐货款抵扣"],
        "淘宝买菜-退货包运费":["代扣款（扣款用途：退货包运费服务费"],
        "淘宝买菜-商家处罚":["代扣款（扣款用途：商家处罚"],
        "淘宝买菜-佣金":["代扣款（扣款用途：佣金"],
        "淘宝买菜-交易收款":["货款{"]
    }

    # 函数来根据关键字确定分类

    # 函数来根据关键字确定分类
    def assign_category(row):
        # 从各个相关列获取字符串
        item_name = str(row['商品名称']).strip()  # 请确认“商品名称”列名
        business_type = str(row['业务类型']).strip()  # 请确认“业务类型”列名
        remark = str(row['备注']).strip()  # 请确认“备注”列名
        counterparty = str(row['对方账号']).strip()  # 请确认“对方账号”列名
        pay_amount = row['支出金额（-元）']; # 支出金额（-元）列名
        income_amount = row['收入金额（+元）']; # 收入金额
        
        # "天猫-保证金充值":["天猫消费者保证金-充值（代扣）", "天猫消费者保证金-充值（代扣）"],
        if remark == "天猫保证金-充值（代扣）" or remark == "天猫消费者保证金-充值":
            return "天猫-保证金充值";
        
        if "万相台无界版扫码充值" in item_name:
            return "淘工厂-直通车充值";
        elif "门道商家助手-基础版-订单付款" in item_name:
            return "门道商家助手-基础版-订单付款";
        
        # 根据业务类型判断
        if business_type == "交易付款" or "基金代发任务" in remark:
            return "天猫-交易收款";
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
        elif counterparty == "杭州昌诚电子商务有限公司(ydbbzj@service.aliyun.com)" and business_type == "转账":
            return "一大包零食交保证金";
        elif counterparty == "杭州淘宝直播严选电子商务有限公司(qdzfb@service.aliyun.com)" and business_type == "转账":
            return "88VIP货款";
        elif counterparty == "**婧(150******97)" and remark == "转账" and pay_amount < 0:
            return "大C店-转出";
        elif counterparty == "**婧(150******97)" and remark == "转账" and income_amount > 0:
            return "大C店-转入";
        
        if "DDD商家结算款" in remark and "扣款用途" not in remark:
            return "天猫超市-交易收款";
        
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

    # 创建新的工作表 “整理” “透视”
    with pd.ExcelWriter(new_file_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='整理', index=False)
        pivot_table.to_excel(writer, sheet_name='透视', index=False)

    print(f"处理完成！新文件生成: {new_file_path}")

# 使用示例
# if __name__ == "__main__":
#     current_dir = os.getcwd()

#     # 读取源Excel文件
#     file_path = os.path.join(current_dir, '天猫-千橙食品店支付宝-10月_副本.xlsx')
#     process_tmall_file(file_path)