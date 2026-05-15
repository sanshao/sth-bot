


### Key Changes:
- **Mappings Added**: 18 specific rules were implemented based on the provided "业务描述" (Business Description).
- **Consolidated Rules**: Categories with multiple source descriptions (like `基础技术服务费` and `售后补偿消费者`) were grouped using the `conditions` array for cleaner organization.
- **Field Matching**: All rules use the `contains` operator on the `业务描述` field, consistent with existing patterns in the project.

### Mapping Summary:
| Category (分类) | Business Description (业务描述) Keywords |
| :--- | :--- |
| **交易收款** | 交易收入-订单收入 |
| **交易收款-优惠券** | 交易收入-优惠券结算 |
| **交易退款** | 交易退款-订单退款 |
| **交易退款-优惠券** | 交易退款-优惠券结算 |
| **基础技术服务费** | 技术服务费-技术服务费, 技术服务费-基础技术服务费 |
| **百亿补贴技术服务费** | 技术服务费-百亿补贴技术服务费 |
| **售后补偿消费者** | 售后费用-小额打款, 售后费用-售后补偿消费者, 售后费用-欺诈发货 |
| **售后费用-运费补偿** | 售后费用-运费补偿 |
| **售后费用-延迟发货** | 售后费用-延迟发货 |
| **售后费用-虚假发货** | 售后费用-虚假发货 |
| **售后费用-缺货** | 售后费用-缺货 |
| **多多进宝** | 营销费用-多多进宝 |
| **推广充值** | 转账-广告账户 |
| **提现** | 提现-提现申请 |
| **交易收款-申诉** | 其他收入-申诉补回 |
