# ⭐Clarity⭐ 梦洁姐采购成本实时看板

## 启动
```bash
cd /root/.hermes/profiles/mengjie/workspace/cost_system
streamlit run app.py --server.address 0.0.0.0 --server.port 8501
```

## 功能
- 填今日营业额后，自动计算 40% 红线。
- 采购逐条录入：食材、数量、单位、批次。
- 自动查报价库，按同品项贵价保守计算。
- 支持“包→件”换算：默认 10 包/件；如果规格写明 8包/件、20包/件，会按规格换算。
- 库里没有或单位不确定时，标待确认，不硬猜。
- 海鲜/临时采购可填“手填整项金额”。
- 自动汇总：已确认采购成本、营业额、实际成本率、是否超过 40%。

## 数据文件
- 采购记录：data/orders.csv
- 营业额：data/sales.csv
- 报价库：优先读取 /root/.hermes/profiles/mengjie/workspace/pricing/ 下的 v6/v5 xlsx。
