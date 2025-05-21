from cozepy import Coze, TokenAuth, Message, ChatEventType, COZE_CN_BASE_URL
import pandas as pd

# 初始化 Coze 客户端
coze_api_token_input = input("请输入您的个人令牌（输入后按回车）：").strip()
coze_api_base = COZE_CN_BASE_URL
coze = Coze(auth=TokenAuth(coze_api_token_input), base_url=coze_api_base)
bot_id_input = input("请输入您的bot_id（输入后按回车）：").strip()

# Excel 配置
# 其中question_column是要进行处理的列，可以根据您的需要进行更改
input_excel = "./data.xlsx"
output_excel = "./answers.xlsx"
id_column = "影像号"
question_column = "报告中描述的整句话"

# 读取原始 Excel 文件
df = pd.read_excel(input_excel, header=2)  # 从第三行开始（前两行为说明，可以按需调整）
df = df[[question_column, id_column]].copy()  # 只保留“问题”和“影像号”列
df = df.dropna(subset=[question_column, id_column])  # 删除空值行

# 创建新的结果列表（用于生成新文档）
results = []

# 开始处理
for idx in range(0, 1):
    question = str(df.iloc[idx][question_column]).strip()
    print(f"\n处理第 {idx + 1} 条：{question}")
    answer_text = ""

    for event in coze.chat.stream(
            bot_id=bot_id_input,
            user_id='user_id',
            additional_messages=[Message.build_user_question_text(question)]
    ):

        if event.event == ChatEventType.CONVERSATION_MESSAGE_DELTA:
            delta = event.message.content
            print(delta, end="")
            answer_text += delta

    # 添加到结果列表
    results.append({
        "影像号": df.iloc[idx]["影像号"],
        question_column: question,
        "智能体诊断结果": answer_text
    })

# 将结果写入新的 Excel 文件
result_df = pd.DataFrame(results)
result_df.to_excel(output_excel, index=False)
print(f"\n回答已保存到文件：{output_excel}")