import os
from cozepy import Coze, TokenAuth, Message, ChatEventType, COZE_CN_BASE_URL
import pandas as pd

# 将API_TOKEN设为环境变量
os.environ["COZE_API_TOKEN"] = ""

# 初始化 Coze 客户端
coze_api_token = os.getenv("COZE_API_TOKEN")
coze_api_base = COZE_CN_BASE_URL
coze = Coze(auth=TokenAuth(coze_api_token), base_url=coze_api_base)

# Excel 配置
# 其中question_column是要进行处理的列，可以根据您的需要进行更改
input_excel = "./data.xlsx"
output_excel = "./answers.xlsx"
question_column = "报告中描述的整句话"

# 读取原始 Excel 文件
df = pd.read_excel(input_excel, header=2)  # 从第三行开始（前两行为说明，可以按需调整）
df = df[[question_column]].copy()  # 只保留“问题”列
df = df.dropna(subset=[question_column])  # 删除空值行

# 创建新的结果列表（用于生成新 DataFrame）
results = []

# 批处理参数
batch_size = 10

# 开始分批处理
for idx in range(0, len(df)):
    question = str(df.iloc[idx][question_column]).strip()
    print(f"\n处理第 {idx + 1} 条：{question}")
    answer_text = ""

    # The return values of the streaming interface can be iterated immediately.
    for event in coze.chat.stream(
            # 这里写您的bot_id
            bot_id='',
            user_id='user_id',
            additional_messages=[Message.build_user_question_text(question)]
    ):

        if event.event == ChatEventType.CONVERSATION_MESSAGE_DELTA:
            print(event.message.content, end="")

    # 添加到结果列表
    results.append({
        question_column: question,
        "智能体诊断结果": answer_text
    })

# 将结果写入新的 Excel 文件
result_df = pd.DataFrame(results)
result_df.to_excel(output_excel, index=False)
print(f"\n回答已保存到文件：{output_excel}")