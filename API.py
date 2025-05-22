from cozepy import Coze, TokenAuth, Message, ChatEventType, COZE_CN_BASE_URL
import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed
from tqdm import tqdm

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

# 用于处理单个问题的函数
def process_question(index, question, image_id):
    answer_text = ""
    try:
        for event in coze.chat.stream(
            bot_id='',
            user_id='user_id',
            additional_messages=[Message.build_user_question_text(question)]
        ):
            if event.event == ChatEventType.CONVERSATION_MESSAGE_DELTA:
                answer_text += event.message.content
    except Exception as e:
        print(f"❌ 第 {index + 1} 条失败：{e}")
        answer_text = "[处理失败]"

    return {
        "index": index,  # 加上原始位置
        id_column: image_id,
        question_column: question,
        "智能体诊断结果": answer_text
    }

# 并发执行所有问题处理，这里线程数暂时设为5
with ThreadPoolExecutor(max_workers=5) as executor:
    futures = []
    for idx, row in df.iterrows():
        question = str(row[question_column]).strip()
        image_id = row[id_column]
        futures.append(executor.submit(process_question, idx, question, image_id))

    for future in tqdm(as_completed(futures), total=len(futures), desc="处理进度"):
        result = future.result()
        results.append(result)

# 排序
results.sort(key=lambda x: x["index"])
for r in results:
    del r["index"]

# 写入结果
result_df = pd.DataFrame(results)
result_df.to_excel(output_excel, index=False)
print(f"\n回答已保存到文件：{output_excel}")