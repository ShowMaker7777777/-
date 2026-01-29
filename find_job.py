from wxauto import WeChat
import pandas as pd
import time
import os
from datetime import datetime

# ================= 配置区域 =================
# 1. 监听列表：请确保这些群在你的PC微信里是【置顶】状态
LISTEN_LIST = ["work1", "work2", "work3", "work4", "work5"] 

# 2. 必选词（白名单）：地域筛选
REQUIRED_KEYWORDS = ["Place1", "Place2", "Place3", "Place4", "Place5", "Place6", "Place7", "Place8", "Place9", "Place10"]

# 3. 拒绝词（黑名单）：性别筛选
BLACK_KEYWORDS = ["1", "2", "3", "4", "5", "6", "7"]

# 4. 复活词（权重最高）：
RESURRECT_KEYWORDS = ["A", "B", "C", "D"]

# 5. 结果保存的文件名
RESULT_FILE = "name.xlsx"
# ===========================================

def get_keywords_status(content):
    """
    判断一条消息是否符合要求
    返回: True (符合/保留), False (不符合/丢弃)
    """
    # 确保内容是字符串
    if not isinstance(content, str):
        return False
        
    # 1. 垃圾信息过滤
    if len(content) < 10: 
        return False

    # 2. 地域筛选
    if not any(place in content for place in REQUIRED_KEYWORDS):
        return False

    # 3. 性别逻辑筛选
    has_black = any(word in content for word in BLACK_KEYWORDS)
    has_resurrect = any(word in content for word in RESURRECT_KEYWORDS)

    if has_black and not has_resurrect:
        return False
    
    return True

def save_to_excel(data_list):
    """
    把抓取到的数据存入 Excel
    """
    df = pd.DataFrame(data_list)
    if not os.path.exists(RESULT_FILE):
        df.to_excel(RESULT_FILE, index=False)
    else:
        with pd.ExcelWriter(RESULT_FILE, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
            try:
                reader = pd.read_excel(RESULT_FILE)
                start_row = len(reader) + 1
            except:
                start_row = 0
            df.to_excel(writer, index=False, header=False, startrow=start_row)
    
    print(f"✅ 已保存 {len(data_list)} 条新单子到表格：{RESULT_FILE}")

def parse_message(msg):
    """
    专门处理新旧版本的消息解析
    返回: (content, sender) 或者 (None, None)
    """
    content = None
    sender = None

    # 情况A：旧版本 (是列表或元组)
    if isinstance(msg, (list, tuple)):
        if len(msg) >= 2:
            sender = msg[0]
            content = msg[1]
    
    # 情况B：新版本 (是对象)
    elif hasattr(msg, 'content'):
        # 排除掉 TimeMessage (时间标签) 和 SystemMessage (系统消息)
        # 如果对象的类名包含 'Time' 或 'System'，通常不是有效聊天
        msg_type = str(type(msg))
        if 'TimeMessage' in msg_type or 'SystemMessage' in msg_type:
            return None, None
            
        content = msg.content
        # 尝试获取发送者，有些对象可能叫 sender
        if hasattr(msg, 'sender'):
            sender = msg.sender
        else:
            sender = "未知发送者"

    # 如果内容不是字符串（比如是图片对象），忽略
    if not isinstance(content, str):
        return None, None
        
    return content, sender

def main():
    try:
        wx = WeChat()
    except Exception as e:
        print(f"❌ 无法连接微信。详细错误信息: {e}")
        return

    print("🚀 监控程序已启动！正在扫描置顶群聊...")
    print(f"📂 筛选结果将保存在：{os.path.abspath(RESULT_FILE)}")
    
    processed_msgs = set()

    while True:
        try:
            # 获取会话列表
            sessions = wx.GetSession()
            new_jobs = []
            
            for session in sessions:
                # 提取会话名称
                if hasattr(session, 'name'):
                    chat_name = session.name
                else:
                    chat_name = str(session)

                if any(keyword in chat_name for keyword in LISTEN_LIST):
                    
                    wx.ChatWith(chat_name) 
                    msgs = wx.GetAllMessage()[-5:] 
                    
                    for msg in msgs:
                        # 【核心修改】使用专门的解析函数
                        content, sender = parse_message(msg)
                        
                        # 如果没解析出内容（比如是时间标签），跳过
                        if not content:
                            continue

                        # 去重
                        if content in processed_msgs:
                            continue
                        
                        processed_msgs.add(content)
                        
                        # 筛选
                        if get_keywords_status(content):
                            print(f"👀 [{chat_name}] 发现目标：{content[:15]}...")
                            new_jobs.append({
                                "抓取时间": datetime.now().strftime("%H:%M:%S"),
                                "来源群": chat_name,
                                "发送者": sender, # 把发送者也记下来
                                "内容": content
                            })
            
            if new_jobs:
                save_to_excel(new_jobs)
                
            time.sleep(5)
            
        except KeyboardInterrupt:
            print("\n🛑 程序已停止")
            break
        except Exception as e:
            # 这里的报错大部分是版本兼容问题，打印出来方便排查
            print(f"⚠️ 扫描中遇到小问题 (自动忽略): {e}")
            time.sleep(5)

if __name__ == "__main__":

    main()
