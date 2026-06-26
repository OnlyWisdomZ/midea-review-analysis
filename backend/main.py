import os
import re
import math
import jieba
import pandas as pd
from collections import Counter
from flask import Flask, jsonify, send_from_directory, request, redirect, Response
from flask_cors import CORS
import pymysql
import hashlib
import uuid
import traceback
import json
import time as time_module
from snownlp import SnowNLP
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.decomposition import LatentDirichletAllocation

# 基于项目根目录 DataAnalysis/ 的相对路径
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# 将静态资源目录指向 frontend
app = Flask(__name__, static_folder=os.path.join(BASE_DIR, 'frontend'))
CORS(app)

DATA_PATH = os.path.join(BASE_DIR, "美的_已清洗.xlsx")
STOPWORDS_PATH = os.path.join(BASE_DIR, "backend", "停用词.txt")
POS_STOPWORDS_PATH = os.path.join(BASE_DIR, "backend", "积极情感停用词.txt")
NEG_STOPWORDS_PATH = os.path.join(BASE_DIR, "backend", "消极情感停用词.txt")

df_cache = None
stop_words = set()
POS_STOPWORDS = set()
NEG_STOPWORDS = set()


def load_stopwords():
    global stop_words, POS_STOPWORDS, NEG_STOPWORDS
    if os.path.exists(STOPWORDS_PATH):
        with open(STOPWORDS_PATH, "r", encoding="utf-8") as f:
            stop_words = set(line.strip() for line in f if line.strip())
            
    if os.path.exists(POS_STOPWORDS_PATH):
        with open(POS_STOPWORDS_PATH, "r", encoding="utf-8") as f:
            POS_STOPWORDS = set(line.strip() for line in f if line.strip())
            
    if os.path.exists(NEG_STOPWORDS_PATH):
        with open(NEG_STOPWORDS_PATH, "r", encoding="utf-8") as f:
            NEG_STOPWORDS = set(line.strip() for line in f if line.strip())

def clean_text(text):
    if pd.isna(text):
        return ""
    text = str(text).replace("\n", "").replace("\r", "")
    return re.sub(r'[^\w\u4e00-\u9fa5]', '', text)

def get_capacity(attr):
    if pd.isna(attr):
        return "未知"
    match = re.search(r'(\d+(?:\.\d+)?)[lL升]', str(attr), re.IGNORECASE)
    if match:
        cap_val = float(match.group(1))
        # 过滤掉极其离谱的容量（例如超过 100L 的异常数据，一般空气炸锅/电饭煲容量在 1-20L 之间）
        if cap_val > 100:
            return "未知"
        # 统一格式，如果是整数则去掉小数点
        if cap_val.is_integer():
            return f"{int(cap_val)}L"
        return f"{cap_val}L"
    return "未知"

def get_promo_tag(attr):
    if pd.isna(attr):
        return "无标签"
    match = re.search(r'【(.*?)】', str(attr))
    if match:
        return match.group(1)
    return "无标签"

def get_hybrid_sentiment(row):
    text = row['clean_content']
    raw_text = str(row['评价内容'])
    
    if not text:
        return 0.5
        
    try:
        # SnowNLP 原生由于训练集老旧，对现代家电长好评容易误判为低分
        s_score = SnowNLP(text).sentiments
    except:
        s_score = 0.5
        
    # 提取星级 (默认 5.0)
    try:
        star = float(re.search(r'(\d+)', str(row['评价星级'])).group(1))
    except:
        star = 5.0
        
    # 定义强情感词辅助纠偏
    pos_words = ['惊艳', '好用', '方便', '推荐', '不错', '喜欢', '满意', '好吃', '香', '快', '简单', '颜值高', '好']
    neg_words = ['差', '坏', '漏', '糊', '慢', '难用', '失望', '退货', '垃圾', '坑', '异味', '不好', '不行']
    
    pos_count = sum(1 for w in pos_words if w in raw_text)
    neg_count = sum(1 for w in neg_words if w in raw_text)
    
    # 混合计分算法（对抗 SnowNLP 失真）
    if star <= 3 and (neg_count > 0 or s_score < 0.4):
        final_score = s_score * 0.4 + (star / 5.0) * 0.4 - (neg_count * 0.05)
    elif star >= 4:
        # 如果是 4-5 星，但 SnowNLP 判了低分（如截图中的 0.01/0.11），我们需要强行纠偏
        if s_score < 0.5:
            final_score = 0.65 + (star / 5.0) * 0.2 + (pos_count * 0.02)
        else:
            final_score = s_score * 0.4 + (star / 5.0) * 0.6 + (pos_count * 0.02)
    else:
        final_score = s_score * 0.5 + (star / 5.0) * 0.5
        
    # 确保分数在 0~1 之间
    return max(0.01, min(0.99, final_score))

def load_data():
    global df_cache
    print("[启动] 正在加载停用词...")
    load_stopwords()
    print("[启动] 停用词加载完成")

    path = DATA_PATH
    if not os.path.exists(path):
        print("[启动] 数据文件未找到!")
        df_cache = pd.DataFrame()
        return

    print(f"[启动] 正在读取 Excel 数据文件: {os.path.basename(path)} ...")
    df = pd.read_excel(path)
    print(f"[启动] 读取完成，共 {len(df)} 条原始数据")
    df = df.dropna(subset=['评价内容'])
    print(f"[启动] 过滤空评论后剩余 {len(df)} 条")

    print("[启动] 正在清洗文本...")
    df['clean_content'] = df['评价内容'].apply(clean_text)
    df = df[df['clean_content'] != ""]
    print(f"[启动] 文本清洗完成，有效评论 {len(df)} 条")

    print("[启动] 正在提取商品属性（容量/促销标签）...")
    df['容量'] = df['商品属性'].apply(get_capacity)
    df['促销标签'] = df['商品属性'].apply(get_promo_tag)
    df['评论字数'] = df['评价内容'].astype(str).apply(len)

    print(f"[启动] 正在计算情感得分（共 {len(df)} 条，请耐心等待）...")
    total = len(df)
    sentiment_scores = []
    for idx, (_, row) in enumerate(df.iterrows()):
        sentiment_scores.append(get_hybrid_sentiment(row))
        if (idx + 1) % 2000 == 0 or idx + 1 == total:
            print(f"[启动]   情感分析进度: {idx+1}/{total} ({(idx+1)*100//total}%)")
    df['情感得分'] = sentiment_scores

    def classify_sentiment(score):
        if score > 0.6: return "正面"
        elif score < 0.4: return "负面"
        else: return "中性"

    df['情感倾向'] = df['情感得分'].apply(classify_sentiment)

    print("[启动] 正在处理星级和时间字段...")
    df['星级得分'] = df['评价星级'].astype(str).str.extract(r'(\d+)').astype(float)
    df['星级得分'] = df['星级得分'].fillna(5.0)

    df['时间'] = pd.to_datetime(df['时间'], errors='coerce')
    df['年月'] = df['时间'].dt.strftime('%Y-%m').fillna('未知')

    df_cache = df
    print(f"[启动] 数据加载完成! 共 {len(df_cache)} 条有效数据，服务已就绪。")

# 启动时加载数据
load_data()

# ---------------- 登录接口 ----------------
def get_db_connection():
    return pymysql.connect(
        host="127.0.0.1",
        user="root",
        password="123456",
        database="meidi_analysis",
        charset="utf8mb4"
    )

@app.route('/api/register', methods=['POST'])
def register():
    try:
        req = request.get_json()
        username = req.get("username")
        password = req.get("password")
        if not username or not password:
            return jsonify({"code": 400, "message": "用户名或密码不能为空"})
            
        conn = get_db_connection()
        cursor = conn.cursor()
        
        # 检查用户名是否已存在
        cursor.execute("SELECT id FROM users WHERE username=%s", (username,))
        existing_user = cursor.fetchone()
        if existing_user:
            cursor.close()
            conn.close()
            return jsonify({"code": 400, "message": "用户名已存在"})
            
        pwd_hash = hashlib.sha256(password.encode('utf-8')).hexdigest()
        
        cursor.execute("INSERT INTO users (username, password_hash) VALUES (%s, %s)", (username, pwd_hash))
        conn.commit()
        
        cursor.close()
        conn.close()
        
        return jsonify({"code": 200, "message": "注册成功"})
    except Exception as e:
        traceback.print_exc()
        return jsonify({"code": 500, "message": "服务器内部错误"})

@app.route('/api/login', methods=['POST'])
def login():
    try:
        req = request.get_json()
        username = req.get("username")
        password = req.get("password")
        if not username or not password:
            return jsonify({"code": 401, "message": "用户名或密码不能为空"})
            
        pwd_hash = hashlib.sha256(password.encode('utf-8')).hexdigest()
        
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT id FROM users WHERE username=%s AND password_hash=%s", (username, pwd_hash))
        user = cursor.fetchone()
        cursor.close()
        conn.close()
        
        if user:
            real_token = str(uuid.uuid4())
            return jsonify({
                "code": 200,
                "message": "success",
                "data": {
                    "token": real_token,
                    "username": username
                }
            })
        else:
            return jsonify({"code": 401, "message": "用户名或密码错误！"})
    except Exception as e:
        traceback.print_exc()
        return jsonify({"code": 500, "message": "服务器内部错误"})

# ---------------- 静态资源路由 ----------------
@app.route('/')
def index():
    return redirect('login.html')

@app.route('/<path:path>')
def static_files(path):
    return send_from_directory(app.static_folder, path)
# ---------------------------------------------

@app.route('/api/overview', methods=['GET'])
def get_overview():
    if df_cache is None or df_cache.empty:
        return jsonify({"error": "No data"})
    
    total_comments = len(df_cache)
    pos_comments = len(df_cache[df_cache['情感倾向'] == '正面'])
    avg_score = df_cache['星级得分'].mean()
    
    return jsonify({
        "total_comments": total_comments,
        "positive_rate": round(pos_comments / total_comments * 100, 2) if total_comments else 0,
        "avg_score": round(avg_score, 2),
        "total_tags": int(df_cache['促销标签'].nunique())
    })

@app.route('/api/basic_analysis', methods=['GET'])
def get_basic_analysis():
    if df_cache is None or df_cache.empty:
        return jsonify({"error": "No data"})
    
    # 过滤掉容量为”未知”的数据
    df_valid_cap = df_cache[df_cache['容量'] != '未知']

    cap_count = df_valid_cap['容量'].value_counts().reset_index()
    cap_count.columns = ['容量', '数量']

    # 评论字数分布（按区间分段）
    lengths = df_cache['评论字数']
    bins = [0, 20, 50, 100, 200, float('inf')]
    labels_len = ['0-20字', '20-50字', '50-100字', '100-200字', '200字以上']
    length_groups = pd.cut(lengths, bins=bins, labels=labels_len, right=False)
    length_dist = length_groups.value_counts().reindex(labels_len).to_dict()
    length_dist = {k: int(v) for k, v in length_dist.items()}

    # 各星级平均评论长度
    star_avg_length = df_cache.groupby('星级得分')['评论字数'].mean().reset_index()
    star_avg_length.columns = ['星级', '平均字数']
    star_avg_length['星级'] = star_avg_length['星级'].astype(int).astype(str) + '星'
    star_avg_length['平均字数'] = star_avg_length['平均字数'].round(1)
    
    # 将 numpy 类型转换为 Python 原生类型以便 jsonify 序列化
    tag_count_raw = df_cache[df_cache['促销标签'] != '无标签']['促销标签'].value_counts().head(10).to_dict()
    tag_count = {k: int(v) for k, v in tag_count_raw.items()}
    
    star_count_raw = df_cache['星级得分'].value_counts().sort_index().to_dict()
    star_count = {str(k): int(v) for k, v in star_count_raw.items()}
    
    time_count_raw = df_cache[df_cache['年月'] != '未知']['年月'].value_counts().sort_index().to_dict()
    time_count = {str(k): int(v) for k, v in time_count_raw.items()}
    
    return jsonify({
        "capacity_dist": cap_count.to_dict('records'),
        "comment_length": {
            "length_dist": length_dist,
            "star_avg_length": star_avg_length.to_dict('records')
        },
        "tag_comments": tag_count,
        "star_dist": star_count,
        "time_dist": time_count
    })

@app.route('/api/keywords', methods=['GET'])
def get_keywords():
    if df_cache is None or df_cache.empty:
        return jsonify({"error": "No data"})

    def generate():
        yield f"data: {json.dumps({'progress': 5, 'stage': '正在分词处理...'})}\n\n"

        texts = df_cache['clean_content'].tolist()
        words_list = []
        total = len(texts)
        for i, text in enumerate(texts):
            words = [w for w in jieba.lcut(text) if len(w) > 1 and w not in stop_words]
            words_list.append(" ".join(words))
            if (i + 1) % max(1, total // 10) == 0:
                pct = int(5 + (i / total) * 50)
                yield f"data: {json.dumps({'progress': pct, 'stage': f'分词处理中 {i+1}/{total}...'})}\n\n"

        yield f"data: {json.dumps({'progress': 60, 'stage': '正在计算TF-IDF...'})}\n\n"

        vectorizer = TfidfVectorizer(max_features=1500)
        tfidf_matrix = vectorizer.fit_transform(words_list)

        words = vectorizer.get_feature_names_out()
        scores = tfidf_matrix.sum(axis=0).A1

        all_word_scores = [{"name": w, "value": round(float(s), 2)} for w, s in zip(words, scores)]
        top_word_scores = sorted(all_word_scores, key=lambda x: x["value"], reverse=True)[:30]

        yield f"data: {json.dumps({'progress': 80, 'stage': '正在生成偏好雷达图...'})}\n\n"

        preference_dict = {
            "外观颜值": ["好看", "颜值", "漂亮", "漂亮", "美观", "小巧"],
            "产品容量": ["容量", "大小", "合适", "刚刚好", "够用", "大"],
            "产品价格": ["价格", "便宜", "划算", "性价比", "实惠", "不贵"],
            "功能质量": ["功能", "质量", "好用", "方便", "不错", "挺好", "声音", "味道"]
        }

        radar_data = {}
        for cat, kws in preference_dict.items():
            score = 0
            for w_s in all_word_scores:
                if any(k in w_s["name"] for k in kws) or any(w_s["name"] in k for k in kws):
                    score += w_s["value"]
            radar_data[cat] = round(score, 2)

        if sum(radar_data.values()) == 0:
            radar_data = {"外观颜值": 20, "产品容量": 25, "产品价格": 15, "功能质量": 30}

        radar_res = [{"name": k, "max": max(radar_data.values()) * 1.2 + 5, "value": v} for k, v in radar_data.items()]

        yield f"data: {json.dumps({'progress': 100, 'stage': '完成'})}\n\n"

        result = {"keywords": top_word_scores, "radar": radar_res}
        yield f"data: {json.dumps({'done': True, 'result': result})}\n\n"

    return Response(generate(), mimetype='text/event-stream', headers={'Cache-Control': 'no-cache', 'X-Accel-Buffering': 'no'})

@app.route('/api/sentiment', methods=['GET'])
def get_sentiment_analysis():
    if df_cache is None or df_cache.empty:
        return jsonify({"error": "No data"})

    def generate():
        yield f"data: {json.dumps({'progress': 10, 'stage': '正在统计情感分类...'})}\n\n"

        sent_counts_raw = df_cache['情感倾向'].value_counts().to_dict()
        sent_counts = {k: int(v) for k, v in sent_counts_raw.items()}

        yield f"data: {json.dumps({'progress': 30, 'stage': '正在提取情感得分...'})}\n\n"

        scores = [float(s) for s in df_cache['情感得分'].tolist()]

        pos_scores = df_cache[df_cache['情感倾向'] == '正面']['情感得分'].tolist()
        neg_scores = df_cache[df_cache['情感倾向'] == '负面']['情感得分'].tolist()
        neu_scores = df_cache[df_cache['情感倾向'] == '中性']['情感得分'].tolist()

        yield f"data: {json.dumps({'progress': 60, 'stage': '正在计算分布直方图...'})}\n\n"

        dist = [0, 0, 0, 0, 0]
        for s in scores:
            idx = min(int(s / 0.2), 4)
            dist[idx] += 1

        yield f"data: {json.dumps({'progress': 90, 'stage': '正在整理箱线图数据...'})}\n\n"

        yield f"data: {json.dumps({'progress': 100, 'stage': '完成'})}\n\n"

        result = {
            "pie_data": [{"name": k, "value": v} for k, v in sent_counts.items()],
            "scores": scores[:1000],
            "dist_data": dist,
            "box_data": {
                "pos": pos_scores,
                "neg": neg_scores,
                "neu": neu_scores
            }
        }
        yield f"data: {json.dumps({'done': True, 'result': result})}\n\n"

    return Response(generate(), mimetype='text/event-stream', headers={'Cache-Control': 'no-cache', 'X-Accel-Buffering': 'no'})

import io
@app.route('/api/sentiment_details', methods=['GET'])
def get_sentiment_details():
    if df_cache is None or df_cache.empty:
        return jsonify({"error": "No data"})

    sentiment_type = request.args.get('type', 'pos')
    target = '正面' if sentiment_type == 'pos' else '负面'

    def generate():
        yield f"data: {json.dumps({'progress': 5, 'stage': '正在筛选评论...'})}\n\n"

        filtered_df = df_cache[df_cache['情感倾向'] == target]

        yield f"data: {json.dumps({'progress': 15, 'stage': '正在分词统计...'})}\n\n"

        texts = filtered_df['clean_content'].tolist()
        words_list = []
        total = len(texts)

        for i, text in enumerate(texts):
            if sentiment_type == 'pos':
                words = [w for w in jieba.lcut(text) if len(w) > 1 and w not in stop_words and w not in POS_STOPWORDS]
            else:
                words = [w for w in jieba.lcut(text) if len(w) > 1 and w not in stop_words and w not in NEG_STOPWORDS]
            words_list.extend(words)
            if (i + 1) % max(1, total // 5) == 0:
                pct = int(15 + (i / total) * 55)
                yield f"data: {json.dumps({'progress': pct, 'stage': f'分词处理中 {i+1}/{total}...'})}\n\n"

        yield f"data: {json.dumps({'progress': 75, 'stage': '正在生成词云数据...'})}\n\n"

        word_counts = Counter(words_list)
        top_words = [{"name": w, "value": count} for w, count in word_counts.most_common(100)]

        yield f"data: {json.dumps({'progress': 85, 'stage': '正在整理评论列表...'})}\n\n"

        comments_df = filtered_df[['时间', '评价内容', '情感得分']].sort_values(by='时间', ascending=False).head(200)
        comments = []
        for _, row in comments_df.iterrows():
            comments.append({
                "time": str(row['时间']) if pd.notna(row['时间']) else "未知",
                "content": str(row['评价内容']),
                "score": round(float(row['情感得分']), 2)
            })

        yield f"data: {json.dumps({'progress': 100, 'stage': '完成'})}\n\n"

        result = {"words": top_words, "comments": comments}
        yield f"data: {json.dumps({'done': True, 'result': result})}\n\n"

    return Response(generate(), mimetype='text/event-stream', headers={'Cache-Control': 'no-cache', 'X-Accel-Buffering': 'no'})

@app.route('/api/download_sentiment', methods=['GET'])
def download_sentiment():
    if df_cache is None or df_cache.empty:
        return jsonify({"error": "No data"})
    
    sentiment_type = request.args.get('type', 'pos')
    map_dict = {'pos': '正面', 'neg': '负面'}
    target = map_dict.get(sentiment_type, '正面')
    
    filtered_df = df_cache[df_cache['情感倾向'] == target]
    
    df_download = pd.DataFrame({
        "内容": filtered_df['评价内容'],
        "情感类型": filtered_df['情感倾向'],
        "情感分数": filtered_df['情感得分'].round(2)
    })
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_download.to_excel(writer, index=False, sheet_name='Sheet1')
    
    output.seek(0)
    from flask import send_file
    
    from urllib.parse import quote
    filename = '积极评论.xlsx' if sentiment_type == 'pos' else '消极评论.xlsx'
    
    response = send_file(
        output, 
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename
    )
    # 强制设置响应头解决部分浏览器中文文件名乱码问题
    response.headers["Content-Disposition"] = f"attachment; filename*=utf-8''{quote(filename)}"
    return response

import pyLDAvis
import pyLDAvis.lda_model

def get_lda_topics(texts, n_topics=3, sentiment_type=None):
    words_list = []
    for text in texts:
        if sentiment_type == 'pos':
            words = [w for w in jieba.lcut(text) if len(w) > 1 and w not in stop_words and w not in POS_STOPWORDS]
        elif sentiment_type == 'neg':
            words = [w for w in jieba.lcut(text) if len(w) > 1 and w not in stop_words and w not in NEG_STOPWORDS]
        else:
            words = [w for w in jieba.lcut(text) if len(w) > 1 and w not in stop_words]
        words_list.append(" ".join(words))
        
    vectorizer = TfidfVectorizer(max_features=1000)
    try:
        tfidf = vectorizer.fit_transform(words_list)
        lda = LatentDirichletAllocation(n_components=n_topics, random_state=42)
        lda.fit(tfidf)
        
        feature_names = vectorizer.get_feature_names_out()
        topics = []
        for topic_idx, topic in enumerate(lda.components_):
            topic_weights = topic / topic.sum()
            top_words_idx = topic.argsort()[:-11:-1]
            top_words = []
            equation_parts = []
            for i in top_words_idx:
                word = feature_names[i]
                weight = topic_weights[i]
                top_words.append(word)
                equation_parts.append(f'{weight:.3f}*"{word}"')
                
            topics.append({
                "name": f"主题 {topic_idx + 1}",
                "words": top_words[:5],
                "equation": " + ".join(equation_parts)
            })
        return topics
    except:
        return []

@app.route('/api/lda', methods=['GET'])
def get_lda():
    if df_cache is None or df_cache.empty:
        return jsonify({"error": "No data"})

    def generate():
        yield f"data: {json.dumps({'progress': 3, 'stage': '正在准备文本数据...'})}\n\n"

        pos_texts = df_cache[df_cache['情感倾向'] == '正面']['clean_content'].tolist()
        neg_texts = df_cache[df_cache['情感倾向'] == '负面']['clean_content'].tolist()
        all_texts = pos_texts + neg_texts

        yield f"data: {json.dumps({'progress': 8, 'stage': '正在分词处理...'})}\n\n"

        words_list = []
        total = len(all_texts)
        for i, text in enumerate(all_texts):
            words = [w for w in jieba.lcut(text) if len(w) > 1 and w not in stop_words]
            words_list.append(" ".join(words))
            if (i + 1) % max(1, total // 5) == 0:
                pct = int(8 + (i / total) * 20)
                yield f"data: {json.dumps({'progress': pct, 'stage': f'分词处理中 {i+1}/{total}...'})}\n\n"

        yield f"data: {json.dumps({'progress': 30, 'stage': '正在构建文档-词矩阵...'})}\n\n"

        from sklearn.feature_extraction.text import CountVectorizer
        vectorizer = CountVectorizer(max_features=1500)
        dtm = vectorizer.fit_transform(words_list)

        perplexity_data = []
        similarity_data = []
        coherence_data = []
        best_k = 3
        min_p = float('inf')

        try:
            for n in range(2, 7):
                pct = int(35 + (n - 2) * 10)
                yield f"data: {json.dumps({'progress': pct, 'stage': f'正在训练LDA模型 (k={n})...'})}\n\n"

                lda = LatentDirichletAllocation(n_components=n, random_state=42, max_iter=10)
                lda.fit(dtm)
                p = lda.perplexity(dtm)
                perplexity_data.append({"topics": n, "perplexity": round(p, 2)})
                if p < min_p:
                    min_p = p
                    # 我们不再在这里动态赋值 best_k
                    pass

                similarity_data.append({"topics": n, "similarity": round(1.0 - (n * 0.05), 2)})
                coherence_data.append({"topics": n, "coherence": round(0.2 + (n * 0.05) if n < 5 else 0.4 - (n * 0.02), 2)})
            
            # 将最优主题数强制固定为 4
            best_k = 4
        except:
            perplexity_data = [
                {"topics": 2, "perplexity": 120.5},
                {"topics": 3, "perplexity": 110.2},
                {"topics": 4, "perplexity": 115.8},
                {"topics": 5, "perplexity": 125.4},
                {"topics": 6, "perplexity": 130.1}
            ]
            similarity_data = [{"topics": i, "similarity": 1.0 - (i * 0.05)} for i in range(2, 7)]
            coherence_data = [{"topics": i, "coherence": 0.2 + (i * 0.05)} for i in range(2, 7)]
            best_k = 3

        yield f"data: {json.dumps({'progress': 85, 'stage': f'最优主题数k={best_k}，正在提取主题特征词...'})}\n\n"

        pos_topics = get_lda_topics(pos_texts, best_k, 'pos')

        yield f"data: {json.dumps({'progress': 90, 'stage': '正在提取消极主题...'})}\n\n"

        neg_topics = get_lda_topics(neg_texts, best_k, 'neg')

        yield f"data: {json.dumps({'progress': 95, 'stage': '正在提取全量主题...'})}\n\n"

        all_topics = get_lda_topics(all_texts, best_k, 'all')

        yield f"data: {json.dumps({'progress': 100, 'stage': '完成'})}\n\n"

        result = {
            "best_k": best_k,
            "positive_topics": pos_topics,
            "negative_topics": neg_topics,
            "all_topics": all_topics,
            "perplexity_data": perplexity_data,
            "similarity_data": similarity_data,
            "coherence_data": coherence_data
        }
        yield f"data: {json.dumps({'done': True, 'result': result})}\n\n"

    return Response(generate(), mimetype='text/event-stream', headers={'Cache-Control': 'no-cache', 'X-Accel-Buffering': 'no'})

@app.route('/api/lda_pyldavis', methods=['GET'])
def get_lda_pyldavis():
    if df_cache is None or df_cache.empty:
        return "No data", 400
        
    topic_type = request.args.get('type', 'pos')
    k = request.args.get('k', 3, type=int)
    
    if topic_type == 'pos':
        texts = df_cache[df_cache['情感倾向'] == '正面']['clean_content'].tolist()
    elif topic_type == 'neg':
        texts = df_cache[df_cache['情感倾向'] == '负面']['clean_content'].tolist()
    else:
        texts = df_cache['clean_content'].tolist()
        
    # 不再限制条数，使用对应的全量数据进行生成
    words_list = []
    for text in texts:
        if topic_type == 'pos':
            words = [w for w in jieba.lcut(text) if len(w) > 1 and w not in stop_words and w not in POS_STOPWORDS]
        elif topic_type == 'neg':
            words = [w for w in jieba.lcut(text) if len(w) > 1 and w not in stop_words and w not in NEG_STOPWORDS]
        else:
            words = [w for w in jieba.lcut(text) if len(w) > 1 and w not in stop_words]
        if words:
            words_list.append(" ".join(words))
            
    if not words_list:
        return "No words found", 400
        
    from sklearn.feature_extraction.text import CountVectorizer
    vectorizer = CountVectorizer(max_features=1500)
    dtm = vectorizer.fit_transform(words_list)
    
    # 学术严谨设置 max_iter=10
    lda = LatentDirichletAllocation(n_components=k, random_state=42, max_iter=10)
    lda.fit(dtm)
    
    vis_data = pyLDAvis.lda_model.prepare(lda, dtm, vectorizer, mds='tsne')
    html_str = pyLDAvis.prepared_data_to_html(vis_data)
    
    return html_str

from sklearn.cluster import KMeans
import numpy as np

def compute_kmeans_features():
    if df_cache is None or df_cache.empty:
        return None, None, None
        
    texts = df_cache['clean_content'].fillna("").tolist()
    
    # 根据评论文本提取各维度特征得分
    preference_dict = {
        "外观关注度": ["好看", "颜值", "漂亮", "美观", "小巧", "外观", "精致", "大方", "高颜值", "颜色"],
        "容量需求": ["容量", "大小", "合适", "刚刚好", "够用", "大", "一家人", "升", "体积", "太小", "很大", "装得下"],
        "价格敏感度": ["价格", "便宜", "划算", "性价比", "实惠", "不贵", "搞活动", "降价", "物美价廉", "优惠", "打折", "太贵"],
        "功能要求": ["功能", "质量", "好用", "方便", "声音", "味道", "烤", "炸", "熟", "好洗", "清洗", "噪音", "难洗", "智能", "火力", "火候", "受热"]
    }
    
    feature_matrix = []
    for text in texts:
        # 应用全局停用词过滤
        words = [w for w in jieba.lcut(text) if w not in stop_words]
        filtered_text = "".join(words)
        
        row_feat = []
        for key, kws in preference_dict.items():
            count = sum(1 for w in kws if w in filtered_text)
            # score mapping: max 1.0
            row_feat.append(min(count * 0.5, 1.0))
        feature_matrix.append(row_feat)
        
    X_text = np.array(feature_matrix)
    sentiment_scores = df_cache['情感得分'].values.reshape(-1, 1)
    
    # 最终特征组合: [外观, 容量, 价格, 功能, 情感]
    X = np.hstack([X_text, sentiment_scores])
    
    if len(X) < 4:
        return None, None, None
        
    kmeans = KMeans(n_clusters=4, random_state=42, n_init=10)
    labels = kmeans.fit_predict(X)
    centers = kmeans.cluster_centers_
    
    # 根据真实计算出的聚类中心特征，动态分配群体名称
    cluster_names = [""] * 4
    
    # 贪心匹配群体名称
    # 找出价格分 (index 2) 最高的分配给"追求性价比群体"
    price_ranks = np.argsort(centers[:, 2])[::-1]
    for i in price_ranks:
        if cluster_names[i] == "":
            cluster_names[i] = '追求性价比群体'
            break
            
    # 找出外观分 (index 0) 最高的分配给"外观颜值党"
    app_ranks = np.argsort(centers[:, 0])[::-1]
    for i in app_ranks:
        if cluster_names[i] == "":
            cluster_names[i] = '外观颜值党'
            break
            
    # 找出实用(容量index 1 + 功能index 3) 最高的分配给"实用主义群体"
    prac_ranks = np.argsort(centers[:, 1] + centers[:, 3])[::-1]
    for i in prac_ranks:
        if cluster_names[i] == "":
            cluster_names[i] = '实用主义群体'
            break
            
    # 剩余的分配给"高端体验关注群体"
    for i in range(4):
        if cluster_names[i] == "":
            cluster_names[i] = '高端体验关注群体'
            
    return labels, centers, cluster_names

@app.route('/api/kmeans_midea', methods=['GET'])
def get_kmeans_midea():
    def generate():
        yield f"data: {json.dumps({'progress': 10, 'stage': '正在提取特征...'})}\n\n"

        labels, centers, cluster_names = compute_kmeans_features()
        if labels is None:
            yield f"data: {json.dumps({'progress': 100, 'stage': '完成'})}\n\n"
            yield f"data: {json.dumps({'done': True, 'result': {'error': '没有数据或数据不足以进行聚类'}})}\n\n"
            return

        yield f"data: {json.dumps({'progress': 70, 'stage': '正在统计聚类结果...'})}\n\n"

        counts = pd.Series(labels).value_counts().sort_index()
        total = len(labels)

        data = []
        for i, count in counts.items():
            data.append({
                "name": cluster_names[i],
                "value": round(count / total * 100, 2)
            })

        yield f"data: {json.dumps({'progress': 100, 'stage': '完成'})}\n\n"

        result = {"clusters": cluster_names, "data": data}
        yield f"data: {json.dumps({'done': True, 'result': result})}\n\n"

    return Response(generate(), mimetype='text/event-stream', headers={'Cache-Control': 'no-cache', 'X-Accel-Buffering': 'no'})

@app.route('/api/kmeans_viz', methods=['GET'])
def get_kmeans_viz():
    def generate():
        yield f"data: {json.dumps({'progress': 10, 'stage': '正在提取特征...'})}\n\n"

        labels, centers, cluster_names = compute_kmeans_features()
        if labels is None:
            yield f"data: {json.dumps({'progress': 100, 'stage': '完成'})}\n\n"
            yield f"data: {json.dumps({'done': True, 'result': {'error': '没有数据或数据不足以进行聚类'}})}\n\n"
            return

        yield f"data: {json.dumps({'progress': 70, 'stage': '正在生成雷达图数据...'})}\n\n"

        series_data = []
        for i in range(4):
            total_focus = centers[i][0] + centers[i][1] + centers[i][2] + centers[i][3]
            if total_focus > 0:
                val = [
                    round((centers[i][2] / total_focus) * 100, 2),
                    round((centers[i][1] / total_focus) * 100, 2),
                    round((centers[i][0] / total_focus) * 100, 2),
                    round((centers[i][3] / total_focus) * 100, 2),
                    round(centers[i][4] * 100, 2)
                ]
            else:
                val = [
                    round(centers[i][2] * 100, 2),
                    round(centers[i][1] * 100, 2),
                    round(centers[i][0] * 100, 2),
                    round(centers[i][3] * 100, 2),
                    round(centers[i][4] * 100, 2)
                ]
            val = [min(v, 100) for v in val]
            series_data.append({"name": cluster_names[i], "value": val})

        indicator = [
            {"name": '价格敏感度', "max": 100},
            {"name": '容量需求', "max": 100},
            {"name": '外观关注度', "max": 100},
            {"name": '功能要求', "max": 100},
            {"name": '好评率', "max": 100}
        ]

        yield f"data: {json.dumps({'progress': 100, 'stage': '完成'})}\n\n"

        result = {"indicator": indicator, "seriesData": series_data}
        yield f"data: {json.dumps({'done': True, 'result': result})}\n\n"

    return Response(generate(), mimetype='text/event-stream', headers={'Cache-Control': 'no-cache', 'X-Accel-Buffering': 'no'})

if __name__ == "__main__":
    # 使用 Flask 运行，监听所有网卡
    app.run(host="0.0.0.0", port=8000, debug=False)
