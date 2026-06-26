import re
import json
import time
import random
import pandas as pd
from DrissionPage import ChromiumPage, ChromiumOptions


class JDCommentCrawler:
    def __init__(self, product_url):
        self.product_url = product_url
        self.product_id = self.extract_product_id(product_url)
        self.page = None

    @staticmethod
    def extract_product_id(url):
        match = re.search(r"/(\d+)\.html", url)
        if not match:
            raise ValueError("无法从 URL 中提取商品 ID")
        return match.group(1)

    def start_browser(self):
        """连接已开启 debug 模式的 Chrome（端口 9222）"""
        co = ChromiumOptions()
        co.set_local_port(9222)
        self.page = ChromiumPage(co)

    def get_one_page(self, page_num=0, score=0, page_size=10, sort_type=5):
        """通过浏览器发起评论 API 请求"""
        ts = int(time.time() * 1000)
        body = json.dumps({
            "productId": self.product_id,
            "score": score,
            "sortType": sort_type,
            "page": page_num,
            "pageSize": page_size,
            "isShadowSku": 0,
            "fold": 1,
        }, separators=(",", ":"))

        # 先访问商品页（首次），让浏览器获得正确的 cookie 和上下文
        if page_num == 0:
            self.page.get(self.product_url)
            time.sleep(3)

        # 通过浏览器执行 JS fetch 来请求评论接口（自动带上 cookie 和签名）
        api_url = (
            f"https://api.m.jd.com/api?"
            f"appid=item-v3&functionId=pc_club_productPageComments"
            f"&client=pc&clientVersion=1.0.0&t={ts}"
            f"&body={body}"
        )

        # 用浏览器内置的 fetch 发请求，绕过 h5st 限制
        js_code = f'''
        return await fetch("{api_url}", {{
            method: "GET",
            credentials: "include"
        }}).then(r => r.text())
        '''
        try:
            text = self.page.run_js(js_code)
        except Exception as e:
            print(f"JS fetch 失败: {e}")
            return []

        if not text or len(text) < 20:
            print(f"响应过短: {text[:100] if text else '空'}")
            return []

        try:
            data = json.loads(text)
        except json.JSONDecodeError:
            print(f"JSON 解析失败，前200字符: {text[:200]}")
            return []

        if "comments" not in data:
            if "login" in str(data).lower():
                raise ValueError("需要登录，请在浏览器中先登录京东账号")
            print(f"警告: 无 comments 字段, keys={list(data.keys())}")
        return data.get("comments", [])

    def crawl(self, max_pages=50, score=0, page_size=10, min_delay=5.0, max_delay=10.0):
        all_rows = []
        for page_num in range(max_pages):
            print(f"正在抓取第 {page_num + 1}/{max_pages} 页...")
            try:
                comments = self.get_one_page(page_num=page_num, score=score, page_size=page_size)
            except Exception as e:
                print(f"第 {page_num + 1} 页失败: {e}")
                break

            if not comments:
                print("没有更多评论，停止抓取")
                break

            for item in comments:
                row = {
                    "商品ID": self.product_id,
                    "用户昵称": item.get("nickname", ""),
                    "评分": item.get("score", ""),
                    "评论内容": item.get("content", "").replace("\n", " ").strip(),
                    "评论时间": item.get("creationTime", ""),
                    "点赞数": item.get("usefulVoteCount", ""),
                    "回复数": item.get("replyCount", ""),
                    "用户级别": item.get("userLevelName", ""),
                    "商品颜色": item.get("productColor", ""),
                    "商品规格": item.get("productSize", ""),
                    "购买时间": item.get("referenceTime", ""),
                    "订单商品名称": item.get("referenceName", ""),
                }
                all_rows.append(row)

            print(f"本页获取 {len(comments)} 条评论")
            sleep_time = random.uniform(min_delay, max_delay)
            print(f"休眠 {sleep_time:.1f}s...")
            time.sleep(sleep_time)

        return all_rows

    @staticmethod
    def save_csv(rows, filename="jd_comments.csv"):
        if not rows:
            print("没有抓到评论")
            return
        df = pd.DataFrame(rows)
        df.drop_duplicates(subset=["用户昵称", "评论内容", "评论时间"], inplace=True)
        df.to_csv(filename, index=False, encoding="utf-8-sig")
        print(f"已保存 {len(df)} 条评论到 {filename}")

    def close(self):
        if self.page:
            self.page.quit()


if __name__ == "__main__":
    urls = [
        "https://item.jd.com/100327958810.html",
        "https://item.jd.com/100258480135.html",
        "https://item.jd.com/100245768279.html",
    ]

    all_comments = []
    crawler = JDCommentCrawler(urls[0])
    crawler.start_browser()

    # 提示：运行前请确保已在 Chrome 浏览器中登录京东账号
    print("请确保 Chrome 中已登录京东，3秒后开始...")
    time.sleep(3)

    for url in urls:
        print(f"\n{'='*50}")
        print(f"开始抓取: {url}")
        print(f"{'='*50}")
        crawler.product_url = url
        crawler.product_id = crawler.extract_product_id(url)

        rows = crawler.crawl(
            max_pages=50,
            score=0,
            page_size=10,
            min_delay=8.0,
            max_delay=15.0,
        )
        all_comments.extend(rows)

        if url != urls[-1]:
            sleep_time = random.uniform(20.0, 40.0)
            print(f"切换商品，休眠 {sleep_time:.1f}s...\n")
            time.sleep(sleep_time)

    JDCommentCrawler.save_csv(all_comments, "jd_combined_comments.csv")
    crawler.close()
