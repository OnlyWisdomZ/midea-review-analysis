import pymysql
import hashlib

def init_database():
    print("正在连接 MySQL 进行数据库初始化...")
    try:
        # 连接 MySQL 服务器（尚未指定库）
        conn = pymysql.connect(
            host="127.0.0.1",
            user="root",
            password="123456",
            charset="utf8mb4"
        )
        cursor = conn.cursor()
        
        # 创建数据库
        cursor.execute("CREATE DATABASE IF NOT EXISTS meidi_analysis DEFAULT CHARACTER SET utf8mb4;")
        cursor.execute("USE meidi_analysis;")
        
        # 创建 users 表
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id INT AUTO_INCREMENT PRIMARY KEY,
                username VARCHAR(50) NOT NULL UNIQUE,
                password_hash VARCHAR(64) NOT NULL
            );
        """)
        
        # 密码 123456 的 SHA-256
        password_str = "123456"
        pwd_hash = hashlib.sha256(password_str.encode('utf-8')).hexdigest()
        
        # 检查是否存在 admin
        cursor.execute("SELECT id FROM users WHERE username='admin';")
        if not cursor.fetchone():
            cursor.execute("INSERT INTO users (username, password_hash) VALUES (%s, %s);", ("admin", pwd_hash))
            print("账号 admin (密码 123456) 已创建。")
        else:
            # 更新密码确保正确
            cursor.execute("UPDATE users SET password_hash=%s WHERE username='admin';", (pwd_hash,))
            print("账号 admin 已存在，密码已重置为 123456。")
            
        conn.commit()
        cursor.close()
        conn.close()
        print("数据库初始化完成。")
    except Exception as e:
        print(f"数据库初始化失败: {e}")

if __name__ == "__main__":
    init_database()
