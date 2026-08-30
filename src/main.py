import os
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
import uvicorn
from dotenv import load_dotenv
from src.routers import download

load_dotenv()

# FastAPIアプリケーションの立ち上げ
app = FastAPI()

# CORSの設定（すべてのオリジンからの通信を許可）
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ルーターの登録
app.include_router(download.router)

# 保存先とツール類のディレクトリ設定
SAVE_DIR = "downloads"

# downloadsフォルダが存在しない場合は自動作成
if not os.path.exists(SAVE_DIR):
    os.makedirs(SAVE_DIR)

if __name__ == "__main__":
    # 環境変数から取得（設定がない場合のデフォルト値も設定）
    host = os.getenv("HOST", "127.0.0.1")
    port = int(os.getenv("PORT", 8749))
    
    uvicorn.run(app, host=host, port=port)