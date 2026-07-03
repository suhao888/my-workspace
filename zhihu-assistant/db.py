# -*- coding: utf-8 -*-
"""
数据库模块 - SQLite 存储，记录已处理的问题，避免重复生成
"""

import sqlite3
import os
from datetime import datetime

DB_PATH = os.path.join(os.path.dirname(__file__), "zhihu_assistant.db")


def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_conn()
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS processed_questions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            question_id TEXT NOT NULL UNIQUE,
            title TEXT NOT NULL,
            url TEXT NOT NULL,
            source TEXT DEFAULT '',
            topic TEXT DEFAULT '',
            hot_score INTEGER DEFAULT 0,
            status TEXT DEFAULT 'draft',
            draft_path TEXT DEFAULT '',
            visit_count INTEGER DEFAULT 0,
            answer_count INTEGER DEFAULT 0,
            follower_count INTEGER DEFAULT 0,
            exposure_score REAL DEFAULT 0,
            created_at TEXT DEFAULT (datetime('now','localtime')),
            updated_at TEXT DEFAULT (datetime('now','localtime'))
        );
        
        CREATE TABLE IF NOT EXISTS drafts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            question_id TEXT NOT NULL,
            content TEXT NOT NULL,
            word_count INTEGER DEFAULT 0,
            created_at TEXT DEFAULT (datetime('now','localtime'))
        );
        
        CREATE INDEX IF NOT EXISTS idx_question_id 
            ON processed_questions(question_id);
        CREATE INDEX IF NOT EXISTS idx_status 
            ON processed_questions(status);
    """)
    conn.commit()
    conn.close()


def is_processed(question_id):
    conn = get_conn()
    row = conn.execute(
        "SELECT id FROM processed_questions WHERE question_id = ?", (question_id,)
    ).fetchone()
    conn.close()
    return row is not None


def save_question(question_id, title, url, source="", topic="", hot_score=0):
    conn = get_conn()
    try:
        conn.execute(
            """
            INSERT OR IGNORE INTO processed_questions 
            (question_id, title, url, source, topic, hot_score)
            VALUES (?, ?, ?, ?, ?, ?)
        """,
            (question_id, title, url, source, topic, hot_score),
        )
        conn.commit()
    except Exception as e:
        print(f"  [DB] 保存问题失败: {e}")
    finally:
        conn.close()


def update_status(question_id, status, draft_path=""):
    conn = get_conn()
    conn.execute(
        """
        UPDATE processed_questions 
        SET status = ?, draft_path = ?, updated_at = datetime('now','localtime')
        WHERE question_id = ?
    """,
        (status, draft_path, question_id),
    )
    conn.commit()
    conn.close()


def save_draft(question_id, content):
    conn = get_conn()
    word_count = len(content)
    conn.execute(
        """
        INSERT INTO drafts (question_id, content, word_count)
        VALUES (?, ?, ?)
    """,
        (question_id, content, word_count),
    )
    conn.commit()
    conn.close()


def get_unprocessed_questions():
    conn = get_conn()
    rows = conn.execute("""
        SELECT * FROM processed_questions 
        WHERE status = 'draft'
        ORDER BY hot_score DESC
        LIMIT 20
    """).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def _save_exposure(
    question_id, visit_count, answer_count, follower_count, exposure_score
):
    """保存曝光数据"""
    conn = get_conn()
    try:
        conn.execute(
            """
            UPDATE processed_questions 
            SET visit_count = ?, answer_count = ?, follower_count = ?,
                exposure_score = ?, updated_at = datetime('now','localtime')
            WHERE question_id = ?
        """,
            (visit_count, answer_count, follower_count, exposure_score, question_id),
        )
        conn.commit()
    except Exception as e:
        print(f"  [DB] 保存曝光数据失败: {e}")
    finally:
        conn.close()


def get_stats():
    conn = get_conn()
    stats = conn.execute("""
        SELECT 
            COUNT(*) as total,
            SUM(CASE WHEN status='draft' THEN 1 ELSE 0 END) as pending,
            SUM(CASE WHEN status='generated' THEN 1 ELSE 0 END) as generated,
            SUM(CASE WHEN status='posted' THEN 1 ELSE 0 END) as posted
        FROM processed_questions
    """).fetchone()
    conn.close()
    return dict(stats)


if __name__ == "__main__":
    init_db()
    print("数据库初始化完成")
