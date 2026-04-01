from datetime import datetime

SERVERS = [
    (1,  "心动",     datetime(2026, 2, 11, 12, 0)),
    (2,  "飞天",     datetime(2026, 1, 28, 12, 0)),
    (3,  "大吉大利", datetime(2026, 1, 14, 12, 0)),
    (4,  "天命",     datetime(2025, 12, 31, 12, 0)),
    (5,  "来财",     datetime(2025, 12, 17, 12, 0)),
    (6,  "红颜",     datetime(2025, 12,  3, 12, 0)),
    (7,  "无双",     datetime(2025, 11, 19, 12, 0)),
    (8,  "赫赫有名", datetime(2025, 11,  5, 12, 0)),
    (9,  "浪浪山",   datetime(2025, 10, 22, 12, 0)),
    (10, "孙悟空",   datetime(2025, 10,  8, 12, 0)),
    (11, "出发",     datetime(2025,  9, 24, 12, 0)),
    (12, "2015",     datetime(2025,  9, 10, 12, 0)),
    (13, "紫禁城",   datetime(2025,  8, 27, 12, 0)),
    (14, "上海滩",   datetime(2025,  8, 27, 12, 0)),
    (15, "广州塔",   datetime(2025,  8, 27, 12, 0)),
    (16, "苍山洱海", datetime(2025,  8, 27, 12, 0)),
    (17, "曲阜孔庙", datetime(2025,  8, 27, 12, 0)),
    (18, "长白山",   datetime(2025,  8, 27, 12, 0)),
]


def select_server() -> tuple:
    """交互式服务器选择，返回 (server_name, open_datetime)。"""
    print("\n请选择服务器：")
    for idx, name, open_dt in SERVERS:
        print(f"  {idx:2d}. {name}  （开服：{open_dt.strftime('%Y-%m-%d')}）")
    while True:
        try:
            raw = input("\n输入序号：").strip()
            choice = int(raw)
            for idx, name, open_dt in SERVERS:
                if idx == choice:
                    print(f"  已选择：{name}")
                    return name, open_dt
            print(f"  无效序号，请输入 1~{len(SERVERS)}")
        except ValueError:
            print("  请输入数字序号。")
        except KeyboardInterrupt:
            print("\n已取消，退出程序。")
            raise SystemExit(0)
