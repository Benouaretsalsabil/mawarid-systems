import sqlite3
import random
from datetime import datetime, timedelta

def build_mawarid_database():
    # الاتصال بقاعدة بيانات موارِد
    conn = sqlite3.connect('smart_meters.db')
    cursor = conn.cursor()

    # إعادة إنشاء جدول القراءات المتكامل
    cursor.execute('DROP TABLE IF EXISTS telemetry')
    cursor.execute('''
        CREATE TABLE telemetry (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            meter_id TEXT NOT NULL,
            utility_type TEXT NOT NULL,
            consumption REAL NOT NULL,
            battery_level INTEGER NOT NULL,
            status TEXT NOT NULL,
            timestamp DATETIME NOT NULL
        )
    ''')

    # قائمة العدادات التجريبية للمرافق الثلاثة
    meters = [
        ('DZ-MTR-W101', 'Water'),
        ('DZ-MTR-W102', 'Water'),
        ('DZ-MTR-E201', 'Electricity'),
        ('DZ-MTR-E202', 'Electricity'),
        ('DZ-MTR-G301', 'Gas')
    ]

    # توليد بيانات لآخر 7 أيام (168 ساعة)
    start_time = datetime.now() - timedelta(days=7)

    for hour in range(168):
        current_time = start_time + timedelta(hours=hour)
        
        for meter_id, utility in meters:
            if utility == 'Water':
                val = round(random.uniform(0.05, 0.75), 2)
                status = 'Normal' if val < 0.65 else 'Leak Warning'
            elif utility == 'Electricity':
                val = round(random.uniform(1.10, 4.20), 2)
                status = 'Normal' if val < 3.80 else 'High Load'
            else:  # Gas
                val = round(random.uniform(0.02, 0.35), 2)
                status = 'Normal' if val < 0.30 else 'Pressure Anomaly'

            battery = max(100 - (hour // 12), 70)

            cursor.execute('''
                INSERT INTO telemetry (meter_id, utility_type, consumption, battery_level, status, timestamp)
                VALUES (?, ?, ?, ?, ?, ?)
            ''', (meter_id, utility, val, battery, status, current_time.strftime('%Y-%m-%d %H:%M:%S')))

    conn.commit()
    conn.close()
    print("✅ تم بفضل الله إنشاء قاعدة بيانات Mawarid Systems بنجاح!")

if __name__ == '__main__':
    build_mawarid_database()