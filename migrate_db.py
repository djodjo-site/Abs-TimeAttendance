import sys
import os

sys.path.append(os.path.abspath(os.path.dirname(__file__)))
from database.db_config import get_db_connection

def init_database():
    try:
        conn = get_db_connection()
        with conn.cursor() as cursor:
            # 1. Table des horaires de travail
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS work_schedules (
                    id SERIAL PRIMARY KEY,
                    name VARCHAR(100) NOT NULL,
                    start_time TIME NOT NULL,
                    end_time TIME NOT NULL,
                    late_tolerance_minutes INT DEFAULT 30
                )
            """)
            
            # 2. Table de liaison pour les horaires multiples par travailleur
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS employee_schedules (
                    badge_or_card_id VARCHAR(100) NOT NULL,
                    schedule_id INT NOT NULL REFERENCES work_schedules(id) ON DELETE CASCADE,
                    PRIMARY KEY (badge_or_card_id, schedule_id)
                )
            """)

            # 3. Table des employés
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS employees (
                    badge_or_card_id VARCHAR(100) PRIMARY KEY,
                    first_name VARCHAR(100),
                    last_name VARCHAR(100),
                    department VARCHAR(100),
                    is_active BOOLEAN DEFAULT TRUE
                )
            """)

            # 4. Table de renommage des appareils / terminaux
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS devices_mapping (
                    device_key VARCHAR(100) PRIMARY KEY,
                    friendly_name VARCHAR(100) NOT NULL
                )
            """)

            # 5. Table des pointages bruts
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS attendance_raw_logs (
                    id SERIAL PRIMARY KEY,
                    badge_or_card_id VARCHAR(100),
                    punch_timestamp TIMESTAMP,
                    device_id VARCHAR(100),
                    punch_type VARCHAR(20) DEFAULT 'IN'
                )
            """)

            # Insertion des horaires par défaut s'il n'en existe aucun
            cursor.execute("SELECT COUNT(*) as cnt FROM work_schedules")
            res = cursor.fetchone()
            if res and res['cnt'] == 0:
                cursor.execute("""
                    INSERT INTO work_schedules (name, start_time, end_time, late_tolerance_minutes)
                    VALUES 
                    ('Standard (08:00 - 17:00)', '08:00:00', '17:00:00', 30),
                    ('Vacation Matin (06:00 - 14:00)', '06:00:00', '14:00:00', 15),
                    ('Vacation Soir (14:00 - 22:00)', '14:00:00', '22:00:00', 15),
                    ('Shift Nuit (22:00 - 06:00)', '22:00:00', '06:00:00', 20)
                """)

            conn.commit()
        conn.close()
        print("Base de données initialisée avec succès (tables vérifiées/créées) !")
    except Exception as e:
        print(f"Erreur lors de l'initialisation BDD : {e}")

if __name__ == "__main__":
    init_database()