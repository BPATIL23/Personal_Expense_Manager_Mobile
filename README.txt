PERSONAL EXPENSE MANAGER - MOBILE + POSTGRES VERSION

LOCAL:
Keep your existing expenses.db. Install requirements and run: python app.py

RENDER:
1. Add DATABASE_URL = PostgreSQL Internal Database URL.
2. Add SECRET_KEY = any long random value.
3. Build Command: pip install -r requirements.txt
4. Start Command: gunicorn app:app

MOBILE:
Open the Render URL in Android Chrome and use Menu -> Add to Home screen / Install app.

IMPORTANT: Do not delete expenses.db. Existing SQLite data is not automatically copied to PostgreSQL.
