# needs: pip install sqlalchemy pymysql bizdays
from sqlalchemy import create_engine, text

engine = create_engine(
    "mysql+pymysql://read.only:read.only@pim-cpt-mysql-prod.prescient.local:3307/prime_profile"
)
with engine.connect() as conn:
    rows = conn.execute(text("SELECT COUNT(*) FROM tmp_funds")).scalar()
    print("rows in tmp_funds:", rows)