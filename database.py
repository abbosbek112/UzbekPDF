from sqlalchemy import create_engine, Column, Integer, String, Boolean
from sqlalchemy.orm import sessionmaker, declarative_base

SQLALCHEMY_DATABASE_URL = "sqlite:///./uzbekpdf.db"

engine = create_engine(
    SQLALCHEMY_DATABASE_URL, connect_args={"check_same_thread": False}
)
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)

Base = declarative_base()

class User(Base):
    __tablename__ = "users"

    id = Column(Integer, primary_key=True, index=True)
    username = Column(String, unique=True, index=True)
    email = Column(String, unique=True, index=True, nullable=True)
    google_id = Column(String, unique=True, index=True, nullable=True)
    avatar_url = Column(String, nullable=True)
    password_hash = Column(String, nullable=True)
    is_premium = Column(Boolean, default=False)
    is_admin = Column(Boolean, default=False)


class Payment(Base):
    __tablename__ = "payments"
    id = Column(Integer, primary_key=True, index=True)
    user_id = Column(Integer)
    amount = Column(Integer)
    plan = Column(String) # 'monthly' or 'yearly'
    status = Column(String, default="pending") # pending, success, failed
    transaction_id = Column(String, unique=True, nullable=True)
    created_at = Column(String)

class PaymeTransaction(Base):
    __tablename__ = "payme_transactions"
    
    id = Column(Integer, primary_key=True, index=True)
    payme_transaction_id = Column(String, unique=True, index=True)
    payme_time = Column(String)
    account_id = Column(Integer)  # User ID
    amount = Column(Integer)      # In tiyin
    state = Column(Integer)       # 1=created, 2=completed, -1=cancelled etc.
    create_time = Column(Integer)
    perform_time = Column(Integer, default=0)
    cancel_time = Column(Integer, default=0)
    reason = Column(Integer, nullable=True)

class Usage(Base):
    __tablename__ = "usage"
    id = Column(Integer, primary_key=True, index=True)
    user_id = Column(Integer, index=True, nullable=True)
    ip_address = Column(String, index=True)
    count = Column(Integer, default=0)
    date = Column(String) # YYYY-MM-DD



Base.metadata.create_all(bind=engine)

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()
