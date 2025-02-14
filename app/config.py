import os


class Config:
    """Application Configuration"""

    # General Settings
    SECRET_KEY = os.getenv("SECRET_KEY", "your_secret_key")
    DEBUG = True
    ENV = "development"

    # File Paths
    BASE_DIR = os.path.abspath(os.path.dirname(__file__))
    EXCEL_FILE = os.path.join(BASE_DIR, "templates", "SampleData.xlsx")
    MODIFIED_FILE = os.path.join(BASE_DIR, "templates", "modified.xlsx")
