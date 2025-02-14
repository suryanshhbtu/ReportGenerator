from flask import Flask
from app.config import Config
from app.routes import api_blueprint

def create_app():
    """Flask App Factory"""
    app = Flask(__name__)
    app.config.from_object(Config)

    # Register Blueprints
    app.register_blueprint(api_blueprint)

    return app
