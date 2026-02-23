import os
import dj_database_url

from config.settings.base import *
from dotenv import load_dotenv
load_dotenv(BASE_DIR / ".env") 

SECRET_KEY = os.getenv('SECRET_KEY')
DEBUG = True
DATABASES = {
    'default': dj_database_url.parse(os.getenv('DATABASE_URL'))
}

# Включаем нужные приложения для разработки
INSTALLED_APPS += [
    "debug_toolbar",
    "django_extensions",
    "silk",
]

MIDDLEWARE += [
    "debug_toolbar.middleware.DebugToolbarMiddleware",
]

STATICFILES_DIRS = [BASE_DIR.parent / "static"]