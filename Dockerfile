FROM python:3.12-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt gunicorn

COPY config.py data_loader.py search_engine.py template_engine.py gcs_helper.py app.py ./
COPY templates/ templates/
COPY print_templates/ print_templates/

RUN mkdir -p uploads user_data

ENV PORT=8080
EXPOSE 8080

CMD exec gunicorn --bind :$PORT --workers 2 --threads 4 --timeout 300 app:app
