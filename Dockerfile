FROM python:3.12-slim
WORKDIR /app
COPY . /app
ENV QUARTAL_HOST=0.0.0.0
ENV QUARTAL_PORT=8765
EXPOSE 8765
VOLUME ["/app/data"]
CMD ["python", "server.py"]
