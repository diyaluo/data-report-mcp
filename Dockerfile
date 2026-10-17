FROM python:3.12-slim
WORKDIR /app
COPY pyproject.toml README.md ./
COPY src ./src
RUN pip install --no-cache-dir mcp openpyxl \
 && pip install --no-cache-dir --no-deps -e .
# stdio by default; http transport set REMOTE_TRANSPORT=http + MCP_PORT
ENV REMOTE_TRANSPORT=stdio
ENTRYPOINT ["python", "-m", "data_report_mcp.server"]
