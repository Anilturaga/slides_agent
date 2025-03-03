# PowerPoint and Excel Assistant

A simple guide to starting the application.

## Setup

1. Create a `.env` file in the root directory with the following format:
```
SSL_CERT_FILE=
OPENAI_API_KEY=
OPENAI_BASE_URL=
```

Fill in your API key and other required values.

## Starting the Application

### Step 1: Start the Temporal Cluster

First, you need to have a Temporal server running. You can set it up using Docker:

```bash
git clone https://github.com/temporalio/docker-compose.git
cd docker-compose
docker-compose -f docker-compose-postgres.yml up
```

This will start the Temporal server and UI (available at http://localhost:8088).

### Step 2: Start the Temporal Worker

In a new terminal:

```bash
python slides_agent/temporal_agent.py
```

You should see output indicating that the worker has started:

```
Running in worker mode...
Worker started. Press Ctrl+C to exit.
```

### Step 3: Start the Streamlit App

In a separate terminal:

```bash
streamlit run slides_agent/streamlit_app.py
```

You should see output indicating that the Streamlit app has started:

```
The web interface will typically be available at http://localhost:8501
```
