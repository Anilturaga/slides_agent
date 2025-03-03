# PowerPoint and Excel Assistant

A simple guide to starting the application.

## Setup

1. Create a `.env` file in the root directory with the following format:
```
SSL_CERT_FILE=
OPENAI_API=
OPENAI_BASE_URL=
```

Fill in your API key and other required values.

## Starting the Application

### Step 1: Start the Temporal Worker

```bash
python3 temporal_agent.worker
```

You should see output indicating that the worker has started:

```Running in worker mode...
Worker started. Press Ctrl+C to exit.
```

### Step 2: Start the Streamlit App

In a separate terminal:

```bash
streamlit run streamlit_app.py
```

You should see output indicating that the Streamlit app has started:

```
The web interface will typically be available at http://localhost:8501
```
