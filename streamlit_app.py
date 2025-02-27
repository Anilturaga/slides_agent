import streamlit as st
import asyncio
from temporalio.client import Client
import time
from datetime import datetime
import os
import uuid
import json

# Set page config first
st.set_page_config(layout="wide", page_title="PowerPoint & Excel Assistant")

# Basic styling
st.markdown("""
<style>
.thread-item {
    background-color: #f0f2f6;
    border-radius: 5px;
    padding: 10px;
    margin-bottom: 5px;
}
.thread-item.active {
    background-color: #e0e5ea;
    border-left: 5px solid #4e8cff;
}
</style>
""", unsafe_allow_html=True)

def init_session_state():
    """Initialize session state variables"""
    if 'threads' not in st.session_state:
        st.session_state.threads = {}
    if 'current_thread' not in st.session_state:
        # Create first thread with a unique workflow ID
        thread_id = "Thread 1"
        workflow_id = f"ppt-agent-{thread_id}-{str(uuid.uuid4())[:8]}"
        st.session_state.threads[thread_id] = {
            "selected_pptx": [],
            "selected_excel": [],
            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "workflow_id": workflow_id
        }
        st.session_state.current_thread = thread_id
    
    # Store files directory in session state
    if 'files_dir' not in st.session_state:
        st.session_state.files_dir = "files"
    
    # Track connection status
    if 'temporal_connected' not in st.session_state:
        st.session_state.temporal_connected = None

def create_new_thread():
    """Create a new thread with empty file selections and a unique workflow ID"""
    thread_id = f"Thread {len(st.session_state.threads) + 1}"
    workflow_id = f"ppt-agent-{thread_id}-{str(uuid.uuid4())[:8]}"
    
    st.session_state.threads[thread_id] = {
        "selected_pptx": [],
        "selected_excel": [],
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "workflow_id": workflow_id
    }
    
    st.session_state.current_thread = thread_id
    return thread_id

async def get_temporal_client():
    """Connect to Temporal server with error handling"""
    try:
        client = await Client.connect("localhost:7233")
        st.session_state.temporal_connected = True
        return client
    except Exception as e:
        st.session_state.temporal_connected = False
        st.error(f"Could not connect to Temporal server: {str(e)}")
        return None

async def get_or_create_workflow(workflow_id):
    """Get or create workflow with specified ID"""
    client = await get_temporal_client()
    if not client:
        return None
    
    try:
        # Try to get existing workflow
        handle = client.get_workflow_handle(workflow_id)
        # Test if workflow is running
        await handle.query("get_conversation_history")
        return handle
    except Exception:
        # If workflow doesn't exist or isn't running, start a new one
        try:
            handle = await client.start_workflow(
                "PPTAgentWorkflow.run",
                id=workflow_id,
                task_queue="ppt-agent-task-queue"
            )
            return handle
        except Exception as e:
            st.error(f"Error creating workflow: {str(e)}")
            return None

async def get_conversation_history(workflow_id):
    """Get conversation history from the workflow"""
    try:
        client = await get_temporal_client()
        if not client:
            return []
        
        handle = client.get_workflow_handle(workflow_id)
        history = await handle.query("get_conversation_history")
        return history
    except Exception as e:
        st.warning(f"Could not fetch conversation history: {str(e)}")
        return []

async def send_user_input(workflow_id, query, pptx_files, excel_files):
    """Send user input to the workflow"""
    try:
        client = await get_temporal_client()
        if not client:
            return False
            
        handle = client.get_workflow_handle(workflow_id)
        input_data = {
            "query": query,
            "pptx_files": pptx_files,
            "excel_files": excel_files
        }
        await handle.signal("user_input", input_data)
        return True
    except Exception as e:
        st.error(f"Error sending message: {str(e)}")
        return False

def poll_for_assistant_response(workflow_id, message_count):
    """Poll until an assistant message is received"""
    progress_bar = st.progress(0)
    message_placeholder = st.empty()
    message_placeholder.info("Waiting for assistant response...")
    
    max_attempts = 60
    delay = 1
    
    for i in range(max_attempts):
        # Update progress bar
        progress_bar.progress((i + 1) / max_attempts)
        
        # Get current conversation history
        history = asyncio.run(get_conversation_history(workflow_id))
        
        # Find the most recent assistant message (ignoring tool messages)
        assistant_messages = [msg for msg in history if msg.get("role") == "assistant"]
        
        # If we have more assistant messages than before
        if len(assistant_messages) > message_count:
            progress_bar.empty()
            message_placeholder.empty()
            return True
            
        # Wait before polling again
        time.sleep(delay)
    
    # Timeout occurred
    progress_bar.empty()
    message_placeholder.warning("Timed out waiting for assistant response.")
    return False

def list_files(directory, extensions):
    """List files with specific extensions in a directory"""
    files = []
    if os.path.exists(directory):
        for file in os.listdir(directory):
            if any(file.lower().endswith(ext) for ext in extensions):
                files.append(os.path.join(directory, file))
    return files

def count_assistant_messages(conversation):
    """Count how many messages from the assistant are in the conversation"""
    return len([msg for msg in conversation if msg.get("role") == "assistant"])

def main():
    # Initialize session state
    init_session_state()
    
    # Main layout
    st.title("PowerPoint & Excel Assistant")
    
    # Create two columns
    left_col, chat_col = st.columns([1, 3])
    
    # Files directory and extensions
    files_dir = st.session_state.files_dir
    pptx_extensions = [".pptx", ".ppt"]
    excel_extensions = [".xlsx", ".xls"]
    
    # Thread management and file selection
    with left_col:
        # Thread section
        col1, col2 = st.columns([3, 1])
        with col1:
            st.markdown("<h3>Threads</h3>", unsafe_allow_html=True)
        with col2:
            if st.button("➕", help="Create a new thread"):
                create_new_thread()
                st.rerun()
        
        # Thread selection
        for thread_id in st.session_state.threads:
            thread_data = st.session_state.threads[thread_id]
            # Create a styled thread box
            active_class = "active" if thread_id == st.session_state.current_thread else ""
            if st.button(
                f"{thread_id}",
                key=f"thread_{thread_id}",
                help=f"Created: {thread_data['created_at']}"
            ):
                st.session_state.current_thread = thread_id
                st.rerun()
        
        # Get current thread data
        current_thread = st.session_state.current_thread
        current_thread_data = st.session_state.threads[current_thread]
        workflow_id = current_thread_data["workflow_id"]
        
        # File selection for current thread
        st.subheader("File Selection")
        
        # List available files
        pptx_files = list_files(files_dir, pptx_extensions)
        excel_files = list_files(files_dir, excel_extensions)
        
        # PowerPoint files
        with st.expander("PowerPoint Files", expanded=True):
            selected_pptx = []
            for pptx_file in pptx_files:
                filename = os.path.basename(pptx_file)
                is_selected = pptx_file in current_thread_data["selected_pptx"]
                if st.checkbox(filename, value=is_selected, key=f"{current_thread}_pptx_{filename}"):
                    selected_pptx.append(pptx_file)
            
            # Update selected files
            current_thread_data["selected_pptx"] = selected_pptx
        
        # Excel files
        with st.expander("Excel Files", expanded=True):
            selected_excel = []
            for excel_file in excel_files:
                filename = os.path.basename(excel_file)
                is_selected = excel_file in current_thread_data["selected_excel"]
                if st.checkbox(filename, value=is_selected, key=f"{current_thread}_excel_{filename}"):
                    selected_excel.append(excel_file)
            
            # Update selected files
            current_thread_data["selected_excel"] = selected_excel
        
        # Connection status
        if st.session_state.temporal_connected is False:
            st.error("Disconnected from Temporal server")
        elif st.session_state.temporal_connected is True:
            st.success("Connected to Temporal server")
    
    # Chat area
    with chat_col:
        # Display current thread info
        st.subheader(f"Thread: {current_thread}")
        
        # Show selected files
        if selected_pptx or selected_excel:
            files_text = []
            if selected_pptx:
                files_text.append(f"**PowerPoint:** {', '.join([os.path.basename(f) for f in selected_pptx])}")
            if selected_excel:
                files_text.append(f"**Excel:** {', '.join([os.path.basename(f) for f in selected_excel])}")
            
            st.info("Selected files: " + " | ".join(files_text))
        
        # Create or get workflow for this thread
        asyncio.run(get_or_create_workflow(workflow_id))
        
        # Get conversation history
        conversation = asyncio.run(get_conversation_history(workflow_id))
        
        # Display chat messages
        st.subheader("Conversation")
        
        if conversation:
            # Skip the system message
            for message in conversation[1:]:
                role = message.get("role", "")
                content = message.get("content", "")
                
                # Only show user and assistant messages
                if role in ["user", "assistant"]:
                    with st.chat_message(role):
                        st.write(content)
        else:
            st.info("Send a message to start the conversation.")
        
        # Chat input
        if prompt := st.chat_input("Ask about your PowerPoint or Excel files..."):
            # Count assistant messages before sending
            assistant_msg_count = count_assistant_messages(conversation)
            
            # Show user message immediately
            with st.chat_message("user"):
                st.write(prompt)
            
            # Send message to the workflow
            success = asyncio.run(send_user_input(workflow_id, prompt, selected_pptx, selected_excel))
            
            if success:
                # Poll until we get a new assistant message
                if poll_for_assistant_response(workflow_id, assistant_msg_count):
                    st.rerun()
            else:
                if st.session_state.temporal_connected is False:
                    st.error("Message not sent: Disconnected from Temporal server")

if __name__ == "__main__":
    main() 