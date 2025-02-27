import streamlit as st
import asyncio
from temporalio.client import Client
import time
from datetime import datetime
import os
import json

async def get_temporal_client():
    """Connect to Temporal server"""
    if 'temporal_client' not in st.session_state:
        st.session_state.temporal_client = await Client.connect("localhost:7233")
    return st.session_state.temporal_client

async def get_workflow_handle():
    """Get handle to the PowerPoint agent workflow"""
    client = await get_temporal_client()
    workflow_id = "ppt-agent-workflow"  # Use a consistent workflow ID
    
    # Check if workflow exists already
    try:
        handle = client.get_workflow_handle(workflow_id)
        # Try a query to see if it's running
        await handle.query("get_conversation_history")
        return handle
    except Exception:
        # If workflow doesn't exist or is not running, start a new one
        handle = await client.start_workflow(
            "PPTAgentWorkflow.run",
            id=workflow_id,
            task_queue="ppt-agent-task-queue"
        )
        return handle

async def get_conversation_history():
    """Get the current conversation history from the workflow"""
    try:
        handle = await get_workflow_handle()
        history = await handle.query("get_conversation_history")
        return history
    except Exception as e:
        st.error(f"Error fetching conversation history: {str(e)}")
        return []

async def send_user_input(query, pptx_files, excel_files):
    """Send user input to the workflow"""
    try:
        handle = await get_workflow_handle()
        input_data = {
            "query": query,
            "pptx_files": pptx_files,
            "excel_files": excel_files
        }
        await handle.signal("user_input", input_data)
        return True
    except Exception as e:
        st.error(f"Error sending user input: {str(e)}")
        return False

def poll_for_assistant_response(message_count):
    """
    Poll the workflow until an assistant response is received
    
    Args:
        message_count: The number of messages before sending user message
    
    Returns:
        True when an assistant response is detected
    """
    # Create a progress bar
    progress_bar = st.progress(0)
    st.write("Waiting for assistant response...")
    
    # Maximum polling attempts and delay between attempts
    max_attempts = 60  # 60 seconds max waiting time
    delay = 1  # 1 second between polls
    
    for i in range(max_attempts):
        # Update progress bar
        progress_bar.progress((i + 1) / max_attempts)
        
        # Get current conversation history
        history = asyncio.run(get_conversation_history())
        
        # If we have more messages than before AND the last message is from the assistant
        if (len(history) > message_count and 
            history[-1].get("role") == "assistant"):
            # Clear progress bar and return success
            progress_bar.empty()
            return True
            
        # Wait before polling again
        time.sleep(delay)
    
    # If we reach here, we timed out waiting for a response
    progress_bar.empty()
    st.warning("Timed out waiting for assistant response.")
    return False

def list_files(directory, extensions):
    """List files with specific extensions in a directory"""
    files = []
    if os.path.exists(directory):
        for file in os.listdir(directory):
            if any(file.lower().endswith(ext) for ext in extensions):
                files.append(os.path.join(directory, file))
    return files

def init_session_state():
    """Initialize session state variables"""
    if 'threads' not in st.session_state:
        st.session_state.threads = {}
    if 'current_thread' not in st.session_state:
        st.session_state.current_thread = "Thread 1"
        st.session_state.threads["Thread 1"] = {
            "selected_pptx": [],
            "selected_excel": [],
            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
    # Store files directory in session state
    if 'files_dir' not in st.session_state:
        st.session_state.files_dir = "files"

def create_new_thread():
    """Create a new thread with empty file selections"""
    # Generate a new thread name
    thread_id = f"Thread {len(st.session_state.threads) + 1}"
    
    # Create thread data
    st.session_state.threads[thread_id] = {
        "selected_pptx": [],
        "selected_excel": [],
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }
    
    # Switch to the new thread
    st.session_state.current_thread = thread_id
    return thread_id

def main():
    # This must be the first Streamlit command
    st.set_page_config(layout="wide", page_title="PowerPoint & Excel Assistant")
    
    # Custom CSS for better styling - moved after set_page_config
    st.markdown("""
    <style>
    .thread-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 1rem;
    }
    .thread-item {
        background-color: #f0f2f6;
        border-radius: 5px;
        padding: 10px;
        margin-bottom: 5px;
        cursor: pointer;
    }
    .thread-item.active {
        background-color: #e0e5ea;
        border-left: 5px solid #4e8cff;
    }
    .file-section {
        background-color: #f9f9f9;
        border-radius: 5px;
        padding: 10px;
        margin-top: 15px;
        border: 1px solid #eaeaea;
    }
    .section-title {
        font-weight: bold;
        margin-bottom: 10px;
        border-bottom: 1px solid #eaeaea;
        padding-bottom: 5px;
    }
    </style>
    """, unsafe_allow_html=True)
    
    # Initialize session state
    init_session_state()
    
    # Main layout
    st.title("PowerPoint & Excel Assistant")
    
    # Create two columns - one for threads/files and one for chat
    left_col, chat_col = st.columns([1, 3])
    
    # Files directory - hidden from UI but accessible in code
    files_dir = st.session_state.files_dir
    pptx_extensions = [".pptx", ".ppt"]
    excel_extensions = [".xlsx", ".xls"]
    
    # Thread management and file selection in the left column
    with left_col:
        # Custom header with Thread title and + button side by side
        col1, col2 = st.columns([3, 1])
        with col1:
            st.markdown("<h3>Threads</h3>", unsafe_allow_html=True)
        with col2:
            if st.button("➕", help="Create a new thread"):
                create_new_thread()
                st.rerun()
        
        # Thread selection with custom styling
        for thread_id in st.session_state.threads:
            thread_data = st.session_state.threads[thread_id]
            # Create a styled thread item
            active_class = "active" if thread_id == st.session_state.current_thread else ""
            thread_html = f"""
            <div class="thread-item {active_class}" id="{thread_id}">
                <b>{thread_id}</b><br>
                <small>{thread_data['created_at']}</small>
            </div>
            """
            if st.markdown(thread_html, unsafe_allow_html=True):
                st.session_state.current_thread = thread_id
                st.rerun()
                
        # Alternative thread selection using selectbox (as backup)
        thread_options = list(st.session_state.threads.keys())
        selected_thread = st.selectbox("Change Thread", thread_options, 
                                       index=thread_options.index(st.session_state.current_thread))
        
        if selected_thread != st.session_state.current_thread:
            st.session_state.current_thread = selected_thread
            st.rerun()
        
        # File selection for current thread
        st.markdown("<div class='file-section'>", unsafe_allow_html=True)
        st.markdown("<div class='section-title'>File Selection</div>", unsafe_allow_html=True)
        
        # List all available files
        pptx_files = list_files(files_dir, pptx_extensions)
        excel_files = list_files(files_dir, excel_extensions)
        
        # Get current thread's selections
        current_thread_data = st.session_state.threads[st.session_state.current_thread]
        
        # PowerPoint files with collapsible section
        with st.expander("PowerPoint Files", expanded=True):
            selected_pptx = []
            for pptx_file in pptx_files:
                filename = os.path.basename(pptx_file)
                # Check if file was previously selected in this thread
                is_selected = pptx_file in current_thread_data["selected_pptx"]
                if st.checkbox(filename, value=is_selected, key=f"{st.session_state.current_thread}_pptx_{filename}"):
                    selected_pptx.append(pptx_file)
            
            # Update current thread's selected PowerPoint files
            current_thread_data["selected_pptx"] = selected_pptx
        
        # Excel files with collapsible section
        with st.expander("Excel Files", expanded=True):
            selected_excel = []
            for excel_file in excel_files:
                filename = os.path.basename(excel_file)
                # Check if file was previously selected in this thread
                is_selected = excel_file in current_thread_data["selected_excel"]
                if st.checkbox(filename, value=is_selected, key=f"{st.session_state.current_thread}_excel_{filename}"):
                    selected_excel.append(excel_file)
            
            # Update current thread's selected Excel files
            current_thread_data["selected_excel"] = selected_excel
        
        st.markdown("</div>", unsafe_allow_html=True)
    
    # Chat area
    with chat_col:
        # Display current thread info in a styled header
        st.markdown(f"""
        <div style="background-color: #f0f2f6; padding: 10px; border-radius: 5px; margin-bottom: 15px;">
            <h3 style="margin: 0; color: #2c3e50;">Thread: {st.session_state.current_thread}</h3>
        </div>
        """, unsafe_allow_html=True)
        
        # Selected files info in a nice formatted box
        if selected_pptx or selected_excel:
            file_html = """
            <div style="background-color: #f8f9fa; padding: 10px; border-radius: 5px; margin-bottom: 15px; border-left: 3px solid #4CAF50;">
                <h4 style="margin-top: 0;">Selected Files</h4>
            """
            
            if selected_pptx:
                file_html += "<p><b>PowerPoint:</b> " + ", ".join([os.path.basename(f) for f in selected_pptx]) + "</p>"
            
            if selected_excel:
                file_html += "<p><b>Excel:</b> " + ", ".join([os.path.basename(f) for f in selected_excel]) + "</p>"
            
            file_html += "</div>"
            st.markdown(file_html, unsafe_allow_html=True)
        
        # Get conversation history
        conversation = asyncio.run(get_conversation_history())
        
        # Display chat messages in a styled container
        st.markdown("<h3>Conversation</h3>", unsafe_allow_html=True)
        chat_container = st.container()
        
        with chat_container:
            # Skip the system message (first message)
            for message in conversation[1:] if len(conversation) > 0 else []:
                role = message.get("role", "")
                content = message.get("content", "")
                
                # Skip tool messages - they're usually verbose and technical
                if role in ["user", "assistant"]:
                    with st.chat_message(role):
                        st.write(content)
        
        # Chat input with better styling
        if prompt := st.chat_input("Ask about your PowerPoint or Excel files...", key="chat_input"):
            # Get current message count
            message_count = len(conversation)
            
            # Show the user message immediately
            with st.chat_message("user"):
                st.write(prompt)
            
            # Send message to workflow with current thread's selected files
            success = asyncio.run(send_user_input(prompt, selected_pptx, selected_excel))
            
            if success:
                # Poll until we get an assistant response
                if poll_for_assistant_response(message_count):
                    st.rerun()
            else:
                st.error("Failed to send message to the workflow")

if __name__ == "__main__":
    main() 