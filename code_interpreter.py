from typing import List, Dict, Any, Optional, Union, Callable
from dataclasses import dataclass, field
import json
import base64
import enum
import time
import sys
import subprocess

class ChartType(str, enum.Enum):
    LINE = "line"
    SCATTER = "scatter"
    BAR = "bar"
    PIE = "pie"
    BOX_AND_WHISKER = "box_and_whisker"
    SUPERCHART = "superchart"
    UNKNOWN = "unknown"

@dataclass
class Chart:
    """Base chart class"""
    type: ChartType
    title: str
    elements: List[Any]

@dataclass
class OutputMessage:
    """Message from stdout/stderr"""
    line: str
    timestamp: int
    error: bool = False

@dataclass
class ExecutionError:
    """Execution error details"""
    name: str
    value: str
    traceback: str

@dataclass
class Result:
    """Execution result with multiple formats"""
    text: Optional[str] = None
    html: Optional[str] = None
    markdown: Optional[str] = None
    svg: Optional[str] = None
    png: Optional[str] = None
    jpeg: Optional[str] = None
    pdf: Optional[str] = None
    latex: Optional[str] = None
    json: Optional[dict] = None
    javascript: Optional[str] = None
    data: Optional[dict] = None
    chart: Optional[Chart] = None
    is_main_result: bool = False

@dataclass
class Execution:
    """Complete execution result"""
    results: List[Result] = field(default_factory=list)
    stdout: List[str] = field(default_factory=list)
    stderr: List[str] = field(default_factory=list)
    error: Optional[ExecutionError] = None
    execution_count: Optional[int] = None

# Code Executor class for running Python code
class CodeExecutor:
    def __init__(self, install_dependencies=True):
        from jupyter_client import KernelManager
        
        if install_dependencies:
            self._ensure_dependencies([
                'matplotlib',
                'pandas',
                'numpy',
                'ipython',
                'pillow',
                'plotly',  # For advanced charts
                'seaborn',  # For statistical visualizations
                'python-pptx',  # For PowerPoint integration
                'openpyxl'   # For Excel integration
            ])
            
        self.km = KernelManager()
        self.km.start_kernel()
        self.client = self.km.client()
        
        # Initialize the kernel with common imports and display setup
        self._initialize_kernel()
        
    def _ensure_dependencies(self, packages: List[str]):
        """Install required packages if they're not already installed"""
        for package in packages:
            try:
                __import__(package.replace('-', '_'))
            except ImportError:
                print(f"Installing {package}...")
                subprocess.check_call([
                    sys.executable, 
                    "-m", 
                    "pip", 
                    "install", 
                    package
                ])
                
    def _initialize_kernel(self):
        """Set up common imports and configurations in the kernel"""
        setup_code = """
        import matplotlib.pyplot as plt
        import numpy as np
        import pandas as pd
        import seaborn as sns
        import plotly.express as px
        from IPython.display import HTML, display, Markdown
        
        # Configure matplotlib for inline plotting
        %matplotlib inline
        plt.style.use('seaborn')
        """
        self.run_code(setup_code)
        
    def parse_output(
        self,
        execution: Execution,
        output: str,
        on_stdout: Optional[Callable[[OutputMessage], Any]] = None,
        on_stderr: Optional[Callable[[OutputMessage], Any]] = None,
        on_result: Optional[Callable[[Result], Any]] = None,
        on_error: Optional[Callable[[ExecutionError], Any]] = None,
    ):
        """Parse a single output line from code execution"""
        try:
            data = json.loads(output)
            data_type = data.pop("type")

            if data_type == "result":
                result = Result(**data)
                execution.results.append(result)
                if on_result:
                    on_result(result)

            elif data_type == "stdout":
                execution.stdout.append(data["text"])
                if on_stdout:
                    on_stdout(OutputMessage(
                        line=data["text"],
                        timestamp=int(time.time() * 1000),
                        error=False
                    ))

            elif data_type == "stderr":
                execution.stderr.append(data["text"])
                if on_stderr:
                    on_stderr(OutputMessage(
                        line=data["text"],
                        timestamp=int(time.time() * 1000),
                        error=True
                    ))

            elif data_type == "error":
                error = ExecutionError(
                    name=data["name"],
                    value=data["value"],
                    traceback=data["traceback"]
                )
                execution.error = error
                if on_error:
                    on_error(error)

            elif data_type == "number_of_executions":
                execution.execution_count = data["execution_count"]

        except Exception as e:
            print(f"Error parsing output: {e}")
            print(f"Raw output: {output}")
        
    def run_code(
        self, 
        code: str, 
        on_stdout: Optional[Callable[[OutputMessage], Any]] = None,
        on_stderr: Optional[Callable[[OutputMessage], Any]] = None,
        on_result: Optional[Callable[[Result], Any]] = None,
        on_error: Optional[Callable[[ExecutionError], Any]] = None,
    ) -> Execution:
        """Execute code and handle all outputs"""
        execution = Execution()
        msg_id = self.client.execute(code, store_history=True)
        
        while True:
            try:
                msg = self.client.get_iopub_msg(timeout=10)
                msg_type = msg['header']['msg_type']
                content = msg['content']
                
                if msg_type == 'stream':
                    output = json.dumps({
                        "type": "stdout" if content['name'] == 'stdout' else "stderr",
                        "text": content['text'],
                        "timestamp": int(time.time() * 1000)
                    })
                    self.parse_output(execution, output, on_stdout, on_stderr, on_result, on_error)
                        
                elif msg_type in ['execute_result', 'display_data']:
                    result_data = {
                        "type": "result",
                        "is_main_result": msg_type == 'execute_result'
                    }
                    
                    data = content['data']
                    for mime, value in data.items():
                        if mime == 'text/plain':
                            result_data['text'] = value
                        elif mime == 'text/html':
                            result_data['html'] = value
                        elif mime == 'image/png':
                            # PNG data comes as base64
                            result_data['png'] = value
                        elif mime == 'image/jpeg':
                            result_data['jpeg'] = value
                        elif mime == 'image/svg+xml':
                            result_data['svg'] = value
                        elif mime == 'application/json':
                            result_data['json'] = value
                        elif mime == 'text/markdown':
                            result_data['markdown'] = value
                    
                    self.parse_output(execution, json.dumps(result_data), on_stdout, on_stderr, on_result, on_error)
                    
                elif msg_type == 'error':
                    error_data = {
                        "type": "error",
                        "name": content['ename'],
                        "value": content['evalue'],
                        "traceback": content['traceback']
                    }
                    self.parse_output(execution, json.dumps(error_data), on_stdout, on_stderr, on_result, on_error)
                    
                elif msg_type == 'status' and content['execution_state'] == 'idle':
                    break
                    
            except Exception as e:
                error_data = {
                    "type": "error",
                    "name": type(e).__name__,
                    "value": str(e),
                    "traceback": []
                }
                self.parse_output(execution, json.dumps(error_data), on_stdout, on_stderr, on_result, on_error)
                break
                
        return execution

    def cleanup(self):
        """Shutdown the kernel"""
        self.client.shutdown()
        self.km.shutdown_kernel()
