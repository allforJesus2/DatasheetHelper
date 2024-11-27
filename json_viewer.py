import tkinter as tk
from tkinter import ttk
from tkinterhtml import TkinterWeb
import json
from pathlib import Path


class JsonViewer3DWindow:
    def __init__(self, parent, json_data):
        self.window = tk.Toplevel(parent)
        self.window.title("3D Results Viewer")
        self.window.geometry("800x600")

        # Load React component
        component_path = Path(__file__).parent / "components" / "JsonViewer3D.jsx"
        if not component_path.parent.exists():
            component_path.parent.mkdir(parents=True)

        with open(component_path, "w") as f:
            f.write("""
import React, { useState, useEffect } from 'react';
import { Card, CardContent } from '@/components/ui/card';
import { FileText } from 'lucide-react';

const JsonViewer3D = ({ data = {} }) => {
    // Component code here - copy the entire component from previous response
};

export default JsonViewer3D;
""")

        # Create web view
        self.web_view = TkinterWeb(self.window)
        self.web_view.pack(fill='both', expand=True)

        # Load component with data
        html_content = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <script src="https://unpkg.com/react@17/umd/react.development.js"></script>
            <script src="https://unpkg.com/react-dom@17/umd/react-dom.development.js"></script>
            <script src="https://cdnjs.cloudflare.com/ajax/libs/babel-standalone/6.26.0/babel.min.js"></script>
        </head>
        <body>
            <div id="root"></div>
            <script>
                window.jsonData = {json.dumps(json_data)};
            </script>
            <script type="text/babel" src="{component_path}"></script>
        </body>
        </html>
        """
        self.web_view.load_html(html_content)