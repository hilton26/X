# source: Gemini in response to "script with python using Flask to have an html page with a row
# of button which, when clicked, execute python scripts"

from flask import Flask, render_template, request, jsonify
import subprocess
import threading
import os

app = Flask(__name__)

# Define a dictionary of scripts with their names and paths
# IMPORTANT: Make sure these scripts are executable and have the necessary
# permissions.
# For demonstration, let's create some dummy scripts.
SCRIPTS = {
    "script1": "python scripts/script1.py",
    "script2": "python scripts/script2.py",
    "script3": "python scripts/script3.py",
    # Add more scripts as needed
}

# Ensure the 'scripts' directory exists
if not os.path.exists("scripts"):
    os.makedirs("scripts")

# Create dummy scripts for demonstration purposes
for i in range(1, 4):
    with open(f"scripts/script{i}.py", "w") as f:
        f.write(f"import time\n")
        f.write(f"print(f'Executing script {i}...')\n")
        f.write(f"time.sleep(2)\n")
        f.write(f"print(f'Script {i} finished.')\n")


@app.route("/")
def index():
    """Renders the main HTML page with the buttons."""
    return render_template("index.html", scripts=SCRIPTS.keys())


@app.route("/execute_script/<script_name>", methods=["POST"])
def execute_script(script_name):
    """
    Executes the specified Python script in a separate thread
    and returns a JSON response.
    """
    if script_name not in SCRIPTS:
        return jsonify(
            {"status": "error", "message": f'Script "{script_name}" not found.'}
        ), 404

    script_command = SCRIPTS[script_name].split()

    def run_script():
        try:
            # Use subprocess.Popen for non-blocking execution
            # You might want to capture stdout/stderr for logging or display later
            process = subprocess.Popen(
                script_command,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
            )
            stdout, stderr = process.communicate()
            if process.returncode == 0:
                print(f"Script '{script_name}' executed successfully.")
                print(f"Stdout: {stdout}")
            else:
                print(
                    f"Script '{script_name}' failed with error code {process.returncode}."
                )
                print(f"Stderr: {stderr}")
        except Exception as e:
            print(f"Error executing script '{script_name}': {e}")

    # Run the script in a separate thread to avoid blocking the Flask app
    thread = threading.Thread(target=run_script)
    thread.start()

    return jsonify(
        {"status": "success", "message": f'Script "{script_name}" execution started.'}
    )


if __name__ == "__main__":
    # Ensure a 'scripts' directory exists and create dummy scripts
    if not os.path.exists("scripts"):
        os.makedirs("scripts")

    for i in range(1, 4):
        with open(f"scripts/script{i}.py", "w") as f:
            f.write(f"import time\n")
            f.write(f"print(f'Executing script {i}...')\n")
            f.write(f"time.sleep(2)\n")
            f.write(f"print(f'Script {i} finished.')\n")

    app.run(debug=True, port = 5002)  # debug=True allows for automatic reloading on code changes
