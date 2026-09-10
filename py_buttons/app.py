# source: Gemini in response to "script with python using Flask to have an
# html page with a row # of button which, when clicked, execute python scripts"

# For easy demonstration, app.py automatically creates a scripts directory and
# some basic Python files (script1.py, etc.) that simply print a message and
# sleep for a bit.

from flask import Flask, render_template, request, jsonify
import subprocess
import threading
import os

app = Flask(__name__)  # initialises the Flask application

# Define a dictionary of scripts with their names and paths
# IMPORTANT: Make sure these scripts are executable and have the necessary
# permissions.
# For demonstration, let's create some dummy scripts.
SCRIPTS = {
    "script1": "python scripts/script1.py",
    "script2": "python scripts/script1.py",
    "script3": "python scripts/script3.py",
    "script4": "python scripts/script4.py",
    "script5": "python eagle_LT_download_and_merge_csv.py",
    # Add more scripts as needed
    # Directly executing user-provided script names or paths can be a huge
    # security risk. In this example, the SCRIPTS dictionary hardcodes the
    # allowed scripts. NEVER allow users to input arbitrary script paths or
    # commands.
}
# This maps friendly names (used in the buttons and URLs) to the actual
# commands to execute the scripts.

# Ensure the 'scripts' directory exists
if not os.path.exists("scripts"):
    os.makedirs("scripts")

# Create dummy scripts for demonstration purposes
for i in range(1, len(SCRIPTS)+1):
    with open(f"scripts/script{i}.py", "w") as f:
        f.write(f"import time\n")
        f.write(f"print(f'Executing script {i}...')\n")
        f.write(f"time.sleep(2)\n")
        f.write(f"print(f'Script {i} finished.')\n")


@app.route("/")
def index():
    """Renders the main HTML page with the buttons."""
    # This route renders the index.html file and passes the names of the 
    # scripts (from the SCRIPTS dictionary) to the template.
    return render_template("index.html", scripts=SCRIPTS.keys())


@app.route("/execute_script/<script_name>", methods=["POST"])
#  This is the endpoint that the JavaScript in your HTML will call when a 
#  button is clicked.
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
            # You might want to capture stdout/stderr for logging
            # or display later
            process = subprocess.Popen(
                script_command,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
            )
            # This is the core of executing the Python script.
            # subprocess.Popen starts a new process.
            # We use stdout=subprocess.PIPE and stderr=subprocess.PIPE to
            # capture the output and errors of the script, respectively.
            # text=True decodes the output as text.
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
    # CRUCIALLY, script execution is wrapped in a threading.Thread. This prevents
    # the Flask web server from blocking while the script runs. If you didn't do
    # this, clicking a button would freeze your web interface until the script
    # finished.
    thread.start()

    return jsonify(
        {
            "status": "success",
            "message": (f'Script "{script_name}" execution started.'),
        }
    )
    # Returns a JSON response to the client (web browser), indicating whether
    # the script execution started successfully.


if __name__ == "__main__":
    # Ensure a 'scripts' directory exists and create dummy scripts
    if not os.path.exists("scripts"):
        os.makedirs("scripts")

    for i in range(1, len(SCRIPTS) + 1):
        with open(f"scripts/script{i}.py", "w") as f:
            f.write(f"import time\n")
            f.write(f"print(f'Executing script {i}...')\n")
            f.write(f"time.sleep(2)\n")
            f.write(f"print(f'Script {i} finished.')\n")

    app.run(
        debug=True, port=5005
    )  # debug=True allows for automatic reloading on code changes
