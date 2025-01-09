const { PythonShell } = require('python-shell');
const fs = require('fs');
const path = require('path');
const os = require('os');
const crypto = require('crypto');

class HTMLToExcelConverter {
    constructor(options = {}) {
        this.pythonScriptPath = path.join(__dirname, 'python', 'html_to_excel.py');
        this.venvPath = options.venvPath || path.join(__dirname, 'venv');
        this.maxChunkSize = options.maxChunkSize || 5 * 1024 * 1024; // 5MB
        this.timeout = options.timeout || 10 * 60 * 1000; // 10 minutes
        this.maxBuffer = options.maxBuffer || 50 * 1024 * 1024; // 50MB
    }

    async convert(html) {
        try {
            // Create a temporary file to store the HTML content
            const tempFile = path.join(os.tmpdir(), crypto.randomBytes(16).toString('hex') + '.html');
            fs.writeFileSync(tempFile, html);

            // Get Python executable path
            const pythonPath = await this.getPythonCommand();
            
            // Configure Python shell options
            const options = {
                mode: 'text',
                pythonPath: pythonPath,
                pythonOptions: ['-u'],
                scriptPath: path.dirname(this.pythonScriptPath),
                args: [tempFile],
                env: process.env
            };

            // Run Python script
            return new Promise((resolve, reject) => {
                let pythonOutput = '';
                let pythonError = '';
                
                const pyshell = new PythonShell('html_to_excel.py', options);

                pyshell.on('message', function (message) {
                    pythonOutput += message + '\n';
                });

                pyshell.on('stderr', function (stderr) {
                    if (stderr.trim()) {
                        pythonError += stderr + '\n';
                    }
                });

                pyshell.end(function (err) {
                    // Clean up temporary file
                    try {
                        fs.unlinkSync(tempFile);
                    } catch (e) {
                        // Ignore cleanup errors
                    }

                    if (err) {
                        reject(new Error(`Python script error: ${err.message}`));
                        return;
                    }

                    try {
                        // Get the last valid JSON object from the output
                        const lines = pythonOutput.trim().split('\n');
                        let lastValidJson = null;
                        
                        for (const line of lines) {
                            try {
                                const json = JSON.parse(line);
                                if (json.success !== undefined) {
                                    lastValidJson = json;
                                }
                            } catch (e) {
                                // Skip non-JSON lines
                            }
                        }

                        if (!lastValidJson) {
                            reject(new Error('No valid JSON output from Python script'));
                            return;
                        }

                        if (!lastValidJson.success) {
                            reject(new Error(lastValidJson.error || 'Unknown error'));
                            return;
                        }

                        resolve(lastValidJson.data);
                    } catch (e) {
                        reject(new Error(`Failed to parse Python output: ${e.message}`));
                    }
                });
            });
        } catch (error) {
            throw new Error(`Conversion failed: ${error.message}`);
        }
    }

    async getPythonCommand() {
        const isWindows = os.platform() === 'win32';
        const pythonExecutable = isWindows ? 'python.exe' : 'python';
        const venvPythonPath = path.join(this.venvPath, 'bin', pythonExecutable);
        
        if (fs.existsSync(venvPythonPath)) {
            return venvPythonPath;
        }
        
        throw new Error('Python virtual environment not found');
    }
}

module.exports = HTMLToExcelConverter; 