const { PythonShell } = require('python-shell');
const fs = require('fs');
const path = require('path');
const os = require('os');
const crypto = require('crypto');

class HTMLToExcelConverter {
    constructor(options = {}) {
        // Set paths relative to the current file location
        const basePath = path.join(__dirname);
        this.pythonPath = path.join(basePath, 'venv', 'bin', 'python3');
        if (os.platform() === 'win32') {
            this.pythonPath = path.join(basePath, 'venv', 'Scripts', 'python.exe');
        }

        this.pythonScriptPath = path.join(basePath, 'python', 'html_to_excel.py');
        this.maxChunkSize = options.maxChunkSize || 5 * 1024 * 1024; // 5MB
        this.timeout = options.timeout || 10 * 60 * 1000; // 10 minutes
        this.maxBuffer = options.maxBuffer || 50 * 1024 * 1024; // 50MB
    }

    async convert(html) {
        if (!html || typeof html !== 'string') {
            throw new Error('HTML content must be a non-empty string');
        }

        try {
            // Create a temporary file to store the HTML content
            const tempFile = path.join(os.tmpdir(), crypto.randomBytes(16).toString('hex') + '.html');
            fs.writeFileSync(tempFile, html, 'utf8');
            
            // Configure Python shell options
            const options = {
                mode: 'text',
                pythonPath: this.pythonPath,
                pythonOptions: ['-u'],
                scriptPath: path.dirname(this.pythonScriptPath),
                args: [tempFile],
                env: {
                    ...process.env,
                    PYTHONPATH: path.join(__dirname, 'python')
                }
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
                        console.warn('Failed to clean up temp file:', e.message);
                    }

                    if (err) {
                        if (pythonError) {
                            reject(new Error(`Python error: ${pythonError.trim()}`));
                        } else {
                            reject(new Error(`Python script error: ${err.message}\nOutput: ${pythonOutput}`));
                        }
                        return;
                    }

                    try {
                        // Get the last valid JSON object from the output
                        const lines = pythonOutput.trim().split('\n');
                        let lastValidJson = null;
                        let parseError = null;
                        
                        for (const line of lines) {
                            try {
                                const json = JSON.parse(line);
                                if (json.success !== undefined) {
                                    lastValidJson = json;
                                }
                            } catch (e) {
                                parseError = e;
                            }
                        }

                        if (!lastValidJson) {
                            if (pythonError) {
                                reject(new Error(`Python error: ${pythonError.trim()}`));
                            } else if (parseError) {
                                reject(new Error(`Failed to parse Python output: ${parseError.message}\nOutput: ${pythonOutput}`));
                            } else {
                                reject(new Error(`No valid JSON output from Python script\nOutput: ${pythonOutput}`));
                            }
                            return;
                        }

                        if (!lastValidJson.success) {
                            reject(new Error(lastValidJson.error || `Unknown error\nOutput: ${pythonOutput}`));
                            return;
                        }

                        resolve(lastValidJson.data);
                    } catch (e) {
                        reject(new Error(`Failed to parse Python output: ${e.message}\nOutput: ${pythonOutput}`));
                    }
                });
            });
        } catch (error) {
            throw new Error(`Conversion failed: ${error.message}`);
        }
    }
}

module.exports = HTMLToExcelConverter; 