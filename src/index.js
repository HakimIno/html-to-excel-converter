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

        this.pythonScriptPath = path.join(basePath, 'python', 'converterv3.py');
        this.maxChunkSize = options.maxChunkSize || 5 * 1024 * 1024; // 5MB
        this.timeout = options.timeout || 10 * 60 * 1000; // 10 minutes
        this.maxBuffer = options.maxBuffer || 50 * 1024 * 1024; // 50MB
        this.debug = options.debug || false;
    }

    async convertToBuffer(html) {
        if (!html || typeof html !== 'string') {
            throw new Error('HTML content must be a non-empty string');
        }

        try {
            // Configure Python shell options
            const options = {
                mode: 'text',
                pythonPath: this.pythonPath,
                pythonOptions: ['-u'],  // Unbuffered output
                scriptPath: path.dirname(this.pythonScriptPath),
                env: {
                    ...process.env,
                    PYTHONPATH: path.join(__dirname, 'python'),
                    PYTHONUNBUFFERED: '1'
                }
            };

            // Run Python script
            return new Promise((resolve, reject) => {
                let pythonOutput = '';
                let pythonError = '';
                
                const pyshell = new PythonShell('converterv3.py', options);

                // Send HTML content and request buffer output
                const input = JSON.stringify({
                    html: html,
                    buffer: true
                });
                
                pyshell.send(input);

                pyshell.on('message', function (message) {
                    if (this.debug) {
                        console.log('[Python Message]:', message);
                    }
                    pythonOutput += message + '\n';
                }.bind(this));

                pyshell.on('stderr', function (stderr) {
                    if (stderr.trim()) {
                        pythonError += stderr + '\n';
                        if (this.debug) {
                            console.log('[Python Debug]:', stderr);
                        }
                    }
                }.bind(this));

                pyshell.end(function (err) {
                    if (err) {
                        reject(new Error(`Python error: ${err.message}\n${pythonError}`));
                        return;
                    }

                    try {
                        // Get the last line that contains valid JSON
                        const lines = pythonOutput.trim().split('\n');
                        let lastValidJson = null;
                        
                        for (const line of lines) {
                            try {
                                const trimmedLine = line.trim();
                                if (!trimmedLine) continue;
                                
                                const json = JSON.parse(trimmedLine);
                                if (json && (json.success !== undefined || json.error)) {
                                    lastValidJson = json;
                                }
                            } catch (e) {
                                if (this.debug) {
                                    console.log('[Parse Debug]:', e.message, 'for line:', line);
                                }
                                continue;
                            }
                        }

                        if (!lastValidJson) {
                            reject(new Error(`No valid JSON response from Python\nOutput: ${pythonOutput}\nError: ${pythonError}`));
                            return;
                        }

                        if (lastValidJson.error) {
                            reject(new Error(lastValidJson.error));
                            return;
                        }

                        if (this.debug) {
                            console.log('[Success]: Conversion completed');
                        }

                        // Return the base64 data
                        resolve(lastValidJson);
                    } catch (e) {
                        reject(new Error(`Failed to process Python output: ${e.message}\nOutput: ${pythonOutput}\nError: ${pythonError}`));
                    }
                }.bind(this));
            });
        } catch (error) {
            throw new Error(`Conversion failed: ${error.message}`);
        }
    }

    async convert(html, outputPath) {
        if (!html || typeof html !== 'string') {
            throw new Error('HTML content must be a non-empty string');
        }

        if (!outputPath || typeof outputPath !== 'string') {
            throw new Error('Output path must be a non-empty string');
        }

        try {
            // Create a temporary file to store the HTML content
            const tempFile = path.join(os.tmpdir(), crypto.randomBytes(16).toString('hex') + '.html');
            fs.writeFileSync(tempFile, html, 'utf8');
            
            // Configure Python shell options
            const options = {
                mode: 'text',
                pythonPath: this.pythonPath,
                pythonOptions: ['-u'],  // Unbuffered output
                scriptPath: path.dirname(this.pythonScriptPath),
                args: [tempFile, outputPath],  // Pass both input and output paths
                env: {
                    ...process.env,
                    PYTHONPATH: path.join(__dirname, 'python'),
                    PYTHONUNBUFFERED: '1'  // Force unbuffered output
                }
            };

            // Run Python script
            return new Promise((resolve, reject) => {
                let pythonOutput = '';
                let pythonError = '';
                
                const pyshell = new PythonShell('html_to_excel.py', options);

                pyshell.on('message', function (message) {
                    if (this.debug) {
                        console.log('[Python Message]:', message);
                    }
                    pythonOutput += message + '\n';
                }.bind(this));

                pyshell.on('stderr', function (stderr) {
                    if (stderr.trim()) {
                        pythonError += stderr + '\n';
                        if (this.debug) {
                            console.log('[Python Debug]:', stderr);
                        }
                    }
                }.bind(this));

                pyshell.end(function (err) {
                    // Clean up temporary file
                    try {
                        fs.unlinkSync(tempFile);
                    } catch (e) {
                        console.warn('Failed to clean up temp file:', e.message);
                    }

                    if (err) {
                        reject(new Error(`Python error: ${err.message}\n${pythonError}`));
                        return;
                    }

                    try {
                        // Get the last line that contains valid JSON
                        const lines = pythonOutput.trim().split('\n');
                        let lastValidJson = null;
                        
                        for (const line of lines) {
                            try {
                                const trimmedLine = line.trim();
                                if (!trimmedLine) continue;
                                
                                const json = JSON.parse(trimmedLine);
                                if (json && (json.success !== undefined || json.error)) {
                                    lastValidJson = json;
                                }
                            } catch (e) {
                                if (this.debug) {
                                    console.log('[Parse Debug]:', e.message, 'for line:', line);
                                }
                                continue;
                            }
                        }

                        if (!lastValidJson) {
                            reject(new Error(`No valid JSON response from Python\nOutput: ${pythonOutput}\nError: ${pythonError}`));
                            return;
                        }

                        if (lastValidJson.error) {
                            reject(new Error(lastValidJson.error));
                            return;
                        }

                        if (this.debug) {
                            console.log('[Success]: Conversion completed');
                        }

                        resolve(lastValidJson);
                    } catch (e) {
                        reject(new Error(`Failed to process Python output: ${e.message}\nOutput: ${pythonOutput}\nError: ${pythonError}`));
                    }
                }.bind(this));
            });
        } catch (error) {
            throw new Error(`Conversion failed: ${error.message}`);
        }
    }
}

module.exports = HTMLToExcelConverter; 