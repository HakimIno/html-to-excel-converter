const path = require('path');
const { spawn } = require('child_process');

class BaseConverter {
    constructor(options = {}) {
        this.maxBuffer = options.maxBuffer || 1024 * 1024;
        this.debug = options.debug || false;
    }

    async _runPythonScript(scriptName, input) {
        return new Promise((resolve, reject) => {
            const pythonPath = 'python3';
            const scriptPath = path.join(__dirname, 'python', scriptName);
            
            const pythonProcess = spawn(pythonPath, [scriptPath], {
                stdio: ['pipe', 'pipe', 'pipe']
            });

            let stdout = '';
            let stderr = '';

            pythonProcess.stdout.on('data', (data) => {
                stdout += data.toString();
            });

            pythonProcess.stderr.on('data', (data) => {
                if (this.debug) {
                    console.log('[Python Debug]:', data.toString());
                }
                stderr += data.toString();
            });

            pythonProcess.on('close', (code) => {
                if (code !== 0) {
                    reject(new Error(`Python process exited with code ${code}\n${stderr}`));
                    return;
                }

                try {
                    const result = JSON.parse(stdout);
                    if (!result.success) {
                        reject(new Error(result.error || 'Unknown error'));
                        return;
                    }
                    resolve(result);
                } catch (error) {
                    reject(new Error('No valid JSON response from Python'));
                }
            });

            pythonProcess.on('error', (error) => {
                reject(error);
            });

            pythonProcess.stdin.write(JSON.stringify(input));
            pythonProcess.stdin.end();
        });
    }
}

class HTMLToExcelConverter extends BaseConverter {
    async convert(html, outputPath) {
        return this._runPythonScript('converter.py', {
            html,
            output_path: outputPath
        });
    }

    async convertToBuffer(html) {
        return this._runPythonScript('converter.py', {
            html,
            return_buffer: true
        });
    }
}

// class HTMLToPDFConverter extends BaseConverter {
//     constructor(options = {}) {
//         super(options);
//         this.pdfOptions = options.pdf || {};
//     }

//     async convert(html, outputPath) {
//         return this._runPythonScript('html_to_pdf_v2.py', {
//             html,
//             output_path: outputPath,
//             options: this.pdfOptions
//         });
//     }
// }

module.exports = {
    HTMLToExcelConverter
}; 