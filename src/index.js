const { PythonShell } = require('python-shell');
const path = require('path');
const { execSync } = require('child_process');
const os = require('os');
const fs = require('fs');
const crypto = require('crypto');

class HTMLToExcelConverter {
  constructor() {
    this.pythonScriptPath = path.join(__dirname, 'python', 'converter.py');
    this.venvPath = this.getVenvPath();
    this.checkPythonDependencies();
  }

  getVenvPath() {
    const envPath = path.join(__dirname, '..', '.env');
    if (fs.existsSync(envPath)) {
      const envContent = fs.readFileSync(envPath, 'utf8');
      const match = envContent.match(/VENV_PATH=(.+)/);
      if (match) return match[1];
    }
    return null;
  }

  getPythonCommand() {
    if (this.venvPath) {
      const isWindows = os.platform() === 'win32';
      return isWindows ?
        path.join(this.venvPath, 'Scripts', 'python.exe') :
        path.join(this.venvPath, 'bin', 'python');
    }

    const platform = os.platform();
    const commands = {
      darwin: ['python3', 'python'],
      linux: ['python3', 'python'],
      win32: ['python', 'py']
    };

    const pythonCommands = commands[platform] || ['python3', 'python'];

    for (const cmd of pythonCommands) {
      try {
        execSync(`${cmd} --version`);
        return cmd;
      } catch (err) {
        continue;
      }
    }

    throw new Error('Python is not installed or not accessible');
  }

  checkPythonDependencies() {
    const pythonCmd = this.getPythonCommand();
    const dependencies = ['bs4', 'xlsxwriter'];
    const missingDeps = [];

    dependencies.forEach(dep => {
      try {
        execSync(`${pythonCmd} -c "import ${dep}"`, {
          stdio: 'ignore',
          env: {
            ...process.env,
            VIRTUAL_ENV: this.venvPath,
            PATH: `${path.dirname(pythonCmd)}${path.delimiter}${process.env.PATH}`
          }
        });
      } catch (err) {
        missingDeps.push(dep);
      }
    });

    if (missingDeps.length > 0) {
      throw new Error(
        `Missing Python dependencies: ${missingDeps.join(', ')}.\n` +
        'Please run: pip install ' + missingDeps.join(' ')
      );
    }
  }

  async convertHtmlToExcel(htmlContent, options = {}) {
    const pythonCmd = this.getPythonCommand();
    const outputPath = path.join(os.tmpdir(), `excel-${crypto.randomBytes(8).toString('hex')}.xlsx`);
    const timeout_ms = options.timeout || 30000;

    return new Promise((resolve, reject) => {
      let pyshell = null;
      let timeout = null;

      const cleanup = () => {
        if (timeout) {
          clearTimeout(timeout);
          timeout = null;
        }
        if (fs.existsSync(outputPath)) {
          try { fs.unlinkSync(outputPath); } catch {}
        }
        if (pyshell) {
          try { pyshell.terminate(); } catch {}
        }
      };

      try {
        pyshell = new PythonShell(this.pythonScriptPath, {
          mode: 'text',
          pythonPath: pythonCmd,
          pythonOptions: ['-u'],
          env: {
            ...process.env,
            VIRTUAL_ENV: this.venvPath,
            PATH: `${path.dirname(pythonCmd)}${path.delimiter}${process.env.PATH}`
          },
          stdin: true
        });

        let errorOutput = '';
        let stdOutput = '';
        let hasSuccessMessage = false;

        pyshell.stderr.on('data', (data) => {
          errorOutput += data;
          console.error('Python stderr:', data);
        });

        pyshell.stdout.on('data', (data) => {
          stdOutput += data;
          if (data.includes('"success": true')) {
            hasSuccessMessage = true;
          }
        });

        pyshell.on('close', () => {
          if (hasSuccessMessage && fs.existsSync(outputPath)) {
            try {
              const buffer = fs.readFileSync(outputPath);
              cleanup();
              resolve(buffer);
            } catch (readErr) {
              cleanup();
              reject(readErr);
            }
          } else {
            cleanup();
            if (errorOutput) {
              try {
                const errorJson = JSON.parse(errorOutput);
                reject(new Error(errorJson.error || 'Unknown error'));
              } catch (e) {
                reject(new Error(errorOutput || 'Unknown error'));
              }
            } else {
              reject(new Error('Conversion failed without error message'));
            }
          }
        });

        pyshell.on('error', (err) => {
          cleanup();
          reject(err);
        });

        timeout = setTimeout(() => {
          cleanup();
          reject(new Error(`Conversion timeout after ${timeout_ms/1000} seconds`));
        }, timeout_ms);

        pyshell.send(JSON.stringify({
          html: htmlContent,
          output: outputPath
        }));
        pyshell.stdin.end();

      } catch (error) {
        cleanup();
        reject(error);
      }
    });
  }
}

module.exports = HTMLToExcelConverter; 