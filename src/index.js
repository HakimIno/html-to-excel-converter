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

    // Fallback to system Python
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
    const dependencies = ['bs4', 'pandas', 'openpyxl', 'cssutils'];
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
        'Please run: npm run install-python-deps'
      );
    }
  }

  async convertHtmlToExcel(htmlContent) {
    const pythonCmd = this.getPythonCommand();
    const outputPath = path.join(os.tmpdir(), `excel-${crypto.randomBytes(8).toString('hex')}.xlsx`);

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
        // สร้าง PythonShell instance พร้อมกับ configuration
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

        // รับ error output
        pyshell.stderr.on('data', (data) => {
          errorOutput += data;
          console.error('Python stderr:', data);
        });

        // รับ standard output
        pyshell.stdout.on('data', (data) => {
          stdOutput += data;
          console.log('Python stdout:', data);
          
          // ตรวจสอบ success message
          if (data.includes('"success": true')) {
            hasSuccessMessage = true;
          }
        });

        // จัดการเมื่อ process จบการทำงาน
        pyshell.on('close', () => {
          console.log('Process closed');
          
          if (hasSuccessMessage && fs.existsSync(outputPath)) {
            try {
              const buffer = fs.readFileSync(outputPath);
              cleanup();
              resolve(buffer);
            } catch (readErr) {
              console.error('Failed to read output file:', readErr);
              cleanup();
              reject(readErr);
            }
          } else {
            console.error('Process error output:', errorOutput);
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

        // จัดการ error จาก PythonShell
        pyshell.on('error', (err) => {
          console.error('PythonShell error:', err);
          cleanup();
          reject(err);
        });

        // ตั้ง timeout
        timeout = setTimeout(() => {
          console.warn('Conversion timeout');
          cleanup();
          reject(new Error('Conversion timeout after 30 seconds'));
        }, 30000);

        // ส่ง HTML content
        console.log('Sending HTML content to Python...');
        pyshell.send(JSON.stringify({
          html: htmlContent,
          output: outputPath
        }));
        pyshell.stdin.end();

      } catch (error) {
        console.error('Setup error:', error);
        cleanup();
        reject(error);
      }
    });
  }
}

module.exports = HTMLToExcelConverter; 