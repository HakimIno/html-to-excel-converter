const { PythonShell } = require('python-shell');
const path = require('path');
const { execSync } = require('child_process');
const os = require('os');
const fs = require('fs');
const crypto = require('crypto');

class HTMLToExcelConverter {
  constructor() {
    this.pythonScriptPath = path.join(__dirname, 'python', 'html_to_excel.py');
    this.venvPath = this.getVenvPath();
    this.tempDir = path.join(os.tmpdir(), 'html-to-excel');
    this.checkPythonDependencies();
    this.createTempDir();
  }

  createTempDir() {
    if (!fs.existsSync(this.tempDir)) {
      fs.mkdirSync(this.tempDir, { recursive: true });
    }
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
    
    // สร้าง temporary file สำหรับ HTML content
    const tempHtmlFile = path.join(
      this.tempDir, 
      `input-${crypto.randomBytes(8).toString('hex')}.html`
    );
    
    try {
      // เขียน HTML content ลง temporary file
      fs.writeFileSync(tempHtmlFile, htmlContent, 'utf8');

      return new Promise((resolve, reject) => {
        let options = {
          pythonPath: pythonCmd,
          args: [tempHtmlFile], // ส่ง file path แทน content
          env: {
            ...process.env,
            VIRTUAL_ENV: this.venvPath,
            PATH: `${path.dirname(pythonCmd)}${path.delimiter}${process.env.PATH}`
          }
        };

        PythonShell.run(this.pythonScriptPath, options).then(messages => {
          try {
            const result = JSON.parse(messages[messages.length - 1]);
            
            if (result.success) {
              const buffer = Buffer.from(result.data, 'base64');
              resolve(buffer);
            } else {
              reject(new Error(result.error));
            }
          } catch (err) {
            reject(new Error('Failed to parse Python output'));
          }
        }).catch(err => {
          reject(err);
        }).finally(() => {
          // ลบ temporary file
          try {
            fs.unlinkSync(tempHtmlFile);
          } catch (err) {
            console.warn('Failed to delete temporary file:', tempHtmlFile);
          }
        });
      });
    } catch (error) {
      // ในกรณีที่มี error ให้ลบ temporary file ด้วย
      try {
        fs.unlinkSync(tempHtmlFile);
      } catch {}
      throw error;
    }
  }
}

module.exports = HTMLToExcelConverter; 