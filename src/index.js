const { PythonShell } = require('python-shell');
const path = require('path');
const { execSync } = require('child_process');
const os = require('os');
const fs = require('fs');

class HTMLToExcelConverter {
  constructor() {
    this.pythonScriptPath = path.join(__dirname, 'python', 'html_to_excel.py');
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
    return new Promise((resolve, reject) => {
      let options = {
        pythonPath: pythonCmd,
        args: [htmlContent],
        env: {
          ...process.env,
          VIRTUAL_ENV: this.venvPath,
          PATH: `${path.dirname(pythonCmd)}${path.delimiter}${process.env.PATH}`
        }
      };

      PythonShell.run(this.pythonScriptPath, options).then(messages => {
        const result = JSON.parse(messages[messages.length - 1]);
        
        if (result.success) {
          const buffer = Buffer.from(result.data, 'base64');
          resolve(buffer);
        } else {
          reject(new Error(result.error));
        }
      }).catch(err => {
        reject(err);
      });
    });
  }
}

module.exports = HTMLToExcelConverter; 