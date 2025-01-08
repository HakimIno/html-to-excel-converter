const { PythonShell } = require('python-shell');
const path = require('path');
const { execSync } = require('child_process');

class HTMLToExcelConverter {
  constructor() {
    this.pythonScriptPath = path.join(__dirname, 'python', 'html_to_excel.py');
    this.checkPythonDependencies();
  }

  getPythonCommand() {
    try {
      execSync('python3 --version');
      return 'python3';
    } catch (err) {
      try {
        execSync('python --version');
        return 'python';
      } catch (err2) {
        throw new Error('Python is not installed or not accessible');
      }
    }
  }

  checkPythonDependencies() {
    const pythonCmd = this.getPythonCommand();
    const dependencies = ['bs4', 'pandas', 'openpyxl', 'cssutils'];
    const missingDeps = [];

    dependencies.forEach(dep => {
      try {
        execSync(`${pythonCmd} -c "import ${dep}"`, { stdio: 'ignore' });
      } catch (err) {
        console.error(`Failed to import ${dep}:`, err.message);
        missingDeps.push(dep);
      }
    });

    if (missingDeps.length > 0) {
      throw new Error(
        `Missing Python dependencies: ${missingDeps.join(', ')}.\n` +
        'Please install them using one of these commands:\n' +
        `${pythonCmd} -m pip install beautifulsoup4 pandas openpyxl cssutils lxml\n` +
        `${pythonCmd} -m pip install --user beautifulsoup4 pandas openpyxl cssutils lxml\n` +
        'pip3 install beautifulsoup4 pandas openpyxl cssutils lxml'
      );
    }
  }

  async convertHtmlToExcel(htmlContent) {
    const pythonCmd = this.getPythonCommand();
    return new Promise((resolve, reject) => {
      let options = {
        pythonPath: pythonCmd,
        args: [htmlContent]
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