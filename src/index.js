const { PythonShell } = require('python-shell');
const path = require('path');
const { execSync } = require('child_process');

class HTMLToExcelConverter {
  constructor() {
    this.pythonScriptPath = path.join(__dirname, 'python', 'html_to_excel.py');
    this.checkPythonDependencies();
  }

  checkPythonDependencies() {
    const dependencies = ['bs4', 'pandas', 'openpyxl', 'cssutils'];
    const missingDeps = [];

    dependencies.forEach(dep => {
      try {
        execSync(`python -c "import ${dep}"`, { stdio: 'ignore' });
      } catch (err) {
        missingDeps.push(dep);
      }
    });

    if (missingDeps.length > 0) {
      throw new Error(
        `Missing Python dependencies: ${missingDeps.join(', ')}. ` +
        'Please run: npm run install-python-deps'
      );
    }
  }

  async convertHtmlToExcel(htmlContent) {
    return new Promise((resolve, reject) => {
      let options = {
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