const { PythonShell } = require('python-shell');
const path = require('path');

class HTMLToExcelConverter {
  constructor() {
    this.pythonScriptPath = path.join(__dirname, 'python', 'html_to_excel.py');
  }

  async convertHtmlToExcel(htmlContent) {
    return new Promise((resolve, reject) => {
      let options = {
        args: [htmlContent]
      };

      PythonShell.run(this.pythonScriptPath, options).then(messages => {
        const result = JSON.parse(messages[messages.length - 1]);
        
        if (result.success) {
          // Decode base64 back to buffer
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