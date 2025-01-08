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
      // สร้าง PythonShell instance พร้อมกับ configuration
      const pyshell = new PythonShell(this.pythonScriptPath, {
        mode: 'json',
        pythonPath: pythonCmd,
        pythonOptions: ['-u'], // unbuffered output
        env: {
          ...process.env,
          VIRTUAL_ENV: this.venvPath,
          PATH: `${path.dirname(pythonCmd)}${path.delimiter}${process.env.PATH}`
        },
        stdin: true
      });

      // ส่ง data ผ่าน stdin
      pyshell.send(JSON.stringify({
        html: htmlContent,
        output: outputPath
      }));

      // รับ response จาก Python script
      pyshell.on('message', (message) => {
        if (message.error) {
          // ถ้ามี error ให้ลบไฟล์ output (ถ้ามี) และ reject
          if (fs.existsSync(outputPath)) {
            try { fs.unlinkSync(outputPath); } catch {}
          }
          reject(new Error(message.error));
        } else if (message.success) {
          // ถ้าสำเร็จ อ่านไฟล์ Excel เป็น buffer
          try {
            const buffer = fs.readFileSync(outputPath);
            resolve(buffer);
          } catch (err) {
            reject(err);
          } finally {
            // ลบไฟล์ output หลังจากอ่านเสร็จ
            try { fs.unlinkSync(outputPath); } catch {}
          }
        }
      });

      // จัดการ error จาก PythonShell
      pyshell.on('error', (err) => {
        if (fs.existsSync(outputPath)) {
          try { fs.unlinkSync(outputPath); } catch {}
        }
        reject(err);
      });

      // จัดการเมื่อ process จบการทำงาน
      pyshell.end((err) => {
        if (err) {
          if (fs.existsSync(outputPath)) {
            try { fs.unlinkSync(outputPath); } catch {}
          }
          reject(err);
        }
      });

      // ตั้ง timeout เพื่อป้องกันการค้าง
      setTimeout(() => {
        pyshell.terminate();
        if (fs.existsSync(outputPath)) {
          try { fs.unlinkSync(outputPath); } catch {}
        }
        reject(new Error('Conversion timeout after 30 seconds'));
      }, 30000);
    });
  }
}

module.exports = HTMLToExcelConverter; 