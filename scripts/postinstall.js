const { execSync } = require('child_process');
const path = require('path');
const os = require('os');
const fs = require('fs');

// สร้าง virtual environment path
const venvPath = path.join(__dirname, '..', 'venv');

function setupVirtualEnv(pythonCmd) {
  try {
    // สร้าง virtual environment
    if (!fs.existsSync(venvPath)) {
      console.log('Creating virtual environment...');
      execSync(`${pythonCmd} -m venv ${venvPath}`);
    }

    // Activate virtual environment และติดตั้ง dependencies
    const isWindows = os.platform() === 'win32';
    const activateCmd = isWindows ? 
      `${path.join(venvPath, 'Scripts', 'activate.bat')}` : 
      `. ${path.join(venvPath, 'bin', 'activate')}`;

    const pipCmd = isWindows ?
      path.join(venvPath, 'Scripts', 'pip') :
      path.join(venvPath, 'bin', 'pip');

    // ติดตั้ง dependencies ใน virtual environment
    const dependencies = [
      'beautifulsoup4>=4.9.3',
      'pandas>=1.3.0',
      'lxml>=4.9.0',
      'openpyxl>=3.0.7',
      'cssutils>=2.3.0'
    ];

    console.log('Installing dependencies in virtual environment...');
    dependencies.forEach(dep => {
      try {
        execSync(`${pipCmd} install ${dep}`, {
          stdio: 'inherit',
          shell: true,
          env: {
            ...process.env,
            VIRTUAL_ENV: venvPath,
            PATH: `${path.dirname(pipCmd)}${path.delimiter}${process.env.PATH}`
          }
        });
      } catch (err) {
        console.warn(`Warning: Failed to install ${dep}`);
      }
    });

    return true;
  } catch (error) {
    console.error('Error setting up virtual environment:', error.message);
    return false;
  }
}

function getPythonCommand() {
  const platform = os.platform();
  const commands = {
    darwin: ['python3', 'python'], // macOS
    linux: ['python3', 'python'],  // Linux
    win32: ['python', 'py']        // Windows
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

try {
  const pythonCmd = getPythonCommand();
  console.log(`Using Python command: ${pythonCmd}`);

  if (setupVirtualEnv(pythonCmd)) {
    console.log('Virtual environment setup completed successfully');
    
    // สร้างไฟล์ .env เพื่อเก็บ path ของ virtual environment
    const envContent = `VENV_PATH=${venvPath}`;
    fs.writeFileSync(path.join(__dirname, '..', '.env'), envContent);
  } else {
    console.error('Failed to setup virtual environment');
  }
} catch (error) {
  console.error('Installation error:', error.message);
  console.error('\nPlease install Python 3.7+ from:');
  console.error('- macOS: brew install python3');
  console.error('- Linux: sudo apt-get install python3');
  console.error('- Windows: https://www.python.org/downloads/');
} 