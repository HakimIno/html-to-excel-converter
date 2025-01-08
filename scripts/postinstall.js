const { execSync } = require('child_process');
const path = require('path');

function getPythonCommand() {
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

function installDependency(pythonCmd, dep) {
  try {
    // ลองติดตั้งแบบ global ก่อน
    execSync(`${pythonCmd} -m pip install ${dep}`, { stdio: 'inherit' });
    return true;
  } catch (err1) {
    try {
      // ถ้าติดตั้ง global ไม่ได้ ลองติดตั้งแบบ --user
      execSync(`${pythonCmd} -m pip install --user ${dep}`, { stdio: 'inherit' });
      return true;
    } catch (err2) {
      try {
        // ถ้ายังไม่ได้ ลองใช้ pip3
        execSync(`pip3 install ${dep}`, { stdio: 'inherit' });
        return true;
      } catch (err3) {
        console.warn(`Warning: Failed to install ${dep}`);
        return false;
      }
    }
  }
}

function checkDependency(pythonCmd, dep) {
  try {
    execSync(`${pythonCmd} -c "import ${dep.split('>')[0].split('=')[0]}"`, { stdio: 'ignore' });
    return true;
  } catch (err) {
    return false;
  }
}

try {
  // ตา Python command ที่ใช้ได้
  const pythonCmd = getPythonCommand();
  console.log(`Using Python command: ${pythonCmd}`);
  
  console.log('Checking Python dependencies...');
  
  const dependencies = {
    'beautifulsoup4': 'bs4',
    'pandas': 'pandas',
    'openpyxl': 'openpyxl',
    'cssutils': 'cssutils',
    'lxml': 'lxml'
  };

  // ตรวจสอบและติดตั้ง dependencies ที่ขาด
  for (const [dep, importName] of Object.entries(dependencies)) {
    if (!checkDependency(pythonCmd, importName)) {
      console.log(`Installing ${dep}...`);
      if (!installDependency(pythonCmd, dep)) {
        console.error(`Failed to install ${dep}. Please install it manually: ${pythonCmd} -m pip install ${dep}`);
      }
    } else {
      console.log(`${dep} is already installed`);
    }
  }
  
  console.log('Python dependencies setup completed');
} catch (error) {
  console.error('Error during installation:', error.message);
  console.error('Please make sure Python 3.7+ is installed and accessible from command line');
  console.error('You can install Python from: https://www.python.org/downloads/');
} 