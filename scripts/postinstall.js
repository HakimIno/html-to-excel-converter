const { execSync } = require('child_process');
const path = require('path');

function installDependency(dep) {
  try {
    // ลองติดตั้งแบบ global ก่อน
    execSync(`python -m pip install ${dep}`, { stdio: 'inherit' });
    return true;
  } catch (err1) {
    try {
      // ถ้าติดตั้ง global ไม่ได้ ลองติดตั้งแบบ --user
      execSync(`python -m pip install --user ${dep}`, { stdio: 'inherit' });
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

function checkDependency(dep) {
  try {
    execSync(`python -c "import ${dep.split('>')[0].split('=')[0]}"`, { stdio: 'ignore' });
    return true;
  } catch (err) {
    return false;
  }
}

try {
  // ตรวจสอบว่ามี Python ติดตั้งหรือไม่
  execSync('python --version');
  
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
    if (!checkDependency(importName)) {
      console.log(`Installing ${dep}...`);
      if (!installDependency(dep)) {
        console.error(`Failed to install ${dep}. Please install it manually: pip install ${dep}`);
      }
    } else {
      console.log(`${dep} is already installed`);
    }
  }
  
  console.log('Python dependencies setup completed');
} catch (error) {
  console.error('Error during installation:', error.message);
  console.error('Please make sure Python 3.7+ is installed and accessible from command line');
  // ไม่ exit process เพื่อให้ npm install ทำงานต่อได้
} 