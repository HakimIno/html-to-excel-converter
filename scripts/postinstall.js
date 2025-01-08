const { execSync } = require('child_process');
const path = require('path');

try {
  // ตรวจสอบว่ามี Python ติดตั้งหรือไม่
  execSync('python --version');
  
  console.log('Installing Python dependencies...');
  
  // ติดตั้ง beautifulsoup4 และ dependencies อื่นๆ โดยตรง
  const dependencies = [
    'beautifulsoup4>=4.9.3',
    'pandas>=1.3.0',
    'lxml>=4.9.0',
    'openpyxl>=3.0.7',
    'cssutils>=2.3.0'
  ];

  // ติดตั้งแต่ละ package
  dependencies.forEach(dep => {
    try {
      console.log(`Installing ${dep}...`);
      execSync(`python -m pip install ${dep}`, { stdio: 'inherit' });
    } catch (err) {
      console.warn(`Warning: Failed to install ${dep}`);
    }
  });
  
  console.log('Successfully installed Python dependencies');
} catch (error) {
  console.error('Error during installation:', error.message);
  console.error('Please make sure Python 3.7+ is installed and accessible from command line');
  // ไม่ exit process เพื่อให้ npm install ทำงานต่อได้
} 