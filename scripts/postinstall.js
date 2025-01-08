const { execSync } = require('child_process');
const path = require('path');

try {
  // ตรวจสอบว่ามี Python ติดตั้งหรือไม่
  execSync('python --version');
  
  // ติดตั้ง Python dependencies
  const requirementsPath = path.join(__dirname, '..', 'src', 'python', 'requirements.txt');
  execSync(`python -m pip install -r ${requirementsPath}`);
  
  console.log('Successfully installed Python dependencies');
} catch (error) {
  console.error('Error: Python is required but not found. Please install Python 3.7 or higher');
  process.exit(1);
} 