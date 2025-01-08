const { execSync } = require('child_process');
const path = require('path');
const fs = require('fs');

function installPythonDependencies() {
    try {
        console.log('Setting up Python virtual environment...');
        
        const rootDir = path.join(__dirname, '..');
        const venvPath = path.join(rootDir, 'venv');
        const requirementsPath = path.join(rootDir, 'src', 'python', 'requirements.txt');
        
        // Create .env file with venv path
        fs.writeFileSync(path.join(rootDir, '.env'), `VENV_PATH=${venvPath}`);
        
        // Check if python3 is available
        try {
            execSync('python3 --version');
        } catch (error) {
            console.error('Python 3 is not installed. Please install Python 3.7 or higher.');
            process.exit(1);
        }
        
        // Create virtual environment
        try {
            execSync(`python3 -m venv "${venvPath}"`);
        } catch (error) {
            console.error('Failed to create virtual environment:', error.message);
            process.exit(1);
        }
        
        // Install dependencies
        const pipCommand = process.platform === 'win32' ? 
            `"${path.join(venvPath, 'Scripts', 'pip.exe')}"` : 
            `"${path.join(venvPath, 'bin', 'pip')}"`;
            
        try {
            execSync(`${pipCommand} install -r "${requirementsPath}"`, { stdio: 'inherit' });
            console.log('Python dependencies installed successfully!');
        } catch (error) {
            console.error('Failed to install Python dependencies:', error.message);
            process.exit(1);
        }
        
    } catch (error) {
        console.error('Error during post-install setup:', error.message);
        process.exit(1);
    }
}

installPythonDependencies(); 