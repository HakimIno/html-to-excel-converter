const { execSync } = require('child_process');
const path = require('path');
const fs = require('fs');

function checkPythonDependency(pythonCmd, dependency) {
    try {
        execSync(`${pythonCmd} -c "import ${dependency}"`, { stdio: 'ignore' });
        return true;
    } catch (error) {
        return false;
    }
}

function installPythonDependencies() {
    try {
        console.log('Setting up Python virtual environment...');
        
        const rootDir = path.join(__dirname, '..');
        const venvPath = path.join(rootDir, 'venv');
        const requirementsPath = path.join(rootDir, 'src', 'python', 'requirements.txt');
        
        // Create .env file with venv path
        fs.writeFileSync(path.join(rootDir, '.env'), `VENV_PATH=${venvPath}`);
        
        // Check if python3 is available
        let pythonCmd = 'python3';
        try {
            execSync(`${pythonCmd} --version`);
        } catch (error) {
            // Try python if python3 is not available
            try {
                pythonCmd = 'python';
                execSync(`${pythonCmd} --version`);
            } catch (error) {
                console.error('Python is not installed. Please install Python 3.7 or higher.');
                process.exit(1);
            }
        }

        // Check if virtual environment exists
        const hasVenv = fs.existsSync(venvPath);
        
        // Check if dependencies are already installed
        if (hasVenv) {
            const pipCmd = process.platform === 'win32' ? 
                path.join(venvPath, 'Scripts', 'python.exe') : 
                path.join(venvPath, 'bin', 'python');
                
            const missingDeps = [];
            ['beautifulsoup4', 'xlsxwriter', 'pandas', 'cssutils'].forEach(dep => {
                const importName = dep === 'beautifulsoup4' ? 'bs4' : dep;
                if (!checkPythonDependency(pipCmd, importName)) {
                    missingDeps.push(dep);
                }
            });
            
            if (missingDeps.length === 0) {
                console.log('All Python dependencies are already installed.');
                return;
            }
        }
        
        // Create virtual environment if it doesn't exist
        if (!hasVenv) {
            try {
                execSync(`${pythonCmd} -m venv "${venvPath}"`);
            } catch (error) {
                console.error('Failed to create virtual environment:', error.message);
                process.exit(1);
            }
        }
        
        // Install dependencies
        const pipCommand = process.platform === 'win32' ? 
            `"${path.join(venvPath, 'Scripts', 'pip.exe')}"` : 
            `"${path.join(venvPath, 'bin', 'pip')}"`;
            
        try {
            // Upgrade pip first
            execSync(`${pipCommand} install --upgrade pip`, { stdio: 'inherit' });
            // Install requirements
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